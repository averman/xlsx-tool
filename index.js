const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');
const fetch = require('node-fetch');

let expired = 0;

function sum(...x) {
    let res = 0;
    for(let n of x) {
        let m = new Number(n);
        res+=Number.isNaN(m)?0:m;
    }
    return res;
}

function serialDate(serial) {
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
 
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
 
    var total_seconds = Math.floor(86400 * fractional_day);
 
    var seconds = total_seconds % 60;
 
    total_seconds -= seconds;
 
    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;
 
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
 }

yargs(hideBin(process.argv)).option('filter', {
    alias: 'f',
    type: 'regex',
    description: 'regex (greplike) string to filter the input filename'
}).option('datasetid', {
    type: 'number',
    description: 'the domo dataset id, empty to create one'
}).option('sheetname', {
    alias: 's',
    type: 'string',
    description: 'the name of the sheet to be processed'
}).option('pk', {
    description: 'primary key for merge',
    type: 'string',
}).option('csv', {
    description: 'set csv output',
    type: 'flag',
}).option('format', {
    description: 'comma delimited list of column',
    type: 'string',
}).command('extract [source] [target] [mappingfile]', 
'extract column(s) from source folder to target folder', (yargs)=>{
    yargs.positional('source', {
        describe: 'source folder',
        default: '.'
    });
    yargs.positional('target', {
        describe: 'target folder',
        default: '..'
    });
    yargs.positional('mappingfile', {
        describe: 'json rule to extract the information(s)',
        default: 'mapping.json'
    });
}, (argv)=>{
    let rule = JSON.parse(fs.readFileSync(argv.mappingfile).toString());
    extract(argv.source, argv.target, rule, argv.filter, argv.sheetname);
}).command('merge [source] [target]', 
'merge xlsx files from source folder to target file', (yargs)=>{
    yargs.positional('source', {
        describe: 'source folder',
        default: '.'
    });
    yargs.positional('target', {
        describe: 'target file',
        default: 'target.xlsx'
    });
}, (argv)=>{
    merge(argv.source, argv.target, argv.filter, argv.sheetname, argv.pk, argv.format, argv.csv);
}).command('upload [source] [name] [schema]', 
'upload xlsx files from source folder to domo', (yargs)=>{
    yargs.positional('source', {
        describe: 'source directory'
    });
    yargs.positional('name', {
        describe: 'dataset name'
    });
    yargs.positional('schema', {
        describe: 'json rule table schema',
        default: 'mapping.json'
    });
}, (argv)=>{
    test(argv.source, argv.name, argv.schema, argv.datasetid, argv.filter, argv.sheetname);
}).argv


async function checkToken(token) {
    if(expired-Date.now()<0){
        console.log("getting new token for domo!",expired,Date.now(),expired-Date.now());
        token = await getDomoToken();
        expired = Date.now()- -(45*60*1000);
        console.log("token will be refreshed at",new Date(expired));
    }
    return token;
}

async function test(source, name, rulefile, datasetid, filterstring, sheetname){
    let rule = JSON.parse(fs.readFileSync(rulefile).toString());
    let filters = rule.map((x,i)=>[x,i]).filter(x=>x[0].filter);
    let doFilter = (r) => filters.map(x=>RegExp(x[0].filter).test(r[x[1]]))
    let token = await checkToken();
    let stats = [['filename', 'row count', ...filters.map(x=>x.colname), 'final count']]
    console.log('Got DOMO token for 1 hour');
    if(!datasetid) {
        datasetid = await createDataset(name, rule, token);
        console.log("your datasetid is",datasetid);
    }
    let url = 'https://api.domo.com/v1/streams/'+datasetid+'/executions';
    let execId = await fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'Authorization': 'Bearer '+token
        }
    }).then(response => response.json())
    .then(json => json.id);
    console.log("uploading from",source,"to",name);
    let files = fs.readdirSync(source);
    let filterRegex = filterstring? new RegExp(filterstring): undefined;
    let filter = filterRegex? x=>filterRegex.test(x) : x=>true;
    let globalPart = 0;
    for(let filename of files){
        let stat = stats[0].map((x,i)=>i==0?filename:0);
        if(fs.lstatSync(path.join(source,filename)).isDirectory() ) continue;
        if(!filename.endsWith(".xlsx")) continue;
        if(!filter(filename)) continue;
        console.log("Reading from", filename);
        let workbook = XLSX.readFile(path.join(source,filename));
        if(sheetname && workbook.SheetNames.indexOf(sheetname)==-1){
            console.error("sheet",sheetname,"does not found in file",filename);
            continue;
        }
        let sheets = sheetname? [sheetname] : workbook.SheetNames;
        console.log("sheets to be uploaded",sheets); 
        for(let activeSheet of sheets) {
            let worksheet = workbook.Sheets[activeSheet];
            let data = XLSX.utils.sheet_to_json(worksheet);
            let bulk = [];
            let count = 1;
            for(let i=0; i<data.length; i++){
                stat[1]++;
                let row = data[i];
                let d = rule.map(x=>{
                    let res;
                    if(x.source) {
                        for(let colMap of x.source) {
                            if(colMap.startsWith("!")) {
                                let scr = colMap.substr(1);
                                res = eval(scr);
                                break;
                            }
                            else if(typeof row[colMap] != "undefined") {
                                res = row[colMap];
                                break;
                            }
                        }
                    } else {
                        res = row[x.colname]
                    }
                    if(x.serialdate) return serialDate(res).toLocaleDateString()
                    return res;
                });
                let flags = doFilter(d);
                stat = stat.map((x,i)=>i>2&&flags[i-2]?x[i]+1:x)
                if(!flags.reduce((p,c)=>p&&c, true)) continue;
                stat[stat.length-1]++;
                bulk.push(d);
                if(bulk.length>=10000 || i == data.length-1){   
                    let part = globalPart+count;
                    console.log("uploading part "+count+"/"+Math.ceil(data.length/10000)+" -- total-part "+part);
                    token = await checkToken(token);
                    await putDataPart(url, execId, part, token, bulk);
                    bulk = [];
                    if(i<data.length-1)
                        count++;
                }
            }
            globalPart = globalPart+count;
        }
        stats.push(stat);
    }
    console.log("committing changes");
    await fetch(url+'/'+execId+"/commit", {
        method: 'PUT',
        headers: {
                'Accept': 'application/json',
                'Content-Type': 'text/csv',
                'Authorization': 'Bearer '+token
            }
        }).then(response => response.json())
        .then(console.log);
    console.log("finished upload to domo dataset with id ",datasetid);
    fs.writeFileSync('stats.csv', stats.map(x=>x.join(",")).join('\n'));
}

async function putDataPart(url,execId, part,token,bulk) {
    await fetch(url+'/'+execId+"/part/"+part, {
        method: 'PUT',
        headers: {
                'Accept': 'application/json',
                'Content-Type': 'text/csv',
                'Authorization': 'Bearer '+token
            },
        body: bulk.map(x=>x.join(',')).join('\n')
        }).catch(async err=>{
            console.log(err);
            await putDataPart(url,execId, part,token,bulk);
        }).then(response => response.json()).catch(async err=>{
            console.log(err);
            await putDataPart(url,execId, part,token,bulk);
        })
        .then(console.log)
}

async function createDataset(name, rule, token){
    let url = 'https://api.domo.com/v1/streams';
    console.log('url is',url)
    return await fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'Authorization': 'Bearer '+token
        },
        body: JSON.stringify({
            "dataSet" : {
              "name" : name,
              "description" : "",
              "schema" : {
                "columns" : rule.map(x=>{return {type: x.type?x.type:"STRING", name: x.colname}})
              }
            },
            "updateMethod" : "APPEND"
          })
    }).then(response => response.json())
    .then(json => json.id);
}

async function getDomoToken(){
    let url = 'https://api.domo.com/oauth/token?grant_type=client_credentials&scope=data';
    let headers = new fetch.Headers();
    let {username, password} = JSON.parse(fs.readFileSync("domo.key").toString());
    headers.set('Authorization', 'Basic ' + Buffer.from((username + ":" + password)).toString('base64'));
    return await fetch(url, {method:'GET',
        headers: headers,
       })
    .then(response => response.json())
    .then(json => json.access_token);
}

function merge(source, target, filterstring, sheetname, pk, cols, csv) {
    console.log("merging from",source,"to",target,"with pk",pk,csv?"to csv format":"");
    let files = fs.readdirSync(source);
    let filterRegex = filterstring? new RegExp(filterstring): undefined;
    let filter = filterRegex? x=>filterRegex.test(x) : x=>true;
    let resBook = XLSX.utils.book_new();
    let resSheets = {};
    for(let filename of files){
        if(fs.lstatSync(path.join(source,filename)).isDirectory() ) continue;
        if(!filename.endsWith(".xlsx")) continue;
        if(!filter(filename)) continue;
        console.log("Reading from", filename);
        let workbook = XLSX.readFile(path.join(source,filename));
        if(sheetname && workbook.SheetNames.indexOf(sheetname)==-1){
            console.error("sheet",sheetname,"does not found in file",filename);
            continue;
        }
        let sheets = sheetname? [sheetname] : workbook.SheetNames;
        console.log("sheets to be merged",sheets); 
        for(let activeSheet of sheets) {
            let worksheet = workbook.Sheets[activeSheet];
            let data = XLSX.utils.sheet_to_json(worksheet);
            if(resSheets[activeSheet]){
                if(pk){
                    for(let d of data){
                        let pkval = d[pk];
                        resSheets[activeSheet][pkval] = resSheets[activeSheet][pkval]?
                        Object.assign(resSheets[activeSheet][pkval],d):d;
                    }
                }else
                    resSheets[activeSheet] = resSheets[activeSheet].concat(data);
            }else{
                if(pk){
                    resSheets[activeSheet] = {};
                    for(let d of data){
                        let pkval = d[pk];
                        resSheets[activeSheet][pkval] = d;
                    }
                }else
                    resSheets[activeSheet] = data;
            }
        }
    }
    for(let sn of Object.keys(resSheets)) {
        let tempData = resSheets[sn]
        if(pk) {
            tempData = [];
            let mark = 0.0;
            let keys = Object.keys(resSheets[sn]);
            for(let i=0; i<keys.length; i++){
                if(resSheets[sn][i]) continue;
                tempData.push(resSheets[sn][keys[i]]);
                let imark = (i*100)/keys.length;
                if(imark - mark >= 5){
                    mark = imark;
                    console.log("merging progress",imark,'%');
                }
            }
        }
        let columns;
        if(!cols){
            console.log("consolidating columns");
            columns = {};
            for(let r of tempData){
                columns = Object.assign(columns,r);
            }
            columns = Object.keys(columns);
        }else {
            columns = cols.split(",");
        }
        let aoa = [columns];
        let mark = 0.0;
        console.log("converting merged data into a sheet");
        for(let i=0; i<tempData.length; i++){
            aoa.push(columns.map(x=>tempData[i][x]));
            let imark = (i*100)/tempData.length;
            if(imark - mark >= 10){
                mark = imark;
                console.log("merging progress",imark,'%');
            }
        }
        if(csv){
            console.log("writing to csv");
            let stringdata = aoa.map(x=>x.join(",")).join("\n");
            fs.writeFileSync(target,stringdata);
        }else{
            let tempSheet = XLSX.utils.aoa_to_sheet(aoa);
            console.log("creating an xlsx based on the sheet");
            XLSX.utils.book_append_sheet(resBook, tempSheet, sn);
        }
    }
    if(!csv){
        console.log("writing to file");
        XLSX.writeFile(resBook, target);
    }

}

function extract(source, target, rule, filterstring, sheetname) {
    console.log("extracting from",source,"to",target);
    let files = fs.readdirSync(source);
    let filterRegex = filterstring? new RegExp(filterstring): undefined;
    let filter = filterRegex? x=>filterRegex.test(x) : x=>true;
    for(let filename of files){
        if(fs.lstatSync(path.join(source,filename)).isDirectory() ) continue;
        if(!filename.endsWith(".xlsx")) continue;
        if(!filter(filename)) continue;
        console.log("Extracting from", filename);
        let workbook = XLSX.readFile(path.join(source,filename));
        if(sheetname && workbook.SheetNames.indexOf(sheetname)==-1){
            console.error("sheet",sheetname,"does not found in file",filename);
            continue;
        }
        let sheets = sheetname? [sheetname] : workbook.SheetNames;
        console.log("sheets to be processed",sheets); 
        for(let activeSheet of sheets) {
            let worksheet = workbook.Sheets[activeSheet];
            let data = XLSX.utils.sheet_to_json(worksheet)
            let result = [];
            for(let row of data){
                let rowres = {};
                for(let colRule of rule){
                    for(let colMap of colRule.source) {
                        if(colMap.startsWith("!")) {
                            let scr = colMap.substr(1);
                            rowres[colRule.colname] = eval(scr);
                        }
                        else if(typeof row[colMap] != "undefined") {
                            rowres[colRule.colname] = row[colMap];
                            break;
                        }
                    }
                }
                result.push(rowres);
            }
            let resBook = XLSX.utils.book_new();
            let resSheet = XLSX.utils.json_to_sheet(result);
            XLSX.utils.book_append_sheet(resBook, resSheet, 'result');
            let sheetExtFn = sheets.length>1?activeSheet:'';
            XLSX.writeFile(resBook, path.join(target,filename.substr(0,filename.length-5)+sheetExtFn+".xlsx"));
        }
    }
}