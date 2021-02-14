const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');

function sum(...x) {
    let res = 0;
    for(let n of x) {
        let m = new Number(n);
        res+=Number.isNaN(m)?0:m;
    }
    return res;
}

yargs(hideBin(process.argv)).option('filter', {
    alias: 'f',
    type: 'regex',
    description: 'regex (greplike) string to filter the input filename'
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
}).argv

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
                        else if(row[colMap]) {
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