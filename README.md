# xlsx-tool

## requirement: 

- nodejs > v.10.x

## installation:

`git clone`
`npm install`

### usage:

#### help

`node index.js --help`

#### extract

`node index.js extract <sourcefolder> <resultfolder> <mappingrule> `

mapping rule are json file (example are in mapping.json file) that specifies the output format. It takes a json of array of column information objects. Where the column information objects is the column definition that needs properties of colname (the column name) and source: an array of source column. If you specify the colname as newcol and source as `["col1", "col2"]` this means that the mapping are defined `col1 -> newcol`, and if there is no col1 - `col2 -> newcol`

#### merge

`node index.js merge <sourcefolder> <resultfile>`

merge all xlsx file in source folder to result file

### options:

#### filter

-f or --filter

a regex string to filter the file to be processed
