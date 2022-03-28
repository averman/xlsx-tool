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

You can also use js to populate the value of the column. Any source with ! at the start will be treated as script, and columns are stored at rowres variable, so a source with value `!rowres['amount']>0?'positive':'negative'` will evaluate the js after !. This also can be used to concat few rows into one new column (for pk in merging). Please note that order matter, a column can only access values the column before it. You can use rowres variable for accessing the columns at the target, or row variable for accessing the columns at source

you can take advantage of helper function sum to sum any number of variables and treat any undefined one as 0

#### merge

`node index.js merge <sourcefolder> <resultfile>`

merge all xlsx file in source folder to result file

use --pk to define the primary key (unique row identifier) so that row with same value in that column will be merged as one, not defining the pk will make merge command only append rows

#### upload

`node index.js upload <xlsxfolderpath> <datasetname> <jsonschemafile>`

please make sure you have domo.key file in the same path as index.js with value
```
{
  "username": "",
  "password": ""
}
```

### options:

#### filter

-f or --filter

a regex string to filter the file to be processed

