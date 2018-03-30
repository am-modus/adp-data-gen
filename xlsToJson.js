var Excel = require("exceljs");
var fs = require("fs");

function churnString(valString){
    let valNum = parseFloat(valString);
    
    if(isNaN(valNum)) {
        return valString;
    }
    else{
        valNum = (valNum * 100) / 100;
        let result = (valNum >= 0 && valNum <= 1) ? (valNum * 100).toFixed(2) + "%" : "$" + valNum;
        return result;
    } 
}

function parseFile(infile, outfile){

    let workbook = new Excel.Workbook();
    let result = [];
    
    console.log("\n**** JSONifying XLXS file ****\n");
    console.log("--> filename:", infile);

    workbook.xlsx.readFile(infile).then(function () {   
        
        let sheets = workbook._worksheets;

        for(var i = 1; i <= sheets.length; i++){
            
            try{
                let sheet = sheets[i];
                if(!sheet) break;

                let worksheet = workbook.getWorksheet(sheet.id);
                let sheetName = ((worksheet.name.split('-'))[1]).toLowerCase().trim();
                let periods = [];
                let columns = [];
                let data = {};

                console.log("----> processing ["+i+"]:", sheetName);

                worksheet.eachRow({includeEmpty: false}, (_row, _rowNumber) => {
                    let row = _row.values;
                    let rowNum = _rowNumber - 1;

                    //let's get rid of the first two columns
                    row = row.slice(3);
                    let newRow = [];

                    //now get rid of all the empty values in the row
                    for (let k = 0; k < row.length; k++){
                        let col = (row[k]) ? ''+row[k] : '';
                        col = col.trim();
                        if(col.length > 0) newRow.push(col);
                    }

                    
                    //now let's populate the fields
                    // 0th row > periods
                    // 1st row will contain duplicates. To get unique elements, get slice of row from 0 to (length of Jth row / number of periods)
                    // all other rows, the first col is the key, other items are the value

                    if (rowNum == 0){
                        periods = newRow.filter(function(item, pos) {
                            return newRow.indexOf(item) == pos;
                        });

                        //console.log("--> periods", '['+periods+']');
                    }
                    else if(rowNum == 1){
                        columns = newRow.slice(0, newRow.length / periods.length)
                        //console.log("---> cols", '['+columns+']');
                    }
                    else if(newRow.length > 0){
                        let key = newRow[0];
                        let val = newRow.slice(1);
                        val = val.map( i => churnString(i));
                        data[key] = val;
                    }
                });

                //now that we have periods, columns and data, create the object
                //for this sheet
                let sheetObject = {
                    type: sheetName,
                    periods: periods,
                    columns: columns,
                    data: data
                };

                result.push(sheetObject);
            }
            catch(e){
                break;
            }
        }
        

        let jsonString = JSON.stringify(result, null, 4);
        fs.writeFileSync(outfile, jsonString);

        console.log("\n*** mock data generated ***\n")
    });

}

module.exports = parseFile;


/*************************************************************
 * main: if this script is run this directly on its own
 * node xlsToJson.js -i "input_path.xlsx" -o "output_path.json"
**************************************************************/

if(!module.parent){

    //process command line args to get in and out paths
    let inputIndex = process.argv.indexOf("-i") + 1;
    let outputIndex = process.argv.indexOf("-o") + 1;
    let inputPath = process.argv[inputIndex];
    let outputPath = process.argv[outputIndex];
    
    if(inputPath && outputPath){
        let isInpValid = fs.existsSync(inputPath) && fs.lstatSync(inputPath).isFile();
        let isOutValid = (outputPath.trim().length > 0) ? true : false;
         
        if(isInpValid && isOutValid){
            parseFile(inputPath, outputPath);
        }
        else console.error("Provide valid paths for input and output files");
    } 
    else console.error("Check command line args. Help: node xlsToJson.js -i 'input_path.xlsx' -o 'output_path.json'");
}