const ExcelGenerator = () => {

    const readXlsxFile = require('read-excel-file/node');
    const wrightXlsxFile = require('excel4node');
    const Readjson = require('fs');
    const constants = require('./constant')

    let DataMap = Readjson.readFileSync(constants.constants.PATHS.MAPPING_FILE);
    DataMap = JSON.parse(DataMap);

    const letters = 'abcdefghijklmnopqrstuvwxyz';
    let MapXYvalues = new Map()

    readXlsxFile(constants.constants.PATHS.TEMPLATE_FILE).then((rows) => {
        for(let i = 0; i < rows.length; i++) {
            for(let j = 0; j < rows[0].length; j++) {
                MapXYvalues.set(rows[i][j], [i+1, j+1])
            }
        }
        
        readXlsxFile(constants.constants.PATHS.DATA_FILE).then((rows) => {

            for(let i = 1; i < rows.length; i++) {
        
                let datafile = new wrightXlsxFile.Workbook()
                let datasheet = datafile.addWorksheet('studentData')
        
                for(let columnNames of MapXYvalues.keys()) {
                    datasheet.cell(MapXYvalues.get(columnNames)[0], MapXYvalues.get(columnNames)[1]).string(columnNames)                
                }
        
                for(let j = 0; j < rows[i].length; j++) {
        
                    var cellNumber = DataMap[rows[0][j].toLowerCase()].replace(/\'/g, '').split(/(\d+)/).filter(Boolean)
                    var columnNumber = cellNumber[0]
                    var rowNumber = cellNumber[1]
        
                    var columnIndex = letters.search(columnNumber);
                    datasheet.cell(rowNumber, columnIndex + 1).string(rows[i][j].toString())
                }
                
                datafile.write(`output data/Datasheet${i}.xlsx`);
            }
        })
    })
}

ExcelGenerator();