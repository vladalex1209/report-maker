const fs = require('fs')
const xlsx = require('write-excel-file/node')

module.exports = {
  toXlsxFile: (reportObj, currentDate) => {
    console.log(reportObj)
    let pathToFolder = './reports'
    if (!fs.existsSync(pathToFolder)) {
      fs.mkdirSync(pathToFolder);
    }

    let filePath = './reports/report - ' + currentDate + '.xlsx'

    function tasks(sumOfTasks) {

      let toReturn = ''

      for (let i = 2; i < Object.keys(sumOfTasks).length; i++) {
        toReturn += Object.keys(sumOfTasks)[i] + ': ' + Object.values(sumOfTasks)[i].time.slice(0, -3)
        toReturn += '\n'
        if (typeof sumOfTasks[Object.keys(sumOfTasks)[i]].notes !== 'undefined') {
          sumOfTasks[Object.keys(sumOfTasks)[i]].notes.forEach((note) => { toReturn += '  ' + note.replace(/\s/g, '  ').replace(/,/, "\n") + '\n' })
        }
      }
      return toReturn;
    }

    let data = [];

    for (let i = 0; i < reportObj.length - 1; i++) {

      let HEADER_ROW = []
      let DATA_ROW_1 = []
      let DATA_ROW_2 = []
      let DATA_ROW_3 = []
      let EMPTY_ROW_4 = []
      let EMPTY_ROW_5 = []


      HEADER_ROW.push(
        {
          value: 'Project',
          topBorderStyle: 'thin',
          leftBorderStyle: 'thin',
          align: 'center',
          color: "#ffffff",
          backgroundColor: '#8e86ae',
          span: 3
        }, null, null,
        {
          value: 'Hours Worked',
          topBorderStyle: 'thin',
          rightBorderStyle: 'thin',
          color: "#ffffff",
          backgroundColor: '#8e86ae'
        })

      DATA_ROW_1.push(
        {
          value: reportObj[i].project,
          backgroundColor: '#ffffff',
          color: "#000000",
          leftBorderStyle: 'thin',
          bottomBorderStyle: 'thin',
          alignVertical: 'center',
          align: 'center',
          rowSpan: 3,
          span: 3
        }, null, null,
        {
          value: reportObj[i].total.slice(0, -3),
          backgroundColor: '#ffffff',
          color: '#000000',
          rightBorderStyle: 'thin',
        })

      DATA_ROW_2.push(null, null,
        null,
        {
          value: '',
          backgroundColor: '#8e86ae',
          color: "#ffffff",
          rightBorderStyle: 'thin',
        })

      DATA_ROW_3.push(null, null,
        null,
        {
          value: tasks(reportObj[i]),
          backgroundColor: '#ffffff',
          color: "#000000",
          rightBorderStyle: 'thin',
          bottomBorderStyle: 'thin',
        })

      EMPTY_ROW_4 = [null, null]

      EMPTY_ROW_5 = [null, null]

      data = [...data, HEADER_ROW, DATA_ROW_1, DATA_ROW_2, DATA_ROW_3, EMPTY_ROW_4, EMPTY_ROW_5]

    }

    let HEADER_ROW = [{ value: "In", backgroundColor: '#8e86ae', topBorderStyle: 'thin', leftBorderStyle: 'thin', color: "#ffffff" }, { value: "Out", color: "#ffffff", backgroundColor: '#8e86ae', topBorderStyle: 'thin' }, { value: "Total", color: "#ffffff", backgroundColor: '#8e86ae', topBorderStyle: 'thin', rightBorderStyle: 'thin' }]
    let DATA_ROW_1 = [{ value: reportObj[reportObj.length - 1].Start.slice(0, -3), color: "#000000", leftBorderStyle: 'thin', backgroundColor: '#000000', }, { value: reportObj[reportObj.length - 1].End.slice(0, -3), color: "#000000", bottomBorderStyle: 'thin', backgroundColor: '#000000', }, { value: reportObj[reportObj.length - 1].Total.slice(0, -3), color: "#000000", rightBorderStyle: 'thin', bottomBorderStyle: 'thin', backgroundColor: '#000000', }]
    let DATA_ROW_2 = [{ value: "Miss", backgroundColor: '#8e86ae', color: "#ffffff", leftBorderStyle: 'thin', rightBorderStyle: 'thin' }]
    let DATA_ROW_3 = [{ value: reportObj[reportObj.length - 1].Miss, color: "#000000", leftBorderStyle: 'thin', rightBorderStyle: 'thin', bottomBorderStyle: 'thin', backgroundColor: '#ffffff', }]

    data = [...data, HEADER_ROW, DATA_ROW_1, DATA_ROW_2, DATA_ROW_3]

    xlsx(data, {
      filePath: filePath,
    })
  }
}
