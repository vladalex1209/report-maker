const fs = require('fs')
const xlsx = require('write-excel-file/node')
require('dotenv').config()

module.exports = {
  toXlsxFile: (reportObj, currentDate) => {

    let filePath = `C://Users//${process.env.USER}//Desktop//`;

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
          topBorderStyle: 'thick',
          leftBorderStyle: 'thick',
          align: 'center',
          color: "#ffffff",
          backgroundColor: '#8e86ae',
          span: 3
        }, null, null,
        {
          value: 'Hours Worked',
          topBorderStyle: 'thick',
          rightBorderStyle: 'thick',
          color: "#ffffff",
          backgroundColor: '#8e86ae'
        })

      DATA_ROW_1.push(
        {
          value: reportObj[i].project,
          backgroundColor: '#ffffff',
          color: "#000000",
          leftBorderStyle: 'thick',
          bottomBorderStyle: 'thick',
          alignVertical: 'center',
          align: 'center',
          rowSpan: 3,
          span: 3
        }, null, null,
        {
          value: reportObj[i].total.slice(0, -3),
          backgroundColor: '#ffffff',
          color: '#000000',
          rightBorderStyle: 'thick',
        })

      DATA_ROW_2.push(null, null,
        null,
        {
          value: '',
          backgroundColor: '#8e86ae',
          color: "#ffffff",
          rightBorderStyle: 'thick',
        })

      DATA_ROW_3.push(null, null,
        null,
        {
          value: tasks(reportObj[i]),
          backgroundColor: '#ffffff',
          color: "#000000",
          rightBorderStyle: 'thick',
          bottomBorderStyle: 'thick',
        })

      EMPTY_ROW_4 = [null, null]

      EMPTY_ROW_5 = [null, null]

      data = [...data, HEADER_ROW, DATA_ROW_1, DATA_ROW_2, DATA_ROW_3, EMPTY_ROW_4, EMPTY_ROW_5]

    }

    let HEADER_ROW = [{ value: "In", backgroundColor: '#8e86ae', topBorderStyle: 'thick', leftBorderStyle: 'thick', color: "#ffffff" }, { value: "Out", color: "#ffffff", backgroundColor: '#8e86ae', topBorderStyle: 'thick' }, { value: "Total", color: "#ffffff", backgroundColor: '#8e86ae', topBorderStyle: 'thick', rightBorderStyle: 'thick' }]
    let DATA_ROW_1 = [{ value: reportObj[reportObj.length - 1].Start.slice(0, -3), color: "#000000", leftBorderStyle: 'thick', backgroundColor: '#ffffff', }, { value: reportObj[reportObj.length - 1].End.slice(0, -3), color: "#000000", bottomBorderStyle: 'thick', backgroundColor: '#ffffff', }, { value: reportObj[reportObj.length - 1].Total.slice(0, -3), color: "#ffffff", rightBorderStyle: 'thick', bottomBorderStyle: 'thick', backgroundColor: '#ffffff', }]
    let DATA_ROW_2 = [{ value: "Miss", backgroundColor: '#8e86ae', color: "#ffffff", leftBorderStyle: 'thick', rightBorderStyle: 'thick' }]
    let DATA_ROW_3 = [{ value: reportObj[reportObj.length - 1].Miss, color: "#000000", leftBorderStyle: 'thick', rightBorderStyle: 'thick', bottomBorderStyle: 'thick', backgroundColor: '#ffffff', }]

    data = [...data, HEADER_ROW, DATA_ROW_1, DATA_ROW_2, DATA_ROW_3]

    try {
      xlsx(data, {
        filePath: filePath,
      })
    } catch (err) { console.log(err) }


  }
}
