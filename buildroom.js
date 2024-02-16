const Excel = require('exceljs')

main()

async function main() {
  try {
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile('C:\\Users\\kabom\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx')

    const ws = workbook.getWorksheet('Documents')
    const table = ws.getTable('Documents')

    //console.log(table.table.columns)

    const columnDict = Object.fromEntries(table.table.columns.map((column, index) => [column.name, index+1]))
    console.log(columnDict)

    console.log(columnDict['Short Title'])

    ws.eachRow((row, rowNumber) => {
      console.log('Row ' + rowNumber + ' = ' + row.values[columnDict['Short Title']])
    })

  } catch (e) {
    console.log(e)
  }
}
