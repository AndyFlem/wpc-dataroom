const Excel = require('exceljs')
const fs = require('fs')


//const catalogFile = 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx'
//const outFolder = 'D:\\temp\\CapLink\\'

const catalogFile = 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx'
const outFolder = 'D:\\OneDrive\\Western Power Company\\WPC Dataroom - Dataroom\\CapLink\\'


const filterColumn = 'CAPLINK Dataroom'

const sps = {
  //WPCWorking: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents',
  //WPCHSESManagementSystem: 'D:\\Western Power Company\\Western Power Company\\WPC HSES Management System - General'
  WPCWorking: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents',
  WPCHSESManagementSystem: 'D:\\OneDrive\\Western Power Company\\WPC HSES Management System - General' 
}

const patt = /\.[^.\\/:*?"<>|\r\n]+$/i

main()
async function main() {
  try {
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(catalogFile)

    const ws = workbook.getWorksheet('Documents')
    const columns = ws.getRow(1).values.map((column, index) => [column, index])
    columns.shift()
    columnDict=Object.fromEntries(columns)

    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return
      if (row.values[columnDict[filterColumn]]) { 
        console.log('Row ' + rowNumber + ' = ' + row.values[columnDict['Short Title']])
        
        let outPath = outFolder + row.values[columnDict['Area']]
        const title = row.values[columnDict['Short Title']]
        const extension = row.values[columnDict['Name']].match(patt)[0]

        const inLocation = sps[row.values[columnDict['SPSite']].result] + '\\' + decodeURI(row.values[columnDict['SPLocation']].result).replace('\General/','/').replace('%26','&').replace('/Shared Documents/', '')
       // console.log(inLocation)
        try {
          if (!fs.existsSync(outPath)) {
            fs.mkdirSync(outPath)
          }
          if (row.values[columnDict['Area2']]) {
            outPath += '\\' + row.values[columnDict['Area2']]
            if (!fs.existsSync(outPath)) {
              fs.mkdirSync(outPath)
            }
          }
          if (row.values[columnDict['Area3']]) {
            outPath += '\\' + row.values[columnDict['Area3']]
            if (!fs.existsSync(outPath)) {
              fs.mkdirSync(outPath)
            }
          }
          const outLocation = outPath + '\\' + title + extension 
          if (!fs.existsSync(outLocation)) {
            if (!fs.existsSync(inLocation)) {
              console.log('Source folder does not exist:', inLocation)
            } else {
              fs.copyFileSync(inLocation, outLocation)
            }
          }

          
        } catch (err) {
          console.error(err)
        }
      } else {
        console.log('Row ' + rowNumber + ' = NOT IN EXPORT [' + row.values[columnDict['Short Title']] + ']')
      }
    })

  } catch (e) {
    console.log(e)
  }
}
