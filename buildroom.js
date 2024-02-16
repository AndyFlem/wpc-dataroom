const Excel = require('exceljs')
const fs = require('fs')
const { DateTime } = require("luxon")

const configs = {
  sable: {
    catalogFile: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx',
    outFolder: 'D:\\temp\\',
    sps: {
      WPCWorking: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents',
      WPCHSESManagementSystem: 'D:\\Western Power Company\\Western Power Company\\WPC HSES Management System - General'
    }
  },
  cloud: {
    catalogFile: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx',
    outFolder: 'D:\\OneDrive\\Western Power Company\\WPC Dataroom - Dataroom\\',
    sps: {
      WPCWorking: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents',
      WPCHSESManagementSystem: 'D:\\OneDrive\\Western Power Company\\WPC HSES Management System - General'
    }
  }  
}

const config = configs.cloud
const catalogFile = config.catalogFile
const filterColumn = ''
const roomName = 'Full'//'CapLink'
const outFolder = config.outFolder + roomName + '\\'

if (!fs.existsSync(outFolder)) {
  fs.mkdirSync(outFolder)
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
      if (row.values[columnDict['Name']] && (filterColumn =='' || row.values[columnDict[filterColumn]])) { 
        console.log('Row ' + rowNumber + ' = ' + row.values[columnDict['Short Title']])
        
        const fileDate = DateTime.fromISO(row.values[columnDict['DateString']])
        let outPath = outFolder + row.values[columnDict['Area']]
        const title = row.values[columnDict['Short Title']]
        const extension = row.values[columnDict['Name']].match(patt)[0]

        const inLocation = config.sps[row.values[columnDict['SPSite']].result] + '\\' + decodeURI(row.values[columnDict['SPLocation']].result).replace('\General/','/').replace('%26','&').replace('/Shared Documents/', '')

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
          const outLocation = outPath + '\\' + fileDate.toFormat('yyyyLLdd') + ' ' + title.trim() + extension.trim()
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
