const Excel = require('exceljs')
const fs = require('fs')
const { DateTime } = require("luxon")

const patt = /\.[^.\\/:*?"<>|\r\n]+$/i

const configs = {
  sable: {
    catalogFile: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\202402 WPC Documents Catalogue.xlsx',
    outFolder: 'D:\\temp\\Datarooms\\',
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

const rooms = {
  HSS: {
    filterColumn: 'HSS Assess Doc',
    roomName:'HSS'
  },
  Caplink: {
    filterColumn: 'CAPLINK Dataroom2',
    roomName: 'CapLink'
  },
  All: {
    filterColumn: false,
    roomName:'All'
  }    

}

const HSSColumns = [
  '0 General',
  '1 Environmental and Social',
  '2 Labour and Working Conditions',
  '3 Water Quality and Sediments',
  '4 Community Impacts and Safety',
  '5 Resettlement',
  '6 Biodiversity and Invasive Species',
  '7 Indigenous Peoples',
  '8 Cultural Heritage',
  '9 Governance and Procurement',
  '10 Communications and Consultation',
  '11 Hydrological Resource',
  '12 Climate Change'
]


// **************************
const config = configs.sable
const room = rooms.All
const folderOnly = true
// **************************

const catalogFile = config.catalogFile

const filterColumn = room.filterColumn
const roomName = room.roomName

const outFolder = config.outFolder + roomName 

if (!fs.existsSync(outFolder)) {
  fs.mkdirSync(outFolder)
}

main()

function makeFolder(folder) {
  if (!fs.existsSync(folder)) {
    console.log('Creating folder:', folder)
    fs.mkdirSync(folder)
  }
}

function makePath(path) {
  let workingPath = ''
  path.forEach((folder) => {
    workingPath += folder + '\\'
    makeFolder(workingPath)
  })
}

function doCopy(inLocation, outPath, fileName, log) {
  makePath(outPath)
  const outLocation = outPath.join('\\') + '\\' + fileName  

  if (!fs.existsSync(outLocation)) {
    if (!fs.existsSync(inLocation)) {
      console.log('Source does not exist:', inLocation)
    } else {
      if (!folderOnly) {
        fs.copyFileSync(inLocation, outLocation)
      }
      console.log(log)      
    }
  }
}

async function main() {

  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(catalogFile)

  const ws = workbook.getWorksheet('Documents')
  const columns = ws.getRow(1).values.map((column, index) => [column, index])
  columns.shift()
  columnDict=Object.fromEntries(columns)

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return

    let filter = filterColumn ? row.values[columnDict[filterColumn]] : true
    if ((typeof filter==='object ') && filter.result) {filter=filter.result}
    
    if (filter) { 
      if (row.values[columnDict['Name']]) {
        let fileDate = DateTime.fromISO(row.values[columnDict['DateString']])
        
        if (!fileDate.isValid) { fileDate = DateTime.fromISO('2024-01-01') }

        const inLocation = config.sps[row.values[columnDict['SPSite']].result] + '\\' + decodeURI(row.values[columnDict['SPLocation']].result).replace('\General/','/').replace('%26','&').replace('/Shared Documents/', '')

        const title = row.values[columnDict['Short Title']]
        const extension = row.values[columnDict['Name']].match(patt)[0]
        const fileName = fileDate.toFormat('yyyyLLdd') + ' ' + title.trim() + extension.trim()

        let outPath = [outFolder, row.values[columnDict['Area']]]
        if (row.values[columnDict['Area2']]) {
          outPath.push(row.values[columnDict['Area2']])
        }
        if (row.values[columnDict['Area3']]) {
          outPath.push(row.values[columnDict['Area3']])
        }
 
        if (roomName === 'HSS') {
          outPath=[outPath[0],'',...outPath.slice(1)]
          HSSColumns.forEach((column) => {
            if (row.values[columnDict[column]] == 1) {
              outPath[1] = column
              doCopy(inLocation, outPath, fileName, 'Row ' + rowNumber + ': ' + fileName)
            }
          })
        } else {
          doCopy(inLocation, outPath, fileName, 'Row ' + rowNumber + ': ' + fileName)
        }
        row.getCell(columnDict['All']).value = DateTime.now().toISO()
      } else {
        console.log('Row ' + rowNumber + ': ' + '!!!!!!! No Name defined')
      }
    } else {
      console.log('Filtered Row ' + rowNumber + '')
    }
  })
  workbook.xlsx.writeFile(catalogFile)
}
