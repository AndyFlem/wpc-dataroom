const configData = require('./config')

const Excel = require('exceljs')
const fs = require('fs')
const { DateTime } = require('luxon')

const patt = /\.[^.\\/:*?"<>|\r\n]+$/i

const config = configData.config
console.log('Room: ', config.room)

const folderSet = config.folderSets[config.folderSet]
const catalogFile = folderSet.catalogFolder + folderSet.catalogFile
console.log('Catalog: ', catalogFile)

const filterColumn = config.rooms[config.room].filterColumn
console.log('Filter column: ', filterColumn)

const outFolder = folderSet.outFolder + config.room

if (!fs.existsSync(outFolder)) {
  fs.mkdirSync(outFolder)
}

main()

function makeFolder (folder) {
  if (!fs.existsSync(folder)) {
    console.log('Creating folder:', folder)
    fs.mkdirSync(folder)
  }
}

function makePath (path) {
  let workingPath = ''
  path.forEach((folder) => {
    workingPath += folder + '\\'
    makeFolder(workingPath)
  })
}

function doCopy (inLocation, outPath, fileName, log) {
  makePath(outPath)
  const outLocation = outPath.join('\\') + '\\' + fileName

  //  if (!fs.existsSync(outLocation)) {
  //    if (!fs.existsSync(inLocation)) {
  //      console.log('Source does not exist:', inLocation)
  //    } else {
  //      if (!config.folderOnly) {
  fs.copyFileSync(inLocation, outLocation)
  //      }
  console.log(log)
//    }
//  }
}

async function main () {
  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(catalogFile)

  const ws = workbook.getWorksheet('Documents')
  const columns = ws.getRow(1).values.map((column, index) => [column, index])
  columns.shift()
  const columnDict = Object.fromEntries(columns)

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return

    if (row.values[columnDict.Name]) {
      let filter = row.values[columnDict[filterColumn]]
      if ((typeof filter === 'object') && filter.result) { filter = filter.result }

      const inLocation = folderSet.sps[row.values[columnDict.SPSite].result] + '\\' + decodeURI(row.values[columnDict.SPLocation].result).replace('General/', '/').replace('%26', '&').replace('/Shared Documents/', '')
      if (!fs.existsSync(inLocation)) {
        console.log('Source does not exist:', inLocation)
        filter = false
      }
      const prevDate = DateTime.fromISO(filter)
      if (prevDate.isValid) {
        const stats = fs.statSync(inLocation)
        if (DateTime.fromJSDate(stats.mtime) > prevDate) {
          console.log(row.values[columnDict['Short Title']], 'File modified, copying again.')
          filter = true
        } else {
          console.log(row.values[columnDict['Short Title']], 'File not modified.')
          filter = false
        }
      }
      // console.log(rowNumber, filter)
      if (filter) {
        let fileDate = DateTime.fromISO(row.values[columnDict.DateString])
        if (!fileDate.isValid) { fileDate = DateTime.fromISO('2024-01-01') }

        const title = row.values[columnDict['Short Title']]
        const extension = row.values[columnDict.Name].match(patt)[0]
        const fileName = fileDate.toFormat('yyyyLLdd') + ' ' + title.trim() + extension.trim()
        let outPath

        if (config.room === 'HSS') {
          // outPath = [outPath[0], '', ...outPath.slice(1)]
          config.HSSColumns.forEach((column) => {
            if (row.values[columnDict[column]] === 1) {
              outPath = [outFolder, column]
              doCopy(inLocation, outPath, fileName, 'Row ' + rowNumber + ': ' + fileName)
            }
          })
        } else {
          outPath = [outFolder, row.values[columnDict.Area]]
          if (row.values[columnDict.Area2]) {
            outPath.push(row.values[columnDict.Area2])
          }
          if (row.values[columnDict.Area3]) {
            outPath.push(row.values[columnDict.Area3])
          }
          doCopy(inLocation, outPath, fileName, 'Row ' + rowNumber + ': ' + fileName)
        }
        row.getCell(columnDict[filterColumn]).value = DateTime.now().toISO()
      } else {
        // console.log('Filtered Row ' + rowNumber + '')
      }
    } else {
      console.log(rowNumber, 'No Name found!!!!')
    }
  })
  workbook.xlsx.writeFile(folderSet.catalogFolder + 'Catalog Out.xlsx')
}
