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

main().catch(console.error)

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

function doCopy (inLocation, outPath, fileName, rowNo) {
  makePath(outPath)
  const outLocation = outPath.join('\\') + '\\' + fileName
  const existed = fs.existsSync(outLocation)

  if (!config.folderOnly) {
    if (!existed && !config.dryrun) {
      fs.copyFileSync(inLocation, outLocation)
      console.log(rowNo + ': ' + fileName)
    } else {
      console.log(rowNo)
    }
  }

  //console.log('.')
  return { outLocation, isNew: !existed }
}

function cellValue (value) {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value
  if (typeof value === 'object') {
    if ('result' in value) return cellValue(value.result)
    if ('text' in value) return value.text
    if (value.richText) return value.richText.map((part) => part.text).join('')
    return ''
  }
  return value
}

async function writeRoomCatalog (entries, columnDict) {
  const columns = config.catalogColumns.filter((column) => {
    if (columnDict[column] === undefined) {
      console.log('Catalog column not found in master catalogue:', column)
      return false
    }
    return true
  })

  const workbook = new Excel.Workbook()
  const ws = workbook.addWorksheet('Catalog', { views: [{ state: 'frozen', ySplit: 1 }] })
  ws.columns = [
    { header: 'Status', key: 'Status', width: 8 },
    { header: 'Name', key: 'Name', width: 60 },
    { header: 'Area', key: 'Area', width: 30 },
    { header: 'Area2', key: 'Area2', width: 30 },
    { header: 'Area3', key: 'Area3', width: 30 },
    ...columns.map((column) => ({ header: column, key: column, width: column === 'Description' ? 60 : 16 }))
  ]
  ws.getRow(1).font = { bold: true }

  const sortKey = (entry) => [0, 1, 2].map((i) => entry.areas[i] || '').concat(entry.name)
  entries.sort((a, b) => {
    const keyA = sortKey(a)
    const keyB = sortKey(b)
    for (let i = 0; i < keyA.length; i++) {
      const diff = keyA[i].localeCompare(keyB[i])
      if (diff) return diff
    }
    return 0
  })
  entries.forEach((entry) => {
    const values = {
      Status: entry.status,
      Name: { text: entry.name, hyperlink: entry.path + '\\' + entry.name },
      Area: entry.areas[0] || '',
      Area2: entry.areas[1] || '',
      Area3: entry.areas[2] || ''
    }
    columns.forEach((column) => {
      values[column] = cellValue(entry.row.values[columnDict[column]])
    })
    const row = ws.addRow(values)
    columns.forEach((column) => {
      if (values[column] instanceof Date) row.getCell(column).numFmt = 'yyyy-mm-dd'
    })
  })
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columns.length } }

  const catalogOut = outFolder + '\\' + DateTime.now().toFormat('yyyyLLdd HHmmss') + ' Catalog.xlsx'
  await workbook.xlsx.writeFile(catalogOut)
  const newCount = entries.filter((entry) => entry.status === 'New').length
  console.log('Catalog written:', catalogOut, '(' + entries.length + ' files, ' + newCount + ' new)')
}

function processLink (rawLink) {
  let link
  if (typeof rawLink === 'object') {
    link = rawLink.text
  } else {
    link = rawLink
  }
  const decoded = decodeURI(link).replace('%26', '&')
  const lastSlash = decoded.lastIndexOf('/')
  const longName = decoded.substring(lastSlash + 1)
  const longPath = decoded.substring(0, lastSlash + 1)
  const query = longName.lastIndexOf('?')
  const name = longName.substring(0, query)

  const path = longPath.substring(longPath.indexOf('/Shared Documents') + 17).replace('/General', '')
  const remPath = longPath.substring(0, longPath.indexOf('/Shared Documents'))
  const site = remPath.substring(remPath.lastIndexOf('/') + 1)
  const extension = name.match(patt)[0]
  return { site, path, name, extension }
}

async function main () {
  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(catalogFile)

  const ws = workbook.getWorksheet('Documents')
  const columns = ws.getRow(1).values.map((column, index) => [column, index])
  columns.shift()
  //console.log(columns)
  const columnDict = Object.fromEntries(columns)
  const entries = []

  const copyAndRecord = (inLocation, outPath, fileName, rowNumber, row) => {
    const { isNew } = doCopy(inLocation, outPath, fileName, rowNumber)
    const areas = outPath.slice(1)
    entries.push({ status: isNew ? 'New' : '', name: fileName, areas, path: areas.join('\\'), row })
  }

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    if (row.values[columnDict.Link]) {
      const fileInfo = processLink(row.values[columnDict.Link])
      let filter = true 
      if (filterColumn!==false) {
        filter = row.values[columnDict[filterColumn]]
        if ((typeof filter === 'object') && filter.result) { filter = filter.result }
      }

      // console.log(fileInfo)
      const inLocation = folderSet.sps[fileInfo.site] + '/' + fileInfo.path + fileInfo.name
      if (!fs.existsSync(inLocation)) {
        console.log(rowNumber, ' Source does not exist:', inLocation)
        filter = false
      }

      /*
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
      */
      // console.log(rowNumber, filter)
      if (filter) {
        let fileDate = DateTime.fromJSDate(row.values[columnDict.Date])
        if (!fileDate.isValid) { 
          console.log(rowNumber + ': Bad date - ', row.values[columnDict.Date])
          fileDate = DateTime.fromISO('2024-01-01') 
        }

        const title = row.values[columnDict['Short Title']]
        const fileName = fileDate.toFormat('yyyyLLdd') + ' ' + title.trim() + fileInfo.extension
        let outPath

        if (config.room === '___HSS') {
          // outPath = [outPath[0], '', ...outPath.slice(1)]
          config.HSSColumns.forEach((column) => {
            if (row.values[columnDict[column]] === 1) {
              outPath = [outFolder, column]
              copyAndRecord(inLocation, outPath, fileName, rowNumber, row)
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
          copyAndRecord(inLocation, outPath, fileName, rowNumber, row)
        }
        //row.getCell(columnDict[filterColumn]).value = DateTime.now().toISO()
      } else {
        // console.log('Filtered Row ' + rowNumber + '')
      }
    } else {
      console.log(rowNumber, 'No Name found!!!!')
    }
  })
  await writeRoomCatalog(entries, columnDict)
  // workbook.xlsx.writeFile(folderSet.catalogFolder + folderSet.catalogFile)
}
