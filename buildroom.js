const configData = require('./config')
const { getToken, resolveShortLink } = require('./graph')

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

function linkUrl (rawLink) {
  if (typeof rawLink !== 'object') return rawLink
  // Short links often display the file name rather than the URL
  if (typeof rawLink.text === 'string' && /^https?:/i.test(rawLink.text)) return rawLink.text
  return rawLink.hyperlink || rawLink.text
}

function isShortLink (link) {
  return /\/:[a-z]+:\/[sg]\//i.test(link) && !/\/Shared( |%20)Documents/i.test(link)
}

// folder is the path within the document library, e.g. '/General/Sub/'
function fileInfoFrom (site, folder, name) {
  const extension = (name.match(patt) || [''])[0]
  return { site, path: folder.replace('/General', ''), name, extension }
}

function sourceLocation (fileInfo) {
  return folderSet.sps[fileInfo.site] + '/' + fileInfo.path + fileInfo.name
}

// Library view links (.../Forms/AllItems.aspx?id=/sites/Site/Shared Documents/folder/name)
// carry the server-relative path in the id parameter
function libraryViewPath (link) {
  try {
    const id = new URL(link).searchParams.get('id')
    return id && id.includes('/Shared Documents/') ? id : null
  } catch (error) {
    return null
  }
}

function processLink (link) {
  const serverPath = libraryViewPath(link)
  if (serverPath) {
    const libraryAt = serverPath.indexOf('/Shared Documents')
    const lastSlash = serverPath.lastIndexOf('/')
    const sitePath = serverPath.substring(0, libraryAt)
    return fileInfoFrom(
      sitePath.substring(sitePath.lastIndexOf('/') + 1),
      serverPath.substring(libraryAt + 17, lastSlash + 1),
      serverPath.substring(lastSlash + 1)
    )
  }

  const decoded = decodeURI(link).replace('%26', '&')
  const lastSlash = decoded.lastIndexOf('/')
  const longName = decoded.substring(lastSlash + 1)
  const longPath = decoded.substring(0, lastSlash + 1)
  const query = longName.lastIndexOf('?')
  const name = longName.substring(0, query)

  const folder = longPath.substring(longPath.indexOf('/Shared Documents') + 17)
  const remPath = longPath.substring(0, longPath.indexOf('/Shared Documents'))
  const site = remPath.substring(remPath.lastIndexOf('/') + 1)
  return fileInfoFrom(site, folder, name)
}

// Short links are cached in the catalog as Folder = 'site/folder/' and FileName = 'name'
function cachedFileInfo (row, columnDict) {
  const folderCell = cellValue(row.values[columnDict.Folder])
  const name = cellValue(row.values[columnDict.FileName])
  if (!folderCell || !name) return null
  const firstSlash = folderCell.indexOf('/')
  if (firstSlash < 0) return null
  return fileInfoFrom(folderCell.substring(0, firstSlash), folderCell.substring(firstSlash), name)
}

async function resolveShortLinks (ws, columnDict) {
  const pending = []
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1 || !row.values[columnDict.Link]) return
    const link = linkUrl(row.values[columnDict.Link])
    if (!isShortLink(link)) return
    const cached = cachedFileInfo(row, columnDict)
    // Re-resolve if the cached file has moved, unless the site isn't synced locally anyway
    if (cached && (folderSet.sps[cached.site] === undefined || fs.existsSync(sourceLocation(cached)))) return
    pending.push({ row, rowNumber, link })
  })
  if (!pending.length) return 0

  console.log('Resolving', pending.length, 'short links via Microsoft Graph')
  const token = await getToken()
  let resolved = 0
  for (const { row, rowNumber, link } of pending) {
    try {
      const { site, folder, name } = await resolveShortLink(link, token)
      row.getCell(columnDict.Folder).value = site + folder
      row.getCell(columnDict.FileName).value = name
      console.log(rowNumber + ': resolved -> ' + site + folder + name)
      resolved++
    } catch (error) {
      console.log(rowNumber, ' Could not resolve short link:', link, error.message)
    }
  }
  return resolved
}

async function saveCatalog (workbook) {
  try {
    fs.copyFileSync(catalogFile, catalogFile.replace(/\.xlsx$/i, '.bak.xlsx'))
    await workbook.xlsx.writeFile(catalogFile)
    console.log('Catalogue updated with resolved short links:', catalogFile)
  } catch (error) {
    console.log('Could not save catalogue (is it open in Excel?):', error.message)
  }
}

async function main () {
  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(catalogFile)

  const ws = workbook.getWorksheet('Documents')
  const columns = ws.getRow(1).values.map((column, index) => [column, index])
  columns.shift()
  //console.log(columns)
  const columnDict = Object.fromEntries(columns)
  if (columnDict.Folder === undefined || columnDict.FileName === undefined) {
    throw new Error('Catalogue needs Folder and FileName columns to cache short links')
  }
  const resolvedCount = await resolveShortLinks(ws, columnDict)
  const entries = []

  const copyAndRecord = (inLocation, outPath, fileName, rowNumber, row) => {
    const { isNew } = doCopy(inLocation, outPath, fileName, rowNumber)
    const areas = outPath.slice(1)
    entries.push({ status: isNew ? 'New' : '', name: fileName, areas, path: areas.join('\\'), row })
  }

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    if (row.values[columnDict.Link]) {
      const link = linkUrl(row.values[columnDict.Link])
      const fileInfo = isShortLink(link) ? cachedFileInfo(row, columnDict) : processLink(link)
      if (!fileInfo) {
        console.log(rowNumber, ' Unresolved short link:', link)
        return
      }
      let filter = true
      if (filterColumn!==false) {
        filter = row.values[columnDict[filterColumn]]
        if ((typeof filter === 'object') && filter.result) { filter = filter.result }
      }

      // console.log(fileInfo)
      const inLocation = sourceLocation(fileInfo)
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
  // Saved even on a dry run so short links don't need resolving again
  if (resolvedCount) await saveCatalog(workbook)
}
