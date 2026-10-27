const msal = require('@azure/msal-node')
const fs = require('fs')
const path = require('path')

const graphConfig = require('./config').config.graph
const graphRoot = 'https://graph.microsoft.com/v1.0'

let pca

function client () {
  if (pca) return pca
  if (!graphConfig || !graphConfig.clientId || graphConfig.clientId.startsWith('<')) {
    throw new Error('config.graph.tenantId / clientId must be set to resolve SharePoint short links')
  }
  const cacheFile = path.join(__dirname, graphConfig.tokenCache)
  const cachePlugin = {
    beforeCacheAccess: async (context) => {
      if (fs.existsSync(cacheFile)) context.tokenCache.deserialize(fs.readFileSync(cacheFile, 'utf8'))
    },
    afterCacheAccess: async (context) => {
      if (context.cacheHasChanged) fs.writeFileSync(cacheFile, context.tokenCache.serialize())
    }
  }
  pca = new msal.PublicClientApplication({
    auth: {
      clientId: graphConfig.clientId,
      authority: 'https://login.microsoftonline.com/' + graphConfig.tenantId
    },
    cache: { cachePlugin }
  })
  return pca
}

async function getToken () {
  const app = client()
  const scopes = graphConfig.scopes
  const accounts = await app.getTokenCache().getAllAccounts()
  if (accounts.length) {
    try {
      const response = await app.acquireTokenSilent({ account: accounts[0], scopes })
      return response.accessToken
    } catch (error) {
      console.log('Silent token refresh failed, signing in again')
    }
  }
  const response = await app.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (deviceCode) => console.log(deviceCode.message)
  })
  return response.accessToken
}

async function graphGet (url, token) {
  const response = await fetch(url, { headers: { Authorization: 'Bearer ' + token } })
  if (!response.ok) {
    throw new Error('Graph ' + response.status + ' ' + response.statusText + ': ' + (await response.text()))
  }
  return response.json()
}

function safeDecode (text) {
  try {
    return decodeURIComponent(text)
  } catch (error) {
    return text
  }
}

// Returns { site, folder, name } where folder is the path within the document library,
// with leading and trailing slashes (e.g. '/PRM Project Management/Catalogue/')
async function resolveShortLink (url, token) {
  const shareId = 'u!' + Buffer.from(url).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')
  const item = await graphGet(graphRoot + '/shares/' + shareId + '/driveItem?$select=name,parentReference', token)

  const parentPath = item.parentReference.path || ''
  const rootRelative = safeDecode(parentPath.substring(parentPath.indexOf('root:') + 5))
  const folder = rootRelative.replace(/\/$/, '') + '/'

  const siteMatch = url.match(/\/:[a-z]+:\/[sg]\/([^/?]+)\//i)
  let site
  if (siteMatch) {
    site = decodeURIComponent(siteMatch[1])
  } else {
    const siteInfo = await graphGet(graphRoot + '/sites/' + item.parentReference.siteId + '?$select=webUrl', token)
    site = decodeURIComponent(siteInfo.webUrl.replace(/\/$/, '').split('/').pop())
  }

  return { site, folder, name: item.name }
}

module.exports = { getToken, resolveShortLink }
