var msal = require('@azure/msal-node');

const config = require(`./config/AAD.json`)
 
//console.log(config)

const confidentialClientApplication = new msal.ConfidentialClientApplication({auth: config.auth});
const clientCredentialRequest = { scopes: config.scopes, skipCache: true }

module.exports = {
  getToken: () => {
    return confidentialClientApplication.acquireTokenByClientCredential(clientCredentialRequest)
    .then((response) => {
      //console.log("Response: ", response)
      return response
    }).catch((error) => {
      console.log(JSON.stringify(error))
    })
  }
}


	
