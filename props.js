const tokenService = require('./token.js')

main()

async function main() {
  try {
    const token = await tokenService.getToken()
    console.log("Token: ", token)
    const accessToken = token.accessToken
    console.log("Access Token: ", accessToken)

    payload = {
      method: 'GET',
      headers: { "Authorization": "Bearer " + accessToken, "Accept": "application/json; odata=verbose" },
    }
  
    //fetch("https://graph.microsoft.com/v1.0/sites/root", payload)
    fetch("https://westernpower.sharepoint.com/sites/WPCWorking/_api/web", payload)
        .then(response => {
            console.log(response)
        }
    ) 
  } catch (e) {
      console.log(e)
  }

}
