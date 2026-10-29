module.exports = {
  config: {
    folderSet: 'sable',
    room: 'Dataroom October 2026',
    folderOnly: false,
    dryrun: false,
    graph: {
      tenantId: '82a04285-0318-423f-b953-bcec333c384b',
      clientId: '8abd6508-fe42-4a13-9c5a-9d48ee4c05f4',
      scopes: ['Files.Read.All'],
      tokenCache: '.msal-cache.json'
    },
    folderSets: {
      laptop: {
        catalogFolder: 'C:\\Users\\AndyFleming\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\WPC Document Catalogue\\',
        catalogFile: '202610 WPC Documents Catalogue.xlsx',
        outFolder: 'C:\\temp\\Datarooms\\',
        sps: {
          WPCWorking: 'C:\\Users\\AndyFleming\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'C:\\Users\\AndyFleming\\Western Power Company\\WPC HSES Management System - Documents\\General'
        }
      },
      sable: {
        catalogFolder: 'D:\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\WPC Document Catalogue\\',
        catalogFile: '202610 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\Western Power Company\\Ngonye Falls Project Finance - Documents\\',
        sps: {
          WPCWorking: 'D:\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\Western Power Company\\WPC HSES Management System - General'
        }
      },
      cloud: {
        catalogFolder: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\WPC Document Catalogue\\',
        catalogFile: '202610 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\Datarooms\\',
        sps: {
          WPCWorking: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\OneDrive\\Western Power Company\\WPC HSES Management System - General'
        }
      }
    },
    rooms: {
      'Dataroom October 2026': {
        filterColumn: 'LenderDD'
      },
      Summary: {
        filterColumn: 'Summary'
      },
      All: {
        filterColumn: false
      }
    },
    catalogColumns: ['Date', 'Update Date', 'Short Title', 'Type', 'State',
      'Organisation', 'Author', 'Version', 'Extent', 'Scheme'],
    HSSColumns: [
      '00 General',
      '01 Environmental and Social',
      '02 Labour and Working Conditions',
      '03 Water Quality and Sediments',
      '04 Community Impacts and Safety',
      '05 Resettlement',
      '06 Biodiversity and Invasive Species',
      '07 Indigenous Peoples',
      '08 Cultural Heritage',
      '09 Governance and Procurement',
      '10 Communications and Consultation',
      '11 Hydrological Resource',
      '12 Climate Change'
    ]
  }
}
