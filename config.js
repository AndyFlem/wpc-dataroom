module.exports = {
  config: {
    folderSet: 'laptop',
    room: 'LenderDD',
    folderOnly: false,
    dryrun: false,
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
        catalogFolder: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\',
        catalogFile: '202610 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\temp\\Datarooms\\',
        sps: {
          WPCWorking: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\Western Power Company\\Western Power Company\\WPC HSES Management System - General'
        }
      },
      cloud: {
        catalogFolder: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\WPC Document Catalogue\\',
        catalogFile: '202610 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\Datarooms\\',//'D:\\OneDrive\\Western Power Company\\Western Power Ngonye Falls Dataroom - General\\',
        sps: {
          WPCWorking: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\OneDrive\\Western Power Company\\WPC HSES Management System - General'
        }
      }
    },
    rooms: {
      LenderDD: {
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
