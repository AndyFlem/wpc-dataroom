module.exports = {
  config: {
    folderSet: 'laptop',
    room: 'HSS',
    folderOnly: true,
    folderSets: {
      laptop: {
        catalogFolder: 'C:\\Users\\kabom\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\',
        catalogFile: '202403 WPC Documents Catalogue.xlsx',
        outFolder: 'C:\\temp\\Datarooms\\',
        sps: {
          WPCWorking: 'C:\\Users\\kabom\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'C:\\Users\\kabom\\Western Power Company\\WPC HSES Management System - Documents\\General'
        }
      },
      sable: {
        catalogFolder: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\',
        catalogFile: '202403 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\temp\\Datarooms\\',
        sps: {
          WPCWorking: 'D:\\Western Power Company\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\Western Power Company\\Western Power Company\\WPC HSES Management System - General'
        }
      },
      cloud: {
        catalogFolder: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents\\PRM Project Management\\Catalogue\\',
        catalogFile: '202403 WPC Documents Catalogue.xlsx',
        outFolder: 'D:\\OneDrive\\Western Power Company\\WPC Dataroom - Dataroom\\',
        sps: {
          WPCWorking: 'D:\\OneDrive\\Western Power Company\\WPC Working - Documents',
          WPCHSESManagementSystem: 'D:\\OneDrive\\Western Power Company\\WPC HSES Management System - General'
        }
      }
    },
    rooms: {
      Test: {
        filterColumn: 'Test'
      },
      HSS: {
        filterColumn: 'HSS Assess Doc'
      },
      CapLink: {
        filterColumn: 'CAPLINK Dataroom2'
      },
      All: {
        filterColumn: false
      }

    },
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
