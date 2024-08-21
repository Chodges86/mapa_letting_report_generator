function generateReport(reportData) {
    const reportTemplate = DriveApp.getFileById('1-6CkRT2DLBEL9CIAyIf1j7yHbJGfCKWwyNBGNcqp01M') // Access main template
    const saveFolder = DriveApp.getFolderById('16c_ktyl13bgIaK4q_iR4d_POsy4uQh5Z') // Access to save location folder
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2023IMPORT")
  
    const monthNames = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const { startRow, lastRow } = reportData;
    const reportDate = new Date(sheet.getRange(startRow, 1).getValue())
  
  
    let currentReport;
    const fileName = `${monthNames[reportDate.getMonth()]} Letting Report ${reportDate.getFullYear()} - MAPA`
    const files = DriveApp.getFilesByName(fileName)
  
    const ui = SpreadsheetApp.getUi()
    if (files.hasNext()) {
      //TODO: Show notice that the file exists and user will be overwriting.  Show with file name.  Give ability to cancel operation
      const file = files.next()
      currentReport = DocumentApp.openById(file.getId())
  
    } else {
      //TODO: Show notice of new file being created with filename.  Give ability to cancel operation
      const copyOfTemplate = reportTemplate.makeCopy(fileName, saveFolder) // Creating the copy of template for manipulation
      currentReport = DocumentApp.openById(copyOfTemplate.getId())
  
    }
  
    const body = currentReport.getBody()
    const header = currentReport.getHeader()
    body.clear()
    const style = {};
    style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#ffffff"
    style[DocumentApp.Attribute.FONT_FAMILY] = "Oswald"
  
    header.getChild(1).setText(reportDate.toLocaleDateString())
    header.getChild(8).setText((`${monthNames[reportDate.getMonth()]} - ${reportDate.getFullYear()}`))
  
    header.getChild(1).setAttributes(style)
    header.getChild(8).setAttributes(style)
  
    const lettingTitles = ["Letting #", "District #", "County", "Project #", "Type Project", "Description", "Neat Tons", "SMA Tons", "Polymer Tons", "Project Total", "Range", "Bid Price", "State Estimate", "%", "Project Completion", "Lowest Bidder"]
    const lettingData = []
    const numOfRows = (lastRow - startRow) + 1
  
    for (i = 0; i < numOfRows; i++) {
      const row = +startRow + i
      const lettingNum = String(sheet.getRange(row, 2).getValue())
      const districtNum = String(sheet.getRange(row, 10).getValue())
      const county = sheet.getRange(row, 3).getValue()
      const projectNum = sheet.getRange(row, 6).getValue()
      const typeProject = sheet.getRange(row, 5).getValue()
      const description = sheet.getRange(row, 4).getValue()
  
      let NumFormat = new Intl.NumberFormat("en-US");
  
      const neatTons = NumFormat.format(sheet.getRange(row, 20).getValue())
      const smaTons = NumFormat.format(sheet.getRange(row, 17).getValue())
      const polymerTons = NumFormat.format(sheet.getRange(row, 21).getValue())
      const projectTotal = NumFormat.format(sheet.getRange(row, 19).getValue())
      const range = sheet.getRange(row, 9).getValue()
  
      let USDollar = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
      });
  
       const bidPrice = USDollar.format(sheet.getRange(row, 22).getValue()) === "$0.00" ? "-" : USDollar.format(sheet.getRange(row, 22).getValue())
      const stateEst = USDollar.format(sheet.getRange(row, 23).getValue()) === "$0.00" ? "-" : USDollar.format(sheet.getRange(row, 23).getValue())
  
      let Percent = new Intl.NumberFormat('en-US', {
        style: 'percent',
        minimumFractionDigits: 2
      });
  
      const percentage = Percent.format(sheet.getRange(row, 24).getValue()) === "NaN%" ? "-" : Percent.format(sheet.getRange(row, 24).getValue())
  
      let projectCompletion;
      if (sheet.getRange(row, 7).getValue() === "" && sheet.getRange(row, 8).getValue() === "") {
        projectCompletion = "Flexible"
      } else if (sheet.getRange(row, 7).getValue() === 0 || sheet.getRange(row, 7).getValue() === "") {
        projectCompletion = new Date(sheet.getRange(row, 8).getValue()).toLocaleDateString()
      } else {
        projectCompletion = sheet.getRange(row, 7).getValue() + " Working Days"
      }
      const lowestBidder = sheet.getRange(row, 25).getValue()
      const lettingArr = [lettingNum, districtNum, county, projectNum, typeProject, description, neatTons, smaTons, polymerTons, projectTotal, range, bidPrice, stateEst, percentage, projectCompletion, lowestBidder]
      lettingData.push(lettingArr)
    }
  
    for (i = 0; i < lettingData.length; i++) {
      // console.log(lettingData[i])
      createLettingTable(i, body, lettingTitles, lettingData[i])
    }
  
    styleTables(body)
  
  }
  
  function createLettingTable(i, body, lettingTitles, lettingData) {
    const section1 = [[lettingTitles[0], lettingTitles[1], lettingTitles[2]], [lettingData[0], lettingData[1], lettingData[2]]]
    const section2 = [[lettingTitles[3], lettingTitles[4]], [lettingData[3], lettingData[4]]]
    const section3 = [[lettingTitles[5]], [lettingData[5]]]
    const section4 = [[lettingTitles[6], lettingTitles[7], lettingTitles[8], lettingTitles[9]], [lettingData[6], lettingData[7], lettingData[8], lettingData[9]]]
  
    // const section5 = [[lettingTitles[10], lettingTitles[11], lettingTitles[12], lettingTitles[13], lettingTitles[14]], [lettingData[10], lettingData[11], lettingData[12], lettingData[13], lettingData[14]]]
  
    const section5 = [[lettingTitles[10], lettingTitles[11], lettingTitles[12], lettingTitles[13]], [lettingData[10], lettingData[11], lettingData[12], lettingData[13]]]
  
    const section6 = [[`${lettingTitles[15]}:  ${lettingData[15]}`], [`${lettingTitles[14]}:  ${lettingData[14]}`]]
  
    const section7 = []
  
    body.insertTable(body.getNumChildren(), section1)
    body.insertTable(body.getNumChildren(), section2)
    body.insertTable(body.getNumChildren(), section3)
    body.insertTable(body.getNumChildren(), section4)
    body.insertTable(body.getNumChildren(), section5)
    body.insertTable(body.getNumChildren(), section6)
    // body.insertTable(body.getNumChildren(), section7)
    body.appendPageBreak()
  
  }
  
  function styleTables(body) {
    const tables = body.getTables()
    const headerStyle = {};
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#454545"
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#FFFFFF"
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = "Oswald"
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 12
    const style = {};
    style[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FFA559"
    style[DocumentApp.Attribute.FONT_FAMILY] = "Oswald"
    style[DocumentApp.Attribute.FONT_SIZE] = 10
  
    const altStyle = {};
    altStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FFA559"
    altStyle[DocumentApp.Attribute.FONT_FAMILY] = "Oswald"
    altStyle[DocumentApp.Attribute.FONT_SIZE] = 8
  
    tables.forEach(table => {
      const numOfRows = table.getNumChildren()
      table.setBorderColor("#FFE6C7")
      for (i = 0; i < numOfRows; i++) {
        const numOfChildren = table.getRow(i).getNumChildren()
        for (c = 0; c < numOfChildren; c++) {
          if (i === 0) {
            table.getRow(i).getChild(c).setAttributes(headerStyle)
          } else {
            const text = table.getRow(i).getChild(c).getText().split("")
            if (text.length < 50) {
              table.getRow(i).getChild(c).setAttributes(style)
            } else {
              table.getRow(i).getChild(c).setAttributes(altStyle)
            }
          }
  
        }
      }
    })
  
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  