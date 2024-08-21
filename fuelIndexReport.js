function generateFuelIndex(reportData) {
    const reportTemplate = DriveApp.getFileById('1zz1tn3Uim8ZRlF4pQQvKjtXrpX4uha8e59jJYDChFrU') // Access main template
    const saveFolder = DriveApp.getFolderById('16c_ktyl13bgIaK4q_iR4d_POsy4uQh5Z') // Access to save location folder
  
    const monthNames = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const reportDate = new Date(reportData.date)
  
    let currentReport;
    const fileName = `${monthNames[reportDate.getMonth()]} Fuel Index Report ${reportDate.getFullYear()} - MAPA`
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
  
    const section1 = [["Fuels", "Per Gallon", "Per Liter"], ["Gasoline", reportData.gasGal, reportData.gasLiter], ["Diesel", reportData.dieselGal, reportData.dieselLiter]]
    const section2 = [
      ["Asphalt Cement", "Per Gallon", "Per Ton", "Per Liter", "Per Metric Ton"],
      ["Viscosity Grade AC-5", reportData.viscosityGradeAC5Gal, reportData.viscosityGradeAC5Ton, reportData.viscosityGradeAC5Liter, reportData.viscosityGradeAC5MetTon],
      ["Viscosity Grade AC-10", reportData.viscosityGradeAC10Gal, reportData.viscosityGradeAC10Ton, reportData.viscosityGradeAC10Liter, reportData.viscosityGradeAC10MetTon],
      ["Viscosity Grade AC-20", reportData.viscosityGradeAC20Gal, reportData.viscosityGradeAC20Ton, reportData.viscosityGradeAC20Liter, reportData.viscosityGradeAC20MetTon],
      ["Viscosity Grade AC-30", reportData.viscosityGradeAC30Gal, reportData.viscosityGradeAC30Ton, reportData.viscosityGradeAC30Liter, reportData.viscosityGradeAC30MetTon],
      ["Grade PG-64-22", reportData.gradePG6422Gal, reportData.gradePG6422Ton, reportData.gradePG6422Liter, reportData.gradePG6422MetTon],
      ["Grade PG-67-22", reportData.gradePG6722Gal, reportData.gradePG6722Ton, reportData.gradePG6722Liter, reportData.gradePG6722MetTon],
      ["Grade PG-76-22", reportData.gradePG7622Gal, reportData.gradePG7622Ton, reportData.gradePG7622Liter, reportData.gradePG7622MetTon],
      ["Grade PG-82-22", reportData.gradePG8222Gal, reportData.gradePG8222Ton, reportData.gradePG8222Liter, reportData.gradePG8222MetTon],
      ]
    const section3 = [
      ["Emulsified Asphalt, Primes, Tack Coats", "Per Gallon", "Per Liter"],
      ["Grade SS-1", reportData.gradeSS1Gal, reportData.gradeSS1Liter],
      ["Grade RS-2C (CRS-2)", reportData.gradeRS2CGal, reportData.gradeRS2CLiter],
      ["Grade CRS-2P", reportData.gradeCRS2PGal, reportData.gradeCRS2PLiter],
      ["Grade EA-1, EPR-1, & AE-P", reportData.gradeEA1Gal, reportData.gradeEA1Liter],
      ["Grade CSS-1 & 1H (Undiluted)", reportData.gradeCSS1UndilutedGal, reportData.gradeCSS1UndilutedLiter],
      ["Grade CSS-1 & 1H", reportData.gradeCSS1Gal, reportData.gradeCSS1Liter],
      ]
  
    body.insertTable(body.getNumChildren(), section1)
    body.insertTable(body.getNumChildren(), section2)
    body.insertTable(body.getNumChildren(), section3)
  
    const bodyStyle = {};
    bodyStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000"
    bodyStyle[DocumentApp.Attribute.FONT_FAMILY] = "Oswald"
    bodyStyle[DocumentApp.Attribute.FONT_SIZE] = 9
  
    body.setAttributes(bodyStyle)
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  