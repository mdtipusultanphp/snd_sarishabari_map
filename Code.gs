/**
 * Apps Script (File: Code.gs)
 * সার্ভার সাইড লজিক
 */

const SPREADSHEET_ID = "1erpnvXFjmDAK8t9-MrWUC2BEp-hIXfZVL7zR1Xj9imk";
const SHEET_NAME = "gps"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('বকেয়া ট্র্যাকার প্রো')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCustomerData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return [];

    const sheetHeaders = values[0].map(h => String(h).trim().toLowerCase()); 
    const dataRows = values.slice(1);
    const headerIndices = {};
    
    const requiredHeaders = ['cust_nam', 'cust_num', 'latitude', 'longitude', 'mobile_no', 'address', 'total_balance', 'book', 'meter_no', 'catagory'];
    
    requiredHeaders.forEach(reqH => {
        const index = sheetHeaders.indexOf(reqH);
        if (index > -1) headerIndices[reqH] = index;
    });

    const dataForMap = dataRows.map(row => {
      const data = {};
      requiredHeaders.forEach(key => {
          if (headerIndices[key] !== undefined) data[key] = row[headerIndices[key]];
      });

      const lat = parseFloat(String(data.latitude).trim());
      const lng = parseFloat(String(data.longitude).trim());
      const balance = parseFloat(String(data.total_balance || '').trim().replace(/,/g, '')) || 0; 

      if (isNaN(lat) || isNaN(lng)) return null; 
      
      let iconColor = 'green'; 
      if (balance >= 15000) iconColor = 'red'; 
      else if (balance >= 5000) iconColor = 'orange'; 

      return {
        lat: lat, lng: lng,
        name: data.cust_nam || 'N/A', 
        custNum: String(data.cust_num || '').trim(),
        mobile: String(data.mobile_no || '').trim() || 'N/A',
        address: data.address || 'N/A',
        bookNo: data.book || 'N/A',
        meterNo: String(data.meter_no || '').trim(),
        tariff: data.catagory || 'N/A',
        balance: balance,
        iconColor: iconColor,
      };
    }).filter(item => item !== null);
    
    return dataForMap;
  } catch (error) {
    Logger.log("ERROR: " + error.toString());
    return [];
  }
}

function updateCustomerBalance(custNum, newBalance) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    const custNumIndex = headers.indexOf('cust_num');
    const balanceIndex = headers.indexOf('total_balance');

    if (custNumIndex === -1 || balanceIndex === -1) return "Error: Column missing";

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][custNumIndex]).trim() === String(custNum).trim()) {
        sheet.getRange(i + 1, balanceIndex + 1).setValue(newBalance);
        return "Success";
      }
    }
    return "Error: Customer not found";
  } catch (e) { return "Error: " + e.toString(); }
}
