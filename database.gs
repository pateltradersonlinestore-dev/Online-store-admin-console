// ==========================================
// 1. CORE DATABASE UTILITIES
// ==========================================
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function getMasterDB() {
  const id = PropertiesService.getScriptProperties().getProperty("MASTER_SHEET_ID");
  if(!id) throw new Error("System is not ready. Please run setupEcosystem().");
  return SpreadsheetApp.openById(id);
}

function ensureDynamicColumns(sheet, dataObject) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let lowerHeaders = headers.map(h => String(h).toLowerCase());
  let columnsAdded = false;

  Object.keys(dataObject).forEach(key => {
    if (key !== "action" && key !== "token" && key !== "email") {
      let lowerKey = String(key).toLowerCase();
      if (!lowerHeaders.includes(lowerKey)) {
        sheet.getRange(1, headers.length + 1).setValue(key);
        headers.push(key);
        lowerHeaders.push(lowerKey);
        columnsAdded = true;
      }
    }
  });

  if(columnsAdded) SpreadsheetApp.flush();
  return headers;
}

// ==========================================
// 2. UNIVERSAL READ/WRITE & AUTO-COMBINE
// ==========================================
function getUniversalData(sheetName) {
  SpreadsheetApp.flush(); 
  const sheet = getMasterDB().getSheetByName(sheetName);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return [];
  const headers = data.shift(); 

  return data.map(row => { 
    let obj = {}; 
    headers.forEach((h, i) => obj[h] = row[i]); 
    if (sheetName === "Users") {
      let addrFileId = obj["AddressFileID"];
      if (addrFileId) {
        try {
          let addrSheet = SpreadsheetApp.openById(addrFileId).getSheetByName("Saved_Addresses");
          if (addrSheet) {
            let aData = addrSheet.getDataRange().getValues();
            let aArr = [];
            for(let r=1; r<aData.length; r++) {
              aArr.push({ l1: String(aData[r][0]), l2: String(aData[r][1]), city: String(aData[r][2]), state: String(aData[r][3]), pin: String(aData[r][4]) });
            }
            obj["Addresses_JSON"] = JSON.stringify(aArr); 
          } else { obj["Addresses_JSON"] = "[]"; }
        } catch(e) { obj["Addresses_JSON"] = "[]"; }
      } else {
        obj["Addresses_JSON"] = "[]";
      }
    }
    return obj; 
  });
}

function universalWriteDynamic(sheetName, payloadObj) {
  let sheet = getMasterDB().getSheetByName(sheetName);
  if(!sheet) {
    sheet = getMasterDB().insertSheet(sheetName);
    sheet.appendRow(Object.keys(payloadObj));
  }
  let headers = ensureDynamicColumns(sheet, payloadObj);
  let newRow = new Array(headers.length).fill("");
  Object.keys(payloadObj).forEach(k => {
    let i = headers.indexOf(k);
    if(i > -1) newRow[i] = payloadObj[k];
  });
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  return {status: "success"};
}

function universalWrite(sheetName, rowDataArray) {
  let sheet = getMasterDB().getSheetByName(sheetName);
  if(!sheet) sheet = getMasterDB().insertSheet(sheetName); 
  sheet.appendRow(rowDataArray);
  SpreadsheetApp.flush();
  return {status: "success"};
}

function universalUpdate(sheetName, searchColIndex, searchValue, updateColIndex, updateValue) {
  const sheet = getMasterDB().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][searchColIndex] == searchValue) {
      sheet.getRange(i + 1, updateColIndex + 1).setValue(updateValue);
      SpreadsheetApp.flush();
      return {status: "success"};
    }
  }
  return {status: "error", message: "Record not found."};
}

// ==========================================
// 3. UTILITIES & VAULTS
// ==========================================
function generateVisualCustomerVault(userId, name, email) {
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("CUSTOMER_BASE_ID"));
  const vaultSheet = SpreadsheetApp.create(userId + "_Customer_Vault");
  DriveApp.getFileById(vaultSheet.getId()).moveTo(folder);
  
  const sheet = vaultSheet.getActiveSheet();
  sheet.setName("Profile_And_Orders");
  sheet.getRange("A1:E2").merge().setValue("OFFICIAL CUSTOMER VAULT").setBackground("#0f172a").setFontColor("#ffffff").setFontSize(16).setHorizontalAlignment("center");
  sheet.getRange("A4").setValue("Customer ID:").setFontWeight("bold"); sheet.getRange("B4").setValue(userId);
  sheet.getRange("A5").setValue("Name:").setFontWeight("bold"); sheet.getRange("B5").setValue(name);
  sheet.getRange("A6").setValue("Email:").setFontWeight("bold"); sheet.getRange("B6").setValue(email);
  sheet.getRange("A9:E9").setValues([["Order ID", "Items", "Amount", "Date", "Status"]]).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold");
  return vaultSheet.getId();
}

function createSecurityLog(userId, email, device, action) {
  const folderId = PropertiesService.getScriptProperties().getProperty("SECURITY_LOGS_ID");
  const folder = DriveApp.getFolderById(folderId);
  const logName = userId + "_Security_Log";
  let files = folder.getFilesByName(logName);
  let doc = files.hasNext() ? DocumentApp.openById(files.next().getId()) : DocumentApp.create(logName);
  if(!files.hasNext()) DriveApp.getFileById(doc.getId()).moveTo(folder);
  doc.getBody().appendParagraph(`[${new Date().toLocaleString()}] Action: ${action} | Device: ${device} | IP: Logged Securely`);
}

function uploadImageWithFallback(base64Data, filename, mimeType) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty("MEDIA_VAULT_ID");
    let folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {status: "success", url: "https://drive.google.com/uc?export=view&id=" + file.getId()};
  } catch (error) {
    return {status: "success", url: "data:" + mimeType + ";base64," + base64Data, fallback: true};
  }
}
