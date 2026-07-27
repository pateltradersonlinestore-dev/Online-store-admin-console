// ==========================================
// 1. PROFILE AND ADDRESS MANAGEMENT
// ==========================================
function findRowIndexByEmailCaseInsensitive(email, sheet) {
  if (!email) return -1;
  const dataArray = sheet.getDataRange().getValues();
  const searchEmail = String(email).toLowerCase().trim();
  let headers = dataArray[0];
  let emailColIndex = headers.indexOf("Email");
  if(emailColIndex === -1) emailColIndex = 1; 

  for (let i = 1; i < dataArray.length; i++) {
    if (dataArray[i][emailColIndex] && String(dataArray[i][emailColIndex]).toLowerCase().trim() === searchEmail) {
      return i + 1;
    }
  }
  return -1;
}

function saveDedicatedAddress(userId, name, addressesJSON, currentAddressFileId) {
  let folderId = PropertiesService.getScriptProperties().getProperty("ADDRESS_VAULT_ID");
  let folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  let fileId = currentAddressFileId;
  let sheet;

  if (!fileId) {
    let file = SpreadsheetApp.create("Addresses_" + name + "_" + userId);
    DriveApp.getFileById(file.getId()).moveTo(folder);
    fileId = file.getId();
    sheet = file.getActiveSheet();
    sheet.setName("Saved_Addresses");
    sheet.appendRow(["Line1", "Line2", "City", "State", "Pincode"]);
  } else {
    try {
      sheet = SpreadsheetApp.openById(fileId).getSheetByName("Saved_Addresses");
      if(!sheet) {
        sheet = SpreadsheetApp.openById(fileId).insertSheet("Saved_Addresses");
        sheet.appendRow(["Line1", "Line2", "City", "State", "Pincode"]);
      }
    } catch(e) {
      let file = SpreadsheetApp.create("Addresses_" + name + "_" + userId);
      DriveApp.getFileById(file.getId()).moveTo(folder);
      fileId = file.getId();
      sheet = file.getActiveSheet();
      sheet.setName("Saved_Addresses");
      sheet.appendRow(["Line1", "Line2", "City", "State", "Pincode"]);
    }
  }

  try {
    let addrs = JSON.parse(addressesJSON);
    if (!Array.isArray(addrs)) addrs = [addrs];
    let lastRow = sheet.getLastRow();
    if(lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    if (addrs.length > 0) {
      let rowsToInsert = addrs.map(a => [a.l1 || "", a.l2 || "", a.city || "", a.state || "", a.pin || ""]);
      sheet.getRange(2, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
    }
  } catch(e) {}
  SpreadsheetApp.flush();
  return fileId;
}

function updateSecureProfile(data) {
  const masterDB = getMasterDB();
  const emailToSearch = String(data.email).toLowerCase().trim();

  const orders = masterDB.getSheetByName("Orders").getDataRange().getValues();
  let oHeaders = orders[0];
  for (let o = 1; o < orders.length; o++) {
    if (orders[o][oHeaders.indexOf("Email")] && String(orders[o][oHeaders.indexOf("Email")]).toLowerCase().trim() === emailToSearch && orders[o][oHeaders.indexOf("Status")] === "Pending") {
      return {status: "error", message: "Profile Locked: You have a pending order. Complete it before updating."};
    }
  }

  const sheet = masterDB.getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(emailToSearch, sheet);
  if(rowIndex === -1) return {status: "error", message: "User not found."};

  const sData = sheet.getDataRange().getValues();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let lowerHeaders = headers.map(h => String(h).toLowerCase());
  let limitCountIndex = headers.indexOf("UpdateCount");
  let limitMonthIndex = headers.indexOf("Month");
  let userIdIndex = headers.indexOf("UserID");

  const currentMonth = new Date().getMonth() + "-" + new Date().getFullYear();
  let limitCount = Number(sData[rowIndex-1][limitCountIndex] || 0); 
  let limitMonth = sData[rowIndex-1][limitMonthIndex] || ""; 
  let userId = sData[rowIndex-1][userIdIndex];
  let trueName = sData[rowIndex-1][headers.indexOf("Name")];

  if (limitMonth !== currentMonth) { limitCount = 0; limitMonth = currentMonth; }
  if (limitCount >= 5) {
    createSecurityLog(userId, data.email, "System", "Profile Update Blocked (Limit Reached)");
    return {status: "error", message: "You cannot update your profile more than 5 times a month."};
  }

  let mappedData = {};
  Object.keys(data).forEach(k => {
    if (k === "action" || k === "token" || k === "email") return;
    let lowerK = String(k).toLowerCase();
    let exactHeaderName = k; 

    if(lowerK === "name") exactHeaderName = "Name";
    else if(lowerK === "phone") exactHeaderName = "Phone";
    else if(lowerK === "profilepic") exactHeaderName = "ProfilePic";
    else if(lowerK === "addressesjson" || lowerK === "address") exactHeaderName = "Addresses_JSON_Payload";
    else {
      let existingIdx = lowerHeaders.indexOf(lowerK);
      if(existingIdx !== -1) exactHeaderName = headers[existingIdx];
    }
    mappedData[exactHeaderName] = data[k];
  });

  let addressPayload = mappedData["Addresses_JSON_Payload"];
  delete mappedData["Addresses_JSON_Payload"]; 

  headers = ensureDynamicColumns(sheet, mappedData);

  Object.keys(mappedData).forEach(key => {
    let colIndex = headers.indexOf(key) + 1;
    if (colIndex > 0) sheet.getRange(rowIndex, colIndex).setValue(mappedData[key]);
  });

  if (addressPayload) {
    let addrFileIdIndex = headers.indexOf("AddressFileID");
    if (addrFileIdIndex === -1) {
      sheet.getRange(1, headers.length + 1).setValue("AddressFileID");
      headers.push("AddressFileID");
      addrFileIdIndex = headers.length - 1;
    }

    let currentAddrFileId = sheet.getRange(rowIndex, addrFileIdIndex + 1).getValue();
    let newName = mappedData["Name"] || trueName;
    let newAddrFileId = saveDedicatedAddress(userId, newName, addressPayload, currentAddrFileId);

    if (currentAddrFileId !== newAddrFileId) {
      sheet.getRange(rowIndex, addrFileIdIndex + 1).setValue(newAddrFileId);
    }
    let legacyAddrIdx = headers.indexOf("Addresses_JSON");
    if (legacyAddrIdx > -1) sheet.getRange(rowIndex, legacyAddrIdx + 1).clearContent();
  }

  sheet.getRange(rowIndex, limitCountIndex + 1).setValue(limitCount + 1);
  sheet.getRange(rowIndex, limitMonthIndex + 1).setValue(currentMonth);

  let vaultIdIndex = headers.indexOf("VaultID");
  let vaultId = sData[rowIndex-1][vaultIdIndex];
  if(vaultId && mappedData["Name"]) {
    try { SpreadsheetApp.openById(vaultId).getSheetByName("Profile_And_Orders").getRange("B5").setValue(mappedData["Name"]); } catch(e) {}
  }

  SpreadsheetApp.flush(); 
  createSecurityLog(userId, data.email, "Browser", "Profile Updated & Address Safely Moved to Vault");
  return {status: "success", message: "Profile updated successfully.", count: limitCount + 1};
}

// ==========================================
// 2. OTP AUTHENTICATION LOGIC
// ==========================================
function requestRegisterOTP(name, email, phone) {
  const sheet = getMasterDB().getSheetByName("Users");
  const searchEmail = String(email).toLowerCase().trim();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    let dbEmail = data[i][headers.indexOf("Email")] ? String(data[i][headers.indexOf("Email")]).toLowerCase().trim() : "";
    let dbPhone = data[i][headers.indexOf("Phone")] ? String(data[i][headers.indexOf("Phone")]).trim() : "";
    if (dbEmail === searchEmail || (phone && dbPhone === phone)) {
      if(data[i][headers.indexOf("Status")] === "Active") return {status: "error", message: "This email or phone number is already registered. Please log in."};
    }
  }

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const newUserId = "PT-" + Math.floor(10000 + Math.random() * 90000);
  let existingRow = findRowIndexByEmailCaseInsensitive(email, sheet);
  
  if(existingRow > 0) {
    sheet.getRange(existingRow, headers.indexOf("OTP") + 1).setValue(otp); 
  } else {
    sheet.appendRow([name, searchEmail, phone || "", otp, "Pending_OTP", "", 0, "Unknown", 0, "", newUserId, "", "", ""]);
  }
  SpreadsheetApp.flush(); 
  sendOTPEmail(email, otp, "Registration");
  return {status: "success", message: "Registration OTP sent successfully."};
}

function verifyRegisterOTP(email, otp, deviceInfo) {
  const sheet = getMasterDB().getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(email, sheet);

  if (rowIndex > 0) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    let otpIndex = headers.indexOf("OTP");
    if(sData[otpIndex] == otp) {
      sheet.getRange(rowIndex, otpIndex + 1).clearContent(); 
      sheet.getRange(rowIndex, headers.indexOf("Status") + 1).setValue("Active"); 
      sheet.getRange(rowIndex, headers.indexOf("Visits") + 1).setValue(1); 
      sheet.getRange(rowIndex, headers.indexOf("Device") + 1).setValue(deviceInfo || "Browser");
      
      const userId = sData[headers.indexOf("UserID")];
      const name = sData[headers.indexOf("Name")];
      const phone = sData[headers.indexOf("Phone")];

      const vaultId = generateVisualCustomerVault(userId, name, email);
      sheet.getRange(rowIndex, headers.indexOf("VaultID") + 1).setValue(vaultId); 

      SpreadsheetApp.flush(); 
      createSecurityLog(userId, email, deviceInfo, "Account Created & Verified");
      sendRegistrationPDF(email, name, phone, userId);
      return {status: "success", message: "Registration successful!", userId: userId};
    }
  }
  return {status: "error", message: "Invalid OTP provided."};
}

function requestLoginOTP(email) {
  const sheet = getMasterDB().getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(email, sheet);

  if (rowIndex > 0) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const status = sheet.getRange(rowIndex, headers.indexOf("Status") + 1).getValue();
    if(status === "Active") {
      const otp = Math.floor(100000 + Math.random() * 900000).toString();
      sheet.getRange(rowIndex, headers.indexOf("OTP") + 1).setValue(otp);
      SpreadsheetApp.flush();
      sendOTPEmail(email, otp, "Login");
      return {status: "success", message: "Login OTP sent successfully."};
    }
  }
  return {status: "error", message: "Account not found! Please register first."};
}

function verifyLoginOTP(email, otp, deviceInfo) {
  const sheet = getMasterDB().getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(email, sheet);

  if (rowIndex > 0) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (sData[headers.indexOf("OTP")] == otp) {
      sheet.getRange(rowIndex, headers.indexOf("OTP") + 1).clearContent(); 
      let visits = Number(sData[headers.indexOf("Visits")] || 0) + 1;
      sheet.getRange(rowIndex, headers.indexOf("Visits") + 1).setValue(visits);
      sheet.getRange(rowIndex, headers.indexOf("Device") + 1).setValue(deviceInfo || "Browser"); 

      SpreadsheetApp.flush();
      const userId = sData[headers.indexOf("UserID")];
      const userName = sData[headers.indexOf("Name")];
      createSecurityLog(userId, email, deviceInfo, "Successful Login");
      sendDeviceAlert(email, deviceInfo, visits); 
      return {status: "success", message: "Login successful!", userId: userId, name: userName};
    }
  }
  return {status: "error", message: "Invalid OTP provided."};
}
