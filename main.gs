/**
* ONLINE STORE BACKEND ENGINE
* Optimized, AI-free, and thoroughly translated for professional e-commerce operations.
*/

// ==========================================
// 1. GLOBAL CONSTANTS & MASTER KEYS
// ==========================================
const MAIN_FOLDER_ID = "1jlHfRHBat1ZlO32uRkBYM2Cw03vhgv86"; 
const RAZORPAY_LIVE_KEY = "rzp_live_SVstISgPrcivjP"; 
const LOGO_URL = "https://i.ibb.co/G4xsrr6j/qafx-Ln-R-md.png"; 
const APP_TOKEN = "PT_SECURE_2026"; 

// ==========================================
// 2. SYSTEM ECOSYSTEM SETUP
// ==========================================
function setupEcosystem() {
  const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);

  const adminVault = getOrCreateFolder(mainFolder, "Admin_Vault");
  const customerBase = getOrCreateFolder(mainFolder, "Customer_Base");
  const securityLogs = getOrCreateFolder(mainFolder, "Security_Logs");
  const mediaVault = getOrCreateFolder(mainFolder, "Media_Vault"); 
  const addressVault = getOrCreateFolder(mainFolder, "Address_Vault"); 

  PropertiesService.getScriptProperties().setProperty("CUSTOMER_BASE_ID", customerBase.getId());
  PropertiesService.getScriptProperties().setProperty("SECURITY_LOGS_ID", securityLogs.getId());
  PropertiesService.getScriptProperties().setProperty("MEDIA_VAULT_ID", mediaVault.getId()); 
  PropertiesService.getScriptProperties().setProperty("ADDRESS_VAULT_ID", addressVault.getId());

  let masterSheetId = PropertiesService.getScriptProperties().getProperty("MASTER_SHEET_ID");
  if (!masterSheetId) {
    const newMaster = SpreadsheetApp.create("MASTER_DB");
    DriveApp.getFileById(newMaster.getId()).moveTo(adminVault);
    masterSheetId = newMaster.getId();
    PropertiesService.getScriptProperties().setProperty("MASTER_SHEET_ID", masterSheetId);

    const ss = SpreadsheetApp.openById(masterSheetId);
    ss.insertSheet("Users").appendRow(["Name", "Email", "Phone", "OTP", "Status", "ProfilePic", "Visits", "Device", "UpdateCount", "Month", "UserID", "VaultID", "AddressFileID", "LastSeen"]);
    ss.insertSheet("Products").appendRow(["ID", "Name", "Category", "Price", "MRP", "Stock", "Image_URL", "Description", "Status"]);
    ss.insertSheet("Categories").appendRow(["CategoryName", "Status"]); 
    ss.insertSheet("Orders").appendRow(["Order_ID", "Email", "Items", "Amount", "Razorpay_ID", "Status", "Date", "Delivery_Address"]);
    ss.insertSheet("Carts").appendRow(["Email", "CartJSON"]);
    ss.insertSheet("Waitlist").appendRow(["ProductID", "Email", "Date"]);
    ss.insertSheet("Reviews").appendRow(["ProductID", "UserName", "Rating", "Comment", "Date"]);

    let defaultSheet = ss.getSheetByName("Sheet1");
    if(defaultSheet) defaultSheet.setName("Dashboard");
  }
  console.log("✅ Ecosystem installed! Address_Vault and Auto-Column are ready.");
}

// ==========================================
// 3. API ROUTING (GET & POST)
// ==========================================
function doGet(e) {
  if (!e.parameter || e.parameter.token !== APP_TOKEN) return jsonResponse({status: "error", message: "🚫 Access Denied!"});
  const action = e.parameter.action;

  if (action === "getProducts") return jsonResponse(getProductsFromSheet());
  if (action === "getOrderHistory") return jsonResponse(getOrderHistory(e.parameter.email));
  if (action === "getCart") return jsonResponse(getCart(e.parameter.email));
  if (action === "getReviews") return jsonResponse(getUniversalData("Reviews"));
  if (action === "getStoreConfig") return jsonResponse({status: "success", rzpKey: RAZORPAY_LIVE_KEY});
  if (action === "readUniversal") return jsonResponse(getUniversalData(e.parameter.sheetName));

  return jsonResponse({status: "error", message: "Invalid GET Action"});
}

function doPost(e) {
  if(!e.postData || !e.postData.contents) return jsonResponse({status: "error", message: "No data received."});
  const data = JSON.parse(e.postData.contents);
  if (data.token !== APP_TOKEN) return jsonResponse({status: "error", message: "🚫 Access Denied!"});
  const action = data.action;

  // AUTH ACTIONS
  if (action === "requestRegisterOTP") return jsonResponse(requestRegisterOTP(data.name, data.email, data.phone));
  if (action === "verifyRegisterOTP") return jsonResponse(verifyRegisterOTP(data.email, data.otp, data.deviceInfo));
  if (action === "requestLoginOTP") return jsonResponse(requestLoginOTP(data.email));
  if (action === "verifyLoginOTP") return jsonResponse(verifyLoginOTP(data.email, data.otp, data.deviceInfo));

  // CUSTOMER ACTIONS
  if (action === "updateProfile") return jsonResponse(updateSecureProfile(data));
  if (action === "placeOrder") return jsonResponse(processOrder(data));
  if (action === "syncCart") return jsonResponse(syncCart(data.email, data.cartItems));
  if (action === "addWaitlist") return jsonResponse(addToWaitlist(data.email, data.productId));
  if (action === "addReview") return jsonResponse(addReview(data));
  if (action === "logActivity") return jsonResponse(logActivity(data.email, data.activityType, data.details)); 

  // ADMIN ACTIONS
  if (action === "updateOrderStatus") return jsonResponse(updateOrderStatusAndAlert(data.orderId, data.status));
  if (action === "universalWrite") return jsonResponse(universalWrite(data.sheetName, data.rowData)); 
  if (action === "universalWriteDynamic") return jsonResponse(universalWriteDynamic(data.sheetName, data.payload)); 
  if (action === "universalUpdate") return jsonResponse(universalUpdate(data.sheetName, data.searchCol, data.searchValue, data.updateCol, data.updateValue));
  if (action === "sendAdminReply") return jsonResponse(sendAdminReply(data.email, data.replyText));
  if (action === "sendBulkPromo") return jsonResponse(sendBulkPromo(data.title, data.message, data.bannerUrl));
  if (action === "uploadImage") return jsonResponse(uploadImageWithFallback(data.base64, data.filename, data.mimeType));
  if (action === "readUniversal") return jsonResponse(getUniversalData(data.sheetName));

  // ENHANCED MANAGEMENT ACTIONS
  if (action === "deleteCategory") return jsonResponse(deleteCategory(data.categoryName));
  if (action === "deleteProduct") return jsonResponse(deleteProduct(data.productId));

  return jsonResponse({status: "error", message: "Invalid POST Action"});
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
