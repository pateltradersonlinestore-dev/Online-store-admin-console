// ==========================================
// 1. STORE FRONT & ORDERS
// ==========================================
function getProductsFromSheet() {
  const data = getUniversalData("Products");
  return data.map(obj => {
    if(obj.Image_URL) {
      let urls = String(obj.Image_URL).split(',');
      let processedUrls = urls.map(u => {
        let trimmed = u.trim();
        if(trimmed.includes("drive.google.com") && !trimmed.includes("uc?export=view")) {
          const fileId = trimmed.match(/[-\w]{25,}/);
          if(fileId) return `https://drive.google.com/uc?export=view&id=${fileId[0]}`;
        }
        return trimmed;
      });
      obj.Image_URL = processedUrls.join(',');
    }
    return obj;
  });
}

function getOrderHistory(email) {
  const data = getUniversalData("Orders");
  const searchEmail = String(email).toLowerCase().trim();
  return data.filter(row => row.Email && String(row.Email).toLowerCase().trim() === searchEmail);
}

function processOrder(orderData) {
  const db = getMasterDB();
  const orderId = "ORD-" + Date.now();
  const orderDate = new Date();
  
  let orderSheet = db.getSheetByName("Orders");
  let oHeaders = ensureDynamicColumns(orderSheet, orderData);
  let newRow = new Array(oHeaders.length).fill("");
  newRow[oHeaders.indexOf("Order_ID")] = orderId;
  newRow[oHeaders.indexOf("Email")] = orderData.email;
  newRow[oHeaders.indexOf("Items")] = orderData.items;
  newRow[oHeaders.indexOf("Amount")] = orderData.amount;
  newRow[oHeaders.indexOf("Razorpay_ID")] = orderData.razorpayId;
  newRow[oHeaders.indexOf("Status")] = "Pending";
  newRow[oHeaders.indexOf("Date")] = orderDate;
  
  Object.keys(orderData).forEach(k => {
    if(!["action", "token", "email", "items", "amount", "razorpayId"].includes(k)) {
      let i = oHeaders.indexOf(k); if(i>-1) newRow[i] = orderData[k];
    }
  });
  
  orderSheet.appendRow(newRow);
  syncCart(orderData.email, "[]");

  const usersSheet = db.getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(orderData.email, usersSheet);
  if(rowIndex > 0) {
    let uHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    let vaultId = usersSheet.getRange(rowIndex, uHeaders.indexOf("VaultID") + 1).getValue();
    if(vaultId) {
      try { SpreadsheetApp.openById(vaultId).getSheetByName("Profile_And_Orders").appendRow([orderId, orderData.items, orderData.amount, orderDate.toLocaleDateString(), "Pending"]); } catch(e) {} 
    }
  }
  SpreadsheetApp.flush();

  sendOrderReceivedEmail(orderData, orderId);
  return {status: "success", orderId: orderId};
}

// ==========================================
// 2. CART & INTERACTION
// ==========================================
function getCart(email) {
  let sheet = getMasterDB().getSheetByName("Carts");
  if (!sheet) return {status: "success", cart: "[]"};
  
  const data = sheet.getDataRange().getValues();
  const searchEmail = String(email).toLowerCase().trim();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).toLowerCase().trim() === searchEmail) {
      return {status: "success", cart: data[i][1] || "[]"};
    }
  }
  return {status: "success", cart: "[]"};
}

function syncCart(email, cartItemsJSON) {
  let sheet = getMasterDB().getSheetByName("Carts");
  const data = sheet.getDataRange().getValues();
  const searchEmail = String(email).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).toLowerCase().trim() === searchEmail) {
      sheet.getRange(i + 1, 2).setValue(cartItemsJSON);
      SpreadsheetApp.flush();
      return {status: "success"};
    }
  }
  sheet.appendRow([searchEmail, cartItemsJSON]);
  SpreadsheetApp.flush();
  return {status: "success"};
}

function logActivity(email, type, details) {
  if(type !== "LastSeen") return {status: "success"};
  const sheet = getMasterDB().getSheetByName("Users");
  const rowIndex = findRowIndexByEmailCaseInsensitive(email, sheet);
  if(rowIndex > 0) {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.getRange(rowIndex, headers.indexOf("LastSeen") + 1).setValue(details);
    SpreadsheetApp.flush();
  }
  return {status: "success"};
}

function addToWaitlist(email, productId) {
  universalWrite("Waitlist", [productId, email, new Date()]);
  return {status: "success"};
}

function addReview(reviewData) {
  universalWrite("Reviews", [reviewData.productId, reviewData.userName, reviewData.rating, reviewData.comment, new Date()]);
  return {status: "success"};
}

// ==========================================
// 3. ADMIN & MANAGEMENT OPERATIONS
// ==========================================
function deleteCategory(categoryName) {
  const sheet = getMasterDB().getSheetByName("Categories");
  if (!sheet) return {status: "error", message: "Categories sheet not found."};
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === String(categoryName).toLowerCase().trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return {status: "success", message: `Category '${categoryName}' deleted successfully.`};
    }
  }
  return {status: "error", message: "Category not found."};
}

function deleteProduct(productId) {
  const db = getMasterDB();
  const productSheet = db.getSheetByName("Products");
  if (!productSheet) return {status: "error", message: "Products sheet not found."};
  
  const data = productSheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf("ID");
  
  if (idIndex === -1) return {status: "error", message: "ID column not found in Products sheet."};
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(productId)) {
      productSheet.deleteRow(i + 1);
      
      // Permanently remove waitlist entries referencing this product
      const waitlistSheet = db.getSheetByName("Waitlist");
      if (waitlistSheet) {
        const wlData = waitlistSheet.getDataRange().getValues();
        for (let j = wlData.length - 1; j > 0; j--) {
          if (String(wlData[j][0]) === String(productId)) {
            waitlistSheet.deleteRow(j + 1);
          }
        }
      }
      
      // Optional: Clean up related reviews here as well, to prevent orphaned data
      const reviewSheet = db.getSheetByName("Reviews");
      if (reviewSheet) {
        const rvData = reviewSheet.getDataRange().getValues();
        for (let k = rvData.length - 1; k > 0; k--) {
          if (String(rvData[k][0]) === String(productId)) {
            reviewSheet.deleteRow(k + 1);
          }
        }
      }
      
      SpreadsheetApp.flush();
      return {status: "success", message: `Product '${productId}' and all related records deleted.`};
    }
  }
  return {status: "error", message: "Product not found."};
}

function updateOrderStatusAndAlert(orderId, newStatus) {
  const sheet = getMasterDB().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf("Order_ID")] === orderId) {
      sheet.getRange(i + 1, headers.indexOf("Status") + 1).setValue(newStatus);
      SpreadsheetApp.flush(); 
      if (newStatus === "Delivered") {
        const email = data[i][headers.indexOf("Email")];
        sendOrderDeliveredEmail(email, orderId);
      }
      return {status: "success", message: `Order updated to ${newStatus}`};
    }
  }
  return {status: "error", message: "Order not found"};
}

function sendBulkPromo(title, message, bannerUrl) {
  const data = getMasterDB().getSheetByName("Users").getDataRange().getValues();
  let emailList = [];
  for(let i = 1; i < data.length; i++) { if(data[i][1] && data[i][1].includes("@")) emailList.push(data[i][1]); }
  if(emailList.length === 0) return {status: "error", message: "No recipient list found."};
  
  const emailsBCC = emailList.join(",");
  sendBulkPromoEmail(emailsBCC, title, message, bannerUrl);
  return {status: "success", message: `Promotional email sent to ${emailList.length} customers!`};
}
