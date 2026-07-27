// ==========================================
// 1. MODERN EMAIL TEMPLATE ENGINE
// ==========================================
const EMAIL_STYLES = `
  body { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #f3f4f6; padding: 20px; color: #1f2937; margin: 0; }
  .container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border: 1px solid #e5e7eb; }
  .header { text-align: center; padding: 32px 20px; background-color: #ffffff; border-bottom: 1px solid #e5e7eb; }
  .header img { border-radius: 12px; max-width: 80px; }
  .content { padding: 32px 24px; line-height: 1.6; }
  .footer { text-align: center; padding: 24px; font-size: 13px; color: #6b7280; background-color: #f9fafb; border-top: 1px solid #e5e7eb; }
  .btn { display: inline-block; padding: 12px 24px; background-color: #0f172a; color: #ffffff; text-decoration: none; border-radius: 6px; font-weight: 500; margin-top: 16px; }
  .highlight-box { background: #f8fafc; border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; margin: 20px 0; }
  h2 { color: #111827; font-size: 22px; font-weight: 600; margin-top: 0; margin-bottom: 16px; }
  p { margin: 0 0 16px 0; }
`;

function buildEmailHtml(contentHtml) {
  return `
    <html>
      <head>
        <style>${EMAIL_STYLES}</style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <img src="${LOGO_URL}" alt="Store Logo">
          </div>
          <div class="content">
            ${contentHtml}
          </div>
          <div class="footer">
            &copy; ${new Date().getFullYear()} Our Online Store. All rights reserved.
          </div>
        </div>
      </body>
    </html>
  `;
}

// ==========================================
// 2. TRANSACTIONAL EMAILS
// ==========================================
function sendOrderReceivedEmail(orderData, orderId) {
  const content = `
    <h2>Order Received! 🎉</h2>
    <p>Hi there, we have successfully received your order and are preparing it for shipment.</p>
    <div class="highlight-box">
      <p style="margin-bottom: 8px;"><strong>Order ID:</strong> ${orderId}</p>
      <p style="margin-bottom: 8px;"><strong>Total Amount:</strong> ₹${orderData.amount}</p>
      <p style="margin-bottom: 8px;"><strong>Payment ID:</strong> ${orderData.razorpayId}</p>
      <p style="margin-bottom: 0;"><strong>Delivery Address:</strong> ${orderData.Delivery_Address || 'Default Profile Address'}</p>
    </div>
    <p>Thank you for shopping with us!</p>
  `;
  MailApp.sendEmail({ to: orderData.email, subject: `Order Confirmation #${orderId}`, htmlBody: buildEmailHtml(content) });
}

function sendOrderDeliveredEmail(email, orderId) {
  const content = `
    <h2>Order Delivered! 🚚</h2>
    <p>Great news! Your order <strong>#${orderId}</strong> has been successfully delivered to your address.</p>
    <p>We hope you enjoy your purchase. If you have any questions or concerns, please don't hesitate to reach out to our support team.</p>
  `;
  MailApp.sendEmail({ to: email, subject: `Your order has been delivered!`, htmlBody: buildEmailHtml(content) });
}

function sendBulkPromoEmail(emailsBCC, title, message, bannerUrl) {
  const bannerHtml = bannerUrl ? `<img src="${bannerUrl}" style="width:100%; border-radius:8px; margin-bottom:24px;">` : "";
  const content = `
    ${bannerHtml}
    <h2>${title}</h2>
    <p style="font-size: 16px; color: #4b5563;">${message}</p>
    <div style="text-align: center; margin-top: 32px;">
      <a href="https://example.com" class="btn">Shop the Collection</a>
    </div>
  `;
  MailApp.sendEmail({ to: "no-reply@store.com", bcc: emailsBCC, subject: `Special Update: ${title}`, htmlBody: buildEmailHtml(content) });
}

function sendAdminReply(email, replyText) {
  const content = `
    <h2>Support Update</h2>
    <p>${replyText}</p>
    <br/>
    <p>Best regards,<br>Customer Support Team</p>
  `;
  MailApp.sendEmail({ to: email, subject: "Support Update regarding your inquiry", htmlBody: buildEmailHtml(content) });
  return {status: "success"};
}

// ==========================================
// 3. AUTHENTICATION & SECURITY EMAILS
// ==========================================
function sendOTPEmail(email, otp, type) {
  const content = `
    <h2>Your ${type} Verification Code</h2>
    <p>Please use the following secure code to proceed with your request:</p>
    <div style="text-align: center; margin: 32px 0;">
      <span style="font-size: 32px; font-weight: 700; letter-spacing: 6px; color: #0f172a; background: #f1f5f9; padding: 16px 24px; border-radius: 8px;">${otp}</span>
    </div>
    <p style="color: #6b7280; font-size: 14px;">This code is valid for 10 minutes. Please do not share it with anyone.</p>
  `;
  MailApp.sendEmail({ to: email, subject: `Your ${type} Code`, htmlBody: buildEmailHtml(content) });
}

function sendRegistrationPDF(email, name, phone, userId) {
  const maskedPhone = phone && phone.length > 4 ? "XXXXXX" + phone.slice(-4) : "XXXX";
  const html = `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 40px; text-align: center; color: #1f2937;">
      <h1 style="color: #0f172a; margin-bottom: 8px;">Premium Account Registration</h1>
      <h2 style="color: #6b7280; font-weight: normal; margin-top: 0;">Registration Successfully Completed</h2>
      <hr style="border:0; border-top: 1px solid #e5e7eb; margin: 32px 0;">
      <div style="text-align: left; max-width: 350px; margin: 0 auto; font-size: 16px; line-height: 2;">
        <p style="margin:0;"><strong>Account Name:</strong> ${name}</p>
        <p style="margin:0;"><strong>Registered Email:</strong> ${email}</p>
        <p style="margin:0;"><strong>Linked Phone:</strong> ${maskedPhone}</p>
        <p style="margin:0;"><strong>Secure User ID:</strong> ${userId}</p>
      </div>
      <br>
      <p style="font-size:12px; color:#9ca3af; margin-top: 48px;">Please securely save this document for your records.</p>
    </div>
  `;
  const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName("Account_Details.pdf");
  MailApp.sendEmail({ 
    to: email, 
    subject: "Welcome! Your Account Registration is Complete", 
    body: "Welcome to our store. Please find your official account registration details securely attached to this email.", 
    attachments: [blob] 
  });
}

function sendDeviceAlert(email, device, visits) {
  const content = `
    <h2 style="color: #dc2626;">Security Alert: New Login Detected</h2>
    <p>A new login was just recorded on your account.</p>
    <div class="highlight-box" style="border-color: #fecaca; background: #fef2f2;">
      <ul style="margin: 0; padding-left: 20px; line-height: 1.8;">
        <li><strong>Device Detected:</strong> ${device || "Browser"}</li>
        <li><strong>Total Account Logins:</strong> ${visits}</li>
        <li><strong>Timestamp:</strong> ${new Date().toLocaleString()}</li>
      </ul>
    </div>
    <p style="font-size: 14px; color: #6b7280;">If this login was authorized by you, feel free to ignore this alert. If you do not recognize this activity, please contact support immediately to secure your account.</p>
  `;
  MailApp.sendEmail({ to: email, subject: "Security Alert - New Login Detected", htmlBody: buildEmailHtml(content) });
}
