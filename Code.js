/** 📤 MAIN ORDER PROCESSOR - ENHANCED & CSV-MATCHED **/
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = e.values;
  const rowIndex = sheet.getLastRow();
  const receiptCol = 14; // Column N (adjusted for 13 columns + timestamp)
  const timestamp = data[0];

  const status = sheet.getRange(rowIndex, receiptCol).getValue();
  if (status && status.toString().toLowerCase().includes("✓")) {
    Logger.log("⏩ Skipping duplicate row: " + timestamp);
    return;
  }

  // CORRECTED: Match actual CSV structure
  const [, name, email, phone, location, room, orderType, orderField1, orderField2, store, paymentMethod, gcashRef, termsAccepted] = data;

  // Smart order field selection - prioritize filled fields
  const rawOrder = determineOrderField(orderField1, orderField2, orderType);

  if (!termsAccepted || !termsAccepted.toLowerCase().includes("agree") || !email.includes('@')) {
    sendError(email, "Oops! Your submission seems incomplete. Please ensure you enter a valid email address and check the 'I agree' box. For help, visit: https://m.me/639741992565449");
    return;
  }

  if (!rawOrder) {
    sendError(email, "No order information provided. Please fill in either 'What will you like to Order' or 'Tell us what you need picked up' field.");
    return;
  }

  // Enhanced order parsing with dual format support
  const parseResult = parseOrderItemsEnhanced(rawOrder);
  
  if (!parseResult.success) {
    sendFormatCorrectionEmail(email, name, parseResult.error, rawOrder);
    return;
  }

  const items = parseResult.items;
  const detectedOrderType = parseResult.type; // 'itemized' or 'prepaid'
  
  const subtotal = items.reduce((sum, item) => sum + (item.quantity * (item.price || 0)), 0);
  const deliveryFee = calculateDeliveryFee(store); // Always calculate service fee regardless of order type
  const total = subtotal + deliveryFee;

  const orderDetails = {
    timestamp, name, email, phone, location, room,
    items, subtotal, deliveryFee, total,
    store, paymentMethod, gcashRef, 
    orderType: detectedOrderType, // Use detected type from parsing
    selectedOrderType: orderType, // Store user's selection for reference
    orderSource: rawOrder === orderField1 ? 'field1' : 'field2' // Track which field was used
  };

  // Try to generate PDF receipt with fallback
  const receiptResult = generateReceiptPDFWithFallback(orderDetails);
  
  // Backup receipt if PDF was generated successfully
  if (receiptResult.pdf) {
    backupReceiptToMonth(receiptResult.pdf, orderDetails);
  }
  
  const success = sendEmailConfirmation(email, { 
    ...orderDetails, 
    receiptPDF: receiptResult.pdf, 
    fallback: receiptResult.fallback 
  });

  if (success) {
    sheet.getRange(rowIndex, receiptCol).setValue("✓ " + new Date().toLocaleString());
    
    // Enhanced admin notification with instant ping
    sendInstantOrderPing(orderDetails);
  } else {
    Logger.log("⚠️ Email failed to send for: " + email);
  }
}


/** 🧠 SMART ORDER FIELD DETERMINATION **/
function determineOrderField(field1, field2, selectedType) {
  // Priority logic for order field selection
  
  // If both fields are empty
  if (!field1 && !field2) {
    return null;
  }
  
  // If only one field is filled, use it
  if (field1 && !field2) {
    Logger.log("📝 Using order field 1: " + field1);
    return field1;
  }
  
  if (field2 && !field1) {
    Logger.log("📝 Using order field 2: " + field2);
    return field2;
  }
  
  // If both fields are filled, make smart decision
  if (field1 && field2) {
    Logger.log("📝 Both order fields filled - making smart selection");
    
    // Use selected order type as hint if available
    if (selectedType) {
      const typeHint = selectedType.toLowerCase();
      
      // If pickup/delivery mentioned, prefer field2
      if (typeHint.includes('pickup') || typeHint.includes('delivery')) {
        Logger.log("📝 Order type suggests pickup/delivery - using field 2");
        return field2;
      }
      
      // If ordering mentioned, prefer field1
      if (typeHint.includes('order')) {
        Logger.log("📝 Order type suggests ordering - using field 1");
        return field1;
      }
    }
    
    // Fallback: Use the longer/more detailed field
    if (field1.length > field2.length) {
      Logger.log("📝 Field 1 is longer - using field 1");
      return field1;
    } else {
      Logger.log("📝 Field 2 is longer or equal - using field 2");
      return field2;
    }
  }
  
  return field1 || field2; // Final fallback
}


/** 🧠 ENHANCED ORDER PARSING - DUAL FORMAT SUPPORT **/
function parseOrderItemsEnhanced(text) {
  if (!text || typeof text !== 'string') {
    return {
      success: false,
      error: "empty_order",
      items: [],
      type: null
    };
  }

  const cleanText = text.trim();
  
  // Pattern 1: Itemized with prices - "Burger 2 @80" or "Burger 2 @ 80"
  const itemizedPattern = /([\w\s]+?)\s+(\d+)\s*@\s*(\d+(?:\.\d{2})?)/gi;
  
  // Pattern 2: Prepaid/Pickup without prices - "Burger 2" or "Fries 1"
  const prepaidPattern = /([\w\s]+?)\s+(\d+)(?!\s*@)/gi;
  
  // Try itemized format first
  const itemizedMatches = [...cleanText.matchAll(itemizedPattern)];
  
  if (itemizedMatches.length > 0) {
    const items = itemizedMatches.map(match => ({
      name: match[1].trim(),
      quantity: parseInt(match[2]),
      price: parseFloat(match[3])
    }));
    
    // Validate parsed items
    const validItems = items.filter(item => 
      item.name && item.quantity > 0 && item.price >= 0
    );
    
    if (validItems.length === items.length) {
      return {
        success: true,
        items: validItems,
        type: 'itemized',
        error: null
      };
    }
  }
  
  // Try prepaid format
  const prepaidMatches = [...cleanText.matchAll(prepaidPattern)];
  
  if (prepaidMatches.length > 0) {
    const items = prepaidMatches.map(match => ({
      name: match[1].trim(),
      quantity: parseInt(match[2]),
      price: 0 // No price for prepaid orders
    }));
    
    // Validate parsed items
    const validItems = items.filter(item => 
      item.name && item.quantity > 0
    );
    
    if (validItems.length === items.length) {
      return {
        success: true,
        items: validItems,
        type: 'prepaid',
        error: null
      };
    }
  }
  
  // Neither format worked - determine specific error
  if (cleanText.includes('@')) {
    return {
      success: false,
      error: "invalid_itemized_format",
      items: [],
      type: null
    };
  } else if (/\d/.test(cleanText)) {
    return {
      success: false,
      error: "invalid_prepaid_format", 
      items: [],
      type: null
    };
  } else {
    return {
      success: false,
      error: "unrecognizable_format",
      items: [],
      type: null
    };
  }
}


/** 📧 FORMAT CORRECTION EMAIL - NO EMOJIS FOR CUSTOMERS **/
function sendFormatCorrectionEmail(email, name, errorType, originalOrder) {
  try {
    if (!email) {
      Logger.log("⚠️ Cannot send format correction email: Email is empty.");
      return;
    }

    let subject = "Campus Cart - Order Format Help";
    let helpMessage = "";
    let examples = "";

    switch (errorType) {
      case "empty_order":
        helpMessage = "We didn't receive any order items from your submission.";
        examples = `
          <p><strong>Please list your items using one of these formats:</strong></p>
          <ul>
            <li><strong>With Prices:</strong> Burger 2 @80 (for 2 burgers at ₱80 each)</li>
            <li><strong>Prepaid/Pickup:</strong> Burger 2 (for 2 burgers, price already arranged)</li>
          </ul>
        `;
        break;
        
      case "invalid_itemized_format":
        helpMessage = "We detected you're trying to include prices, but the format isn't quite right.";
        examples = `
          <p><strong>Correct Format with Prices:</strong></p>
          <ul>
            <li>Burger 2 @80</li>
            <li>Fries 1 @45</li>
            <li>Coke 3 @25</li>
          </ul>
          <p><em>Format: ItemName Quantity @Price</em></p>
        `;
        break;
        
      case "invalid_prepaid_format":
        helpMessage = "We detected quantities but couldn't understand the item names clearly.";
        examples = `
          <p><strong>Correct Format for Prepaid/Pickup Orders:</strong></p>
          <ul>
            <li>Burger 2</li>
            <li>Fries 1</li>
            <li>Coke 3</li>
          </ul>
          <p><em>Format: ItemName Quantity</em></p>
        `;
        break;
        
      default:
        helpMessage = "We couldn't understand your order format.";
        examples = `
          <p><strong>Please use one of these formats:</strong></p>
          <ul>
            <li><strong>With Prices:</strong> Burger 2 @80, Fries 1 @45</li>
            <li><strong>Prepaid/Pickup:</strong> Burger 2, Fries 1</li>
          </ul>
        `;
    }

    const htmlBody = `
      <p>Hi ${name || 'there'}!</p>
      
      <p>${helpMessage}</p>
      
      ${examples}
      
      <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #dc3545; margin: 20px 0;">
        <p><strong>Your Original Order:</strong></p>
        <p><em>"${originalOrder || 'No order text received'}"</em></p>
      </div>
      
      <p><strong>Tips for Success:</strong></p>
      <ul>
        <li>Make sure to fill in either the "What will you like to Order" field OR the "Tell us what you need picked up" field</li>
        <li>List each item on a separate line or separate with commas</li>
        <li>Use numbers for quantities (not words like "two")</li>
        <li>Double-check spelling of item names</li>
        <li>If including prices, use @ symbol before the price</li>
      </ul>
      
      <p><strong>Need More Help?</strong></p>
      <p>Message us directly: <a href="https://m.me/639741992565449">Customer Support</a></p>
      
      <p>Please correct your order format and submit again. We're here to help!</p>
      
      <p>Best regards,<br>Campus Cart AUP Team</p>
    `;

    GmailApp.sendEmail(email, subject, '', { htmlBody });
    Logger.log(`📧 Format correction email sent to: ${email} (Error: ${errorType})`);
    
  } catch (err) {
    Logger.log("❌ Error sending format correction email: " + err.message);
  }
}


/** 🔔 INSTANT ORDER PING NOTIFICATION - KEEP EMOJIS FOR ADMIN **/
function sendInstantOrderPing(orderDetails) {
  try {
    const { name, email, phone, location, room, items, total, store, paymentMethod, orderType, selectedOrderType, orderSource } = orderDetails;
    
    // Generate delivery code for tracking
    const deliveryCode = `CC-${Math.random().toString(36).substring(2, 10).toUpperCase()}-${Math.random().toString(36).substring(2, 4).toUpperCase()}`;
    
    // Format items for quick reading
    const itemSummary = items.map(item => {
      if (orderType === 'itemized') {
        return `${item.name} x${item.quantity} @₱${item.price}`;
      } else {
        return `${item.name} x${item.quantity} (prepaid)`;
      }
    }).join(', ');
    
    // Determine urgency based on order value
    const urgencyEmoji = total >= 500 ? "🔥" : total >= 200 ? "⚡" : "📦";
    
    const subject = `${urgencyEmoji} NEW ORDER ALERT - ${name} (₱${total})`;
    
    const message = `
🆕 NEW CAMPUS CART ORDER RECEIVED!

👤 CUSTOMER: ${name}
📱 PHONE: ${phone}
📧 EMAIL: ${email}
📍 LOCATION: ${location}, ${room}

🛍️ ITEMS: ${itemSummary}
🏪 STORE: ${store}
💳 PAYMENT: ${paymentMethod}
💰 TOTAL: ₱${total} (${orderType === 'prepaid' ? 'PREPAID ORDER' : 'COLLECT ON DELIVERY'})

📋 ORDER TYPE SELECTED: ${selectedOrderType || 'Not specified'}
📝 ORDER SOURCE: ${orderSource === 'field1' ? 'Order Field' : 'Pickup Field'}

🎫 DELIVERY CODE: ${deliveryCode}
⏰ ORDER TIME: ${new Date().toLocaleString()}

${urgencyEmoji === "🔥" ? "🔥 HIGH VALUE ORDER - PRIORITY PROCESSING!" : ""}

Ready for dispatch processing!
    `;

    // Send to both admin emails for redundancy
    GmailApp.sendEmail("campuscart59@gmail.com", subject, message);
    GmailApp.sendEmail("gwanmesiamalcomp@gmail.com", `[BACKUP] ${subject}`, message);
    
    Logger.log(`🔔 Instant order ping sent for: ${name} (₱${total})`);
    
  } catch (error) {
    Logger.log("❌ Error sending instant order ping: " + error);
  }
}


/** 📧 ENHANCED EMAIL CONFIRMATION - NO EMOJIS FOR CUSTOMERS **/
function sendEmailConfirmation(email, data) {
  try {
    const { name, location, room, items, subtotal, deliveryFee, total, receiptPDF, paymentMethod, fallback, orderType, selectedOrderType } = data;
    const attachments = receiptPDF ? [receiptPDF] : [];

    // Generate delivery code
    const deliveryCode = `CC-${Math.random().toString(36).substring(2, 10).toUpperCase()}-${Math.random().toString(36).substring(2, 4).toUpperCase()}`;

    // Format items based on order type
    const itemLines = items.map(item => {
      if (orderType === 'itemized') {
        return `- ${item.name} x${item.quantity} @ ₱${item.price} = ₱${item.quantity * item.price}`;
      } else {
        return `- ${item.name} x${item.quantity} (prepaid)`;
      }
    }).join("<br>");

    const fallbackNote = fallback
      ? `<p><strong>Note:</strong> We were unable to generate your official receipt PDF due to a system error. This email serves as your confirmation.</p>`
      : "";

    // Order type specific messaging - Updated for service fee
    const orderTypeMessage = orderType === 'prepaid' 
      ? `<p><strong>Payment Status:</strong> Items are prepaid. Please prepare ₱${deliveryFee} for the service fee upon delivery.</p>`
      : `<p><strong>Payment:</strong> Please prepare the exact amount of ₱${total} for collection upon delivery.</p>`;

    // Show selected order type if available
    const orderTypeInfo = selectedOrderType 
      ? `<p><strong>Order Type Selected:</strong> ${selectedOrderType}</p>`
      : "";

    const htmlBody = `
      <p>Dear ${name},</p>
      <p>Thank you for your order! Your order has been received and is being processed.</p>
      
      <div style="background-color: #e8f5e8; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <p><strong>Your Delivery Code: ${deliveryCode}</strong></p>
        <p><em>Please keep this code handy for delivery verification.</em></p>
      </div>
      
      ${orderTypeInfo}
      
      <p><strong>Order Details:</strong></p>
      <ul>
        <li><strong>Items:</strong><br>${itemLines}</li>
        ${orderType === 'itemized' ? `<li><strong>Items Total:</strong> ₱${subtotal}</li>` : `<li><strong>Items:</strong> PREPAID (already settled)</li>`}
        <li><strong>Service Fee:</strong> ₱${deliveryFee}</li>
        ${orderType === 'itemized' ? `<li><strong>Grand Total:</strong> ₱${total}</li>` : `<li><strong>Service Fee to Collect:</strong> ₱${deliveryFee}</li>`}
      </ul>
      
      ${orderTypeMessage}
      
      <p><strong>Delivery Address:</strong><br>${location} - ${room}</p>
      
      ${fallbackNote}
      
      <div style="background-color: #f0f8ff; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <p><strong>Help Us Improve!</strong></p>
        <p>After you receive your order, please take a moment to share your feedback:</p>
        <p><a href="https://forms.gle/GM2TF7asYSC3aPKW7" style="background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Leave Feedback</a></p>
      </div>
      
      <p><strong>Need Support?</strong></p>
      <p>Message us at: <a href="https://m.me/639741992565449">Customer Support</a></p>
      
      <p>Best regards,<br>Campus Cart AUP Team</p>
    `;

    GmailApp.sendEmail(email, `Campus Cart Order Confirmation - ${deliveryCode}`, '', {
      htmlBody,
      attachments
    });
    return true;
  } catch (err) {
    Logger.log("❌ Error sending confirmation email: " + err.message);
    return false;
  }
}


/** 🚚 DELIVERY FEE CALCULATOR **/
function calculateDeliveryFee(store) {
  const near = [
    "Store Near AUP", "Dali Putting Kahoy", "Local Sari Sari Store", "Dante's",
    "Canteen", "AUP Cafeteria", "Loading Snacks", "Kusina ni Ate",
    "Chooks-to-Go (Puting Kahoy)", "Burgers & Fries Stalls (near trike terminal)"
  ];
  const far = [
    "Santa Rosa Heights", "Wet Market", "Dali Santa Rosa", "Safe More",
    "Andok's Santa Rosa", "Mini Stop", "7-Eleven", "7 Eleven"
  ];
  const veryFar = [
    "Paseo", "Nuvali", "Fresh Market", "South Mall", "Robinson",
    "All Home", "SM City Sta. Rosa"
  ];

  if (!store || typeof store !== 'string') return 0;

  const normalized = store.toLowerCase();
  if (veryFar.some(zone => normalized.includes(zone.toLowerCase()))) return 199;
  if (far.some(zone => normalized.includes(zone.toLowerCase()))) return 99;
  if (near.some(zone => normalized.includes(zone.toLowerCase()))) return 69;
  return 0;
}


/** 📟 PDF RECEIPT GENERATOR WITH FALLBACK **/
function generateReceiptPDFWithFallback(orderDetails = {}) {
  try {
    const pdf = generateReceiptPDF(orderDetails);
    return { pdf: pdf, fallback: false };
  } catch (error) {
    Logger.log("❌ Error generating receipt PDF: " + error);
    // Notify admin about PDF generation failure
    GmailApp.sendEmail("gwanmesiamalcomp@gmail.com", "Campus Cart - PDF Generation Error", 
      `PDF generation failed for ${orderDetails.name} (${orderDetails.email}).\n\nError: ${error}`);
    return { pdf: null, fallback: true };
  }
}


/** 📟 ENHANCED PDF RECEIPT GENERATOR - KEEP EMOJIS **/
function generateReceiptPDF(orderDetails = {}) {
  const {
    timestamp = "Unknown", name = "Customer", email = "", phone = "", location = "", room = "",
    items = [], subtotal = 0, deliveryFee = 0, total = 0,
    store = "", paymentMethod = "", gcashRef = "", orderType = "itemized",
    selectedOrderType = ""
  } = orderDetails;

  const doc = DocumentApp.create(`CampusCart_Receipt_${name}_${new Date().getTime()}`);
  const body = doc.getBody();

  // Generate delivery code
  const deliveryCode = `CC-${Math.random().toString(36).substring(2, 10).toUpperCase()}-${Math.random().toString(36).substring(2, 4).toUpperCase()}`;

  // Header
  body.appendParagraph("🎓 Campus Cart Receipt").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`📅 Order Date: ${new Date(timestamp).toLocaleDateString()} ${new Date(timestamp).toLocaleTimeString()}`);
  body.appendParagraph(`🎫 Delivery Code: ${deliveryCode}`);
  body.appendParagraph(`📦 Order Type: ${orderType === 'prepaid' ? 'PREPAID ORDER' : 'ITEMIZED ORDER'}`);
  if (selectedOrderType) {
    body.appendParagraph(`📋 Selected Type: ${selectedOrderType}`);
  }
  body.appendParagraph(""); // Empty line

  // Customer Details
  body.appendParagraph("👤 CUSTOMER DETAILS").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(`• Customer Name: ${name}`);
  body.appendParagraph(`• Phone Number: ${phone}`);
  body.appendParagraph(`• Email Address: ${email}`);
  body.appendParagraph(`• Delivery Location: ${location}, ${room}`);
  body.appendParagraph(`• Store: ${store}`);
  body.appendParagraph(`• Payment Method: ${paymentMethod}${gcashRef ? ` (Ref: ${gcashRef})` : ""}`);
  body.appendParagraph(""); // Empty line

  // Ordered Items
  body.appendParagraph("🛒 ORDERED ITEMS").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  items.forEach(item => {
    if (orderType === 'itemized') {
      const itemTotal = (item.quantity || 0) * (item.price || 0);
      body.appendParagraph(`• ${item.name || "Unnamed Item"} x${item.quantity || 0} @ ₱${(item.price || 0).toFixed(2)} = ₱${itemTotal.toFixed(2)}`);
    } else {
      body.appendParagraph(`• ${item.name || "Unnamed Item"} x${item.quantity || 0} (prepaid)`);
    }
  });
  body.appendParagraph(""); // Empty line

  // Financial Summary - Show for both order types but with different messaging
  body.appendParagraph("💰 FINANCIAL SUMMARY").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  if (orderType === 'itemized') {
    body.appendParagraph(`• Items Total: ₱${subtotal.toFixed(2)}`);
    body.appendParagraph(`• Service Fee: ₱${deliveryFee.toFixed(2)}`);
    body.appendParagraph(`• GRAND TOTAL: ₱${total.toFixed(2)}`);
  } else {
    body.appendParagraph(`• Items: PREPAID (already settled)`);
    body.appendParagraph(`• Service Fee: ₱${deliveryFee.toFixed(2)}`);
    body.appendParagraph(`• TOTAL SERVICE FEE TO COLLECT: ₱${deliveryFee.toFixed(2)}`);
  }
  body.appendParagraph(""); // Empty line

  // Footer with feedback link
  body.appendParagraph("✅ Terms: Customer agreed to service conditions");
  body.appendParagraph("📞 Support: https://m.me/639741992565449");
  body.appendParagraph("📋 Feedback: https://forms.gle/GM2TF7asYSC3aPKW7");
  body.appendParagraph("💙 Thank you for choosing Campus Cart!");

  doc.saveAndClose();
  return DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
}


/** 📁 BACKUP SYSTEM - CREATE FOLDER STRUCTURE **/
function createBackupFolders() {
  try {
    // Check if main backup folder exists
    let mainFolder;
    const existingFolders = DriveApp.getFoldersByName("Campus Cart Backups");
    
    if (existingFolders.hasNext()) {
      mainFolder = existingFolders.next();
      Logger.log("📁 Main backup folder found");
    } else {
      mainFolder = DriveApp.createFolder("Campus Cart Backups");
      Logger.log("📁 Main backup folder created");
    }

    // Create receipts subfolder
    let receiptsFolder;
    const existingReceiptsFolder = mainFolder.getFoldersByName("Receipts");
    if (existingReceiptsFolder.hasNext()) {
      receiptsFolder = existingReceiptsFolder.next();
    } else {
      receiptsFolder = mainFolder.createFolder("Receipts");
      Logger.log("📁 Receipts folder created");
    }

    // Create dispatch summaries subfolder  
    let dispatchFolder;
    const existingDispatchFolder = mainFolder.getFoldersByName("Dispatch Summaries");
    if (existingDispatchFolder.hasNext()) {
      dispatchFolder = existingDispatchFolder.next();
    } else {
      dispatchFolder = mainFolder.createFolder("Dispatch Summaries");
      Logger.log("📁 Dispatch Summaries folder created");
    }

    return {
      main: mainFolder,
      receipts: receiptsFolder,
      dispatch: dispatchFolder
    };
  } catch (error) {
    Logger.log("❌ Error creating backup folders: " + error);
    return null;
  }
}


/** 🗂️ ORGANIZE BACKUPS BY MONTH **/
function organizeBackupsByMonth() {
  try {
    const folders = createBackupFolders();
    if (!folders) return;

    const currentDate = new Date();
    const monthYear = currentDate.toLocaleDateString('en-US', { 
      year: 'numeric', 
      month: 'long' 
    });

    // Create month folders in both receipts and dispatch folders
    const monthFolders = {};
    
    // Receipts month folder
    const existingReceiptsMonth = folders.receipts.getFoldersByName(monthYear);
    if (existingReceiptsMonth.hasNext()) {
      monthFolders.receipts = existingReceiptsMonth.next();
    } else {
      monthFolders.receipts = folders.receipts.createFolder(monthYear);
    }

    // Dispatch month folder  
    const existingDispatchMonth = folders.dispatch.getFoldersByName(monthYear);
    if (existingDispatchMonth.hasNext()) {
      monthFolders.dispatch = existingDispatchMonth.next();
    } else {
      monthFolders.dispatch = folders.dispatch.createFolder(monthYear);
    }

    return monthFolders;
    
  } catch (error) {
    Logger.log("❌ Error organizing monthly folders: " + error);
    return null;
  }
}


/** 💾 ENHANCED BACKUP RECEIPT WITH MONTHLY ORGANIZATION **/
function backupReceiptToMonth(receiptPDF, orderDetails) {
  try {
    const monthFolders = organizeBackupsByMonth();
    if (!monthFolders) return;

    const date = new Date(orderDetails.timestamp);
    const dateStr = date.toLocaleDateString('en-US').replace(/\//g, '-');
    const timeStr = date.toLocaleTimeString('en-US', { 
      hour: '2-digit', 
      minute: '2-digit' 
    }).replace(/:/g, '');
    
    const backupName = `Receipt_${orderDetails.name}_${dateStr}_${timeStr}`;
    
    // Create backup in monthly receipts folder
    const tempDoc = DocumentApp.create(backupName);
    const tempFile = DriveApp.getFileById(tempDoc.getId());
    const backupFile = tempFile.makeCopy(backupName, monthFolders.receipts);
    
    // Clean up temp file
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
    
    Logger.log(`💾 Receipt backed up to monthly folder: ${backupName}`);
    
  } catch (error) {
    Logger.log("❌ Error backing up receipt to monthly folder: " + error);
  }
}


/** 💾 ENHANCED BACKUP DISPATCH WITH MONTHLY ORGANIZATION **/
function backupDispatchToMonth(dispatchPDF, dateStr) {
  try {
    const monthFolders = organizeBackupsByMonth();
    if (!monthFolders) return;

    const backupName = `Dispatch_${dateStr.replace(/\s/g, '_').replace(/,/g, '')}_${Date.now()}`;
    
    // Create backup in monthly dispatch folder
    const tempDoc = DocumentApp.create(backupName);
    const tempFile = DriveApp.getFileById(tempDoc.getId());
    const backupFile = tempFile.makeCopy(backupName, monthFolders.dispatch);
    
    // Clean up temp file
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
    
    Logger.log(`💾 Dispatch summary backed up to monthly folder: ${backupName}`);
    
  } catch (error) {
    Logger.log("❌ Error backing up dispatch to monthly folder: " + error);
  }
}


/** 📊 DISPATCH LIST PDF GENERATOR WITH FALLBACK **/
function generateDispatchPDFWithFallback(summaryData, dateStr) {
  try {
    const pdf = generateDispatchPDF(summaryData, dateStr);
    return { pdf: pdf, fallback: false };
  } catch (error) {
    Logger.log("❌ Error generating dispatch PDF: " + error);
    GmailApp.sendEmail("gwanmesiamalcomp@gmail.com", "Campus Cart - Dispatch PDF Error", 
      `Dispatch PDF generation failed for ${dateStr}.\n\nError: ${error}`);
    return { pdf: null, fallback: true };
  }
}


/** 📋 ENHANCED DISPATCH LIST PDF GENERATOR - KEEP EMOJIS **/
function generateDispatchPDF(summaryData, dateStr) {
  const { totalOrders, totalRevenue, totalDelivery, orders, storeGroups } = summaryData;
  
  const doc = DocumentApp.create(`CampusCart_DispatchList_${dateStr}_${new Date().getTime()}`);
  const body = doc.getBody();

  // Header
  body.appendParagraph("🚛 Campus Cart Dispatch List").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`📅 Date: ${dateStr} | ⏰ Generated: ${new Date().toLocaleString()}`);
  
  // Summary
  body.appendParagraph("📊 DISPATCH SUMMARY").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  const summaryTable = body.appendTable();
  const summaryHeaderRow = summaryTable.appendTableRow();
  summaryHeaderRow.appendTableCell("Total Orders");
  summaryHeaderRow.appendTableCell("Total Customers");
  summaryHeaderRow.appendTableCell("Total Revenue");
  summaryHeaderRow.appendTableCell("Total Collection");
  
  const summaryDataRow = summaryTable.appendTableRow();
  summaryDataRow.appendTableCell(totalOrders.toString());
  summaryDataRow.appendTableCell(totalOrders.toString());
  summaryDataRow.appendTableCell(`₱${totalRevenue.toFixed(2)}`);
  summaryDataRow.appendTableCell(`₱${(totalRevenue + totalDelivery).toFixed(2)}`);
  
  body.appendParagraph(""); // Empty line

  if (totalOrders === 0) {
    body.appendParagraph("📭 No orders to dispatch today.");
    doc.saveAndClose();
    return DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  }

  // Orders table
  body.appendParagraph("👥 ORDERS BY CUSTOMER").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  const ordersTable = body.appendTable();
  const orderHeaderRow = ordersTable.appendTableRow();
  orderHeaderRow.appendTableCell("Customer Name");
  orderHeaderRow.appendTableCell("Items Ordered");
  orderHeaderRow.appendTableCell("Delivery Location");
  orderHeaderRow.appendTableCell("Time");
  orderHeaderRow.appendTableCell("Phone");
  orderHeaderRow.appendTableCell("Store");
  orderHeaderRow.appendTableCell("Payment");
  orderHeaderRow.appendTableCell("Total");
  orderHeaderRow.appendTableCell("Type");
  orderHeaderRow.appendTableCell("Order Source");
  
  // Sort orders by time for logical flow
  orders.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  
  orders.forEach(order => {
    const time = new Date(order.timestamp).toLocaleTimeString('en-US', { 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    // Format items based on order type
    const itemSummary = order.items.map(item => {
      if (order.orderType === 'prepaid') {
        return `${item.name}(${item.quantity})`;
      } else {
        return `${item.name}(${item.quantity})`;
      }
    }).join(', ');
    
    const fullLocation = `${order.location}, ${order.room}`;
    const paymentInfo = order.paymentMethod + (order.gcashRef ? ` (${order.gcashRef})` : '');
    const orderTypeDisplay = order.orderType === 'prepaid' ? 'PREPAID' : 'ITEMIZED';
    const orderSourceDisplay = order.orderSource === 'field1' ? 'Order Field' : 'Pickup Field';
    
    const row = ordersTable.appendTableRow();
    row.appendTableCell(order.name);
    row.appendTableCell(itemSummary);
    row.appendTableCell(fullLocation);
    row.appendTableCell(time);
    row.appendTableCell(order.phone);
    row.appendTableCell(order.store);
    row.appendTableCell(paymentInfo);
    row.appendTableCell(`₱${(order.subtotal + order.deliveryFee).toFixed(2)}`);
    row.appendTableCell(orderTypeDisplay);
    row.appendTableCell(orderSourceDisplay);
  });
  
  body.appendParagraph(""); // Empty line

  // Footer notes
  body.appendParagraph("📋 DISPATCH NOTES").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("✅ All orders are listed chronologically for efficient processing");
  body.appendParagraph("📱 Contact customers using provided phone numbers for delivery coordination");
  body.appendParagraph("💰 ITEMIZED orders: collect full amount | PREPAID orders: collect service fee only");
  body.appendParagraph("📝 Order Source indicates which form field contained the order information");
  body.appendParagraph("🚛 Ready for dispatch - Campus Cart Team!");

  doc.saveAndClose();
  return DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
}


/** 🚛 ENHANCED DAILY DISPATCH LIST GENERATOR **/
function sendDailySummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const today = new Date();
  const dateStr = today.toLocaleDateString('en-US', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  const start = new Date(today.setHours(0, 0, 0, 0));
  const end = new Date(today.setHours(23, 59, 59, 999));
  const rows = sheet.getDataRange().getValues().slice(1);

  const todayOrders = rows.filter(row => {
    const time = new Date(row[0]);
    return time >= start && time <= end;
  });

  const summary = {
    totalOrders: todayOrders.length,
    totalRevenue: 0,
    totalDelivery: 0,
    orders: [],
    storeGroups: {}
  };

  // Process orders and group by store
  todayOrders.forEach(row => {
    // CORRECTED: Match new CSV structure
    const [timestamp, name, email, phone, location, room, orderType, orderField1, orderField2, store, paymentMethod, gcashRef] = row;
    
    // Smart order field selection
    const rawOrder = determineOrderField(orderField1, orderField2, orderType);
    if (!rawOrder) return; // Skip if no order data
    
    // Parse order using enhanced parser
    const parseResult = parseOrderItemsEnhanced(rawOrder);
    if (!parseResult.success) return; // Skip failed parsing
    
    const items = parseResult.items;
    const detectedOrderType = parseResult.type;
    const subtotal = items.reduce((sum, item) => sum + (item.quantity * (item.price || 0)), 0);
    const deliveryFee = calculateDeliveryFee(store); // Always calculate service fee regardless of order type
    
    summary.totalRevenue += subtotal;
    summary.totalDelivery += deliveryFee;
    
    const orderData = { 
      timestamp, name, email, phone, location, room, store, 
      items, subtotal, deliveryFee, paymentMethod, gcashRef, 
      orderType: detectedOrderType,
      selectedOrderType: orderType,
      orderSource: rawOrder === orderField1 ? 'field1' : 'field2'
    };
    
    summary.orders.push(orderData);
    
    // Group by store for dispatch
    if (!summary.storeGroups[store]) {
      summary.storeGroups[store] = [];
    }
    summary.storeGroups[store].push(orderData);
  });

  // Generate PDF dispatch list with fallback
  const pdfResult = generateDispatchPDFWithFallback(summary, dateStr);

  // Backup dispatch summary if PDF was generated successfully
  if (pdfResult.pdf) {
    backupDispatchToMonth(pdfResult.pdf, dateStr);
  }

  // Create dispatch text and HTML for email
  const dispatchText = generateDispatchText(summary, dateStr);
  const dispatchHTML = generateDispatchHTML(summary, dateStr, pdfResult.fallback);

  // Send dispatch email
  try {
    const emailOptions = {
      htmlBody: dispatchHTML
    };
    
    if (pdfResult.pdf) {
      emailOptions.attachments = [pdfResult.pdf];
    }

    GmailApp.sendEmail(
      "campuscart59@gmail.com", 
      `🚛 Campus Cart Dispatch List - ${dateStr}`, 
      dispatchText, 
      emailOptions
    );
    
    Logger.log(`✅ Dispatch list sent successfully for ${dateStr} ${pdfResult.fallback ? '(without PDF)' : '(with PDF)'}`);
  } catch (err) {
    Logger.log("❌ Error sending dispatch email: " + err.message);
    
    // Fallback: Send basic text email
    try {
      GmailApp.sendEmail(
        "campuscart59@gmail.com", 
        `🚛 Campus Cart Dispatch List - ${dateStr} (Fallback)`, 
        `${dispatchText}\n\n⚠️ Note: This is a fallback email due to email delivery issues.`
      );
      Logger.log("✅ Fallback dispatch list sent successfully");
    } catch (fallbackErr) {
      Logger.log("❌ Even fallback email failed: " + fallbackErr.message);
    }
  }
}


/** 📱 MANUAL DISPATCH LIST SENDER **/
function sendManualDispatch() {
  Logger.log("🔧 Manual dispatch list triggered");
  sendDailySummary();
}


/** 📝 ENHANCED DISPATCH TEXT GENERATOR - KEEP EMOJIS FOR ADMIN **/
function generateDispatchText(summary, dateStr) {
  const { totalOrders, totalRevenue, totalDelivery, orders } = summary;
  
  let text = `🚛 CAMPUS CART DISPATCH LIST - ${dateStr}\n`;
  text += `📊 SUMMARY: ${totalOrders} Orders | ${totalOrders} Customers | ₱${totalRevenue.toFixed(2)} Revenue | ₱${(totalRevenue + totalDelivery).toFixed(2)} Total Collection\n\n`;

  if (totalOrders === 0) {
    text += "📭 No orders to dispatch today.\n";
    return text;
  }

  text += "👥 ORDERS BY CUSTOMER:\n";
  text += "=" + "=".repeat(60) + "\n";
  
  // Sort orders by time
  orders.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  
  orders.forEach((order, index) => {
    const time = new Date(order.timestamp).toLocaleTimeString('en-US', { 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    const itemSummary = order.items.map(item => {
      if (order.orderType === 'prepaid') {
        return `${item.name}(${item.quantity})`;
      } else {
        return `${item.name}(${item.quantity})`;
      }
    }).join(', ');
    
    const paymentInfo = order.paymentMethod + (order.gcashRef ? ` (${order.gcashRef})` : '');
    const orderTypeDisplay = order.orderType === 'prepaid' ? 'PREPAID' : 'ITEMIZED';
    const orderSourceDisplay = order.orderSource === 'field1' ? 'Order Field' : 'Pickup Field';
    
    text += `\n${index + 1}. ${order.name} | ${time} | ${order.phone}\n`;
    text += `   📍 ${order.location}, ${order.room}\n`;
    text += `   🛍️ ${itemSummary}\n`;
    text += `   🏪 ${order.store} | 💳 ${paymentInfo}\n`;
    text += `   📦 ${orderTypeDisplay} ORDER | 📝 Source: ${orderSourceDisplay}\n`;
    
    if (order.orderType === 'itemized') {
      text += `   💰 COLLECT: ₱${(order.subtotal + order.deliveryFee).toFixed(2)}\n`;
    } else {
      text += `   💰 COLLECT SERVICE FEE: ₱${order.deliveryFee.toFixed(2)}\n`;
    }
  });

  text += `\n📊 TOTALS: ${totalOrders} Orders | ₱${(totalRevenue + totalDelivery).toFixed(2)} Total Collection\n`;
  text += "🚛 Ready for dispatch!\n";

  return text;
}


/** 🌐 ENHANCED DISPATCH HTML GENERATOR - KEEP EMOJIS FOR ADMIN **/
function generateDispatchHTML(summary, dateStr, fallback) {
  const { totalOrders, totalRevenue, totalDelivery, orders } = summary;
  
  const fallbackNote = fallback
    ? `<p><strong>⚠️ Note:</strong> PDF dispatch list could not be generated. All information is provided below.</p>`
    : `<p><strong>📄 PDF Dispatch List:</strong> Complete dispatch list with detailed tables is attached.</p>`;

  let html = `
    <h2>🚛 Campus Cart Dispatch List</h2>
    <p><strong>📅 Date:</strong> ${dateStr}</p>
    
    <table border="1" style="border-collapse: collapse; margin: 20px 0;">
      <tr style="background-color: #f0f0f0;">
        <th style="padding: 10px;">Total Orders</th>
        <th style="padding: 10px;">Total Customers</th>
        <th style="padding: 10px;">Total Revenue</th>
        <th style="padding: 10px;">Total Collection</th>
      </tr>
      <tr>
        <td style="padding: 10px; text-align: center;">${totalOrders}</td>
        <td style="padding: 10px; text-align: center;">${totalOrders}</td>
        <td style="padding: 10px; text-align: center;">₱${totalRevenue.toFixed(2)}</td>
        <td style="padding: 10px; text-align: center;">₱${(totalRevenue + totalDelivery).toFixed(2)}</td>
      </tr>
    </table>
    
    ${fallbackNote}
  `;

  if (totalOrders === 0) {
    html += `<p><strong>📭 No orders to dispatch today.</strong></p>`;
    return html;
  }

  html += `<h3>👥 Orders by Customer</h3>`;
  
  // Sort orders by time
  orders.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  
  orders.forEach((order, index) => {
    const time = new Date(order.timestamp).toLocaleTimeString('en-US', { 
      hour: '2-digit', 
      minute: '2-digit' 
    });
    
    const itemSummary = order.items.map(item => {
      if (order.orderType === 'prepaid') {
        return `${item.name}(${item.quantity})`;
      } else {
        return `${item.name}(${item.quantity})`;
      }
    }).join(', ');
    
    const paymentInfo = order.paymentMethod + (order.gcashRef ? ` (${order.gcashRef})` : '');
    const orderTypeDisplay = order.orderType === 'prepaid' ? 'PREPAID' : 'ITEMIZED';
    const orderSourceDisplay = order.orderSource === 'field1' ? 'Order Field' : 'Pickup Field';
    const collectAmount = order.orderType === 'itemized' 
      ? `₱${(order.subtotal + order.deliveryFee).toFixed(2)}` 
      : `₱${order.deliveryFee.toFixed(2)} (Service Fee Only)`;
    
    html += `
      <div style="margin-bottom: 15px; padding: 12px; border-left: 4px solid #007bff; background-color: #f8f9fa;">
        <strong>${index + 1}. ${order.name}</strong> | ${time} | ${order.phone}<br>
        📍 ${order.location}, ${order.room}<br>
        🛍️ ${itemSummary}<br>
        🏪 ${order.store} | 💳 ${paymentInfo}<br>
        📦 ${orderTypeDisplay} ORDER | 📝 Source: ${orderSourceDisplay}<br>
        <strong>💰 COLLECT: ${collectAmount}</strong>
      </div>
    `;
  });

  html += `
    <p><em>Dispatch list generated on: ${new Date().toLocaleString()}</em></p>
    <p><strong>🚛 Campus Cart Team - Ready to Dispatch!</strong></p>
  `;

  return html;
}


/** ❌ ENHANCED ERROR HANDLING - NO EMOJIS FOR CUSTOMERS **/
function sendError(email, message) {
  try {
    if (!email) {
      Logger.log("⚠️ Cannot send error message: Email is empty.");
      return;
    }

    GmailApp.sendEmail(email, "Campus Cart Submission Error", '', {
      htmlBody: `<p>Hi there!</p>
        <p>We couldn't process your order because:</p>
        <p><strong>${message}</strong></p>
        <p>Please double-check your form and try again.</p>
        <p>Need help? Message us here: <a href="https://m.me/639741992565449">Customer Support</a></p>`
    });
  } catch (err) {
    Logger.log("❌ Error sending error email: " + err.message);
  }
}


/** 🗂️ BACKUP MAINTENANCE FUNCTIONS **/
function organizeExistingBackups() {
  try {
    Logger.log("🔧 Starting manual backup organization...");
    
    const folders = createBackupFolders();
    if (!folders) {
      Logger.log("❌ Could not create/access backup folders");
      return;
    }

    // Get all files in receipts folder that aren't already organized
    const receiptFiles = folders.receipts.getFiles();
    let receiptCount = 0;
    
    while (receiptFiles.hasNext()) {
      const file = receiptFiles.next();
      const fileName = file.getName();
      
      // Skip if it's already in a monthly folder or is a folder itself
      if (fileName.includes('_') && !fileName.includes('January') && !fileName.includes('February')) {
        try {
          // Extract date from filename to determine month
          const dateMatch = fileName.match(/(\d{1,2}-\d{1,2}-\d{4})/);
          if (dateMatch) {
            const dateParts = dateMatch[1].split('-');
            const fileDate = new Date(`${dateParts[2]}-${dateParts[0]}-${dateParts[1]}`);
            const monthYear = fileDate.toLocaleDateString('en-US', { 
              year: 'numeric', 
              month: 'long' 
            });
            
            // Create monthly folder if it doesn't exist
            let monthFolder;
            const existingMonth = folders.receipts.getFoldersByName(monthYear);
            if (existingMonth.hasNext()) {
              monthFolder = existingMonth.next();
            } else {
              monthFolder = folders.receipts.createFolder(monthYear);
            }
            
            // Move file to monthly folder
            file.moveTo(monthFolder);
            receiptCount++;
            Logger.log(`📁 Moved receipt to ${monthYear}: ${fileName}`);
          }
        } catch (error) {
          Logger.log(`⚠️ Could not organize file ${fileName}: ${error}`);
        }
      }
    }

    // Get all files in dispatch folder that aren't already organized
    const dispatchFiles = folders.dispatch.getFiles();
    let dispatchCount = 0;
    
    while (dispatchFiles.hasNext()) {
      const file = dispatchFiles.next();
      const fileName = file.getName();
      
      // Skip if it's already in a monthly folder or is a folder itself
      if (fileName.includes('_') && !fileName.includes('January') && !fileName.includes('February')) {
        try {
          // Extract date from filename
          const dateMatch = fileName.match(/Dispatch_(\w+_\d+_\d{4})/);
          if (dateMatch) {
            const dateStr = dateMatch[1].replace(/_/g, ' ');
            const fileDate = new Date(dateStr);
            if (!isNaN(fileDate.getTime())) {
              const monthYear = fileDate.toLocaleDateString('en-US', { 
                year: 'numeric', 
                month: 'long' 
              });
              
              // Create monthly folder if it doesn't exist
              let monthFolder;
              const existingMonth = folders.dispatch.getFoldersByName(monthYear);
              if (existingMonth.hasNext()) {
                monthFolder = existingMonth.next();
              } else {
                monthFolder = folders.dispatch.createFolder(monthYear);
              }
              
              // Move file to monthly folder
              file.moveTo(monthFolder);
              dispatchCount++;
              Logger.log(`📁 Moved dispatch to ${monthYear}: ${fileName}`);
            }
          }
        } catch (error) {
          Logger.log(`⚠️ Could not organize file ${fileName}: ${error}`);
        }
      }
    }

    Logger.log(`✅ Backup organization complete! Moved ${receiptCount} receipts and ${dispatchCount} dispatch files`);
    
  } catch (error) {
    Logger.log("❌ Error organizing existing backups: " + error);
  }
}


/** 🧹 CLEANUP OLD BACKUPS **/
function cleanupOldBackups() {
  try {
    Logger.log("🧹 Starting backup cleanup...");
    
    const folders = createBackupFolders();
    if (!folders) return;

    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 6); // Keep 6 months of backups
    
    let deletedCount = 0;

    // Cleanup receipts
    const receiptSubfolders = folders.receipts.getFolders();
    while (receiptSubfolders.hasNext()) {
      const monthFolder = receiptSubfolders.next();
      const folderName = monthFolder.getName();
      
      try {
        const folderDate = new Date(folderName + ' 1'); // Add day to make it a valid date
        if (folderDate < cutoffDate) {
          const fileCount = countFilesInFolder(monthFolder);
          monthFolder.setTrashed(true);
          deletedCount += fileCount;
          Logger.log(`🗑️ Deleted old receipt folder: ${folderName} (${fileCount} files)`);
        }
      } catch (error) {
        Logger.log(`⚠️ Could not process receipt folder ${folderName}: ${error}`);
      }
    }

    // Cleanup dispatch summaries
    const dispatchSubfolders = folders.dispatch.getFolders();
    while (dispatchSubfolders.hasNext()) {
      const monthFolder = dispatchSubfolders.next();
      const folderName = monthFolder.getName();
      
      try {
        const folderDate = new Date(folderName + ' 1');
        if (folderDate < cutoffDate) {
          const fileCount = countFilesInFolder(monthFolder);
          monthFolder.setTrashed(true);
          deletedCount += fileCount;
          Logger.log(`🗑️ Deleted old dispatch folder: ${folderName} (${fileCount} files)`);
        }
      } catch (error) {
        Logger.log(`⚠️ Could not process dispatch folder ${folderName}: ${error}`);
      }
    }

    Logger.log(`✅ Cleanup complete! Deleted ${deletedCount} old backup files`);
    
  } catch (error) {
    Logger.log("❌ Error during backup cleanup: " + error);
  }
}


/** 📊 COUNT FILES IN FOLDER **/
function countFilesInFolder(folder) {
  let count = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    files.next();
    count++;
  }
  return count;
}


/** 📈 BACKUP STATUS REPORT **/
function generateBackupStatusReport() {
  try {
    Logger.log("📈 Generating backup status report...");
    
    const folders = createBackupFolders();
    if (!folders) return;

    let report = "📊 CAMPUS CART BACKUP STATUS REPORT\n";
    report += "=" + "=".repeat(50) + "\n\n";
    report += `📅 Generated: ${new Date().toLocaleString()}\n\n`;

    // Receipts status
    report += "📄 RECEIPTS BACKUP STATUS:\n";
    const receiptSubfolders = folders.receipts.getFolders();
    let totalReceiptFiles = 0;
    
    while (receiptSubfolders.hasNext()) {
      const monthFolder = receiptSubfolders.next();
      const folderName = monthFolder.getName();
      const fileCount = countFilesInFolder(monthFolder);
      totalReceiptFiles += fileCount;
      report += `   ${folderName}: ${fileCount} files\n`;
    }
    
    report += `   TOTAL RECEIPTS: ${totalReceiptFiles} files\n\n`;

    // Dispatch status
    report += "🚛 DISPATCH SUMMARIES BACKUP STATUS:\n";
    const dispatchSubfolders = folders.dispatch.getFolders();
    let totalDispatchFiles = 0;
    
    while (dispatchSubfolders.hasNext()) {
      const monthFolder = dispatchSubfolders.next();
      const folderName = monthFolder.getName();
      const fileCount = countFilesInFolder(monthFolder);
      totalDispatchFiles += fileCount;
      report += `   ${folderName}: ${fileCount} files\n`;
    }
    
    report += `   TOTAL DISPATCH FILES: ${totalDispatchFiles} files\n\n`;
    report += `💾 GRAND TOTAL: ${totalReceiptFiles + totalDispatchFiles} backup files\n`;
    report += "\n✅ Backup system operational!";

    // Send report via email
    GmailApp.sendEmail(
      "campuscart59@gmail.com", 
      "📊 Campus Cart Backup Status Report", 
      report
    );
    
    Logger.log("✅ Backup status report sent successfully");
    Logger.log(report);
    
  } catch (error) {
    Logger.log("❌ Error generating backup status report: " + error);
  }
}


/** ⏱ TRIGGER SETUP FUNCTIONS **/
function createDailyTriggers() {
  // Remove existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendDailySummary") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger at 2 PM
  ScriptApp.newTrigger("sendDailySummary")
    .timeBased()
    .everyDays(1)
    .atHour(14) // 2 PM
    .create();
    
  Logger.log("✅ Daily dispatch trigger created for 2 PM");
}


function createBackupMaintenanceTriggers() {
  // Remove existing backup triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "cleanupOldBackups" || 
        trigger.getHandlerFunction() === "generateBackupStatusReport") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Weekly backup cleanup (every Sunday at 3 AM)
  ScriptApp.newTrigger("cleanupOldBackups")
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  // Monthly backup status report (every 30 days at 9 AM)
  ScriptApp.newTrigger("generateBackupStatusReport")
    .timeBased()
    .everyDays(30)
    .atHour(9)
    .create();
    
  Logger.log("✅ Backup maintenance triggers created");
  Logger.log("   - Weekly cleanup: Sundays at 3 AM");
  Logger.log("   - Monthly report: Every 30 days at 9 AM");
}


/** 🚀 ENHANCED SYSTEM SETUP **/
function setupCampusCartSystem() {
  Logger.log("🚀 Setting up Enhanced Campus Cart system...");
  
  // Create backup folders
  createBackupFolders();
  
  // Set up daily dispatch triggers
  createDailyTriggers();
  
  // Set up backup maintenance triggers
  createBackupMaintenanceTriggers();
  
  // Organize any existing backups
  organizeExistingBackups();
  
  // Generate initial status report
  generateBackupStatusReport();
  
  Logger.log("✅ Enhanced Campus Cart system setup complete!");
  Logger.log("📁 Backup folders created and organized");
  Logger.log("⏰ Triggers set up for automation");
  Logger.log("📊 Initial backup report generated");
  Logger.log("🔧 Enhanced features: Smart parsing, instant pings, feedback integration");
  Logger.log("📝 CSV Structure: Corrected for dual order fields and proper column mapping");
  Logger.log("📧 Customer emails: Emojis removed for professional appearance");
}


/** 🧪 TESTING FUNCTIONS **/
function testOrderFieldDetermination() {
  Logger.log("🧪 Testing order field determination...");
  
  // Test cases
  const testCases = [
    {
      field1: "Burger 2 @80",
      field2: "",
      orderType: "Order",
      expected: "field1"
    },
    {
      field1: "",
      field2: "Pick up Burger 2",
      orderType: "Pickup",
      expected: "field2"
    },
    {
      field1: "Burger 2 @80",
      field2: "Fries 1 @45",
      orderType: "Order",
      expected: "field1"
    },
    {
      field1: "Short",
      field2: "Much longer description with more details",
      orderType: "",
      expected: "field2"
    }
  ];
  
  testCases.forEach((testCase, index) => {
    const result = determineOrderField(testCase.field1, testCase.field2, testCase.orderType);
    const resultField = result === testCase.field1 ? "field1" : "field2";
    const passed = resultField === testCase.expected;
    
    Logger.log(`Test ${index + 1}: ${passed ? '✅ PASS' : '❌ FAIL'} - Expected: ${testCase.expected}, Got: ${resultField}`);
  });
}


function testOrderParsing() {
  Logger.log("🧪 Testing order parsing...");
  
  const testOrders = [
    "Burger 2 @80",
    "Burger 2 @ 80, Fries 1 @ 45",
    "Burger 2",
    "Burger 2, Fries 1",
    "Invalid format here",
    ""
  ];
  
  testOrders.forEach((order, index) => {
    const result = parseOrderItemsEnhanced(order);
    Logger.log(`Test ${index + 1}: "${order}" -> ${result.success ? '✅ SUCCESS' : '❌ FAILED'} (${result.type || result.error})`);
    if (result.success) {
      Logger.log(`   Items: ${JSON.stringify(result.items)}`);
    }
  });
}


/** 🔧 MANUAL TESTING FUNCTIONS **/
function simulateFormSubmission() {
  Logger.log("🔧 Simulating form submission for testing...");
  
  // Sample data matching the CSV structure
  const sampleData = [
    new Date(), // Timestamp
    "John Doe", // Name
    "john.doe@example.com", // Email
    "09123456789", // Phone
    "Dorm A", // Delivery Location
    "Room 123", // Room Number
    "Order", // Choose your Order type
    "Burger 2 @80, Fries 1 @45", // What will you like to Order
    "", // Tell us what you need picked up
    "Store Near AUP", // Where are you buying from
    "GCash", // How will you pay
    "REF123456", // GCash Reference
    "I agree to all terms and conditions" // Terms agreement
  ];
  
  // Create mock event object
  const mockEvent = {
    values: sampleData
  };
  
  try {
    onFormSubmit(mockEvent);
    Logger.log("✅ Form submission simulation completed successfully");
  } catch (error) {
    Logger.log("❌ Form submission simulation failed: " + error);
  }
}


/** 📋 SYSTEM DIAGNOSTIC **/
function runSystemDiagnostic() {
  Logger.log("🔍 Running Campus Cart system diagnostic...");
  
  try {
    // Test 1: Check spreadsheet access
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    if (sheet) {
      Logger.log("✅ Spreadsheet access: OK");
      Logger.log(`   Rows: ${sheet.getLastRow()}, Columns: ${sheet.getLastColumn()}`);
    } else {
      Logger.log("❌ Spreadsheet access: FAILED");
    }
    
    // Test 2: Check backup folders
    const folders = createBackupFolders();
    if (folders) {
      Logger.log("✅ Backup system: OK");
    } else {
      Logger.log("❌ Backup system: FAILED");
    }
    
    // Test 3: Test email system
    try {
      GmailApp.sendEmail("campuscart59@gmail.com", "🔍 Campus Cart Diagnostic Test", "System diagnostic test email - ignore this message.");
      Logger.log("✅ Email system: OK");
    } catch (emailError) {
      Logger.log("❌ Email system: FAILED - " + emailError);
    }
    
    // Test 4: Check triggers
    const triggers = ScriptApp.getProjectTriggers();
    const dailyTriggers = triggers.filter(t => t.getHandlerFunction() === "sendDailySummary");
    Logger.log(`✅ Triggers: ${triggers.length} total, ${dailyTriggers.length} daily dispatch`);
    
    // Test 5: Test order parsing
    testOrderParsing();
    
    // Test 6: Test order field determination
    testOrderFieldDetermination();
    
    Logger.log("🏁 System diagnostic complete!");
    
  } catch (error) {
    Logger.log("❌ System diagnostic error: " + error);
  }
}