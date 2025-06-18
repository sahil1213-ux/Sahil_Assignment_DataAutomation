/**
 * Gmail Quotation Extractor to Existing Google Sheet
 * Updates existing "Price Quotations" spreadsheet, "Quotations" sheet
 * Processes emails with "Quotation" in subject line
 */

/**
 * Main function to process quotation emails and update the sheet
 */
function processQuotationEmails() {
  try {
    // Get the existing spreadsheet and sheet
    const sheet = getExistingSheet();
    if (!sheet) {
      console.error('❌ Could not find the required sheet');
      return;
    }
    
    // Search for emails with "Quotation" in subject
    const emails = getQuotationEmails();
    
    if (emails.length === 0) {
      console.log('📧 No emails found with "Quotation" in subject');
      return;
    }
    
    console.log(`📬 Found ${emails.length} emails to process`);
    
    // Process each email
    let processedCount = 0;
    emails.forEach(email => {
      try {
        const extractedData = extractQuotationData(email);
        if (extractedData) {
          appendDataToSheet(sheet, extractedData);
          console.log(`✅ Processed: ${email.getSubject()}`);
          processedCount++;
        } else {
          console.log(`⚠️ No quotation data found in: ${email.getSubject()}`);
        }
      } catch (error) {
        console.error(`❌ Error processing email "${email.getSubject()}":`, error);
      }
    });
    
    console.log(`🎉 Successfully processed ${processedCount} out of ${emails.length} emails`);
    
  } catch (error) {
    console.error('❌ Error in processQuotationEmails:', error);
  }
}

/**
 * Get the existing "Price Quotations" spreadsheet and "Quotations" sheet
 */
function getExistingSheet() {
  try {
    // Get all spreadsheets to find "Price Quotations"
    const files = DriveApp.getFilesByName('Price Quotations');
    
    if (!files.hasNext()) {
      console.error('❌ Spreadsheet "Price Quotations" not found');
      return null;
    }
    
    const file = files.next();
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    
    // Get the "Quotations" sheet
    const sheet = spreadsheet.getSheetByName('Quotations');
    
    if (!sheet) {
      console.error('❌ Sheet "Quotations" not found in spreadsheet "Price Quotations"');
      return null;
    }
    
    console.log('✅ Found existing sheet: Price Quotations > Quotations');
    return sheet;
    
  } catch (error) {
    console.error('❌ Error accessing existing sheet:', error);
    return null;
  }
}

/**
 * Get emails with "Quotation" in subject line
 */
function getQuotationEmails() {
  try {
    // Search Gmail for emails with "Quotation" in subject
    const query = 'subject:Quotation is:unread newer_than:30d';
    const threads = GmailApp.search(query, 0, 50); // Limit to 50 most recent
    
    const emails = [];
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        // Process only unread messages
        if (!message.isUnread()) return;
        emails.push(message);
        message.markRead();
      });
    });
    
    return emails;
    
  } catch (error) {
    console.error('❌ Error fetching emails:', error);
    return [];
  }
}

/**
 * Extract quotation data from email body
 */
function extractQuotationData(email) {
  const body = email.getPlainBody();
  const sender = email.getFrom();
  const subject = email.getSubject();
  const date = email.getDate();
  
  console.log(`🔍 Extracting data from: ${subject}`);
  
  // Extract all required fields
  const product = extractProduct(body);
  const quantity = extractQuantity(body);
  const unitPrice = extractUnitPrice(body);
  const totalPrice = extractTotalPrice(body);
  const deliveryTime = extractDeliveryTime(body);
  const validTill = extractValidTill(body);
  
  // Check if we have minimum required data
  if (product || quantity || unitPrice || totalPrice) {
    return {
      date: date,
      sender: cleanSenderEmail(sender),
      subject: subject,
      product: product || 'N/A',
      quantity: quantity || 'N/A',
      unitPrice: unitPrice || 'N/A',
      totalPrice: totalPrice || 'N/A',
      deliveryTime: deliveryTime || 'N/A',
      validTill: validTill || 'N/A'
    };
  }
  
  return null;
}

/**
 * Extract product information from email body
 */
function extractProduct(body) {
  const productPatterns = [
    /Product:\s*([^\n\r]+)/i,
    /Item:\s*([^\n\r]+)/i,
    /Product Name:\s*([^\n\r]+)/i,
    /Item Name:\s*([^\n\r]+)/i,
    /Description:\s*([^\n\r]+)/i
  ];
  
  for (const pattern of productPatterns) {
    const match = body.match(pattern);
    if (match && match[1].trim()) {
      return match[1].trim();
    }
  }
  
  return null;
}

/**
 * Extract quantity from email body
 */
function extractQuantity(body) {
  const quantityPatterns = [
    /Quantity:\s*(\d+)\s*units?/i,
    /Qty:\s*(\d+)\s*units?/i,
    /Quantity:\s*(\d+)/i,
    /Qty:\s*(\d+)/i,
    /(\d+)\s*units?/i,
    /(\d+)\s*pieces?/i,
    /(\d+)\s*pcs?/i
  ];
  
  for (const pattern of quantityPatterns) {
    const match = body.match(pattern);
    if (match && match[1]) {
      return parseInt(match[1]);
    }
  }
  
  return null;
}

/**
 * Extract unit price from email body
 */
function extractUnitPrice(body) {
  const unitPricePatterns = [
    /Unit Price:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Price per unit:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Per unit:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Unit cost:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i
  ];
  
  for (const pattern of unitPricePatterns) {
    const match = body.match(pattern);
    if (match && match[2]) {
      const price = parseFloat(match[2].replace(/,/g, ''));
      const currency = match[1] || match[3] || '';
      return currency + price.toLocaleString();
    }
  }
  
  return null;
}

/**
 * Extract total price from email body
 */
function extractTotalPrice(body) {
  const totalPricePatterns = [
    /Total Price:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Total:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Total Amount:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Grand Total:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i,
    /Amount:\s*([₹$€£¥]?)\s*([\d,]+(?:\.\d{2})?)\s*([₹$€£¥]?)/i
  ];
  
  for (const pattern of totalPricePatterns) {
    const match = body.match(pattern);
    if (match && match[2]) {
      const price = parseFloat(match[2].replace(/,/g, ''));
      const currency = match[1] || match[3] || '';
      return currency + price.toLocaleString();
    }
  }
  
  return null;
}

/**
 * Extract delivery time from email body
 */
function extractDeliveryTime(body) {
  const deliveryPatterns = [
    /Delivery time:\s*([^\n\r]+)/i,
    /Delivery:\s*([^\n\r]+)/i,
    /Shipping time:\s*([^\n\r]+)/i,
    /Lead time:\s*([^\n\r]+)/i,
    /Delivery period:\s*([^\n\r]+)/i,
    /Timeline:\s*([^\n\r]+)/i,
    /(\d+[-\s]?\d*\s*(?:days?|weeks?|months?|business days?))/i
  ];
  
  for (const pattern of deliveryPatterns) {
    const match = body.match(pattern);
    if (match && match[1] && match[1].trim()) {
      return match[1].trim();
    }
  }
  
  return null;
}

/**
 * Extract validity date from email body
 */
function extractValidTill(body) {
  const validityPatterns = [
    /Valid till:\s*([^\n\r]+)/i,
    /Valid until:\s*([^\n\r]+)/i,
    /Expires on:\s*([^\n\r]+)/i,
    /Expiry:\s*([^\n\r]+)/i,
    /Quote valid till:\s*([^\n\r]+)/i,
    /Validity:\s*([^\n\r]+)/i,
    /Valid through:\s*([^\n\r]+)/i
  ];
  
  for (const pattern of validityPatterns) {
    const match = body.match(pattern);
    if (match && match[1] && match[1].trim()) {
      return match[1].trim();
    }
  }
  
  return null;
}

/**
 * Clean sender email to extract just the email address
 */
function cleanSenderEmail(sender) {
  // Extract email from "Name <email@domain.com>" format
  const emailMatch = sender.match(/<([^>]+)>/);
  if (emailMatch) {
    return emailMatch[1];
  }
  return sender;
}

/**
 * Append extracted data to the existing sheet
 */
function appendDataToSheet(sheet, data) {
  try {
    // Find the next empty row
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    
    // Prepare the row data in the order of expected columns
    const rowData = [
      data.date,
      data.sender,
      data.subject,
      data.product,
      data.quantity,
      data.unitPrice,
      data.totalPrice,
      data.deliveryTime,
      data.validTill
    ];
    
    // Append the data
    const range = sheet.getRange(nextRow, 1, 1, rowData.length);
    range.setValues([rowData]);
    
    console.log(`📝 Added row ${nextRow} with data:`, data.product);
    
  } catch (error) {
    console.error('❌ Error appending data to sheet:', error);
  }
}

/**
 * Test function to check email extraction without updating the sheet
 */
function testEmailExtraction() {
  try {
    console.log('🧪 Testing email extraction...');
    
    const emails = getQuotationEmails();
    console.log(`Found ${emails.length} emails with "Quotation" in subject`);
    
    if (emails.length > 0) {
      // Test with the first email
      const testEmail = emails[0];
      console.log(`\n📧 Testing with email: "${testEmail.getSubject()}"`);
      console.log(`From: ${testEmail.getFrom()}`);
      console.log(`Date: ${testEmail.getDate()}`);
      
      const extractedData = extractQuotationData(testEmail);
      
      if (extractedData) {
        console.log('\n✅ Extracted data:');
        console.log('Product:', extractedData.product);
        console.log('Quantity:', extractedData.quantity);
        console.log('Unit Price:', extractedData.unitPrice);
        console.log('Total Price:', extractedData.totalPrice);
        console.log('Delivery Time:', extractedData.deliveryTime);
        console.log('Valid Till:', extractedData.validTill);
      } else {
        console.log('❌ No data could be extracted from this email');
      }
    }
    
  } catch (error) {
    console.error('❌ Error in test function:', error);
  }
}

/**
 * Utility function to check if the required sheet exists
 */
function checkSheetExists() {
  const sheet = getExistingSheet();
  if (sheet) {
    console.log('✅ Sheet found successfully');
    console.log('Sheet name:', sheet.getName());
    console.log('Last row with data:', sheet.getLastRow());
    return true;
  } else {
    console.log('❌ Required sheet not found');
    return false;
  }
}