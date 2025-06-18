# 📩 Gmail Quotation Tracker Using Google Apps Script
Link: https://docs.google.com/spreadsheets/d/11cHTSTjVQC96aTKilPms0tdPbugQRn5uU7GCfCYIh6o/edit?usp=sharing

![Image](https://github.com/user-attachments/assets/e02beabc-d4d0-4429-88e3-054a434f09b6)

## 📝 Objective

Automate the process of extracting quotation details from emails and logging them into a Google Sheet, without processing the same email message twice, even across multiple runs.

---

## ✅ Functional Requirements

1. **Email Source**:
   - Target Gmail inbox.
   - Only emails with `"Quotation"` in the subject line.
   - Only include emails received **within the last 30 days**.
   - Only consider emails that were **never processed before**.

2. **Data Extraction & Logging**:
   - For each new (unseen) message, extract the following:
     - Date
     - Sender
     - Subject
     - Product
     - Quantity
     - Unit Price
     - Total Price
     - Delivery Time
     - Valid Till
   - Append this data to a Google Sheet named: `Price Quotations`
   - Sheet tab name: `Quotations`

3. **Trigger Mechanism**:
   - A custom menu in the spreadsheet:
     - Menu Name: **📩 Email Actions**
     - Item: **🔄 Refresh List**
     - Action: Runs the function `processQuotationEmails`
   - No timed or installable triggers (handled manually by user).

4. **Mark Emails as Read**:
   - After processing, mark emails as read to avoid visual clutter in the inbox.

---

## 🛠️ Script Outline

### Gmail Search Query:
```js
const query = 'subject:Quotation newer_than:30d is:unread';
