# **📌 Google Sheets Accounting System**

## **📖 Overview**
This project is a **Google Sheets-based Accounting System** built using **Google Apps Script**. It allows users to:

✅ **Manage Payments**: Add payments, update statuses, and convert foreign currencies automatically.
✅ **Track Employee Salaries**: Deduct salaries, mark as paid, and track salary payments.
✅ **Manage Expenses**: Add, pay, and track expenses with automatic Running Total updates.
✅ **Automated Financial Dashboard**: Automatically calculates **Running Total, Partner Share, and Expenses.**

The system ensures real-time financial tracking, eliminating manual calculations, and **keeps Running Total updated dynamically**.

---

## **🚀 Features**
### ✅ **Payments Management**
- Add new payments via **Payment Form**.
- Update payment status to **"Received"** and **automatically update Running Total**.
- Convert **foreign currencies** to PKR using **GOOGLEFINANCE()**.

### ✅ **Employee Salary Management**
- Select employees for salary deduction.
- Deduct salaries and **update Running Total**.
- Prevent duplicate deductions for salaries already paid.

### ✅ **Expense Management**
- Add expenses via **Expense Form**.
- If the expense is marked as **"Paid"**, it is **immediately deducted from Running Total**.
- If the expense is marked as **"Unpaid"**, it will be **deducted when processed via Pay Unpaid Expenses**.

### ✅ **Automated Financial Dashboard**
- Displays **Running Total** with real-time updates.
- Tracks **Partner Share (50%)** and **Partner Share Amount**.
- Updates dynamically when **payments, salaries, or expenses change**.

---

## **⚙️ Setup Instructions**

### 📌 **1. Open Google Sheets & Enable Apps Script**
1. Open **Google Sheets**.
2. Click on **Extensions** > **Apps Script**.
3. Delete any existing script.
4. Copy and paste the provided `Code.js` file into the Apps Script editor.
5. Click **Save** (`Ctrl + S`).

### 📌 **2. Create Necessary Sheets**
Make sure your **Google Sheet** has the following **sheets (tabs):**

- **Payments** → To track payments.
- **Employees** → To track employees and salary status.
- **Expenses** → To track expenses.
- **Dashboard** → To display financial summary.

### 📌 **3. Set Up Menu for UI**
Once the script is deployed, reload your Google Sheets, and you’ll see a new **Accounting Menu** with options:

✅ **Add Payment Record**  
✅ **Add Employee Record**  
✅ **Deduct Salaries & Notify Owner**  
✅ **Add Expense Record**  
✅ **Pay Unpaid Expenses**  
✅ **Update Payment Record**  

---

## **🛠️ How to Use**

### **➤ 1. Adding a Payment**
1. Click **Accounting > Add Payment Record**.
2. Fill in the payment details.
3. If the payment is later marked as "Received", it **adds the amount to Running Total**.

### **➤ 2. Managing Employee Salaries**
1. Click **Accounting > Deduct Salaries & Notify Owner**.
2. Automatically deducts salaries for **employees marked as "Not Received"**.
3. Updates **Running Total** on the Dashboard.
4. Sends an **email notification** of salaries paid.

### **➤ 3. Managing Expenses**
1. Click **Accounting > Add Expense Record**.
2. If the expense is marked as **"Paid"**, the amount is **immediately deducted** from Running Total.
3. If the expense is **"Unpaid"**, use **"Pay Unpaid Expenses"** to deduct it later.

### **➤ 4. Viewing Financial Summary (Dashboard)**
1. Go to the **Dashboard Sheet**.
2. See the **Running Total, Partner Share, and Deducted Expenses.**

---

## **📌 Notes & Best Practices**

- **🚨 Do not manually edit the Running Total cell in the Dashboard Sheet.**
- **💰 Ensure currency conversion formulas are working properly in the Payments sheet.**
- **🔄 Always use the provided menu options for proper updates.**
- **📧 The system automatically sends email notifications for salary payments.**
- **✅ Make sure your spreadsheet follows the expected column structure.**

---

## **📩 Need Help?**
If you need any **modifications, bug fixes, or feature requests**, feel free to reach out!

🚀 **Enjoy seamless Accounting with Google Sheets!** 🚀

