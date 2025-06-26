# 📄 Invoice Consolidator App

### 🧰 Built with: **Python**, **Tkinter**, **Pandas**, **ReportLab**

---

## ✨ What is this?

The **Invoice Consolidator App** is a simple, elegant **desktop tool** designed to help teams effortlessly **combine multiple Excel-based invoices** into a **single, professionally formatted PDF summary**.

Perfect for:
- ✅ Finance or admin teams juggling multiple invoices
- ✅ Interns or junior analysts streamlining reporting
- ✅ Anyone tired of **manual copy-paste** madness!

---

## 🧠 Key Features

- 📂 **Add multiple Excel invoice files**
- 📋 **Automatically extracts relevant invoice data**
- 🔍 **Groups and summarizes data by account**
- 🧮 **Calculates total extended price and unique accounts**
- 🖨️ **Generates a beautiful, print-ready PDF report**
- 🖱️ **Simple point-and-click interface**

---

## 🚀 How It Works

### 1️⃣ Launch the App  
Double-click the app or run the script in Python. A clean, user-friendly window will appear.

### 2️⃣ Select Files  
Click **“Add File”** to select one or more Excel invoice files (`.xlsx` or `.xls`). You’ll see them listed clearly on the screen.

### 3️⃣ Process Invoices  
Once your files are selected, hit **“Process Files”**. The app will:
- Read all sheets from each Excel file
- Extract key fields like **Account Name**, **Invoice Number**, **Extended Price**, etc.
- Combine them into a **neatly grouped table**
- Append a **Total row** with a summary of charges

### 4️⃣ Save As PDF  
Choose a location to save the output as a **PDF file** – perfect for sending or printing!

---

## 📸 Screenshot (Optional)
> *Add an image here if you’d like, e.g. the UI or a sample PDF.*

---

## 📦 Requirements

- Python 3.8+
- Libraries used:
  - `pandas`
  - `tkinter`
  - `reportlab`
  - `pygame` *(optional, can be removed if not used)*

You can install the required packages using:
```bash
pip install pandas reportlab
