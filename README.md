# 🧰 Multi-Shop Stock Check Toolkit  
**Excel + VBA toolkit for multi-branch retail stock checking and supplier ordering**

---

### 📌 Overview  
This toolkit is a production-used Excel + VBA system built for a small retail chain to **standardise stock checks**, **automate checklist exports**, and **generate supplier orders** across multiple branches, without needing EPOS integration.

It replaces a slow, error-prone manual workflow with a reliable and structured process.

---

## 🚀 Features

### **📦 Central Product Catalogue**
Single source of truth for all products:
- Product ID  
- Category  
- Brand  
- Display name  
- Order-unit size  
- Global/Local activeness  

---

### **🏪 Per-Shop Smart Stock Sheets**
Each shop has its own `ShopName_Stock` sheet with:
- Current stock (units)  
- Par levels  
- Order-unit size  
- Last counted date  
- Local vs Global status  
- Auto-sorting  
- Quick-toggle filters (hide inactive / show all)  

All sheets follow an *identical structure*, enabling automation.

---

### **📝 One-Click Checklist Export**
Exports a clean stock-counting checklist for any shop:

- Only active items  
- Grouped by category  
- Auto-sorted  
- Shop-branded header  
- Saved automatically to:  
  **`/Exports/Checklists/ShopName_StockCheck_YYYY-MM-DD.xlsx`**

---

### **📥 One-Click Checklist Import**
After a shop returns their completed checklist:

- Reads counted units  
- Updates the correct shop sheet  
- Stamps "Last Counted" date  
- Ignores invalid or missing rows  
- Works even if importing only one shop at a time  

---

### **📦 Automated Supplier Order Generator**
Aggregates all shops’ checklists and produces a ready-to-paste WhatsApp order message:

- Totals required units/boxes across all branches  
- Uses each product’s order-unit size  
- Saves output to:  
  **`/Exports/Supplier_Order_YYYY-MM-DD.txt`**

Example output:
- Gotek X – Blue (3 Units)
- Gotek Pro – Titanium (2 Boxes of 10)


---

## 🔧 Reliability & Compatibility

### ✔ OneDrive-Safe File Handling  
Automatically resolves whether Excel is opened via:
- Local OneDrive path (`C:\Users\...\OneDrive\...`)
- Cloud URL (`https://d.docs.live.net/...`)

Ensures exports/imports always save to the correct folders.

### ✔ Password-Protected Workflow  
All sheets are protected with `UserInterfaceOnly`, allowing:
- Sorting  
- Filtering  
- Automation  
while preventing accidental edits.

### ✔ Robust Error Handling  
Includes:
- Safe folder creation  
- Graceful failures on missing checklists  
- Automatic filename conflict resolution  
- Protection recovery  

---

## 📁 Folder Structure

```
Vapers Paradise /
├── Stocklists /
│   ├── Exports /
│   │   └── Checklists /
│   │       └── Supplier_Order_YYYY-MM-DD.txt
│   ├── Imports /
│   │   └── Checklists /   (place returned files here)
│   └── Master Stocklist.xlsm
```

## 🖥 Technology

| Component | Purpose |
|----------|---------|
| **Excel (XLSM)** | Interface + storage |
| **VBA** | Automation + file I/O + sorting + validation |
| **OneDrive Sync** | Cross-machine access |
| **Structured Tables** | Consistent schema across shops |

---

## 📜 Licence
Internal business tool — not licensed for external distribution.

---

## 👤 Author
**Danny Hussain**  
Retail Ops, Automation & Data Specialist  
