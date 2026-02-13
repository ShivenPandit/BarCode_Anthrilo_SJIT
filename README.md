# Anthrilo Barcode Label Generator

Professional barcode label generator for Anthrilo Design House products.

## ⚡ Critical Updates

- ✅ Fixed label size to **5cm × 9.5cm** (590 × 1122 pixels @ 300 DPI)
- ✅ Barcode generated from **SKU Code** (NOT Style Code)
- ✅ Text below barcode shows **SKU Code only**
- ✅ Single desktop application - no Python installation required

## Key Features

✅ **Correct Label Size** - Exactly 5cm × 9.5cm  
✅ **SKU-Based Barcodes** - Barcode from SKU Code column  
✅ **Professional Layout** - All required fields with proper spacing  
✅ **Standalone App** - No dependencies to install  
✅ **Excel/CSV Support** - Works with your spreadsheets  
✅ **PDF Output** - All labels in single file  

## Quick Start

**Option 1: Use the Executable (Recommended)**
```bash
.\APP\AnthriloLabelGenerator.exe
```

**Option 2: Run with Python**
```bash
python label_generator_gui.py
```

## Excel File Requirements

Your Excel file must have these columns:

| Column | Purpose | Alternative Names |
|--------|---------|-------------------|
| Vendor Article Name | Product name | Product Name |
| Size | Product size | Sizes |
| Brand Name | Brand | Brand |
| Vendor Article No | Style Code (text only) | Style Code |
| **SKU Code** | **For barcode generation** | SKU, SKU_CODE |
| MRP | Retail price | Price, Retail Price |

### Example Excel Data:
```
Vendor Article Name              Size  Brand Name  Vendor Article No  SKU Code       MRP
DINOSAUR FUN PYJAMA SET         7-8Y  Anthrilo    AG-001            8901234567890  599
LONG SLEEVE SHIRT               5-6Y  Anthrilo    AG-002            8901234567891  399
...
```

## How to Use

1. **Launch the Application**
   - Double-click `AnthriloLabelGenerator.exe`

2. **Select Excel File**
   - Click "Browse" next to "Excel/CSV File"
   - Choose your product data file

3. **Choose Output Folder**
   - Click "Browse" next to "Output Folder"
   - Select destination for PDF

4. **Generate Labels**
   - Click "Generate Labels" button
   - PDF will be created as `labels.pdf`

## Label Layout

```
┌────────────────────────────────────┐
│ PRODUCT NAME - SIZE                │  ← Center aligned
│                                    │
│ Brand : [Brand Name]               │
│ Style code : [Vendor Article No]   │
│ Size : [Size]                      │
│                                    │
│ Month/Year of Manufacture: MM/YYYY │
│ Marketed by : Anthrilo Design House│
│ KH400/414 Rahon Road, Punjab 141007│
│ customercare@anthrilo.com          │
│ +919888264040                      │
│                                    │
│ MRP: [Price] (inclusive of all    │
│      taxes)                        │
│                                    │
│  ┌──────────────────────────────┐ │
│  │  ▌▌ ▌ ▌▌▌ ▌▌ ▌ ▌▌▌ ▌▌ ▌     │ │
│  └──────────────────────────────┘ │
│        [SKU CODE TEXT]             │
└────────────────────────────────────┘
```

## Label Specifications

- **Size**: 5cm width × 9.5cm height
- **Pixels**: 590 × 1122 (at 300 DPI)
- **Format**: PDF with all labels
- **Barcode Type**: Code128 from SKU Code
- **Barcode Settings** (Optimized for Laser Scanners):
  - Module Width (X-dimension): 0.40mm (bar thickness)
  - Module Height: 20mm
  - Quiet Zone: 5mm on each side
  - Scale: 140% width for reliable scanning
  - Color: Pure black (#000000)
  - DPI: 300
- **One Row = One Label**

## ⚠️ Important Notes

### Barcode Source: SKU Code
- Barcode is generated from the **SKU Code** column ONLY
- Style Code is shown as text on the label (NOT used for barcode)
- If SKU Code is missing → Row is skipped automatically

### Text Below Barcode
- Shows **SKU Code value only**
- Not the barcode symbology/machine value
- Center-aligned, single line

## Troubleshooting

### "Could not find required columns"
Check that your Excel file has:
- Vendor Article Name / Product Name
- Size / Sizes
- Brand Name / Brand
- Vendor Article No / Style Code
- **SKU Code / SKU** (REQUIRED)
- MRP / Price

### "Skipped N rows with missing SKU Code"
Add SKU Code values to those rows to generate their labels.

### Labels cut off or spacing issues
Ensure no extremely long product names (max ~60 characters recommended).

## System Requirements

- Windows 7 or later
- OR Python 3.6+ with packages from `requirements.txt`
- Excel file in .xlsx, .xls, or .csv format

## Development

If you want to modify the code:

```bash
pip install -r requirements.txt
python label_generator_gui.py
```

Build new executable:
```bash
python build_exe_simple.py
```

## Support

Contact Anthrilo Design House:
- Email: customercare@anthrilo.com
- Phone: +919888264040
- Address: KH400/414 Rahon Road, Punjab 141007

---

**Version:** 3.1 (Professional Barcode Scanning)  
**Updated:** February 12, 2026  
**Key Improvements:** Optimized barcode for laser scanners with wider bars (140%), proper module width (0.40mm), and adequate quiet zones (5mm) for reliable scanning  
Anthrilo Design House

