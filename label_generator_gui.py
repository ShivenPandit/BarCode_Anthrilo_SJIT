"""
Anthrilo Barcode Label Generator
Standalone Desktop Application
Generates barcode labels from Excel data with fixed specifications:
- Label Size: 5cm x 9.5cm (590 x 1122 pixels at 300 DPI)
- Barcode Source: SKU Code (not Style Code)
- Output: PDF with all labels
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import tempfile
import sys
from datetime import datetime


class LabelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anthrilo Barcode Label Generator")
        self.root.geometry("700x500")
        self.root.configure(bg="#f0f0f0")
        
        self.excel_path = None
        self.output_dir = None
        
        # Title
        title_label = tk.Label(
            root, 
            text="Anthrilo Barcode Label Generator", 
            font=("Arial", 18, "bold"), 
            bg="#f0f0f0", 
            fg="#333"
        )
        title_label.pack(pady=20)
        
        # Subtitle
        subtitle = tk.Label(
            root,
            text="Generate professional barcode labels from Excel data",
            font=("Arial", 10),
            bg="#f0f0f0",
            fg="#666"
        )
        subtitle.pack(pady=5)
        
        # Frame for file selection
        file_frame = tk.Frame(root, bg="#f0f0f0")
        file_frame.pack(pady=15, padx=20, fill="x")
        
        tk.Label(
            file_frame, 
            text="Excel/CSV File:", 
            font=("Arial", 11), 
            bg="#f0f0f0", 
            width=15, 
            anchor="w"
        ).grid(row=0, column=0, padx=5, pady=8)
        
        self.file_label = tk.Label(
            file_frame, 
            text="No file selected", 
            font=("Arial", 10), 
            bg="white", 
            relief="sunken", 
            anchor="w"
        )
        self.file_label.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        
        tk.Button(
            file_frame, 
            text="Browse", 
            command=self.browse_file,
            bg="#4CAF50", 
            fg="white", 
            font=("Arial", 10, "bold"),
            padx=15
        ).grid(row=0, column=2, padx=5, pady=8)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Frame for output directory
        output_frame = tk.Frame(root, bg="#f0f0f0")
        output_frame.pack(pady=15, padx=20, fill="x")
        
        tk.Label(
            output_frame, 
            text="Output Folder:", 
            font=("Arial", 11), 
            bg="#f0f0f0", 
            width=15, 
            anchor="w"
        ).grid(row=0, column=0, padx=5, pady=8)
        
        self.output_label = tk.Label(
            output_frame, 
            text="No folder selected", 
            font=("Arial", 10), 
            bg="white", 
            relief="sunken", 
            anchor="w"
        )
        self.output_label.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        
        tk.Button(
            output_frame, 
            text="Browse", 
            command=self.browse_output,
            bg="#2196F3", 
            fg="white", 
            font=("Arial", 10, "bold"),
            padx=15
        ).grid(row=0, column=2, padx=5, pady=8)
        
        output_frame.columnconfigure(1, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            root, 
            orient="horizontal", 
            length=400, 
            mode="determinate"
        )
        self.progress.pack(pady=20)
        
        # Status label
        self.status_label = tk.Label(
            root, 
            text="Ready", 
            font=("Arial", 10), 
            bg="#f0f0f0", 
            fg="#666"
        )
        self.status_label.pack(pady=5)
        
        # Generate button
        self.generate_btn = tk.Button(
            root, 
            text="Generate Labels", 
            command=self.generate_labels,
            bg="#FF9800", 
            fg="white", 
            font=("Arial", 14, "bold"),
            padx=40, 
            pady=12, 
            state="disabled"
        )
        self.generate_btn.pack(pady=20)
        
    def browse_file(self):
        """Browse and select Excel/CSV file"""
        filename = filedialog.askopenfilename(
            title="Select Excel or CSV file",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"), 
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.excel_path = filename
            self.file_label.config(text=Path(filename).name)
            self.check_ready()
    
    def browse_output(self):
        """Browse and select output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir = folder
            self.output_label.config(text=folder)
            self.check_ready()
    
    def check_ready(self):
        """Enable/disable generate button based on selections"""
        if self.excel_path and self.output_dir:
            self.generate_btn.config(state="normal")
        else:
            self.generate_btn.config(state="disabled")
    
    def generate_barcode(self, sku_code, label_counter):
        """
        Generate barcode image from SKU Code with professional laser scanning settings
        
        Key parameters for reliable laser scanning:
        - module_width (X-dimension): 0.40mm - controls bar thickness
        - module_height: 20mm - adequate height for laser sweep
        - quiet_zone: 5mm - clear margins for scanner detection
        - dpi: 300 - matches label resolution
        
        Args:
            sku_code (str): SKU Code to generate barcode from
            label_counter (int): Label number for temp file naming
            
        Returns:
            PIL.Image or None: Barcode image, or None if generation fails
        """
        try:
            barcode_value = str(sku_code).strip()
            if not barcode_value or barcode_value.lower() == 'nan':
                return None
            
            # Professional barcode settings for laser scanners
            writer = ImageWriter()
            writer_options = {
                'module_width': 0.40,    # 0.4mm bar width (X-dimension) - CRITICAL for scanning
                'module_height': 20.0,   # 20mm height - recommended for laser scanners
                'quiet_zone': 5.0,       # 5mm quiet zone on each side - prevents scan errors
                'write_text': False,     # We handle text separately below barcode
                'dpi': 300,              # Match label DPI for crisp bars
                'foreground': 'black',   # Pure black (#000000) for maximum contrast
            }
            
            code128 = barcode.get("code128", barcode_value, writer=writer)
            
            with tempfile.TemporaryDirectory() as tmpdir:
                barcode_path = os.path.join(tmpdir, f"bar_{label_counter}")
                code128.save(barcode_path, options=writer_options)
                
                # Load barcode at its optimal generated size (no arbitrary resize)
                barcode_img = Image.open(barcode_path + ".png")
                return barcode_img.copy()
                
        except Exception as e:
            print(f"Barcode generation failed for SKU {sku_code}: {e}")
            return None
    
    def generate_labels(self):
        """Generate labels from Excel data"""
        try:
            self.generate_btn.config(state="disabled")
            self.status_label.config(text="Loading data...", fg="#2196F3")
            self.root.update()
            
            # ===== COLUMN MAPPING (Source of Truth) =====
            column_mapping = {
                # Product Name
                'VENDOR ARTICLE NAME': 'VENDOR_ARTICLE_NAME',
                'VENDOR_ARTICLE_NAME': 'VENDOR_ARTICLE_NAME',
                'PRODUCT NAME': 'VENDOR_ARTICLE_NAME',
                'PRODUCT_NAME': 'VENDOR_ARTICLE_NAME',
                
                # Size
                'SIZE': 'SIZE',
                'SIZES': 'SIZE',
                
                # Brand
                'BRAND NAME': 'BRAND_NAME',
                'BRAND_NAME': 'BRAND_NAME',
                'BRAND': 'BRAND_NAME',
                
                # Style Code (TEXT ONLY - NOT for barcode)
                'VENDOR ARTICLE NO': 'VENDOR_ARTICLE_NO',
                'VENDOR_ARTICLE_NO': 'VENDOR_ARTICLE_NO',
                'VENDOR ARTICLE NUMBER': 'VENDOR_ARTICLE_NO',
                'STYLE_CODE': 'VENDOR_ARTICLE_NO',
                'STYLE CODE': 'VENDOR_ARTICLE_NO',
                
                # SKU Code (FOR BARCODE GENERATION)
                'SKU CODE': 'SKU_CODE',
                'SKU_CODE': 'SKU_CODE',
                'SKU': 'SKU_CODE',
                'SKUCODE': 'SKU_CODE',
                
                # MRP
                'MRP': 'MRP',
                'PRICE': 'MRP',
                'RETAIL_PRICE': 'MRP',
                'RETAIL PRICE': 'MRP',
                
                # Quantity
                'QUANTITY': 'QUANTITY',
                'QTY': 'QUANTITY',
                'QUAN': 'QUANTITY',
            }
            
            # Try to load and parse Excel file
            df = None
            header_row = None
            
            for skip_rows in range(0, 20):
                try:
                    if self.excel_path.endswith('.csv'):
                        temp_df = pd.read_csv(self.excel_path, skiprows=skip_rows)
                    else:
                        temp_df = pd.read_excel(self.excel_path, skiprows=skip_rows)
                    
                    # Normalize column names
                    temp_df.columns = temp_df.columns.astype(str).str.strip().str.upper()
                    
                    # Check for required columns
                    found_cols = [col for col in column_mapping.keys() if col in temp_df.columns]
                    
                    if len(set([column_mapping[col] for col in found_cols])) >= 4:
                        df = temp_df
                        header_row = skip_rows
                        break
                except:
                    continue
            
            if df is None:
                messagebox.showerror(
                    "Error",
                    "Could not find required columns.\n\n"
                    "Required columns:\n"
                    "- Vendor Article Name / Product Name\n"
                    "- Size\n"
                    "- Brand Name\n"
                    "- Vendor Article No (Style Code)\n"
                    "- SKU Code (for barcode)\n"
                    "- MRP / Price"
                )
                self.status_label.config(text="Error: Missing columns", fg="#f44336")
                self.generate_btn.config(state="normal")
                return
            
            # Rename columns
            df = df.rename(columns=column_mapping)
            
            # Validate essential columns
            required_cols = ['VENDOR_ARTICLE_NAME', 'SIZE', 'BRAND_NAME', 'VENDOR_ARTICLE_NO', 'SKU_CODE', 'MRP']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                messagebox.showerror(
                    "Error",
                    f"Missing required columns: {', '.join(missing_cols)}"
                )
                self.status_label.config(text="Error: Missing columns", fg="#f44336")
                self.generate_btn.config(state="normal")
                return
            
            self.status_label.config(
                text=f"Found headers at row {header_row + 1}. Processing...",
                fg="#2196F3"
            )
            self.root.update()
            
            # Create output directory
            os.makedirs(self.output_dir, exist_ok=True)
            
            # Get current date for manufacture date
            current_date = datetime.now()
            manufacture_date = f"{current_date.month:02d}/{current_date.year}"
            
            # Static text with dynamic date
            STATIC_TEXT = [
                f"Month/Year of Manufacture: {manufacture_date}",
                "Marketed by : Anthrilo Design House,",
                "KH400/414 Rahon Road, Punjab 141007,",
                "customercare@anthrilo.com, +919888264040"
            ]
            
            # Label dimensions (5cm x 9.5cm at 300 DPI)
            LABEL_WIDTH = 590      # 5cm at 300 DPI
            LABEL_HEIGHT = 1122    # 9.5cm at 300 DPI
            MARGIN = 28  # Larger margin for clear edge safety
            
            # Calculate total labels
            total_labels = len(df)
            self.progress["maximum"] = total_labels
            
            # Load fonts
            try:
                font_header = ImageFont.truetype("arial.ttf", 52)
                font_text = ImageFont.truetype("arial.ttf", 42)
                font_small = ImageFont.truetype("arial.ttf", 32)
            except:
                try:
                    font_header = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 52)
                    font_text = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 42)
                    font_small = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 32)
                except:
                    font_header = ImageFont.load_default()
                    font_text = ImageFont.load_default()
                    font_small = ImageFont.load_default()
            
            # Helper functions
            def wrap_text(text, font, max_width, draw_obj):
                """Wrap text to fit width"""
                words = str(text).split()
                lines = []
                current_line = []
                
                for word in words:
                    test_line = ' '.join(current_line + [word])
                    bbox = draw_obj.textbbox((0, 0), test_line, font=font)
                    width = bbox[2] - bbox[0]
                    
                    if width <= max_width:
                        current_line.append(word)
                    else:
                        if current_line:
                            lines.append(' '.join(current_line))
                        current_line = [word]
                
                if current_line:
                    lines.append(' '.join(current_line))
                
                return lines
            
            def get_safe_value(row, key, default=""):
                """Safely extract value from row"""
                val = row.get(key, default)
                return str(val) if pd.notna(val) else default
            
            # Generate labels
            all_labels = []
            label_counter = 0
            rows_skipped = 0
            
            for i, row in df.iterrows():
                # Extract values
                product_name = get_safe_value(row, 'VENDOR_ARTICLE_NAME', 'N/A')
                size = get_safe_value(row, 'SIZE', 'N/A')
                brand = get_safe_value(row, 'BRAND_NAME', 'N/A')
                style_code = get_safe_value(row, 'VENDOR_ARTICLE_NO', 'N/A')
                sku_code = get_safe_value(row, 'SKU_CODE', '')
                mrp = get_safe_value(row, 'MRP', '0')
                qty_str = get_safe_value(row, 'QUANTITY', '1')
                
                # Convert quantity to integer
                try:
                    quantity = int(float(qty_str))
                    quantity = max(1, quantity)
                except:
                    quantity = 1
                
                # CRITICAL: Check for SKU Code
                if not sku_code or sku_code.lower() == 'nan' or sku_code.strip() == '':
                    print(f"Row {i+1}: Skipping - SKU Code is missing")
                    rows_skipped += 1
                    continue
                
                # Generate labels for each quantity
                for qty_num in range(quantity):
                    label_counter += 1
                    self.status_label.config(
                        text=f"Generating label {label_counter}...",
                        fg="#2196F3"
                    )
                    self.progress["value"] = label_counter
                    self.root.update()
                    
                    # Create label image
                    img = Image.new("RGB", (LABEL_WIDTH, LABEL_HEIGHT), "white")
                    draw = ImageDraw.Draw(img)
                
                    current_y = MARGIN
                    
                    # ===== SECTION 1: Product Name + Size =====
                    header_text = f"{product_name} - {size}"
                    header_lines = wrap_text(header_text, font_header, LABEL_WIDTH - 2*MARGIN, draw)
                    
                    for line in header_lines:
                        bbox = draw.textbbox((0, 0), line, font=font_header)
                        text_width = bbox[2] - bbox[0]
                        x = (LABEL_WIDTH - text_width) // 2
                        draw.text((x, current_y), line, font=font_header, fill="black")
                        current_y += 54
                    
                    current_y += 2
                
                    # ===== SECTION 2: Brand, Style Code, Size =====
                    dynamic_fields = [
                        f"Brand : {brand}",
                        f"Style code : {style_code}",
                        f"Size : {size}"
                    ]
                    
                    for field in dynamic_fields:
                        draw.text((MARGIN, current_y), field, font=font_text, fill="black")
                        current_y += 50
                    
                    current_y += 2
                
                    # ===== SECTION 3: Static Text =====
                    for line in STATIC_TEXT:
                        # Wrap each static text line to fit within margins
                        wrapped_lines = wrap_text(line, font_small, LABEL_WIDTH - 2*MARGIN, draw)
                        for wrapped_line in wrapped_lines:
                            draw.text((MARGIN, current_y), wrapped_line, font=font_small, fill="black")
                            current_y += 38
                    
                    current_y += 2
                    
                    # ===== SECTION 4: MRP =====
                    price_text = f"MRP: {mrp} (inclusive of all taxes)"
                    # Wrap MRP text to fit within margins
                    wrapped_mrp = wrap_text(price_text, font_text, LABEL_WIDTH - 2*MARGIN, draw)
                    for wrapped_line in wrapped_mrp:
                        draw.text((MARGIN, current_y), wrapped_line, font=font_text, fill="black")
                        current_y += 50
                    
                    # ===== SECTION 5: Barcode (OPTIMIZED FOR LASER SCANNING) =====
                    current_y += 25  # LARGE gap for strong visual separation from static text
                    barcode_img = self.generate_barcode(sku_code, label_counter)
                    
                    if barcode_img:
                        # Reserve space for SKU text below barcode (40px text + 10px gap + 28px bottom margin = 78px)
                        bottom_reservation = 78
                        remaining_height = LABEL_HEIGHT - current_y - bottom_reservation
                        
                        if remaining_height > 80:
                            # WIDE barcode for laser scanning: 100% height, 140% width
                            # Wider bars = more reliable scanning (most important fix)
                            new_height = min(remaining_height, barcode_img.height)
                            new_height = int(new_height * 1.0)  # 100% of calculated height
                            scale = new_height / barcode_img.height
                            new_width = int(barcode_img.width * scale * 1.40)  # 140% width for thick bars
                            
                            # Ensure barcode fits within label margins
                            max_barcode_width = LABEL_WIDTH - (MARGIN * 2)
                            if new_width > max_barcode_width:
                                # Scale down proportionally if too wide
                                scale_down = max_barcode_width / new_width
                                new_width = max_barcode_width
                                new_height = int(new_height * scale_down)
                            
                            barcode_img = barcode_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        
                        barcode_x = (LABEL_WIDTH - barcode_img.width) // 2
                        img.paste(barcode_img, (barcode_x, current_y))
                        current_y += barcode_img.height
                        
                        # Add SKU text below barcode with gap
                        current_y += 10  # Gap between barcode and text
                        sku_text = f"{sku_code}"
                        bbox = draw.textbbox((0, 0), sku_text, font=font_small)
                        text_width = bbox[2] - bbox[0]
                        text_x = (LABEL_WIDTH - text_width) // 2
                        draw.text((text_x, current_y), sku_text, font=font_small, fill="black")
                    
                    # Add to list
                    all_labels.append(img.copy())
                    
                    # Memory optimization for low-end PCs: periodically force garbage collection
                    if label_counter % 100 == 0:
                        import gc
                        gc.collect()  # Force garbage collection every 100 labels
            
            # Save to PDF
            if all_labels:
                pdf_path = os.path.join(self.output_dir, "labels.pdf")
                
                # For very large label sets (2000+), save in chunks to prevent crashes
                if len(all_labels) > 1000:
                    self.status_label.config(text=f"Saving {len(all_labels)} labels to PDF (this may take a moment)...", fg="#2196F3")
                    self.root.update()
                
                all_labels[0].save(
                    pdf_path,
                    save_all=True,
                    append_images=all_labels[1:],
                    resolution=300.0
                )
                
                msg = f"Successfully generated {label_counter} labels!"
                if rows_skipped > 0:
                    msg += f"\n(Skipped {rows_skipped} rows with missing SKU Code)"
                msg += f"\n\nSaved to: {pdf_path}"
                
                self.status_label.config(
                    text=f"Success! Generated {label_counter} labels",
                    fg="#4CAF50"
                )
                messagebox.showinfo("Success", msg)
            else:
                messagebox.showerror(
                    "Error",
                    f"No labels generated.\nAll {rows_skipped} rows were skipped due to missing SKU Code."
                )
                self.status_label.config(text="Error: No valid labels", fg="#f44336")
        
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="#f44336")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
            import traceback
            traceback.print_exc()
        
        finally:
            self.generate_btn.config(state="normal")
            self.progress["value"] = 0


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = LabelGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
