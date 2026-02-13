"""
Anthrilo Barcode Label Generator - Core Script
Generates barcode labels from Excel data
Label Size: 5cm x 9.5cm (590 x 1122 pixels at 300 DPI)
Barcode Source: SKU Code (NOT Style Code)
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os
import tempfile
import logging
import sys
from datetime import datetime

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s'
)
logger = logging.getLogger(__name__)


class AnthriloLabelGenerator:
    """Generate barcode labels for Anthrilo Design House"""
    
    # Label dimensions (5cm x 9.5cm at 300 DPI)
    LABEL_WIDTH = 590
    LABEL_HEIGHT = 1122
    MARGIN = 28  # Larger margin for clear edge safety
    
    # Static content
    STATIC_TEXT = [
        "Month/Year of Manufacture: 10/2025",
        "Marketed by : Anthrilo Design House,",
        "KH400/414 Rahon Road, Punjab 141007,",
        "customercare@anthrilo.com, +919888264040"
    ]
    
    # Column mapping
    COLUMN_MAPPING = {
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
    }
    
    def __init__(self, excel_path, output_dir="generated_labels"):
        """
        Initialize label generator
        
        Args:
            excel_path (str): Path to Excel/CSV file
            output_dir (str): Output directory for PDF
        """
        self.excel_path = excel_path
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        
        # Load fonts
        self.fonts = self._load_fonts()
        
        logger.info(f"Initialized generator with file: {excel_path}")
    
    def _load_fonts(self):
        """Load fonts with fallback support"""
        fonts = {}
        font_sizes = {
            'header': 52,
            'text': 42,
            'small': 32
        }
        
        font_paths = [
            "arial.ttf",
            "C:\\Windows\\Fonts\\arial.ttf"
        ]
        
        for name, size in font_sizes.items():
            fonts[name] = ImageFont.load_default()
            
            for path in font_paths:
                try:
                    fonts[name] = ImageFont.truetype(path, size)
                    break
                except:
                    continue
        
        return fonts
    
    def _get_safe_value(self, row, key, default=""):
        """Safely extract value from DataFrame row"""
        val = row.get(key, default)
        return str(val) if pd.notna(val) else default
    
    def _wrap_text(self, text, font, max_width, draw_obj):
        """Wrap text to fit within max_width"""
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
    
    def _generate_barcode(self, sku_code, label_counter):
        """
        Generate barcode from SKU Code with professional laser scanning settings
        
        Key parameters for reliable laser scanning:
        - module_width (X-dimension): 0.40mm - controls bar thickness
        - module_height: 20mm - adequate height for laser sweep
        - quiet_zone: 5mm - clear margins for scanner detection
        - dpi: 300 - matches label resolution
        
        Args:
            sku_code (str): SKU Code for barcode
            label_counter (int): Label number for naming
            
        Returns:
            PIL.Image or None: Barcode image
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
            logger.error(f"Barcode generation failed for SKU {sku_code}: {e}")
            return None
    
    def _load_and_validate_data(self):
        """
        Load and validate Excel data
        
        Returns:
            pd.DataFrame or None: Valid DataFrame, or None if validation fails
        """
        logger.info(f"Loading data from {self.excel_path}...")
        
        df = None
        header_row = None
        
        # Try to find header row
        for skip_rows in range(0, 20):
            try:
                if self.excel_path.lower().endswith('.csv'):
                    temp_df = pd.read_csv(self.excel_path, skiprows=skip_rows)
                else:
                    temp_df = pd.read_excel(self.excel_path, skiprows=skip_rows)
                
                # Normalize column names
                temp_df.columns = temp_df.columns.astype(str).str.strip().str.upper()
                
                # Check for required columns
                found_cols = [col for col in self.COLUMN_MAPPING.keys() if col in temp_df.columns]
                
                if len(set([self.COLUMN_MAPPING[col] for col in found_cols])) >= 4:
                    df = temp_df
                    header_row = skip_rows
                    break
            except Exception as e:
                continue
        
        if df is None:
            logger.error("Could not find required columns in file")
            return None
        
        logger.info(f"Found headers at row {header_row + 1}")
        
        # Rename columns
        df = df.rename(columns=self.COLUMN_MAPPING)
        
        # Validate essential columns
        required_cols = ['VENDOR_ARTICLE_NAME', 'SIZE', 'BRAND_NAME', 'VENDOR_ARTICLE_NO', 'SKU_CODE', 'MRP']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            logger.error(f"Missing required columns: {missing_cols}")
            return None
        
        logger.info(f"Loaded {len(df)} rows from Excel")
        return df
    
    def generate(self):
        """Generate all labels from Excel data with flow-based layout"""
        # Load and validate data
        df = self._load_and_validate_data()
        if df is None:
            logger.error("Failed to load and validate data")
            return False
        
        all_labels = []
        label_counter = 0
        rows_skipped = 0
        
        logger.info("Starting label generation...")
        
        for i, row in df.iterrows():
            # Extract values
            product_name = self._get_safe_value(row, 'VENDOR_ARTICLE_NAME', 'N/A')
            size = self._get_safe_value(row, 'SIZE', 'N/A')
            brand = self._get_safe_value(row, 'BRAND_NAME', 'N/A')
            style_code = self._get_safe_value(row, 'VENDOR_ARTICLE_NO', 'N/A')
            sku_code = self._get_safe_value(row, 'SKU_CODE', '')
            mrp = self._get_safe_value(row, 'MRP', '0')
            qty_str = self._get_safe_value(row, 'QUANTITY', '1')
            
            # Convert quantity to integer
            try:
                quantity = int(float(qty_str))
                quantity = max(1, quantity)
            except:
                quantity = 1
            
            # CRITICAL: Check for SKU Code
            if not sku_code or sku_code.lower() == 'nan' or sku_code.strip() == '':
                logger.warning(f"Row {i+1}: Skipping - SKU Code is missing")
                rows_skipped += 1
                continue
            
            # Generate labels for each quantity
            for qty_num in range(quantity):
                label_counter += 1
                
                # Create label image
                img = Image.new("RGB", (self.LABEL_WIDTH, self.LABEL_HEIGHT), "white")
            draw = ImageDraw.Draw(img)
            
            current_y = self.MARGIN
            
            # ===== SECTION 1: Product Name + Size (TOP) =====
            header_text = f"{product_name} - {size}"
            header_lines = self._wrap_text(
                header_text,
                self.fonts['header'],
                self.LABEL_WIDTH - 2*self.MARGIN,
                draw
            )
            
            for line in header_lines:
                bbox = draw.textbbox((0, 0), line, font=self.fonts['header'])
                text_width = bbox[2] - bbox[0]
                x = (self.LABEL_WIDTH - text_width) // 2
                draw.text((x, current_y), line, font=self.fonts['header'], fill="black")
                current_y += 54
            
            current_y += 2
            
            # ===== SECTION 2: Brand, Style Code, Size =====
            dynamic_fields = [
                f"Brand : {brand}",
                f"Style code : {style_code}",
                f"Size : {size}"
            ]
            
            for field in dynamic_fields:
                draw.text((self.MARGIN, current_y), field, font=self.fonts['text'], fill="black")
                current_y += 50
            
            current_y += 2
            
            # ===== SECTION 3: Static Text =====
            for line in self.STATIC_TEXT:
                # Wrap each static text line to fit within margins
                wrapped_lines = self._wrap_text(line, self.fonts['small'], self.LABEL_WIDTH - 2*self.MARGIN, draw)
                for wrapped_line in wrapped_lines:
                    draw.text((self.MARGIN, current_y), wrapped_line, font=self.fonts['small'], fill="black")
                    current_y += 38
            
            current_y += 2
            
            # ===== SECTION 4: MRP =====
            price_text = f"MRP: {mrp} (inclusive of all taxes)"
            # Wrap MRP text to fit within margins
            wrapped_mrp = self._wrap_text(price_text, self.fonts['text'], self.LABEL_WIDTH - 2*self.MARGIN, draw)
            for wrapped_line in wrapped_mrp:
                draw.text((self.MARGIN, current_y), wrapped_line, font=self.fonts['text'], fill="black")
                current_y += 50
            
            # ===== SECTION 5: Barcode (OPTIMIZED FOR LASER SCANNING) =====
            current_y += 25  # LARGE gap for strong visual separation from static text
            barcode_img = self._generate_barcode(sku_code, label_counter)
            
            if barcode_img:
                # Reserve space for SKU text below barcode (40px text + 10px gap + 28px bottom margin = 78px)
                bottom_reservation = 78
                remaining_height = self.LABEL_HEIGHT - current_y - bottom_reservation
                
                if remaining_height > 80:
                    # WIDE barcode for laser scanning: 100% height, 140% width
                    # Wider bars = more reliable scanning (most important fix)
                    new_height = min(remaining_height, barcode_img.height)
                    new_height = int(new_height * 1.0)  # 100% of calculated height
                    scale = new_height / barcode_img.height
                    new_width = int(barcode_img.width * scale * 1.40)  # 140% width for thick bars
                    
                    # Ensure barcode fits within label margins
                    max_barcode_width = self.LABEL_WIDTH - (self.MARGIN * 2)
                    if new_width > max_barcode_width:
                        # Scale down proportionally if too wide
                        scale_down = max_barcode_width / new_width
                        new_width = max_barcode_width
                        new_height = int(new_height * scale_down)
                    
                    barcode_img = barcode_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                barcode_x = (self.LABEL_WIDTH - barcode_img.width) // 2
                img.paste(barcode_img, (barcode_x, current_y))
                current_y += barcode_img.height
                
                # Add SKU text below barcode
                current_y += 10
                sku_text = f"{sku_code}"
                bbox = draw.textbbox((0, 0), sku_text, font=self.fonts['small'])
                text_width = bbox[2] - bbox[0]
                text_x = (self.LABEL_WIDTH - text_width) // 2
                draw.text((text_x, current_y), sku_text, font=self.fonts['small'], fill="black")
            
            # Add to list
            all_labels.append(img.copy())
            logger.info(f"Generated label {label_counter}: {product_name[:40]}... (SKU: {sku_code})")
            
            # Memory optimization: force garbage collection every 100 labels
            if label_counter % 100 == 0:
                import gc
                gc.collect()
        
        # Save to PDF
        if all_labels:
            pdf_path = os.path.join(self.output_dir, "labels.pdf")
            
            # For large label sets, log progress
            if len(all_labels) > 1000:
                logger.info(f"Saving {len(all_labels)} labels to PDF...")
            
            all_labels[0].save(
                pdf_path,
                save_all=True,
                append_images=all_labels[1:],
                resolution=300.0
            )
            
            logger.info(f"[OK] Successfully generated {label_counter} labels")
            logger.info(f"[OK] Output file: {pdf_path}")
            
            if rows_skipped > 0:
                logger.warning(f"[WARNING] Skipped {rows_skipped} rows with missing SKU Code")
            
            return True
        else:
            logger.error(f"No labels generated - all {rows_skipped} rows were skipped")
            return False


def main():
    """Command-line entrypoint"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python label_generator.py <excel_file> [output_dir]")
        print("Example: python label_generator.py data.xlsx ./labels")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "generated_labels"
    
    if not os.path.exists(excel_file):
        logger.error(f"File not found: {excel_file}")
        sys.exit(1)
    
    generator = AnthriloLabelGenerator(excel_file, output_dir)
    success = generator.generate()
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
