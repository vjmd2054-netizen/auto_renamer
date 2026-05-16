# Working properly

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from collections import Counter
import concurrent.futures
import threading
import sys


# **PORTABLE TESSERACT PATH SETUP**
def get_tesseract_path():
    """Get the correct Tesseract path whether running as script or executable"""
    if getattr(sys, 'frozen', False):
        # Running as executable (PyInstaller bundle)
        base_path = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable)
        # Check for Tesseract in the same folder as executable
        exe_dir = os.path.dirname(sys.executable)
        tesseract_paths = [
            os.path.join(exe_dir, 'Tesseract', 'tesseract.exe'),
            os.path.join(exe_dir, 'tesseract.exe'),
            os.path.join(base_path, 'Tesseract', 'tesseract.exe'),
        ]

        for path in tesseract_paths:
            if os.path.exists(path):
                return path

        # If not found, try to find it in PATH or common locations
        for common_path in [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        ]:
            if os.path.exists(common_path):
                return common_path

        # Last resort: Check system PATH
        import shutil
        tesseract_in_path = shutil.which('../PDF_Renamer_Portable/Tesseract')
        if tesseract_in_path:
            return tesseract_in_path

        return None
    else:
        return r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Set Tesseract path
tesseract_path = get_tesseract_path()
if tesseract_path and os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
else:
    print("Warning: Tesseract not found. OCR functionality will not work.")


class PDFRenamerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF AWBs Renamer")
        self.root.geometry("700x850")

        # SET WINDOW ICON
        try:
            self.root.iconbitmap("folder_purple.ico")
        except:
            try:
                # Try alternative icon paths for portable version
                icon_path = os.path.join(os.path.dirname(sys.executable), "folder_purple.ico")
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
            except:
                pass

        # Variables
        self.folder_path = tk.StringVar()
        self.excel_file_path = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.page_option = tk.StringVar(value="first")  # Default to first page
        self.append_to_excel = tk.BooleanVar(value=True)
        self.is_processing = False
        self.stop_processing = False

        self.setup_gui()

    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(main_frame, text="PDF AWBs Renamer",
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))

        # Folder selection
        ttk.Label(main_frame, text="Select PDF Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)

        folder_entry = ttk.Entry(main_frame, textvariable=self.folder_path, width=50)
        folder_entry.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)

        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_folder)
        browse_btn.grid(row=2, column=1, padx=5, pady=5)

        # Excel file selection
        ttk.Label(main_frame, text="Excel File (Optional):").grid(row=3, column=0, sticky=tk.W, pady=(10, 5))

        excel_frame = ttk.Frame(main_frame)
        excel_frame.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_path, width=45)
        excel_entry.grid(row=0, column=0, sticky=tk.W)

        excel_browse_btn = ttk.Button(excel_frame, text="Browse", command=self.browse_excel_file)
        excel_browse_btn.grid(row=0, column=1, padx=5)

        # Append to Excel checkbox
        self.append_checkbox = ttk.Checkbutton(excel_frame, text="Append to this Excel file",
                                               variable=self.append_to_excel, state="disabled",
                                               command=self.update_excel_options)
        self.append_checkbox.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))

        # Auto-generate note
        auto_note = ttk.Label(main_frame,
                              text="Note: If no Excel file is selected, one will be auto-generated in the PDF folder",
                              font=("Arial", 9), foreground="blue")
        auto_note.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(5, 10))

        # Page selection option - 3 OPTIONS
        ttk.Label(main_frame, text="Select Page to Check:").grid(row=6, column=0, sticky=tk.W, pady=(10, 5))

        page_frame = ttk.Frame(main_frame)
        page_frame.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Page selection radio buttons - 3 OPTIONS NOW
        ttk.Radiobutton(page_frame, text="First Page Only", variable=self.page_option,
                        value="first").grid(row=0, column=0, padx=10, sticky=tk.W)
        ttk.Radiobutton(page_frame, text="Last Page Only", variable=self.page_option,
                        value="last").grid(row=0, column=1, padx=10, sticky=tk.W)
        ttk.Radiobutton(page_frame, text="Second-to-Last Page", variable=self.page_option,
                        value="second_last").grid(row=0, column=2, padx=10, sticky=tk.W)

        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress, maximum=100)
        self.progress_bar.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=9, column=0, columnspan=2, pady=10)

        self.process_btn = ttk.Button(button_frame, text="Process PDFs", command=self.process_pdfs)
        self.process_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_processing_func, state="disabled")
        self.stop_btn.grid(row=0, column=1, padx=5)

        # Text area for output
        ttk.Label(main_frame, text="Output:").grid(row=10, column=0, sticky=tk.W, pady=(10, 0))

        self.output_text = tk.Text(main_frame, height=15, width=70)
        self.output_text.grid(row=11, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.output_text.yview)
        scrollbar.grid(row=11, column=2, sticky=(tk.N, tk.S), pady=5)
        self.output_text.configure(yscrollcommand=scrollbar.set)

        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(11, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def is_excel_file_locked(self, filepath):
        """
        Check if an Excel file is open/locked by another process.
        Returns True if file is locked/open, False if available.
        """
        if not os.path.exists(filepath):
            return False  # File doesn't exist, so not locked

        try:
            # Method 1: Try to open the file in exclusive mode
            with open(filepath, 'a', encoding='utf-8') as f:
                pass
            return False  # File is not locked
        except (IOError, PermissionError) as e:
            # File is locked or inaccessible
            error_msg = str(e).lower()
            if any(keyword in error_msg for keyword in ['permission denied', 'being used', 'locked', 'access denied']):
                return True
            return True  # Assume locked on any error
        except Exception:
            return False  # Other errors, assume not locked

    def update_excel_options(self):
        """Enable/disable append checkbox based on Excel file selection"""
        excel_path = self.excel_file_path.get()

        if excel_path:
            if os.path.exists(excel_path):
                # Check if file is not locked before enabling append
                if not self.is_excel_file_locked(excel_path):
                    self.append_checkbox.config(state="normal")
                else:
                    self.append_checkbox.config(state="disabled")
                    self.append_to_excel.set(False)
                    messagebox.showwarning(
                        "Excel File Open",
                        f"Cannot append to '{os.path.basename(excel_path)}' because it's open in another program.\n"
                        f"Please close the file to enable append mode."
                    )
            else:
                # File doesn't exist, disable append
                self.append_checkbox.config(state="disabled")
                self.append_to_excel.set(False)
        else:
            self.append_checkbox.config(state="disabled")
            self.append_to_excel.set(False)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log_output(f"Selected folder: {folder_selected}")

    def browse_excel_file(self):
        file_selected = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_selected:
            self.excel_file_path.set(file_selected)
            self.log_output(f"Selected Excel file: {os.path.basename(file_selected)}")
            self.update_excel_options()

    def log_output(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()

    def extract_12_digit_number_not_near_tin(self, text):
        """
        Extract 12-digit numbers that are NOT near TIN text - IMPROVED VERSION
        """
        # First, find all 12-digit numbers
        all_numbers = re.findall(r'\b\d{12}\b', text)

        if not all_numbers:
            return []

        # Filter out numbers that end with "000"
        filtered_numbers = []
        for number in all_numbers:
            # Skip numbers ending with "000"
            if number.endswith('000'):
                continue
            filtered_numbers.append(number)

        if not filtered_numbers:
            return []

        # Find numbers that are near TIN (within 100 characters or same/next line)
        numbers_near_tin = []

        # Pattern to find TIN followed by number within 100 chars
        tin_near_pattern = r'TIN.{0,100}?(\d{12})'
        for match in re.finditer(tin_near_pattern, text, re.IGNORECASE):
            if match.group(1):
                numbers_near_tin.append(match.group(1))

        # Pattern to find number followed by TIN within 100 chars
        near_tin_pattern = r'(\d{12}).{0,100}?TIN'
        for match in re.finditer(near_tin_pattern, text, re.IGNORECASE):
            if match.group(1):
                numbers_near_tin.append(match.group(1))

        # Also check line-by-line for TIN on same line or next line
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if re.search(r'TIN', line, re.IGNORECASE):
                # Check current line for numbers
                line_numbers = re.findall(r'\b\d{12}\b', line)
                numbers_near_tin.extend(line_numbers)

                # Check next line if exists
                if i + 1 < len(lines):
                    next_line_numbers = re.findall(r'\b\d{12}\b', lines[i + 1])
                    numbers_near_tin.extend(next_line_numbers)

        # Remove duplicates
        numbers_near_tin = list(set(numbers_near_tin))

        # Return numbers that are NOT near TIN and don't end with "000"
        numbers_not_near_tin = [num for num in filtered_numbers if num not in numbers_near_tin]

        return numbers_not_near_tin

    def extract_valid_12_digit_numbers(self, text):
        """
        Extract valid 12-digit numbers with multiple filters:
        1. Must be exactly 12 digits
        2. Cannot end with "000"
        3. Cannot be near TIN text
        4. Additional validation rules
        """
        # First pass: Get all 12-digit numbers
        all_12_digit = re.findall(r'\b\d{12}\b', text)

        if not all_12_digit:
            return []

        # Apply filters
        valid_numbers = []
        rejected_numbers = []

        for number in all_12_digit:
            # FILTER 1: Skip numbers ending with "000"
            if number.endswith('000'):
                rejected_numbers.append((number, "ends with 000"))
                continue

            # FILTER 2: Check for suspicious patterns
            # Check if all digits are the same (e.g., 111111111111)
            if len(set(number)) == 1:
                rejected_numbers.append((number, "all digits same"))
                continue

            # FILTER 3: Check for sequential patterns (e.g., 123456789012)
            if number in ['123456789012', '012345678901']:
                rejected_numbers.append((number, "sequential pattern"))
                continue

            # If passed all filters, add to valid numbers
            valid_numbers.append(number)

        # FILTER 4: Remove numbers near TIN
        if valid_numbers:
            numbers_near_tin = []

            # Find TIN references
            lines = text.split('\n')
            for i, line in enumerate(lines):
                if re.search(r'\bTIN\b', line, re.IGNORECASE):
                    # Check current line for numbers
                    line_numbers = re.findall(r'\b\d{12}\b', line)
                    numbers_near_tin.extend(line_numbers)

                    # Check adjacent lines (previous and next)
                    if i > 0:
                        prev_line_numbers = re.findall(r'\b\d{12}\b', lines[i - 1])
                        numbers_near_tin.extend(prev_line_numbers)
                    if i + 1 < len(lines):
                        next_line_numbers = re.findall(r'\b\d{12}\b', lines[i + 1])
                        numbers_near_tin.extend(next_line_numbers)

            # Remove duplicates
            numbers_near_tin = list(set(numbers_near_tin))

            # Filter out numbers near TIN
            final_numbers = [num for num in valid_numbers if num not in numbers_near_tin]

            # Log rejection reasons if debugging
            if rejected_numbers and len(final_numbers) == 0:
                # Only log if we filtered everything out
                for num, reason in rejected_numbers:
                    self.log_output(f"  Filtered out: {num} ({reason})")

            return final_numbers

        return []

    def get_most_frequent_number(self, numbers):
        """Find the most frequent number from a list"""
        if not numbers:
            return None
        counter = Counter(numbers)
        return counter.most_common(1)[0][0]

    def create_excel_record(self, records, output_folder, append_to_existing=False):
        """Create or append to Excel file with renaming records"""
        if not records:
            return None

        excel_path = self.excel_file_path.get() if self.excel_file_path.get() else None

        try:
            if append_to_existing and excel_path and os.path.exists(excel_path):
                # DOUBLE-CHECK FILE IS NOT LOCKED BEFORE READING
                if self.is_excel_file_locked(excel_path):
                    raise Exception(
                        f"Excel file is open in another program. Please close '{os.path.basename(excel_path)}' and try again.")

                # Read existing data
                existing_df = pd.read_excel(excel_path)
                new_df = pd.DataFrame(records)
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)

                # Save back to the same file
                combined_df.to_excel(excel_path, index=False)

                self.log_output(f"Appended to existing Excel file: {os.path.basename(excel_path)}")
                self.log_output(f"Added {len(records)} new records (Total: {len(combined_df)} records)")

                return excel_path

            else:
                # Create new Excel file
                if excel_path and self.append_to_excel.get():
                    # Use selected path for new file
                    save_path = excel_path
                else:
                    # ALWAYS CREATE NEW FILE IN PDF FOLDER WHEN NO EXCEL SELECTED
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    excel_filename = f"PDF_Renaming_Record_{timestamp}.xlsx"
                    save_path = os.path.join(output_folder, excel_filename)

                # Check if file already exists and is locked
                if os.path.exists(save_path) and self.is_excel_file_locked(save_path):
                    # Generate alternative filename with milliseconds
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                    excel_filename = f"PDF_Renaming_Record_{timestamp}.xlsx"
                    save_path = os.path.join(output_folder, excel_filename)

                df = pd.DataFrame(records)
                # Add more detailed columns
                if 'Processing Date' not in df.columns:
                    df['Processing Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                df.to_excel(save_path, index=False)

                if excel_path and self.append_to_excel.get():
                    self.log_output(f"Created new Excel file: {os.path.basename(save_path)}")
                else:
                    self.log_output(f"Excel record saved: {os.path.basename(save_path)}")
                    self.log_output(f"Location: {save_path}")

                return save_path

        except PermissionError as e:
            self.log_output(f"❌ Permission error saving Excel: {e}")
            self.log_output(f"   The file might be open in Excel. Please close it and try again.")

            # Try to save with a different name
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                alt_filename = f"PDF_Renaming_Record_{timestamp}_BACKUP.xlsx"
                alt_path = os.path.join(output_folder, alt_filename)

                df = pd.DataFrame(records)
                df['Processing Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.to_excel(alt_path, index=False)

                self.log_output(f"💾 Saved backup Excel file: {os.path.basename(alt_path)}")
                self.log_output(f"   Backup location: {alt_path}")
                return alt_path
            except Exception as backup_error:
                self.log_output(f"❌ Failed to save backup Excel: {backup_error}")
                return None

        except Exception as e:
            self.log_output(f"❌ Error saving Excel file: {e}")
            return None

    def process_pages_for_ocr(self, doc):
        """Get the pages to process based on user selection"""
        total_pages = len(doc)
        page_option = self.page_option.get()

        if page_option == "first":
            return [0]  # First page only
        elif page_option == "last":
            return [total_pages - 1] if total_pages > 0 else []  # Last page only
        elif page_option == "second_last":
            # Check second-to-last page, but only if there are at least 2 pages
            if total_pages >= 2:
                return [total_pages - 2]  # Second-to-last page
            elif total_pages == 1:
                # If only 1 page, use that page
                return [0]
            else:
                return []  # Empty PDF
        else:
            return [0]  # Default to first page

    def extract_numbers_from_page(self, doc, page_num):
        """Extract numbers from a single page"""
        page = doc[page_num]
        images = page.get_images(full=True)

        all_numbers = []
        found_tin = False

        for img_index, img in enumerate(images):
            if self.stop_processing:
                break

            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]

            image = Image.open(io.BytesIO(image_bytes))

            # Use 2 OCR configurations for balance of speed and accuracy
            try:
                # Configuration 1: Optimized for numbers
                text1 = pytesseract.image_to_string(
                    image,
                    config='--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ/: '
                )

                # Configuration 2: Default for fallback
                text2 = pytesseract.image_to_string(image, config='--psm 6')

                # Use the improved extraction method
                numbers1 = self.extract_valid_12_digit_numbers(text1)
                numbers2 = self.extract_valid_12_digit_numbers(text2)

                combined_numbers = numbers1 + numbers2

                if combined_numbers:
                    all_numbers.extend(combined_numbers)

                    # Check if TIN was detected for logging
                    if not found_tin and (
                            re.search(r'\bTIN\b', text1, re.IGNORECASE) or re.search(r'\bTIN\b', text2, re.IGNORECASE)):
                        found_tin = True

            except Exception as e:
                continue

        return all_numbers, found_tin

    def process_single_pdf_optimized(self, pdf_info):
        """Process single PDF with page selection option"""
        pdf_path, pdf_file = pdf_info

        if self.stop_processing:
            return None

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)

            # Log PDF info
            if total_pages == 0:
                doc.close()
                self.log_output(f"{pdf_file}: Empty PDF (0 pages)")
                return None

            pages_to_check = self.process_pages_for_ocr(doc)

            if not pages_to_check:
                doc.close()
                self.log_output(f"{pdf_file}: No valid pages to check")
                return None

            all_numbers = []
            found_tin = False

            for page_index in pages_to_check:
                if self.stop_processing:
                    break

                page_numbers, page_found_tin = self.extract_numbers_from_page(doc, page_index)
                all_numbers.extend(page_numbers)
                found_tin = found_tin or page_found_tin

            doc.close()

            # Find most frequent number
            if all_numbers and not self.stop_processing:
                most_frequent = self.get_most_frequent_number(all_numbers)

                # Additional validation: Check if number ends with "000" (should be filtered already)
                if most_frequent and most_frequent.endswith('000'):
                    self.log_output(f"⚠️ Warning: Selected number {most_frequent} ends with '000' but was chosen")
                    # You could choose to skip this file here
                    # return None

                # Replace filename with just the number
                directory = os.path.dirname(pdf_path)
                new_filename = f"{most_frequent}.pdf"
                new_path = os.path.join(directory, new_filename)

                # Check if file already exists with this name
                if os.path.exists(new_path):
                    # Add timestamp to avoid overwriting
                    timestamp = datetime.now().strftime("%H%M%S")
                    new_filename = f"{most_frequent}_{timestamp}.pdf"
                    new_path = os.path.join(directory, new_filename)

                os.rename(pdf_path, new_path)

                result = {
                    'Original Name': pdf_file,
                    'New Name': most_frequent,
                    'New Filename': new_filename,
                    'Page Checked': pages_to_check[0] + 1,  # 1-based for display
                    'Total Pages': total_pages,
                    'Page Option': self.page_option.get(),
                    'Processing Time': datetime.now().strftime("%H:%M:%S"),
                    'Found TIN': 'Yes' if found_tin else 'No',
                    'Ends with 000': 'Yes' if most_frequent.endswith('000') else 'No'
                }

                # Log the result with page info
                page_info = self.get_page_info_text(result['Page Option'], result['Page Checked'], total_pages)

                self.log_output(f"✓ {pdf_file} → {most_frequent}.pdf{page_info}")
                if found_tin:
                    self.log_output(f"  TIN text was filtered out")
                if most_frequent.endswith('000'):
                    self.log_output(f"⚠️  Note: Number ends with '000'")

                return result
            else:
                # Log why no numbers were found
                if not all_numbers:
                    page_info = self.get_page_info_text(self.page_option.get(), pages_to_check[0] + 1, total_pages)
                    self.log_output(f"❌ {pdf_file}: No valid 12-digit numbers found{page_info}")
                    self.log_output(f"   (Numbers ending with '000' are filtered out)")
                    if found_tin:
                        self.log_output(f"   Numbers may have been filtered due to TIN proximity")

        except Exception as e:
            self.log_output(f"❌ {pdf_file}: Error - {str(e)}")

        return None

    def get_page_info_text(self, page_option, page_checked, total_pages):
        """Generate text describing which page was checked"""
        if page_option == "first":
            return " (first page)"
        elif page_option == "last":
            return " (last page)"
        elif page_option == "second_last":
            if total_pages >= 2:
                return " (second-to-last page)"
            else:
                return " (only page)"  # For single-page PDFs
        return ""

    def process_pdfs_parallel(self, folder_path):
        """Process PDFs using parallel processing"""
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

        if not pdf_files:
            return []

        page_option = self.page_option.get()
        page_info = self.get_page_info_text(page_option, 1, 0)

        self.log_output(f"🔄 Processing {len(pdf_files)} PDF files{page_info} in parallel...")
        self.log_output(f"   Filtering: Numbers ending with '000' will be ignored")

        # Prepare file list (skip already processed files)
        pdf_info_list = []
        for f in pdf_files:
            # Skip files that are already named with 12 digits
            if not re.match(r'^\d{12}\.pdf$', f):
                pdf_info_list.append((os.path.join(folder_path, f), f))

        if not pdf_info_list:
            self.log_output("ℹ All files already appear to be processed")
            return []

        records = []
        completed = 0

        # Use ThreadPoolExecutor for parallel processing
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, len(pdf_info_list))) as executor:
            future_to_pdf = {executor.submit(self.process_single_pdf_optimized, pdf_info): pdf_info for pdf_info in
                             pdf_info_list}

            for future in concurrent.futures.as_completed(future_to_pdf):
                if self.stop_processing:
                    break

                result = future.result()
                if result:
                    records.append(result)

                completed += 1
                progress = (completed / len(pdf_info_list)) * 100
                self.progress.set(progress)

        return records

    def stop_processing_func(self):
        """Stop the current processing"""
        self.stop_processing = True
        self.process_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.log_output("🛑 Processing stopped by user")

    def process_pdfs(self):
        """Main processing function with Excel file lock checking"""
        if self.is_processing:
            return

        folder_path = self.folder_path.get()

        if not folder_path:
            messagebox.showerror("Error", "Please select a PDF folder first!")
            return

        if not os.path.exists(folder_path):
            messagebox.showerror("Error", "Selected folder does not exist!")
            return

        # CHECK IF EXCEL FILE IS OPEN/LOCKED BEFORE STARTING
        excel_path = self.excel_file_path.get()
        if excel_path:
            if not os.path.exists(excel_path):
                # If append is checked but file doesn't exist, warn user
                if self.append_to_excel.get():
                    response = messagebox.askyesno(
                        "Excel File Not Found",
                        f"The Excel file '{os.path.basename(excel_path)}' does not exist.\n"
                        f"Do you want to create a new file instead?"
                    )
                    if response:
                        self.append_to_excel.set(False)
                    else:
                        return
            else:
                # File exists, check if it's locked
                if self.is_excel_file_locked(excel_path):
                    response = messagebox.askyesno(
                        "Excel File is Open",
                        f"The Excel file '{os.path.basename(excel_path)}' is currently open in another program.\n\n"
                        f"Please close the file in Excel and try again.\n\n"
                        f"Do you want to continue without Excel output?",
                        icon='warning'
                    )
                    if not response:
                        return
                    else:
                        # Continue without Excel output
                        self.excel_file_path.set("")
                        self.append_to_excel.set(False)

        # Reset state
        self.is_processing = True
        self.stop_processing = False
        self.process_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        # Clear previous output
        self.output_text.delete(1.0, tk.END)

        # Start processing in separate thread
        def processing_thread():
            try:
                self.log_output(f"📁 Processing started for: {folder_path}")
                # self.log_output(f"🔍 Looking for 12-digit numbers NOT near TIN text")
                # self.log_output(f"🚫 Filtering out numbers that end with '000'")

                # Get page info for logging
                page_option = self.page_option.get()
                page_info = ""
                if page_option == "first":
                    page_info = "First page only"
                elif page_option == "last":
                    page_info = "Last page only"
                elif page_option == "second_last":
                    page_info = "Second-to-last page"

                self.log_output(f"📄 {page_info}")

                # Log Excel options
                excel_path = self.excel_file_path.get()
                if excel_path:
                    if self.append_to_excel.get():
                        self.log_output(f"🗁 Will append results to: {os.path.basename(excel_path)}")
                    else:
                        self.log_output(f"🗁 Will save results to: {os.path.basename(excel_path)}")
                else:
                    self.log_output(f"🗁 Excel file: Auto-generate in PDF folder")

                self.log_output("⚡ Using optimized parallel processing...")
                self.log_output("=" * 60)

                # Get all PDF files
                pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

                if not pdf_files:
                    self.log_output("❌ No PDF files found in the selected folder")
                    return

                # Process files
                if len(pdf_files) > 5:
                    records = self.process_pdfs_parallel(folder_path)
                else:
                    # Sequential processing for small batches
                    records = []
                    for i, pdf_file in enumerate(pdf_files):
                        if self.stop_processing:
                            break
                        pdf_path = os.path.join(folder_path, pdf_file)
                        if not re.match(r'^\d{12}\.pdf$', pdf_file):
                            self.log_output(f"\n🔧 Processing: {pdf_file}")
                            result = self.process_single_pdf_optimized((pdf_path, pdf_file))
                            if result:
                                records.append(result)
                        else:
                            self.log_output(f"⏭️ Skipping (already processed): {pdf_file}")

                        progress = ((i + 1) / len(pdf_files)) * 100
                        self.progress.set(progress)

                # Summary
                self.log_output("\n" + "=" * 60)
                self.log_output("📋 BATCH PROCESSING SUMMARY")
                self.log_output("=" * 60)

                successful = len(records)
                total_pdfs = len(pdf_files)
                skipped = total_pdfs - successful - len([f for f in pdf_files if re.match(r'^\d{12}\.pdf$', f)])

                if successful > 0:
                    for record in records:
                        page_info = self.get_page_info_text(
                            record['Page Option'],
                            record['Page Checked'],
                            record['Total Pages']
                        )
                        warning = " ⚠️" if record.get('Ends with 000') == 'Yes' else ""
                        self.log_output(f"✓ {record['Original Name']} → {record['New Name']}.pdf{page_info}{warning}")

                    self.log_output(f"\n✅ Successfully processed: {successful}/{total_pdfs} files")

                    if skipped > 0:
                        self.log_output(f"⏭️ Skipped: {skipped} files (no valid numbers found or filtered)")

                    # ALWAYS CREATE EXCEL FILE
                    saved_excel_path = None

                    if excel_path and records:
                        # Double-check if file is still locked before writing
                        if os.path.exists(excel_path) and self.is_excel_file_locked(excel_path):
                            self.log_output(f"\n⚠️ Excel file '{os.path.basename(excel_path)}' is now open/locked")
                            self.log_output(f"   Will create a new Excel file in PDF folder instead.")

                            # Create new Excel file in PDF folder
                            saved_excel_path = self.create_excel_record(records, folder_path, False)
                        else:
                            # File is not locked, proceed with save/append
                            append_existing = self.append_to_excel.get() and os.path.exists(excel_path)
                            saved_excel_path = self.create_excel_record(records, folder_path, append_existing)
                    else:
                        # No Excel selected, ALWAYS create one in PDF folder
                        self.log_output(f"\nCreating Excel file in PDF folder...")
                        saved_excel_path = self.create_excel_record(records, folder_path, False)

                    # Show completion message
                    if saved_excel_path:
                        excel_basename = os.path.basename(saved_excel_path)
                        messagebox.showinfo(
                            "Complete",
                            f"✅ Processing complete!\n\n"
                            f"📁 Files processed: {successful}/{total_pdfs}\n"
                            f"📄 Page option: {page_info}\n"
                            f"📊 Excel file: {excel_basename}\n"
                            f"📂 Location: {os.path.dirname(saved_excel_path)}"
                        )
                    else:
                        messagebox.showinfo(
                            "Complete",
                            f"✅ Processing complete!\n\n"
                            f"📁 Files processed: {successful}/{total_pdfs}\n"
                            f"📄 Page option: {page_info}\n"
                            f"⚠️ Note: Could not save Excel file"
                        )
                else:
                    self.log_output("❌ No files were successfully renamed")
                    self.log_output(f"   Total PDFs checked: {total_pdfs}")
                    self.log_output(f"   Note: Numbers ending with '000' are automatically filtered out")

                    if total_pdfs > 0:
                        already_processed = len([f for f in pdf_files if re.match(r'^\d{12}\.pdf$', f)])
                        if already_processed > 0:
                            self.log_output(f"   Already processed: {already_processed} files")

                    # Even if no files renamed, create empty Excel file for record
                    if records or total_pdfs > 0:
                        self.log_output(f"\n📊 Creating Excel file for record...")
                        empty_records = [{
                            'Original Name': 'No files renamed',
                            'New Name': 'N/A',
                            'Page Checked': 'N/A',
                            'Total Pages': 'N/A',
                            'Page Option': page_option,
                            'Processing Time': datetime.now().strftime("%H:%M:%S"),
                            'Found TIN': 'N/A',
                            'Note': 'Filtered: Numbers ending with 000 ignored'
                        }]
                        saved_excel_path = self.create_excel_record(empty_records, folder_path, False)

                        if saved_excel_path:
                            messagebox.showinfo(
                                "Complete",
                                f"Processing complete!\n\n"
                                f"No files were successfully renamed.\n"
                                f"Page option: {page_info}\n"
                                f"Note: Numbers ending with '000' filtered out\n"
                                f"Excel record saved: {os.path.basename(saved_excel_path)}\n"
                                f"Check the output log for details."
                            )
                        else:
                            messagebox.showinfo(
                                "Complete",
                                f"Processing complete!\n\n"
                                f"No files were successfully renamed.\n"
                                f"Page option: {page_info}\n"
                                f"Note: Numbers ending with '000' filtered out\n"
                                f"Check the output log for details."
                            )

            except Exception as e:
                self.log_output(f"❌ Error during processing: {e}")
                messagebox.showerror("Error", f"Processing failed: {e}")
            finally:
                # Reset UI state
                self.is_processing = False
                self.stop_processing = False
                self.progress.set(0)
                self.process_btn.config(state="normal")
                self.stop_btn.config(state="disabled")

        # Start processing thread
        thread = threading.Thread(target=processing_thread)
        thread.daemon = True
        thread.start()


def main():
    root = tk.Tk()
    app = PDFRenamerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
