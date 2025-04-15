import tkinter as tk
from tkinter import filedialog, ttk
import os
import hashlib
import csv
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import threading

class DOCXAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Metadata Analyzer by helpdesk8675")
        
        # Configure root window
        self.root.geometry("600x300")
        
        # Create frame for paths
        paths_frame = ttk.Frame(root, padding="10")
        paths_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input path selection
        ttk.Label(paths_frame, text="Input Folder:").grid(row=0, column=0, sticky=tk.W)
        self.input_path = tk.StringVar()
        ttk.Entry(paths_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(paths_frame, text="Browse", command=self.browse_input).grid(row=0, column=2)
        
        # Output path selection
        ttk.Label(paths_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W)
        self.output_path = tk.StringVar()
        ttk.Entry(paths_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(paths_frame, text="Browse", command=self.browse_output).grid(row=1, column=2)
        
        # Process button
        ttk.Button(paths_frame, text="Process", command=self.start_processing, style='Green.TButton').grid(row=2, column=1, pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(paths_frame, length=400, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=3, pady=10)
        
        # Status label
        self.status_var = tk.StringVar()
        ttk.Label(paths_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3)
        
        # Configure green style for process button
        style = ttk.Style()
        style.configure('Green.TButton', foreground='green')
        
        # Define XML namespaces used across functions
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
            'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
        }

    def browse_input(self):
        directory = filedialog.askdirectory()
        self.input_path.set(directory)

    def browse_output(self):
        directory = filedialog.askdirectory()
        self.output_path.set(directory)

    def calculate_md5(self, filepath):
        hash_md5 = hashlib.md5()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def start_processing(self):
        """Start the processing in a separate thread to keep UI responsive"""
        input_folder = self.input_path.get()
        output_folder = self.output_path.get()
        
        if not input_folder or not output_folder:
            self.status_var.set("Please select both input and output folders")
            return
            
        # Reset progress bar
        self.progress['value'] = 0
        
        # Start processing in a separate thread
        threading.Thread(target=self.process_files, daemon=True).start()

    def update_progress(self, value=None, status=None):
        """Update the progress bar and status text from a worker thread"""
        if value is not None:
            self.progress['value'] = value
        if status is not None:
            self.status_var.set(status)
        self.root.update_idletasks()

    def get_empty_metadata(self):
        """Return an empty metadata dictionary with all fields initialized"""
        return {
            # Core file info
            'Filepath': '',
            'MD5': 'Not Found',
            'Created Date': 'Not Found',
            'File Size (bytes)': 'Not Found',
            
            # Core Properties
            'Title': 'Not Found',
            'Creator': 'Not Found',
            'Description': 'Not Found',
            'Subject': 'Not Found',
            'Last Modified By': 'Not Found',
            'Revision': 'Not Found',
            'Created DateTime': 'Not Found',
            'Modified DateTime': 'Not Found',
            'Category': 'Not Found',
            'Keywords': 'Not Found',
            
            # App Properties
            'Application': 'Not Found',
            'App Version': 'Not Found',
            'Company': 'Not Found',
            'Total Editing Time': 'Not Found',
            'Pages': 'Not Found',
            'Words': 'Not Found',
            'Characters': 'Not Found',
            'Lines': 'Not Found',
            'Paragraphs': 'Not Found',
            'Template': 'Not Found',
            'Last Printed': 'Not Found',
            'Characters with Spaces': 'Not Found',
            'Manager': 'Not Found',
            'Document Security': 'Not Found',  # Renamed from DocSecurity
            'Shared Doc': 'Not Found',        # Renamed from SharedDoc
            
            # Document Settings
            'Default Tab Stop': 'Not Found',
            'Zoom Level': 'Not Found',
            'Proof State': 'Not Found',
            'Auto Hyphenation': 'Not Found',
            'Track Revisions': 'Not Found',
            'Attached Template': 'Not Found',
            
            # Theme Info
            'Theme Name': 'Not Found',
            'Color Scheme': 'Not Found',
            'Font Scheme': 'Not Found',
            
            # Security
            'Encryption Status': 'Not Protected',
            'Password Protected': 'No',
            'Document Restrictions': 'None',
            
            # Document Statistics
            'Compression Ratio': 'Not Calculated',
            'Embedded Objects Count': 0,
            'External Links Count': 0,
            
            # Original IDs
            'w14:docId': 'Not Found',
            'w15:docId': 'Not Found',
            'w:rsidRoot': 'Not Found',
            'w:rsid Values': '',
            'RSID Count': 0,  # Added to track number of RSID values
            
            # Comments and Revisions
            'Comments Count': 0,
            'Revisions Count': 0,
            'Comment Authors': '',
            
            # Fonts
            'Used Fonts': '',
            'Font Substitutions': '',
            'Embedded Fonts': '',
            
            # Error fields
            'Error': '',
            'Error: Core Properties': '',
            'Error: App Properties': '',
            'Error: Settings': '',
            'Error: Theme Info': '',
            'Error: Font Info': '',
            'Error: Comments': ''
        }

    def extract_docx_metadata(self, docx_path):
        """Extract metadata from a DOCX file"""
        metadata = self.get_empty_metadata()
        metadata['Filepath'] = docx_path
        
        if not os.path.exists(docx_path):
            metadata['Error'] = 'File not found'
            return metadata
            
        # Set basic file information
        try:
            metadata['MD5'] = self.calculate_md5(docx_path)
            metadata['Created Date'] = datetime.fromtimestamp(os.path.getctime(docx_path)).strftime('%Y-%m-%d %H:%M:%S')
            metadata['File Size (bytes)'] = os.path.getsize(docx_path)
        except Exception as e:
            metadata['Error'] = f'Error accessing file: {str(e)}'
            return metadata

        try:
            with zipfile.ZipFile(docx_path, 'r') as docx:
                # Extract file content once and reuse
                file_contents = {}
                for file_path in ['docProps/core.xml', 'docProps/app.xml', 'word/settings.xml', 
                                 'word/theme/theme1.xml', 'word/fontTable.xml', 'word/comments.xml']:
                    if file_path in docx.namelist():
                        file_contents[file_path] = docx.read(file_path)

                # Process core properties
                self._extract_core_properties(metadata, file_contents.get('docProps/core.xml'))
                
                # Process app properties
                self._extract_app_properties(metadata, file_contents.get('docProps/app.xml'))
                
                # Process settings and IDs
                settings_root = self._extract_settings(metadata, file_contents.get('word/settings.xml'))
                
                # Process theme information
                self._extract_theme_info(metadata, file_contents.get('word/theme/theme1.xml'))
                
                # Process font information
                self._extract_font_info(metadata, file_contents.get('word/fontTable.xml'))
                
                # Process comments
                self._extract_comments(metadata, file_contents.get('word/comments.xml'))
                
                # Check for document protection
                if settings_root is not None and settings_root.find('.//w:documentProtection', self.namespaces) is not None:
                    metadata['Document Restrictions'] = 'Protected'
                    metadata['Password Protected'] = 'Yes'
                
                # Calculate compression ratio
                uncompressed_size = sum(info.file_size for info in docx.infolist())
                compressed_size = os.path.getsize(docx_path)
                metadata['Compression Ratio'] = f"{(compressed_size / uncompressed_size * 100):.2f}%" if uncompressed_size > 0 else "N/A"

        except zipfile.BadZipFile:
            metadata['Error'] = 'Invalid or corrupted DOCX file'
        except ET.ParseError:
            metadata['Error'] = 'Invalid XML in DOCX file'
        except Exception as e:
            metadata['Error'] = f'Error processing file: {str(e)}'

        return metadata
        
    def _extract_core_properties(self, metadata, xml_content):
        """Extract core properties from core.xml"""
        if not xml_content:
            return
            
        try:
            core_root = ET.fromstring(xml_content)
            
            for prop, xpath in {
                'Title': './/dc:title',
                'Creator': './/dc:creator',
                'Description': './/dc:description',
                'Subject': './/dc:subject',
                'Last Modified By': './/cp:lastModifiedBy',
                'Revision': './/cp:revision',
                'Created DateTime': './/dcterms:created',
                'Modified DateTime': './/dcterms:modified',
                'Category': './/cp:category',
                'Keywords': './/cp:keywords'
            }.items():
                elem = core_root.find(xpath, self.namespaces)
                if elem is not None and elem.text:
                    metadata[prop] = elem.text
        except Exception as e:
            metadata['Error: Core Properties'] = str(e)
            
    def _extract_app_properties(self, metadata, xml_content):
        """Extract app properties from app.xml"""
        if not xml_content:
            return
            
        try:
            app_root = ET.fromstring(xml_content)
            
            # Define mapping with proper namespace prefix and matching metadata field names
            app_properties = {
                'Template': 'Template',
                'TotalTime': 'Total Editing Time',
                'Pages': 'Pages',
                'Words': 'Words',
                'Characters': 'Characters',
                'Application': 'Application',
                'DocSecurity': 'Document Security',  # Changed from DocSecurity
                'Lines': 'Lines',
                'Paragraphs': 'Paragraphs',
                'Manager': 'Manager',
                'Company': 'Company',
                'CharactersWithSpaces': 'Characters with Spaces',
                'SharedDoc': 'Shared Doc',  # Changed from SharedDoc
                'AppVersion': 'App Version',
                'LastPrinted': 'Last Printed'
            }
            
            # Look for each property with the proper namespace
            for xml_prop, metadata_field in app_properties.items():
                xpath = f'.//ep:{xml_prop}'
                elem = app_root.find(xpath, self.namespaces)
                if elem is not None and elem.text:
                    metadata[metadata_field] = elem.text
        except Exception as e:
            metadata['Error: App Properties'] = str(e)
            
    def _extract_settings(self, metadata, xml_content):
        """Extract settings and IDs from settings.xml"""
        if not xml_content:
            return None
            
        try:
            settings_root = ET.fromstring(xml_content)
            
            # Extract IDs
            w14_docid = settings_root.find('.//w14:docId', self.namespaces)
            if w14_docid is not None:
                metadata['w14:docId'] = w14_docid.get(f"{{{self.namespaces['w14']}}}val")

            w15_docid = settings_root.find('.//w15:docId', self.namespaces)
            if w15_docid is not None:
                metadata['w15:docId'] = w15_docid.get(f"{{{self.namespaces['w15']}}}val")

            rsid_root = settings_root.find('.//w:rsidRoot', self.namespaces)
            if rsid_root is not None:
                metadata['w:rsidRoot'] = rsid_root.get(f"{{{self.namespaces['w']}}}val")

            # Extract all RSID values (revision identifier values)
            rsid_vals = []
            
            # Look for rsids in the rsids element
            rsids_elem = settings_root.find('.//w:rsids', self.namespaces)
            if rsids_elem is not None:
                for rsid in rsids_elem.findall('.//w:rsid', self.namespaces):
                    val = rsid.get(f"{{{self.namespaces['w']}}}val")
                    if val:
                        rsid_vals.append(val)
            
            # Store the RSID values
            metadata['w:rsid Values'] = ';'.join(rsid_vals) if rsid_vals else 'Not Found'
            metadata['RSID Count'] = len(rsid_vals)

            # Extract Settings
            tab_stop = settings_root.find('.//w:defaultTabStop', self.namespaces)
            metadata['Default Tab Stop'] = tab_stop.get(f"{{{self.namespaces['w']}}}val") if tab_stop is not None else 'Not Found'

            zoom = settings_root.find('.//w:zoom', self.namespaces)
            metadata['Zoom Level'] = zoom.get(f"{{{self.namespaces['w']}}}percent") if zoom is not None else 'Not Found'

            metadata['Track Revisions'] = 'Yes' if settings_root.find('.//w:trackRevisions', self.namespaces) is not None else 'No'
            
            return settings_root
        except Exception as e:
            metadata['Error: Settings'] = str(e)
            return None
            
    def _extract_theme_info(self, metadata, xml_content):
        """Extract theme information from theme1.xml"""
        if not xml_content:
            return
            
        try:
            theme_root = ET.fromstring(xml_content)
            
            theme = theme_root.find('.//a:theme', self.namespaces)
            if theme is not None:
                metadata['Theme Name'] = theme.get('name', 'Not Found')

            clr_scheme = theme_root.find('.//a:clrScheme', self.namespaces)
            if clr_scheme is not None:
                metadata['Color Scheme'] = clr_scheme.get('name', 'Not Found')

            font_scheme = theme_root.find('.//a:fontScheme', self.namespaces)
            if font_scheme is not None:
                metadata['Font Scheme'] = font_scheme.get('name', 'Not Found')
        except Exception as e:
            metadata['Error: Theme Info'] = str(e)
            
    def _extract_font_info(self, metadata, xml_content):
        """Extract font information from fontTable.xml"""
        if not xml_content:
            return
            
        try:
            font_root = ET.fromstring(xml_content)
            
            fonts = font_root.findall('.//w:font', self.namespaces)
            metadata['Used Fonts'] = ';'.join(font.get(f"{{{self.namespaces['w']}}}name", '') for font in fonts)
        except Exception as e:
            metadata['Error: Font Info'] = str(e)
            
    def _extract_comments(self, metadata, xml_content):
        """Extract comments information from comments.xml"""
        if not xml_content:
            return
            
        try:
            comments_root = ET.fromstring(xml_content)
            comments = comments_root.findall('.//w:comment', self.namespaces)
            metadata['Comments Count'] = len(comments)
            metadata['Comment Authors'] = ';'.join(set(comment.get(f"{{{self.namespaces['w']}}}author", '') for comment in comments))
        except Exception as e:
            metadata['Error: Comments'] = str(e)

    def process_files(self):
        """Process all DOCX files and write metadata to CSV"""
        input_folder = self.input_path.get()
        output_folder = self.output_path.get()
        
        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)

        output_file = os.path.join(output_folder, f"docx_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        
        # Get list of all DOCX files
        docx_files = []
        for root, _, files in os.walk(input_folder):
            for file in files:
                if file.lower().endswith('.docx'):
                    docx_files.append(os.path.join(root, file))

        if not docx_files:
            self.update_progress(status="No DOCX files found in input folder")
            return

        self.update_progress(status=f"Found {len(docx_files)} DOCX files to process")
        self.progress['maximum'] = len(docx_files)

        # Get fieldnames from the empty metadata template
        fieldnames = list(self.get_empty_metadata().keys())

        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                for idx, docx_file in enumerate(docx_files):
                    self.update_progress(idx, f"Processing: {os.path.basename(docx_file)}")

                    metadata = self.extract_docx_metadata(docx_file)
                    writer.writerow(metadata)

            self.update_progress(len(docx_files), f"Processing complete! Output saved to: {output_file}")
            
        except PermissionError:
            self.update_progress(status="Error: Unable to write to output file - Permission denied")
        except Exception as e:
            self.update_progress(status=f"Error processing files: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DOCXAnalyzerGUI(root)
    root.mainloop()
