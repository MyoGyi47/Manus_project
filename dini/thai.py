import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from pyproj import Transformer, CRS
import os

class CoordinateConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("WGS84 to Indian 1975 Converter - Excel Support")
        self.root.geometry("700x550")
        
        # Indian 1975 - Thailand PROJ string
        #self.PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.80170000 +towgs84=109.814,-622.629,106.882,4.431552,-0.135253,-12.971781,-26.816613 +units=m +no_defs"
        self.PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.801698010253 +towgs84=-246.632,-784.833,-276.923,0,0,0,0 +units=m +no_defs"

        # WGS84 UTM zone 47N
        self.WGS84_47 = "+proj=utm +zone=47 +datum=WGS84 +units=m +no_defs"
        
        # Initialize transformers
        self.transformer_geographic_to_local = None
        self.transformer_utm_to_local = None
        
        # Supported file extensions
        self.input_extensions = [('Excel files', '*.xlsx *.xls'), 
                                 ('CSV files', '*.csv'), 
                                 ('All files', '*.*')]
        self.output_extensions = [('Excel files', '*.xlsx'), 
                                  ('CSV files', '*.csv'), 
                                  ('All files', '*.*')]
        
        self.setup_transformers()
        self.create_widgets()
        
    def setup_transformers(self):
        """Setup coordinate transformers"""
        try:
            # WGS84 geographic to Indian 1975 UTM
            crs_wgs84_geo = CRS.from_epsg(4326)  # WGS84 geographic
            crs_indian_utm = CRS.from_proj4(self.PROJ_47)
            self.transformer_geographic_to_local = Transformer.from_crs(
                crs_wgs84_geo, crs_indian_utm, always_xy=True
            )
            
            # WGS84 UTM to Indian 1975 UTM (if needed)
            crs_wgs84_utm = CRS.from_proj4(self.WGS84_47)
            self.transformer_utm_to_local = Transformer.from_crs(
                crs_wgs84_utm, crs_indian_utm, always_xy=True
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to setup transformers: {str(e)}")
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="WGS84 to Indian 1975 Converter", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file section
        input_frame = ttk.LabelFrame(main_frame, text="Input File", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(input_frame, text="File Path:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.input_path = tk.StringVar()
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=60)
        input_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(input_frame, text="Browse Excel/CSV", command=self.browse_input).grid(
            row=1, column=1, sticky=tk.W)
        
        # Sheet selection (for Excel files)
        ttk.Label(input_frame, text="Excel Sheet Name (optional):", font=("Arial", 9)).grid(
            row=2, column=0, sticky=tk.W, pady=(10, 5))
        
        self.sheet_name = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.sheet_name, width=30).grid(
            row=2, column=1, sticky=tk.W, padx=(0, 10))
        
        # Input format selection
        format_frame = ttk.LabelFrame(main_frame, text="Input Format", padding="10")
        format_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.input_format = tk.StringVar(value="geographic")
        
        ttk.Radiobutton(format_frame, text="Geographic (Longitude/Latitude in degrees)", 
                       variable=self.input_format, value="geographic").grid(
            row=0, column=0, sticky=tk.W, pady=2)
        
        ttk.Radiobutton(format_frame, text="UTM (Easting/Northing in meters - WGS84)", 
                       variable=self.input_format, value="utm").grid(
            row=1, column=0, sticky=tk.W, pady=2)
        
        # Column names
        col_frame = ttk.LabelFrame(main_frame, text="Column Names in Input File", padding="10")
        col_frame.grid(row=3, column=0, columnspan=3, pady=(0, 15), sticky=(tk.W, tk.E))
        
        ttk.Label(col_frame, text="X/Longitude/Easting Column:").grid(
            row=0, column=0, sticky=tk.W, pady=5)
        self.lon_col = tk.StringVar(value="lon")
        ttk.Entry(col_frame, textvariable=self.lon_col, width=25).grid(
            row=0, column=1, padx=(10, 0), sticky=tk.W)
        
        ttk.Label(col_frame, text="Y/Latitude/Northing Column:").grid(
            row=1, column=0, sticky=tk.W, pady=5)
        self.lat_col = tk.StringVar(value="lat")
        ttk.Entry(col_frame, textvariable=self.lat_col, width=25).grid(
            row=1, column=1, padx=(10, 0), sticky=tk.W)
        
        # Output file section
        output_frame = ttk.LabelFrame(main_frame, text="Output File", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Output format selection
        ttk.Label(output_frame, text="Output Format:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.output_format = tk.StringVar(value="excel")
        ttk.Radiobutton(output_frame, text="Excel (.xlsx)", 
                       variable=self.output_format, value="excel", command=self.update_output_extension).grid(
            row=1, column=0, sticky=tk.W, pady=2)
        ttk.Radiobutton(output_frame, text="CSV (.csv)", 
                       variable=self.output_format, value="csv", command=self.update_output_extension).grid(
            row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(output_frame, text="File Path:").grid(
            row=2, column=0, sticky=tk.W, pady=(10, 5))
        
        self.output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=60)
        output_entry.grid(row=3, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="Browse Save Location", command=self.browse_output).grid(
            row=3, column=1, sticky=tk.W)
        
        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="Convert Coordinates", 
                                      command=self.convert_coordinates, style="Accent.TButton")
        self.convert_btn.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="", font=("Arial", 9))
        self.status_label.grid(row=6, column=0, columnspan=3)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=400)
        self.progress.grid(row=7, column=0, columnspan=3, pady=(10, 0))
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        input_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)
        
        # Style for accent button
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10, "bold"), padding=10)
    
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select input Excel or CSV file",
            filetypes=self.input_extensions
        )
        if filename:
            self.input_path.set(filename)
            # Auto-suggest output filename
            self.suggest_output_filename(filename)
            
            # If Excel file, try to detect sheet names
            if filename.lower().endswith(('.xlsx', '.xls')):
                try:
                    xl = pd.ExcelFile(filename)
                    if len(xl.sheet_names) == 1:
                        self.sheet_name.set(xl.sheet_names[0])
                    self.status_label.config(
                        text=f"Excel file loaded. Available sheets: {', '.join(xl.sheet_names)}",
                        foreground="blue"
                    )
                except:
                    pass
    
    def suggest_output_filename(self, input_path):
        base_name = os.path.splitext(input_path)[0]
        if self.output_format.get() == "excel":
            self.output_path.set(f"{base_name}_indian1975.xlsx")
        else:
            self.output_path.set(f"{base_name}_indian1975.csv")
    
    def update_output_extension(self):
        # Update output file extension based on format selection
        current_path = self.output_path.get()
        if current_path:
            base_name = os.path.splitext(current_path)[0]
            if self.output_format.get() == "excel":
                self.output_path.set(f"{base_name}.xlsx")
            else:
                self.output_path.set(f"{base_name}.csv")
    
    def browse_output(self):
        filetypes = [('Excel files', '*.xlsx')] if self.output_format.get() == "excel" else [('CSV files', '*.csv')]
        
        filename = filedialog.asksaveasfilename(
            title="Save output file",
            defaultextension=".xlsx" if self.output_format.get() == "excel" else ".csv",
            filetypes=filetypes
        )
        if filename:
            self.output_path.set(filename)
    
    def convert_coordinates(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        lon_col = self.lon_col.get()
        lat_col = self.lat_col.get()
        sheet = self.sheet_name.get() if self.sheet_name.get() else 0
        
        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select input and output files")
            return
        
        if not lon_col or not lat_col:
            messagebox.showerror("Error", "Please specify column names")
            return
        
        try:
            # Show progress
            self.progress.start()
            self.convert_btn.config(state='disabled')
            self.status_label.config(text="Reading input file...", foreground="blue")
            self.root.update()
            
            # Read input file based on extension
            if input_file.lower().endswith(('.xlsx', '.xls')):
                # Read Excel file
                df = pd.read_excel(input_file, sheet_name=sheet if sheet else 0)
                input_type = "Excel"
            else:
                # Read CSV file
                df = pd.read_csv(input_file)
                input_type = "CSV"
            
            self.status_label.config(text=f"Read {len(df)} rows from {input_type} file", 
                                     foreground="blue")
            self.root.update()
            
            if lon_col not in df.columns or lat_col not in df.columns:
                available_cols = ', '.join(df.columns.tolist())
                messagebox.showerror("Error", 
                    f"Columns '{lon_col}' or '{lat_col}' not found in file.\n"
                    f"Available columns: {available_cols}")
                return
            
            # Get coordinates
            lon_values = pd.to_numeric(df[lon_col], errors='coerce')
            lat_values = pd.to_numeric(df[lat_col], errors='coerce')
            
            # Check for NaN values
            nan_count = lon_values.isna().sum() + lat_values.isna().sum()
            if nan_count > 0:
                messagebox.showwarning("Warning", 
                    f"Found {nan_count} non-numeric values in coordinate columns. "
                    f"These rows will be skipped.")
                # Remove rows with NaN coordinates
                valid_mask = lon_values.notna() & lat_values.notna()
                lon_values = lon_values[valid_mask]
                lat_values = lat_values[valid_mask]
                df = df[valid_mask].copy()
            
            self.status_label.config(text="Converting coordinates...", foreground="blue")
            self.root.update()
            
            # Convert based on input format
            if self.input_format.get() == "geographic":
                # Geographic (lat/lon in degrees) to Indian 1975 UTM
                if self.transformer_geographic_to_local is None:
                    self.setup_transformers()
                
                # Transform coordinates
                eastings, northings = self.transformer_geographic_to_local.transform(
                    lon_values.values, lat_values.values
                )
                
            else:  # UTM format
                # WGS84 UTM to Indian 1975 UTM
                if self.transformer_utm_to_local is None:
                    self.setup_transformers()
                
                # Transform coordinates
                eastings, northings = self.transformer_utm_to_local.transform(
                    lon_values.values, lat_values.values
                )
            
            # Add converted coordinates to dataframe
            df['easting_indian1975'] = eastings
            df['northing_indian1975'] = northings
            df['zone'] = 47  # Add zone information
            
            # Round coordinates to reasonable precision
            df['easting_indian1975'] = df['easting_indian1975'].round(3)
            df['northing_indian1975'] = df['northing_indian1975'].round(3)
            
            # Add original coordinates if they were transformed
            if self.input_format.get() == "geographic":
                df['lon_original'] = lon_values
                df['lat_original'] = lat_values
            
            self.status_label.config(text="Saving output file...", foreground="blue")
            self.root.update()
            
            # Save to output file based on format
            if self.output_format.get() == "excel":
                # Save as Excel
                df.to_excel(output_file, index=False)
                output_type = "Excel"
            else:
                # Save as CSV
                df.to_csv(output_file, index=False)
                output_type = "CSV"
            
            # Stop progress bar
            self.progress.stop()
            self.convert_btn.config(state='normal')
            
            # Update status
            self.status_label.config(
                text=f"Successfully converted {len(df)} points. Saved as {output_type}: {output_file}",
                foreground="green"
            )
            
            # Show success message with details
            success_msg = (
                f"Conversion Completed Successfully!\n\n"
                f"Input: {input_type} file with {len(df)} points\n"
                f"Output: {output_type} file\n"
                f"Location: {output_file}\n\n"
                f"New columns added:\n"
                f"- easting_indian1975\n"
                f"- northing_indian1975\n"
                f"- zone (always 47)"
            )
            
            messagebox.showinfo("Success", success_msg)
            
        except Exception as e:
            self.progress.stop()
            self.convert_btn.config(state='normal')
            error_msg = str(e)
            self.status_label.config(text=f"Error: {error_msg}", foreground="red")
            messagebox.showerror("Conversion Error", 
                f"Failed to convert coordinates:\n\n{error_msg}")
    
    def run(self):
        self.root.mainloop()

def main():
    # Check if required packages are installed
    try:
        import pandas
        import pyproj
        from openpyxl import Workbook  # Check for Excel support
    except ImportError as e:
        missing_pkg = str(e).split("'")[1] if "'" in str(e) else "unknown"
        if missing_pkg == "openpyxl":
            print("Note: openpyxl is required for Excel support")
            print("Install with: pip install openpyxl")
        print(f"\nRequired packages: pip install pandas pyproj openpyxl")
        return
    
    root = tk.Tk()
    app = CoordinateConverterGUI(root)
    app.run()

if __name__ == "__main__":
    main()