import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from pyproj import CRS, Transformer
import os

# PROJ strings
PROJ_46 = "+proj=utm +zone=46 +a=6377276.345 +rf=300.8017 +towgs84=109.814,-622.629,-106.882,4.431552,0.135253,12.971781,26.816613 +units=m +no_defs"
PROJ_47 = "+proj=utm +zone=47 +a=6377276.34518 +rf=300.8016996980544 +towgs84=-109.814,622.629,106.882,4.431552,-0.135253,-12.971781,26.816613 +units=m +no_defs"
#PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.8017 +towgs84=-109.814,622.629,106.882,-4.431552,0.135253,12.971781,26.816613 +units=m +no_defs"
PROJ_MMD2000_GEO = "+proj=longlat +a=6377276.345 +rf=300.8017 +towgs84=-109.814,622.629,106.882,4.431552,0.135253,12.971781,26.816613 +no_defs"

# CRS Definitions
crs_wgs84 = CRS.from_epsg(4326)           # Global GCS (WGS84)
crs_mmd2000_geo = CRS.from_proj4(PROJ_MMD2000_GEO)  # Local GCS (MMD2000)
crs_46 = CRS.from_proj4(PROJ_46)          # Local PCS Zone 46
crs_47 = CRS.from_proj4(PROJ_47)          # Local PCS Zone 47
crs_utm46 = CRS.from_epsg(32646)          # Global UTM Zone 46N (WGS84)
crs_utm47 = CRS.from_epsg(32647)          # Global UTM Zone 47N (WGS84)

# Transformers
global_geo_to_local_pcs_46 = Transformer.from_crs(crs_wgs84, crs_46, always_xy=True)
local_pcs_to_global_geo_46 = Transformer.from_crs(crs_46, crs_wgs84, always_xy=True)
global_geo_to_local_pcs_47 = Transformer.from_crs(crs_wgs84, crs_47, always_xy=True)
local_pcs_to_global_geo_47 = Transformer.from_crs(crs_47, crs_wgs84, always_xy=True)
global_geo_to_local_geo = Transformer.from_crs(crs_wgs84, crs_mmd2000_geo, always_xy=True)
local_geo_to_global_geo = Transformer.from_crs(crs_mmd2000_geo, crs_wgs84, always_xy=True)
wgs84_to_utm46 = Transformer.from_crs(crs_wgs84, crs_utm46, always_xy=True)
utm46_to_wgs84 = Transformer.from_crs(crs_utm46, crs_wgs84, always_xy=True)
wgs84_to_utm47 = Transformer.from_crs(crs_wgs84, crs_utm47, always_xy=True)
utm47_to_wgs84 = Transformer.from_crs(crs_utm47, crs_wgs84, always_xy=True)

class MultiConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Coordinate Converter - All in One")
        
        # Input file
        tk.Label(root, text="Input File (Excel/CSV):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.input_entry = tk.Entry(root, width=60)
        self.input_entry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)
        tk.Button(root, text="Browse", command=self.browse_input, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        # Output file
        tk.Label(root, text="Output File (Auto: input_datum.xlsx):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.output_entry = tk.Entry(root, width=60)
        self.output_entry.grid(row=1, column=1, padx=5, pady=5, columnspan=2)
        tk.Button(root, text="Browse", command=self.browse_output, width=10).grid(row=1, column=3, padx=5, pady=5)
        
        # Zone selection
        tk.Label(root, text="Zone:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.zone_var = tk.StringVar(value="46")
        tk.Radiobutton(root, text="Zone 46", variable=self.zone_var, value="46").grid(row=2, column=1, sticky="w")
        tk.Radiobutton(root, text="Zone 47", variable=self.zone_var, value="47").grid(row=2, column=2, sticky="w")
        
        # Source type
        tk.Label(root, text="Source Coordinate Type:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.source_type = tk.StringVar(value="global_geo")
        
        source_options = [
            ("Global Geo (WGS84 Lon/Lat)", "global_geo"),
            ("Local Geo (MMD2000 Lon/Lat)", "local_geo"),
            ("Global PCS (UTM EN)", "global_pcs"),
            ("Local PCS (MMD2000 EN)", "local_pcs")
        ]
        
        self.source_menu = tk.OptionMenu(root, self.source_type, *[opt[0] for opt in source_options])
        self.source_menu.grid(row=3, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.source_mapping = {opt[0]: opt[1] for opt in source_options}
        
        # Output options
        tk.Label(root, text="Output Includes:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        
        self.output_global_geo = tk.BooleanVar(value=True)
        self.output_local_geo = tk.BooleanVar(value=True)
        self.output_global_pcs = tk.BooleanVar(value=True)
        self.output_local_pcs = tk.BooleanVar(value=True)
        
        tk.Checkbutton(root, text="Global Geo", variable=self.output_global_geo).grid(row=4, column=1, sticky="w")
        tk.Checkbutton(root, text="Local Geo", variable=self.output_local_geo).grid(row=4, column=2, sticky="w")
        tk.Checkbutton(root, text="Global PCS", variable=self.output_global_pcs).grid(row=5, column=1, sticky="w")
        tk.Checkbutton(root, text="Local PCS", variable=self.output_local_pcs).grid(row=5, column=2, sticky="w")
        
        # Manual column entry (NEW)
        tk.Label(root, text="OR Manual Column Names:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        
        frame_manual = tk.Frame(root)
        frame_manual.grid(row=6, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        
        tk.Label(frame_manual, text="X/Easting Column:").pack(side=tk.LEFT)
        self.manual_x_col = tk.Entry(frame_manual, width=20)
        self.manual_x_col.pack(side=tk.LEFT, padx=5)
        
        tk.Label(frame_manual, text="Y/Northing Column:").pack(side=tk.LEFT)
        self.manual_y_col = tk.Entry(frame_manual, width=20)
        self.manual_y_col.pack(side=tk.LEFT, padx=5)
        
        # Detect columns button
        tk.Button(root, text="Detect Columns", command=self.detect_and_fill_columns, 
                 bg="lightyellow", width=15).grid(row=7, column=1, pady=5)
        
        # Run button
        tk.Button(root, text="Convert All", command=self.convert_all, 
                 bg="lightgreen", width=20, height=2, font=("Arial", 10, "bold")).grid(row=8, column=1, columnspan=2, pady=15)

    def browse_input(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)
            
            # Auto-generate output filename
            base_name = os.path.splitext(os.path.basename(filename))[0]
            output_dir = os.path.dirname(filename)
            output_path = os.path.join(output_dir, f"{base_name}_converted.xlsx")
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)
            
            # Try to detect columns automatically
            self.detect_and_fill_columns()

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def detect_and_fill_columns(self):
        """Detect columns from file and fill manual entry boxes"""
        infile = self.input_entry.get()
        if not infile:
            return
            
        try:
            if infile.endswith('.csv'):
                df = pd.read_csv(infile, nrows=1)  # Read just first row for column names
            else:
                df = pd.read_excel(infile, nrows=1)
            
            # Show available columns
            col_list = ", ".join(df.columns)
            messagebox.showinfo("Columns in File", f"Available columns:\n{col_list}")
            
            # Try to auto-detect and fill
            for col in df.columns:
                col_lower = col.lower()
                
                # Look for easting/X columns
                if any(keyword in col_lower for keyword in ['easting', 'east', 'x', 'xcoord']):
                    self.manual_x_col.delete(0, tk.END)
                    self.manual_x_col.insert(0, col)
                
                # Look for northing/Y columns
                if any(keyword in col_lower for keyword in ['northing', 'north', 'y', 'ycoord']):
                    self.manual_y_col.delete(0, tk.END)
                    self.manual_y_col.insert(0, col)
                
                # Look for longitude columns
                if any(keyword in col_lower for keyword in ['longitude', 'lon', 'long']):
                    self.manual_x_col.delete(0, tk.END)
                    self.manual_x_col.insert(0, col)
                    # Set source type to geo
                    self.source_type.set("Global Geo (WGS84 Lon/Lat)")
                
                # Look for latitude columns
                if any(keyword in col_lower for keyword in ['latitude', 'lat']):
                    self.manual_y_col.delete(0, tk.END)
                    self.manual_y_col.insert(0, col)
            
        except Exception as e:
            messagebox.showwarning("Warning", f"Cannot read file for column detection: {str(e)}")

    def convert_all(self):
        infile = self.input_entry.get()
        outfile = self.output_entry.get()
        
        if not infile:
            messagebox.showerror("Error", "Please select input file")
            return
            
        if not outfile:
            base_name = os.path.splitext(os.path.basename(infile))[0]
            output_dir = os.path.dirname(infile)
            outfile = os.path.join(output_dir, f"{base_name}_converted.xlsx")
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, outfile)

        # Read input file
        try:
            if infile.endswith('.csv'):
                df = pd.read_csv(infile)
            else:
                df = pd.read_excel(infile)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read file: {str(e)}")
            return

        # Get manual column entries
        manual_x = self.manual_x_col.get().strip()
        manual_y = self.manual_y_col.get().strip()
        
        # Use manual columns if provided
        if manual_x and manual_y:
            if manual_x not in df.columns:
                messagebox.showerror("Error", f"Column '{manual_x}' not found in file")
                return
            if manual_y not in df.columns:
                messagebox.showerror("Error", f"Column '{manual_y}' not found in file")
                return
            
            src_lon_col, src_lat_col = manual_x, manual_y
            source_type_display = self.source_type.get()
            source_type = self.source_mapping.get(source_type_display, "unknown")
            
            # Try to guess source type from column names if not specified
            if source_type == "unknown":
                col_lower_x = manual_x.lower()
                if any(keyword in col_lower_x for keyword in ['longitude', 'lon', 'long']):
                    source_type = "global_geo"
                    self.source_type.set("Global Geo (WGS84 Lon/Lat)")
                elif any(keyword in col_lower_x for keyword in ['easting', 'east', 'x']):
                    # Check if it's local or global by value range
                    sample_val = df[manual_x].iloc[0] if len(df) > 0 else 0
                    if sample_val > 1000000:  # UTM values are large
                        source_type = "global_pcs"
                        self.source_type.set("Global PCS (UTM EN)")
                    else:
                        source_type = "local_pcs"
                        self.source_type.set("Local PCS (MMD2000 EN)")
                else:
                    # Default to global geo
                    source_type = "global_geo"
                    self.source_type.set("Global Geo (WGS84 Lon/Lat)")
            
        else:
            # Try auto-detection
            detected = self.detect_source_columns(df)
            source_type_display = self.source_type.get()
            source_type = self.source_mapping.get(source_type_display)
            
            if source_type not in detected:
                # Try to find any coordinate columns
                all_detected = self.find_any_coordinate_columns(df)
                if all_detected:
                    src_lon_col, src_lat_col, guessed_type = all_detected
                    source_type = guessed_type
                    # Update GUI to show what we found
                    for display, stype in self.source_mapping.items():
                        if stype == guessed_type:
                            self.source_type.set(display)
                            break
                else:
                    messagebox.showerror("Error", 
                        f"No coordinate columns found.\n"
                        f"Please enter column names manually.\n"
                        f"Available columns: {', '.join(df.columns)}")
                    return
            else:
                src_lon_col, src_lat_col = detected[source_type]
        
        # Create result dataframe starting with source columns
        result_df = df.copy()
        result_df["Source_Type"] = self.source_type.get()
        result_df[f"Source_X"] = df[src_lon_col]
        result_df[f"Source_Y"] = df[src_lat_col]
        
        # Get source coordinates
        x_src = df[src_lon_col].values.astype(float)
        y_src = df[src_lat_col].values.astype(float)
        
        try:
            zone = self.zone_var.get()
            
            # ============================================
            # CONVERT FROM SOURCE TO ALL TARGETS
            # ============================================
            
            if source_type == "global_geo":
                # Source: WGS84 Lon/Lat
                lon_wgs84, lat_wgs84 = x_src, y_src
                
                # To Local Geo
                if self.output_local_geo.get():
                    lon_local, lat_local = global_geo_to_local_geo.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Longitude"] = lon_local
                    result_df["Local_Latitude"] = lat_local
                
                # To Local PCS
                if self.output_local_pcs.get():
                    if zone == "46":
                        easting_local, northing_local = global_geo_to_local_pcs_46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_local, northing_local = global_geo_to_local_pcs_47.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Easting"] = easting_local
                    result_df["Local_Northing"] = northing_local
                
                # To Global PCS (UTM)
                if self.output_global_pcs.get():
                    if zone == "46":
                        easting_utm, northing_utm = wgs84_to_utm46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_utm, northing_utm = wgs84_to_utm47.transform(lon_wgs84, lat_wgs84)
                    result_df["Global_Easting"] = easting_utm
                    result_df["Global_Northing"] = northing_utm
                
                # Keep Global Geo
                if self.output_global_geo.get():
                    result_df["Global_Longitude"] = lon_wgs84
                    result_df["Global_Latitude"] = lat_wgs84
            
            elif source_type == "local_geo":
                # Source: MMD2000 Lon/Lat
                lon_local, lat_local = x_src, y_src
                
                # To Global Geo
                if self.output_global_geo.get():
                    lon_wgs84, lat_wgs84 = local_geo_to_global_geo.transform(lon_local, lat_local)
                    result_df["Global_Longitude"] = lon_wgs84
                    result_df["Global_Latitude"] = lat_wgs84
                else:
                    # Still need WGS84 for other conversions
                    lon_wgs84, lat_wgs84 = local_geo_to_global_geo.transform(lon_local, lat_local)
                
                # To Local PCS
                if self.output_local_pcs.get():
                    if zone == "46":
                        easting_local, northing_local = global_geo_to_local_pcs_46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_local, northing_local = global_geo_to_local_pcs_47.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Easting"] = easting_local
                    result_df["Local_Northing"] = northing_local
                
                # To Global PCS
                if self.output_global_pcs.get():
                    if zone == "46":
                        easting_utm, northing_utm = wgs84_to_utm46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_utm, northing_utm = wgs84_to_utm47.transform(lon_wgs84, lat_wgs84)
                    result_df["Global_Easting"] = easting_utm
                    result_df["Global_Northing"] = northing_utm
                
                # Keep Local Geo
                if self.output_local_geo.get():
                    result_df["Local_Longitude"] = lon_local
                    result_df["Local_Latitude"] = lat_local
            
            elif source_type == "global_pcs":
                # Source: UTM Easting/Northing
                easting_utm, northing_utm = x_src, y_src
                
                # To Global Geo (WGS84)
                if zone == "46":
                    lon_wgs84, lat_wgs84 = utm46_to_wgs84.transform(easting_utm, northing_utm)
                else:
                    lon_wgs84, lat_wgs84 = utm47_to_wgs84.transform(easting_utm, northing_utm)
                
                if self.output_global_geo.get():
                    result_df["Global_Longitude"] = lon_wgs84
                    result_df["Global_Latitude"] = lat_wgs84
                
                # To Local Geo
                if self.output_local_geo.get():
                    lon_local, lat_local = global_geo_to_local_geo.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Longitude"] = lon_local
                    result_df["Local_Latitude"] = lat_local
                
                # To Local PCS
                if self.output_local_pcs.get():
                    if zone == "46":
                        easting_local, northing_local = global_geo_to_local_pcs_46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_local, northing_local = global_geo_to_local_pcs_47.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Easting"] = easting_local
                    result_df["Local_Northing"] = northing_local
                
                # Keep Global PCS
                if self.output_global_pcs.get():
                    result_df["Global_Easting"] = easting_utm
                    result_df["Global_Northing"] = northing_utm
            
            elif source_type == "local_pcs":
                # Source: MMD2000 Easting/Northing
                easting_local, northing_local = x_src, y_src
                
                # To Global Geo
                if zone == "46":
                    lon_wgs84, lat_wgs84 = local_pcs_to_global_geo_46.transform(easting_local, northing_local)
                else:
                    lon_wgs84, lat_wgs84 = local_pcs_to_global_geo_47.transform(easting_local, northing_local)
                
                if self.output_global_geo.get():
                    result_df["Global_Longitude"] = lon_wgs84
                    result_df["Global_Latitude"] = lat_wgs84
                
                # To Local Geo
                if self.output_local_geo.get():
                    lon_local, lat_local = global_geo_to_local_geo.transform(lon_wgs84, lat_wgs84)
                    result_df["Local_Longitude"] = lon_local
                    result_df["Local_Latitude"] = lat_local
                
                # To Global PCS
                if self.output_global_pcs.get():
                    if zone == "46":
                        easting_utm, northing_utm = wgs84_to_utm46.transform(lon_wgs84, lat_wgs84)
                    else:
                        easting_utm, northing_utm = wgs84_to_utm47.transform(lon_wgs84, lat_wgs84)
                    result_df["Global_Easting"] = easting_utm
                    result_df["Global_Northing"] = northing_utm
                
                # Keep Local PCS
                if self.output_local_pcs.get():
                    result_df["Local_Easting"] = easting_local
                    result_df["Local_Northing"] = northing_local
            
            # Clean any NaN or infinite values
            result_df = result_df.replace([np.inf, -np.inf], np.nan)
            
            # Save output
            if outfile.endswith('.csv'):
                result_df.to_csv(outfile, index=False)
            else:
                if not outfile.endswith(('.xlsx', '.xls')):
                    outfile = outfile + '.xlsx'
                result_df.to_excel(outfile, index=False)
            
            # Show summary
            output_cols = [col for col in result_df.columns if col not in ['Source_Type', 'Source_X', 'Source_Y'] and not col.startswith('Unnamed')]
            messagebox.showinfo("Success", 
                f"Conversion complete!\n\n"
                f"Source: {self.source_type.get()}\n"
                f"Source Columns: {src_lon_col}, {src_lat_col}\n"
                f"Zone: {zone}\n"
                f"Output columns: {len(output_cols)}\n"
                f"Saved to: {outfile}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}\n\nError type: {type(e).__name__}")

    def detect_source_columns(self, df):
        """Detect source coordinate columns - IMPROVED VERSION"""
        columns = {}
        
        # Check each column name
        for col in df.columns:
            col_lower = col.lower().strip()
            
            # Global Geo detection
            if any(keyword in col_lower for keyword in ['longitude', 'lon', 'long']):
                # Find matching latitude
                for col2 in df.columns:
                    if col2 != col and any(keyword in col2.lower() for keyword in ['latitude', 'lat']):
                        columns['global_geo'] = (col, col2)
                        break
            
            # Local Geo detection  
            elif any(keyword in col_lower for keyword in ['local_longitude', 'local_lon', 'local_long']):
                for col2 in df.columns:
                    if col2 != col and any(keyword in col2.lower() for keyword in ['local_latitude', 'local_lat']):
                        columns['local_geo'] = (col, col2)
                        break
            
            # Easting/Northing detection (more flexible)
            elif any(keyword in col_lower for keyword in ['easting', 'east', 'xcoord', 'x']):
                # Try to find matching northing
                for col2 in df.columns:
                    col2_lower = col2.lower()
                    if col2 != col and any(keyword in col2_lower for keyword in ['northing', 'north', 'ycoord', 'y']):
                        # Check if it's likely to be local or global by column name
                        if 'local' in col_lower or 'local' in col2_lower:
                            columns['local_pcs'] = (col, col2)
                        elif 'global' in col_lower or 'global' in col2_lower or 'utm' in col_lower or 'utm' in col2_lower:
                            columns['global_pcs'] = (col, col2)
                        else:
                            # Default to local_pcs for Myanmar context
                            columns['local_pcs'] = (col, col2)
                        break
        
        return columns

    def find_any_coordinate_columns(self, df):
        """Find any coordinate columns regardless of naming"""
        # First try exact matches
        for col1 in df.columns:
            col1_lower = col1.lower()
            for col2 in df.columns:
                if col1 == col2:
                    continue
                    
                col2_lower = col2.lower()
                
                # Check for Lon/Lat pairs
                if (('lon' in col1_lower and 'lat' in col2_lower) or 
                    ('long' in col1_lower and 'lat' in col2_lower) or
                    ('lon' in col1_lower and 'latitude' in col2_lower)):
                    # Check if local or global
                    if 'local' in col1_lower or 'local' in col2_lower:
                        return col1, col2, "local_geo"
                    else:
                        return col1, col2, "global_geo"
                
                # Check for Easting/Northing pairs
                elif (('east' in col1_lower and 'north' in col2_lower) or
                      ('x' in col1_lower and 'y' in col2_lower) or
                      ('easting' in col1_lower and 'northing' in col2_lower)):
                    # Check value range to guess local vs global
                    try:
                        sample1 = float(df[col1].iloc[0]) if len(df) > 0 else 0
                        if sample1 > 100000:  # Likely UTM/global
                            return col1, col2, "global_pcs"
                        else:  # Likely local
                            return col1, col2, "local_pcs"
                    except:
                        return col1, col2, "local_pcs"  # Default to local
        
        return None

if __name__ == "__main__":
    root = tk.Tk()
    app = MultiConverterGUI(root)
    root.mainloop()