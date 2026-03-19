import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pyproj import CRS, Transformer

# PROJ strings
PROJ_46 = "+proj=utm +zone=46 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"
PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"
PROJ_MMD2000_GEO = "+proj=longlat +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +no_defs"

# Define CRS
crs_wgs84 = CRS.from_epsg(4326)   # Global GCS (WGS84)
crs_mmd2000_geo = CRS.from_proj4(PROJ_MMD2000_GEO)  # Local GCS (MMD2000 datum)
crs_46 = CRS.from_proj4(PROJ_46)  # Local PCS Zone 46
crs_47 = CRS.from_proj4(PROJ_47)  # Local PCS Zone 47

# Transformers
to_local_46 = Transformer.from_crs(crs_wgs84, crs_46, always_xy=True)
to_global_46 = Transformer.from_crs(crs_46, crs_wgs84, always_xy=True)

to_local_47 = Transformer.from_crs(crs_wgs84, crs_47, always_xy=True)
to_global_47 = Transformer.from_crs(crs_47, crs_wgs84, always_xy=True)

to_local_geo = Transformer.from_crs(crs_wgs84, crs_mmd2000_geo, always_xy=True)
to_global_geo = Transformer.from_crs(crs_mmd2000_geo, crs_wgs84, always_xy=True)

class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GCS ↔ PCS Converter")

        # Input file
        tk.Label(root, text="Input File (CSV/Excel):").grid(row=0, column=0, sticky="w")
        self.input_entry = tk.Entry(root, width=50)
        self.input_entry.grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.browse_input).grid(row=0, column=2)

        # Output file
        tk.Label(root, text="Output File:").grid(row=1, column=0, sticky="w")
        self.output_entry = tk.Entry(root, width=50)
        self.output_entry.grid(row=1, column=1)
        tk.Button(root, text="Browse", command=self.browse_output).grid(row=1, column=2)

        # Conversion direction
        tk.Label(root, text="Conversion Direction:").grid(row=2, column=0, sticky="w")
        self.direction = tk.StringVar(value="g2l46")
        tk.OptionMenu(root, self.direction, "g2l46", "l2g46", "g2l47", "l2g47").grid(row=2, column=1)

        # Transformer format (GCS, PCS, Both)
        tk.Label(root, text="Transformer Output Format:").grid(row=3, column=0, sticky="w")
        self.format_choice = tk.StringVar(value="both")
        tk.OptionMenu(root, self.format_choice, "gcs", "pcs", "both").grid(row=3, column=1)

        # Run button
        tk.Button(root, text="Convert", command=self.convert).grid(row=4, column=1)

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV/Excel files", "*.csv *.xlsx")])
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def convert(self):
        infile = self.input_entry.get()
        outfile = self.output_entry.get()
        if not infile or not outfile:
            messagebox.showerror("Error", "Please select input and output files")
            return

        # Read input
        if infile.endswith(".csv"):
            df = pd.read_csv(infile)
        else:
            df = pd.read_excel(infile)

        # Detect columns
        if "Longitude" in df.columns and "Latitude" in df.columns:
            xcol, ycol = "Longitude", "Latitude"
        elif "Easting" in df.columns and "Northing" in df.columns:
            xcol, ycol = "Easting", "Northing"
        else:
            messagebox.showerror("Error", "Input must have Longitude/Latitude or Easting/Northing columns")
            return

        fmt = self.format_choice.get()
        dirn = self.direction.get()

        # Conversion logic
        if dirn == "l2g46":
            lon, lat = to_global_46.transform(df[xcol].values, df[ycol].values)
            e_global, n_global = Transformer.from_crs(crs_wgs84, CRS.from_epsg(32646), always_xy=True).transform(lon, lat)

            if fmt in ["pcs", "both"]:
                df["Local_Easting"], df["Local_Northing"] = df[xcol], df[ycol]
                df["Global_Easting"], df["Global_Northing"] = e_global, n_global
            if fmt in ["gcs", "both"]:
                df["Longitude"], df["Latitude"] = lon, lat

        elif dirn == "g2l46":
            e_local, n_local = to_local_46.transform(df[xcol].values, df[ycol].values)
            lon_local, lat_local = to_local_geo.transform(df[xcol].values, df[ycol].values)

            if fmt in ["pcs", "both"]:
                df["Local_Easting"], df["Local_Northing"] = e_local, n_local
            if fmt in ["gcs", "both"]:
                df["Global_Longitude"], df["Global_Latitude"] = df[xcol], df[ycol]
                df["Local_Longitude"], df["Local_Latitude"] = lon_local, lat_local

        elif dirn == "l2g47":
            lon, lat = to_global_47.transform(df[xcol].values, df[ycol].values)
            e_global, n_global = Transformer.from_crs(crs_wgs84, CRS.from_epsg(32647), always_xy=True).transform(lon, lat)

            if fmt in ["pcs", "both"]:
                df["Local_Easting"], df["Local_Northing"] = df[xcol], df[ycol]
                df["Global_Easting"], df["Global_Northing"] = e_global, n_global
            if fmt in ["gcs", "both"]:
                df["Longitude"], df["Latitude"] = lon, lat

        elif dirn == "g2l47":
            e_local, n_local = to_local_47.transform(df[xcol].values, df[ycol].values)
            lon_local, lat_local = to_local_geo.transform(df[xcol].values, df[ycol].values)

            if fmt in ["pcs", "both"]:
                df["Local_Easting"], df["Local_Northing"] = e_local, n_local
            if fmt in ["gcs", "both"]:
                df["Global_Longitude"], df["Global_Latitude"] = df[xcol], df[ycol]
                df["Local_Longitude"], df["Local_Latitude"] = lon_local, lat_local

        # Save output
        if outfile.endswith(".csv"):
            df.to_csv(outfile, index=False)
        else:
            df.to_excel(outfile, index=False)

        messagebox.showinfo("Success", f"Conversion complete!\nSaved to {outfile}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConverterGUI(root)
    root.mainloop()
