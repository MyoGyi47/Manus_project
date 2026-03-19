"""
Core processing logic for Geo Processor
"""
import re
import pandas as pd
from pyproj import Transformer, CRS
from config import (
    PROJ_46, PROJ_47, PROTECTED_SPLIT_NAMES_WORD, 
    ROAD_KEYWORDS, ALLOWED_KEYWORDS, POLYGON_KEYWORDS
)
from utils import (
    get_all_vertices, move_del_to_remark, detect_zone_from_filename,
    normalize_for_comparison, round_coordinate_for_phrase,
    extract_kml_from_kmz, kml_coords_to_list
)
import xml.etree.ElementTree as ET
from shapely.geometry import Point, LineString, Polygon

# ==================== COORDINATE TRANSFORMATION ====================
def get_transformer(lon, proj_choice, zone_mode="auto", manual_zone=None, filename=None):
    """Get appropriate transformer based on projection and zone"""
    # Determine zone
    if zone_mode == "manual" and manual_zone in [46, 47]:
        zone = manual_zone
    elif zone_mode == "auto" and filename:
        zone = detect_zone_from_filename(filename, lon, proj_choice)
    else:
        # Default auto detection
        zone = 47 if (lon >= 96) else 46
        if proj_choice == "Custom_UTM":
            zone = 47 if (lon >= 95.9968784133) else 46
    
    # Create transformer
    if proj_choice == "Custom_UTM":
        proj_string = PROJ_46 if zone == 46 else PROJ_47
        return Transformer.from_crs("EPSG:4326", CRS.from_proj4(proj_string), always_xy=True)
    elif proj_choice == "WGS84_UTM":
        epsg_code = 32600 + zone
        return Transformer.from_crs("EPSG:4326", f"EPSG:{epsg_code}", always_xy=True)
    elif proj_choice == "WGS84_LatLon":
        return Transformer.from_crs("EPSG:4326", "EPSG:4326", always_xy=True)
    
    return Transformer.from_crs("EPSG:4326", "EPSG:4326", always_xy=True)

# ==================== NAME PROCESSING ====================
class NameProcessor:
    """Process and split feature names"""
    
    def __init__(self):
        self.protected_names = PROTECTED_SPLIT_NAMES_WORD
        self.road_keywords = ROAD_KEYWORDS
        self.allowed_keywords = ALLOWED_KEYWORDS
    
    def has_change_pattern(self, name):
        """Check if name matches change pattern with allowed keywords"""
        if not isinstance(name, str):
            return False
        
        name_lower = name.lower().strip()
        if ' to ' in name_lower or ' from ' in name_lower:
            return any(keyword in name_lower for keyword in self.allowed_keywords)
        return False
    
    def extract_change_keywords(self, row):
        """Extract change keywords from name"""
        name = str(row['Name']).strip()
        original_remark = row.get('Remark', '')
        
        patterns = [
            (r'^Change\s+(\w+)\s+(to|from)\s+(\w+)(?:\s+(\d+))?\s*(?:\((.*?)\))?\s*$', True),
            (r'^(\w+)\s+(to|from)\s+(\w+)(?:\s+(\d+))?\s*(?:\((.*?)\))?\s*$', False)
        ]
        
        for pattern_str, has_change_prefix in patterns:
            match = re.match(pattern_str, name, re.IGNORECASE)
            if match:
                groups = match.groups()
                if has_change_prefix:
                    from_keyword, direction, to_keyword, number, description = groups[0], groups[1], groups[2], groups[3], groups[4]
                else:
                    from_keyword, direction, to_keyword, number, description = groups[0], groups[1], groups[2], groups[3], groups[4]
                
                if (from_keyword.lower() in self.allowed_keywords and 
                    to_keyword.lower() in self.allowed_keywords):
                    
                    base_name = f"Change {from_keyword} {direction} {to_keyword}" if has_change_prefix else f"{from_keyword} {direction} {to_keyword}"
                    new_name = f"{base_name} {number}" if number else base_name
                    new_object = description if description else ""
                    
                    return new_name, new_object, original_remark
        
        return name, '', original_remark
    
    def process(self, row, source_type):
        """Main processing method for name splitting"""
        name = str(row['Name']).strip()
        original_remark = row.get('Remark', '')
        has_underscore = '_' in name
        #remark_to_return = '' if has_underscore else original_remark
        # အခု:
        remark_to_return = original_remark  # underscore ပါရင်လည်း remark ထားခဲ့မယ်

        # Priority 1: Protected names
        if has_underscore:
            for p_name in sorted(self.protected_names, key=len, reverse=True):
                name_normalized = normalize_for_comparison(name)
                p_normalized = normalize_for_comparison(p_name)
                
                if name_normalized.startswith(p_normalized + '_'):
                    pattern = r'^' + re.escape(p_name).replace('_', r'\s*_\s*') + r'(?=_|\s|$)'
                    match = re.match(pattern, name, re.IGNORECASE)
                    
                    if match:
                        remaining = name[match.end():].strip()
                        if remaining.startswith('_'):
                            remaining = remaining[1:].strip()
                        
                        if remaining:
                            # Remove del content
                            if 'del' in remaining.lower():
                                idx = remaining.lower().find('del')
                                remaining = remaining[:idx].strip()
                            return p_name, remaining, remark_to_return
        
        # Priority 2: Generic underscore splitting        
        # Priority 2: Generic underscore splitting
        if has_underscore and not any(p.lower() in name.lower() for p in self.protected_names):
            parts = name.split('_', 1)
            if len(parts) == 2:
                object_part = parts[1].strip()
                object_lower = object_part.lower()
                
                # Define bua patterns
                bua_patterns = ['delete bua', 'deleted bua', 'deleted bua area',
                            'bua deleted area', 'bua deleted', 'bua delete']
                
                # Check for bua patterns
                if any(pattern in object_lower for pattern in bua_patterns):
                    return parts[0].strip(), object_part, remark_to_return
                elif 'del' in object_lower:
                    idx = object_lower.find('del')
                    return parts[0].strip(), object_part[:idx].strip(), remark_to_return
                else:
                    return parts[0].strip(), object_part, remark_to_return
        
        # Priority 3: Change keywords
        change_result = self.extract_change_keywords(row)
        if change_result[0] != name:
            return change_result
        
        # Priority 4: Road keywords
        matched_keyword = None
        keyword_position = -1
        
        for keyword in sorted(self.road_keywords, key=len, reverse=True):
            pattern = r'\b' + re.escape(keyword) + r'\b'
            match = re.search(pattern, name, re.IGNORECASE)
            if match:
                matched_keyword = match.group()
                keyword_position = match.start()
                break
        
        if matched_keyword and keyword_position != -1:
            base_name_segment = name[:keyword_position + len(matched_keyword)].strip()
            suffix_full = name[keyword_position + len(matched_keyword):].strip()
            
            if suffix_full:
                # Check for repair/alignment patterns
                repair_pattern = r'^\s*(?:(?:\(?(repair|check|alignment|dual\s+lane|under\s+construction)\)?)\s*(?:\d+\s*)?\s*)+'
                repair_match = re.match(repair_pattern, suffix_full, re.IGNORECASE)
                
                if repair_match:
                    full_match_text = repair_match.group(0).strip()
                    remaining = suffix_full[repair_match.end():].strip()
                    new_name = f"{base_name_segment} {full_match_text}".strip()
                    
                    if remaining.startswith('('):
                        paren_match = re.search(r'\(\s*(.*?)\s*\)', remaining)
                        if paren_match and not remaining[paren_match.end():].strip():
                            return new_name, paren_match.group(1).strip(), remark_to_return
                    
                    return new_name, remaining, remark_to_return if remaining else (new_name, '', remark_to_return)
                
                # Check for parentheses
                if '(' in suffix_full:
                    paren_match = re.search(r'\(\s*(.*?)\s*\)', suffix_full)
                    if paren_match:
                        return base_name_segment, paren_match.group(1).strip(), remark_to_return
                
                return base_name_segment, suffix_full.strip(), remark_to_return
        
        return name, '', remark_to_return

# ==================== KML/KMZ PROCESSING ====================
def parse_kml(kml_content, proj_choice, zone_mode="auto", manual_zone=None, filename=None):
    """Parse KML content and extract features"""
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    root = ET.fromstring(kml_content)
    placemarks = root.findall(".//kml:Placemark", ns)
    rows = []
    no_counter = 1
    
    for pm in placemarks:
        name_elem = pm.find("kml:name", ns)
        name = name_elem.text.strip() if name_elem is not None else ""
        desc_elem = pm.find("kml:description", ns)
        remark = desc_elem.text.strip() if desc_elem is not None else ""
        
        geometry = None
        geom_type_str = None
        
        point_elem = pm.find(".//kml:Point/kml:coordinates", ns)
        line_elem = pm.find(".//kml:LineString/kml:coordinates", ns)
        poly_elem = pm.find(".//kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates", ns)
        
        if point_elem is not None and point_elem.text.strip():
            coords_list = kml_coords_to_list(point_elem.text)
            geom_type_str = "point"
            if coords_list:
                geometry = Point(coords_list[0])
        elif line_elem is not None and line_elem.text.strip():
            coords_list = kml_coords_to_list(line_elem.text)
            geom_type_str = "line"
            if coords_list:
                geometry = LineString(coords_list)
        elif poly_elem is not None and poly_elem.text.strip():
            coords_list = kml_coords_to_list(poly_elem.text)
            geom_type_str = "polygon"
            if coords_list:
                geometry = Polygon(coords_list)
        
        if geometry is None:
            continue
        
        all_vertices = get_all_vertices(geometry)
        if not all_vertices:
            continue
        
        needs_rounding = proj_choice in ["Custom_UTM", "WGS84_UTM"]
        lon_main = all_vertices[0][0]
        transformer = get_transformer(lon_main, proj_choice, zone_mode, manual_zone, filename)
        feature_name, feature_remark = move_del_to_remark(name, remark)
        
        for lon, lat in all_vertices:
            x_out, y_out = transformer.transform(lon, lat)
            x_coord = round(x_out) if needs_rounding else x_out
            y_coord = round(y_out) if needs_rounding else y_out
            
            rows.append([no_counter, feature_name, x_coord, y_coord, feature_remark, geom_type_str])
        
        no_counter += 1
    
    return rows

# ==================== VECTOR FILE PROCESSING ====================
def process_vector(input_path, proj_choice, zone_mode="auto", manual_zone=None):
    """Process geospatial vector files"""
    if input_path.lower().endswith(('.kmz', '.kml')):
        if input_path.lower().endswith('.kmz'):
            kml_content = extract_kml_from_kmz(input_path)
        else:
            with open(input_path, 'r', encoding='utf-8') as f:
                kml_content = f.read()
        return parse_kml(kml_content, proj_choice, zone_mode, manual_zone, input_path)
    else:
        import geopandas as gpd
        gdf = gpd.read_file(input_path)
        rows = []
        no_counter = 1
        needs_rounding = proj_choice in ["Custom_UTM", "WGS84_UTM"]
        
        for _, row in gdf.iterrows():
            geometry = row.geometry
            if geometry is None or geometry.is_empty:
                continue
            
            all_vertices = get_all_vertices(geometry)
            if not all_vertices:
                continue
            
            geom_type_str = geometry.geom_type.replace('Multi', '').replace('String', '').lower()
            lon_main = all_vertices[0][0]
            transformer = get_transformer(lon_main, proj_choice, zone_mode, manual_zone, input_path)
            
            name = row.get('Name', row.get('name', ''))
            remark = row.get('Remark', row.get('description', ''))
            feature_name, feature_remark = move_del_to_remark(name, remark)
            
            for lon, lat in all_vertices:
                x_out, y_out = transformer.transform(lon, lat)
                x_coord = round(x_out) if needs_rounding else x_out
                y_coord = round(y_out) if needs_rounding else y_out
                rows.append([no_counter, feature_name, x_coord, y_coord, feature_remark, geom_type_str])
            
            no_counter += 1
        
        return rows

# ==================== EXCEL EXPORT ====================
class ExcelExporter:
    """Export data to Excel with multiple sheets"""
    
    def __init__(self, proj_choice):
        self.proj_choice = proj_choice
        self.name_processor = NameProcessor()
        self._setup_headers()
    
    def _setup_headers(self):
        """Setup column headers based on projection choice"""
        if self.proj_choice in ["Custom_UTM", "WGS84_UTM"]:
            self.coord_cols = ["Easting", "Northing"]
        elif self.proj_choice == "WGS84_LatLon":
            self.coord_cols = ["Longitude", "Latitude"]
        else:
            self.coord_cols = ["X_Coord", "Y_Coord"]
        
        self.other_headers = ["No.", "Name", self.coord_cols[0], self.coord_cols[1], "Remark", "Source"]
        self.word_headers = ["No.", "Name", "Object", self.coord_cols[0], self.coord_cols[1], 
                            "EastingMM", "NorthingMM", "Remark", "Source"]
    
    def export(self, all_rows, output_path):
        """Main export method"""
        df_all_raw = pd.DataFrame(all_rows, columns=self.other_headers)
        
        # Create different sheets
        df_word = self._create_word_sheet(df_all_raw)
        df_aa = self._create_aa_sheet(df_all_raw)
        df_excel = self._create_excel_sheet(df_all_raw)
        df_raw = self._create_raw_sheet(df_all_raw)
        df_pts = self._create_points_sheet(df_all_raw)
        
        # Write to Excel
        self._write_to_excel(output_path, df_word, df_aa, df_excel, df_raw, df_pts)
    
    def _add_mm_columns(self, df):
        """Add Burmese digit columns to DataFrame"""
        df = df.copy()
        df["EastingMM"] = df[self.coord_cols[0]].astype(str).apply(
            lambda x: round_coordinate_for_phrase(x, "easting")
        )
        df["NorthingMM"] = df[self.coord_cols[1]].astype(str).apply(
            lambda x: round_coordinate_for_phrase(x, "northing")
        )
        return df
    
    def _reindex_sequential(self, df):
        """Reindex DataFrame with sequential numbering"""
        df = df.copy()
        current_no = 1
        new_numbers = []
        
        for _, row in df.iterrows():
            if pd.notna(row['No.']) and str(row['No.']).strip() != '':
                new_numbers.append(current_no)
                current_no += 1
            else:
                new_numbers.append('')
        
        df['No.'] = new_numbers
        return df
    
    def _create_word_sheet(self, df_all_raw):
        """Create 'To Word' sheet"""
        # Group by feature number and process
        grouped = df_all_raw.groupby('No.')
        new_rows = []
        
        for feature_no, group in grouped:
            if group.empty:
                continue
            
            source_type = group['Source'].iloc[0].lower()
            feature_name = group['Name'].iloc[0].lower()
            
            # Determine row selection logic
            has_road_keyword = any(keyword in feature_name for keyword in ROAD_KEYWORDS)
            has_polygon_keyword = any(keyword in feature_name for keyword in POLYGON_KEYWORDS)
            
            if source_type == 'point':
                new_rows.append(group.iloc[0].to_dict())
            elif source_type == 'polygon':
                new_rows.append(group.iloc[0].to_dict())
            elif source_type == 'line':
                if has_road_keyword:
                    # Start and end rows for roads
                    new_rows.append(group.iloc[0].to_dict())
                    if len(group) > 1:
                        end_row = group.iloc[-1].to_dict()
                        end_row['No.'] = ''
                        end_row['Remark'] = ''
                        new_rows.append(end_row)
                elif has_polygon_keyword:
                    new_rows.append(group.iloc[0].to_dict())
                else:
                    new_rows.append(group.iloc[0].to_dict())
                    if len(group) > 1:
                        end_row = group.iloc[-1].to_dict()
                        end_row['No.'] = ''
                        end_row['Remark'] = ''
                        new_rows.append(end_row)
        
        # Create DataFrame and process
        df_word = pd.DataFrame(new_rows, columns=self.other_headers)
        df_word.insert(2, 'Object', '')
        df_word = self._add_mm_columns(df_word)
        df_word = df_word[self.word_headers]
        
        # Apply name splitting
        self._apply_name_splitting(df_word)
        
        # Reindex
        df_word = self._reindex_sequential(df_word)
        
        return df_word
    
    def should_move_te_zu(self, name):
        """Check if te zu should be moved to object column"""
        name_str = str(name).strip().lower()
        
        # Normalize variations: tezu, te-zu, te  zu → te zu
        normalized = re.sub(r'te\s*[-]?\s*zu', 'te zu', name_str)
        
        # မရွှေ့ရန် cases (normalized version နဲ့စစ်):
        # 1. "te zu_" (underscore ပါ)
        if 'te zu_' in normalized:
            return False
            
        # 2. "xxxxx te zu" (ရှေ့မှာ စာရှိ)  
        if normalized != 'te zu' and ' te zu' in normalized:
            return False
            
        # 3. "te zu xxx" (နောက်မှာ စာရှိ)
        if normalized != 'te zu' and 'te zu ' in normalized:
            return False
        
        # ရွှေ့ရန် cases:
        # "te zu" သီးသန့် သို့မဟုတ် variations (tezu, te-zu, te  zu)
        te_zu_patterns = ['te zu', 'tezu', 'te-zu', 'te  zu']
        return any(pattern == normalized for pattern in te_zu_patterns)
    
    def move_te_zu_to_object(self, df):
        """Move te zu to object column for special cases"""
        mask = df['Name'].apply(self.should_move_te_zu)
        
        if not mask.any():
            return df

        df_to_process = df[mask].copy()
        
        def apply_move(row):
            name = str(row['Name']).strip()
            obj = str(row['Object']).strip()
            remark = str(row['Remark']).strip()
            
            # Extract del/delete content for remark
            del_match = re.search(r'\b(del|delete)\s+(.*)', name, re.IGNORECASE)
            del_content = del_match.group(0) if del_match else ""
            
            # Remove del content from name
            clean_name = name
            if del_match:
                clean_name = re.sub(r'\b(del|delete)\s+.*', '', clean_name, flags=re.IGNORECASE).strip()
            
            # Build new values
            new_name = ''  # Name ကို empty လုပ်မယ်
            new_obj = f"{obj} {clean_name}".strip() if obj else clean_name
            new_remark = f"{remark} | {del_content}".strip() if del_content else remark
            
            # Remove leading "|" if any
            if new_remark.startswith('|'):
                new_remark = new_remark[1:].strip()
                
            return pd.Series([new_name, new_obj, new_remark])
        
        new_splits = df_to_process.apply(apply_move, axis=1)
        new_splits.columns = ['Name_Moved', 'Object_Moved', 'Remark_Moved']
        
        df.loc[mask, 'Name'] = new_splits['Name_Moved']
        df.loc[mask, 'Object'] = new_splits['Object_Moved']
        df.loc[mask, 'Remark'] = new_splits['Remark_Moved']
        
        return df
    
    def _apply_name_splitting(self, df):
        """Apply name splitting logic to DataFrame"""
        protected_lower = [name.lower() for name in PROTECTED_SPLIT_NAMES_WORD]
        
        has_underscore = df['Name'].str.contains('_', na=False)
        has_road_suffix = df['Name'].apply(self._has_road_keyword_and_suffix)
        has_change = df['Name'].apply(self.name_processor.has_change_pattern)
        
        mask_to_process = (has_underscore | has_road_suffix | has_change) & ~df['Name'].str.lower().isin(protected_lower)
        
        if mask_to_process.any():
            split_results = df.loc[mask_to_process].apply(
                lambda row: self.name_processor.process(row, row['Source']),
                axis=1,
                result_type='expand'
            )
            split_results.columns = ['Name_New', 'Object_New', 'Remark_New']
            df.loc[mask_to_process, 'Name'] = split_results['Name_New']
            df.loc[mask_to_process, 'Object'] = split_results['Object_New']
            df.loc[mask_to_process, 'Remark'] = split_results['Remark_New']
        
        # ADD THIS: Apply te zu movement after regular name splitting
        df = self.move_te_zu_to_object(df)
        
        return df
    
    def _has_road_keyword_and_suffix(self, name):
        """Check if name has road keyword with suffix"""
        name_lower = str(name).strip().lower()
        if not name_lower:
            return False
        
        for keyword in ROAD_KEYWORDS:
            keyword_lower = keyword.lower()
            pos = name_lower.find(keyword_lower)
            while pos != -1:
                # Check if it's a whole word match
                before_ok = (pos == 0 or not name_lower[pos-1].isalnum() or name_lower[pos-1] == ' ')
                after_pos = pos + len(keyword_lower)
                after_ok = (after_pos >= len(name_lower) or not name_lower[after_pos].isalnum() or name_lower[after_pos] == ' ')
                
                if before_ok and after_ok:
                    remaining = name_lower[after_pos:].strip()
                    if remaining:
                        return True
                
                pos = name_lower.find(keyword_lower, pos + 1)
        
        return False
    
    def _create_aa_sheet(self, df_all_raw):
        """Create 'aa' sheet"""
        df_aa = df_all_raw.copy()
        df_aa.insert(2, 'Object', '')
        df_aa = self._add_mm_columns(df_aa)
        df_aa = df_aa[self.word_headers]
        
        # Apply name splitting
        self._apply_name_splitting(df_aa)
        
        # Reindex with blank rows for same features
        df_aa['No.'] = df_aa['No.'].astype(object)
        mask = df_aa['No.'].ne(df_aa['No.'].shift()).fillna(True)
        df_aa.loc[~mask, ['No.', 'Name']] = ''
        
        df_aa = self._reindex_sequential(df_aa)
        
        return df_aa
    
    def _create_excel_sheet(self, df_all_raw):
        """Create 'Excel' sheet"""
        # Filter logic (simplified version)
        is_line_or_polygon = df_all_raw['Source'].str.lower().isin(['line', 'polygon'])
        has_underscore = df_all_raw['Name'].str.contains('_', na=False)
        
        # Helper function for protected underscore check
        def has_only_protected_underscore(name):
            name_str = str(name)
            name_without_brackets = re.sub(r'\([^)]*\)', '', name_str).strip()
            temp_name = name_without_brackets
            
            for protected_name in PROTECTED_SPLIT_NAMES_WORD:
                if temp_name.lower().startswith(protected_name.lower()):
                    temp_name = temp_name[len(protected_name):].strip()
                    break
            
            return temp_name.count('_') == 0
        
        has_only_protected = df_all_raw['Name'].apply(has_only_protected_underscore)
        
        # Keep logic
        keep_mask = (
            is_line_or_polygon |
            (~is_line_or_polygon & has_underscore & ~has_only_protected)
        )
        
        df_excel = df_all_raw[keep_mask].copy()
        
        # Reindex with blank rows
        df_excel['No.'] = df_excel['No.'].astype(object)
        mask = df_excel['No.'].ne(df_excel['No.'].shift()).fillna(True)
        df_excel.loc[~mask, ['No.', 'Name']] = ''
        
        df_excel = self._reindex_sequential(df_excel)
        
        return df_excel
    
    def _create_raw_sheet(self, df_all_raw):
        """Create 'Raw Data' sheet"""
        df_raw = df_all_raw.copy()
        df_raw['No.'] = df_raw['No.'].astype(object)
        mask = df_raw['No.'].ne(df_raw['No.'].shift()).fillna(True)
        df_raw.loc[~mask, ['No.', 'Name']] = ''
        return df_raw
    
    def _create_points_sheet(self, df_all_raw):
        """Create 'Pts' sheet (points with underscores)"""
        df_points_raw = df_all_raw[df_all_raw['Source'].str.lower() == 'point'].copy()
        has_underscore = df_points_raw['Name'].str.contains('_', na=False)
        
        # Helper function for protected underscore check
        def has_only_protected_underscore(name):
            name_str = str(name)
            name_without_brackets = re.sub(r'\([^)]*\)', '', name_str).strip()
            temp_name = name_without_brackets
            
            for protected_name in PROTECTED_SPLIT_NAMES_WORD:
                if temp_name.lower().startswith(protected_name.lower()):
                    temp_name = temp_name[len(protected_name):].strip()
                    break
            
            return temp_name.count('_') == 0
        
        has_only_protected = df_points_raw['Name'].apply(has_only_protected_underscore)
        
        # Filter points with underscores that are not protected-only
        filter_mask = has_underscore & ~has_only_protected
        df_points = df_points_raw[filter_mask].copy()
        
        # Reindex sequentially
        df_points['No.'] = range(1, len(df_points) + 1)
        
        return df_points
    
    def _write_to_excel(self, output_path, df_word, df_aa, df_excel, df_raw, df_pts):
        """Write all sheets to Excel file"""
        from openpyxl.styles import Font, Alignment
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write each sheet
            df_word.to_excel(writer, sheet_name='To Word', index=False, header=self.word_headers)
            df_aa.to_excel(writer, sheet_name='aa', index=False, header=self.word_headers)
            df_excel.to_excel(writer, sheet_name='Excel', index=False, header=self.other_headers)
            df_raw.to_excel(writer, sheet_name='Raw Data', index=False, header=self.other_headers)
            df_pts.to_excel(writer, sheet_name='Pts', index=False, header=self.other_headers)
            
            # Apply styling to all sheets
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                self._apply_styling(ws)
    
    def _apply_styling(self, ws):
        """Apply styling to worksheet"""
        from openpyxl.styles import Font, Alignment
        
        # Set font for all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Pyidaungsu", size=13)
                if cell.row == 1:
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="left")
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            header_name = col[0].value
            
            for cell in col:
                length = len(str(cell.value)) if cell.value is not None else 0
                max_length = max(max_length, length)
            
            if header_name == "No.":
                ws.column_dimensions[col_letter].width = max(5, max_length + 2)
            else:
                ws.column_dimensions[col_letter].width = max(10, max_length + 3)