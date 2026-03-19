#perfect and tt46.py ထဲမှာ "_" ပါတဲ့ row တွေမှာ del, delete တွေကို remark ကို ပြန်ထည့်ပေးထားပြီးရွှေ့ပေးတယ်။
import pandas as pd
import pandas as pd
import numpy as np
from io import StringIO
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from shapely.geometry import Point, LineString, Polygon, MultiPoint, MultiLineString, MultiPolygon
import sys

# PyQt5 imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QFileDialog, 
                             QMessageBox, QProgressBar, QTextEdit, QTabWidget, QGroupBox, QRadioButton)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor

# Geospatial libraries
try:
    import geopandas as gpd
    from pyproj import Transformer, CRS
    from openpyxl.styles import Font, Alignment
    GEOSPATIAL_LIBS_READY = True
except ImportError:
    GEOSPATIAL_LIBS_READY = False
    
# Custom Datum PROJ strings
PROJ_46 = "+proj=utm +zone=46 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"
PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"

# Global definitions for splitting logic
ALLOWED_KEYWORDS = ['ew', 'mr', 'sr', 'or', 'ct', 'pt', 'fp']
POLYGON_KEYWORDS = ['bua', 'lake', 'pond', 'religious area',
                    'sport field', 'martyrs temple', 'dam', 'fish farm', 'swamp area',
                     'cultivation area', 'reservoir', 'highway bus terminal compound', 'myit kyo in',
                      'solar panel', 'spill way', 'river', 'cemetery area', 'water area repair', 'golf course',
                       'livestock farm' ]
PROTECTED_SPLIT_NAMES_WORD = ['Kyauk_O',  'U_yin', 'Nga ku_Oh', 'Ta da_U',
                              'Tha bye_U', 'Kyauk_O(Kyauk kon)', 'Kyun_U', 'San_U', 'Le gyin_U',
                                'Kan_U', 'Laung daw_U', 'O_gyi gwe', 'O_yin', 'Daung_U', 'Chaung_U',
                                'O_pon daw', 'Chaung zon', "Chaung gwa", "Chaung bat", "U_hmin", "TADA_U",
                                "Ta da_U",] 
ROAD_KEYWORDS = ["main road", "main rd", "mainroad", "main-road",
    "secondary road", "secondary rd", "secondary-road",
    "cart track", "cart-track", "carttrack",
    "other road", "other rd", "other-road",
    "pack track", "packtrack", "pack-track",
    "footpath", "foot path", "foot-path",'canal', 'stream', 'chaung', 'embankment',
    'river', 'fish farm', "express way", "expressway", "express-way", "zaung dan",]

def convert_to_mm_digits(number_str):
    """ဂဏန်းတွေကို မြန်မာလို ပြောင်းပေးတဲ့ function"""
    mm_digits = {
        '0': '၀', '1': '၁', '2': '၂', '3': '၃', '4': '၄',
        '5': '၅', '6': '၆', '7': '၇', '8': '၈', '9': '၉'
    }
    return ''.join(mm_digits.get(char, char) for char in str(number_str))

def round_coordinate_for_phrase(coord_str, coord_type=None):
    """
    Applies custom rounding and slicing logic for coordinates.
    """
    coord_str = str(coord_str).strip()
    extracted_four_digits = None

    if coord_type == "easting":
        if len(coord_str) < 5:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[1:5]
    elif coord_type == "northing":
        if len(coord_str) < 6:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[2:6]
    else:
        if len(coord_str) < 6:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[2:6]
        
    if extracted_four_digits is None or not extracted_four_digits.isdigit():
        return convert_to_mm_digits(coord_str)

    try:
        num_with_decimal = float(extracted_four_digits[:-1] + '.' + extracted_four_digits[-1])
    except (ValueError, IndexError):
        return convert_to_mm_digits(coord_str)
    
    rounded_val = round(num_with_decimal)
    result_str = str(int(rounded_val)).zfill(3)
    return convert_to_mm_digits(result_str)

def detect_zone(lon, proj_choice):
    """Longitude အလိုက် UTM Zone 46 သို့ 47 ကို စစ်ဆေးသည်။"""
    if proj_choice == "Custom_UTM":  # MMD2000
        return 47 if lon >= 95.9968784133 else 46
    else:  # WGS84 UTM or WGS84 LatLon
        return 47 if lon >= 96 else 46
    
def detect_zone_from_filename(filename, lon, proj_choice):
    """
    Filename ကနေ zone ကို detect လုပ်သည်။
    Patterns:
    1. 2195 15, 2195-15, 2195 15 MZO, 2195-15MZO
    2. 219515, 219515MZO
    3. 95 15, 95-15 (2-digit year)
    4. 95/15, 95_15
    
    Logic:
    - 2195 ဖိုင်အတွက်: 13,14,15,16 ဆိုရင် Zone 46
    - 2196 ဖိုင်အတွက်: 01,1,02,2,03,3,04,4 ဆိုရင် Zone 47
    - တခြား case အားလုံးအတွက် default auto detection (lon နဲ့ဆုံးဖြတ်)
    """
    import re
    
    # Filename ကနေ base name ကို ယူပါ
    base_name = os.path.basename(filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    # Remove non-alphanumeric for better matching
    clean_name = re.sub(r'[^a-zA-Z0-9\s\-]', ' ', name_without_ext)
    
    # Different patterns to try
    patterns = [
        r'(\d{4})[\s\-_](\d{1,2})',      # 2195 15, 2195-15, 2195_15
        r'(\d{4})(\d{2})',               # 219515
        r'(\d{2})[\s\-_](\d{1,2})',      # 95 15, 95-15
        r'(\d{2})(\d{2})',               # 9515
    ]
    
    for pattern in patterns:
        match = re.search(pattern, clean_name)
        if match:
            year_part = match.group(1)  # နှစ် part
            number_part = match.group(2)  # နံပါတ် part
            
            # Determine if year is 2-digit or 4-digit
            if len(year_part) == 4:
                year_last_two = year_part[2:]  # 2195 -> 95
            else:  # 2-digit
                year_last_two = year_part  # 95 -> 95
            
            try:
                year_num = int(year_last_two)
                num = int(number_part)
                
                # Zone detection logic - ရှင်းလင်းအောင်ပြင်ထားတယ်
                if year_num == 95:
                    # 2195 ဖိုင်: 13,14,15,16 ဆိုရင် Zone 46
                    if num in [13, 14, 15, 16]:
                        return 46
                    # တခြား 01-12 အတွက် default auto detection
                    else:
                        return detect_zone(lon, proj_choice)
                
                elif year_num == 96:
                    # 2196 ဖိုင်: 01-04 ဆိုရင် Zone 47
                    if num in [1, 2, 3, 4]:
                        return 47
                    # တခြား 05-16 အတွက် default auto detection
                    else:
                        return detect_zone(lon, proj_choice)
                
                # တခြား year အတွက် default auto detection
                else:
                    return detect_zone(lon, proj_choice)
                        
            except ValueError:
                continue
    
    # Filename မှာ pattern မတွေ့ရင် original logic ကို သုံးမယ်
    return detect_zone(lon, proj_choice)

def get_transformer(lon, proj_choice, zone_mode="auto", manual_zone=None, filename=None):
    """ရွေးချယ်ထားသော Projection စနစ်အတွက် Transformer ကို ပြန်ပေးသည်။"""
    
    # Determine zone based on mode
    if zone_mode == "manual" and manual_zone in [46, 47]:
        zone = manual_zone
    elif zone_mode == "auto" and filename:
        # NEW: Filename-based auto detection
        zone = detect_zone_from_filename(filename, lon, proj_choice)
    else:
        # Original auto detection
        zone = detect_zone(lon, proj_choice)
    
    if proj_choice == "Custom_UTM":
        return Transformer.from_crs("EPSG:4326", CRS.from_proj4(PROJ_46 if zone == 46 else PROJ_47), always_xy=True)
    elif proj_choice == "WGS84_UTM":
        epsg_code = 32600 + zone
        return Transformer.from_crs("EPSG:4326", f"EPSG:{epsg_code}", always_xy=True)
    elif proj_choice == "WGS84_LatLon":
        return Transformer.from_crs("EPSG:4326", "EPSG:4326", always_xy=True)
    return get_transformer(lon, "Custom_UTM")

def extract_kml_from_kmz(kmz_path):
    """KMZ ဖိုင်အတွင်းမှ KML content ကို ထုတ်ယူသည်။"""
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        for file_name in kmz.namelist():
            if file_name.endswith('.kml'):
                return kmz.read(file_name).decode('utf-8')
    raise FileNotFoundError("KMZ ဖိုင်အတွင်း KML ဖိုင်ကို ရှာမတွေ့ပါ။")

def move_del_to_remark(name, remark):
    """Name ထဲမှ 'del' ကို Remark သို့ ရွှေ့သည်။"""
    #match = re.search(r'del', name, re.IGNORECASE)
    match = re.search(r'\b(del|delete)\b', name, re.IGNORECASE)
    if match:
        split_index = match.start()
        moved_text = name[split_index:].strip()
        name = name[:split_index].strip()
        remark = f"{remark.strip()} | {moved_text}" if remark else moved_text
    return name, remark

def get_all_vertices(geometry):
    """Shapely Geometry မှ Vertex အားလုံးကို ထုတ်ယူသည်။"""
    if geometry is None: return []
    geom_type = geometry.geom_type
    
    if geom_type in ['Point', 'MultiPoint']:
        return [(p.x, p.y) for p in (geometry.geoms if geom_type == 'MultiPoint' else [geometry])]
    
    elif geom_type in ['LineString', 'MultiLineString']:
        all_coords = []
        geometries = geometry.geoms if geom_type == 'MultiLineString' else [geometry]
        for line in geometries:
            coords = list(line.coords)
            all_coords.extend(coords[:-1] if len(coords) > 1 and coords[0] == coords[-1] else coords)
        return all_coords
        
    elif geom_type in ['Polygon', 'MultiPolygon']:
        all_coords = []
        geometries = geometry.geoms if geom_type == 'MultiPolygon' else [geometry]
        for poly in geometries:
            coords = list(poly.exterior.coords)
            all_coords.extend(coords[:-1] if coords and coords[0] == coords[-1] else coords)
        return all_coords
    return []

def parse_kml(kml_content, proj_choice, zone_mode="auto", manual_zone=None, filename=None):
    """KML Content မှ Placemark များကို Extract လုပ်ပြီး Coordinate စနစ် ပြောင်းလဲသည်။"""
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    root = ET.fromstring(kml_content)
    placemarks = root.findall(".//kml:Placemark", ns)
    rows = []
    no_counter = 1
    
    def kml_coords_to_list(coords_str):
        coords = []
        for point in coords_str.strip().split(' '):
            if point:
                lon, lat, *_ = map(float, point.split(","))
                coords.append((lon, lat))
        return coords

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
            if coords_list: geometry = Point(coords_list[0])
        elif line_elem is not None and line_elem.text.strip():
            coords_list = kml_coords_to_list(line_elem.text)
            geom_type_str = "line"
            if coords_list: geometry = LineString(coords_list)
        elif poly_elem is not None and poly_elem.text.strip():
            coords_list = kml_coords_to_list(poly_elem.text)
            geom_type_str = "polygon"
            if coords_list: geometry = Polygon(coords_list)
        
        if geometry is None: continue

        all_vertices = get_all_vertices(geometry)
        if not all_vertices: continue
        
        needs_rounding = proj_choice in ["Custom_UTM", "WGS84_UTM"]
        lon_main = all_vertices[0][0]
        
        # NEW: Pass filename to get_transformer
        transformer = get_transformer(lon_main, proj_choice, zone_mode, manual_zone, filename)
        
        feature_name, feature_remark = move_del_to_remark(name, remark)

        for lon, lat in all_vertices:
            x_out, y_out = transformer.transform(lon, lat)
            x_coord = round(x_out) if needs_rounding else x_out
            y_coord = round(y_out) if needs_rounding else y_out
            
            rows.append([no_counter, feature_name, x_coord, y_coord, feature_remark, geom_type_str])
            
        no_counter += 1

    return rows

def process_vector(input_path, proj_choice, zone_mode="auto", manual_zone=None):
    """Geospatial ဖိုင်များကို ဖတ်ပြီး Rows များအဖြစ် ပြောင်းလဲသည်။"""
    if input_path.lower().endswith(('.kmz', '.kml')):
        kml_content = extract_kml_from_kmz(input_path) if input_path.lower().endswith('.kmz') else open(input_path,'r', encoding='utf-8').read()
        return parse_kml(kml_content, proj_choice, zone_mode, manual_zone, input_path)  # filename pass
    else:
        gdf = gpd.read_file(input_path)
        rows = []
        no_counter = 1
        needs_rounding = proj_choice in ["Custom_UTM", "WGS84_UTM"]
        
        for idx, row in gdf.iterrows():
            geometry = row.geometry
            if geometry is None or geometry.is_empty: continue

            all_vertices = get_all_vertices(geometry)
            if not all_vertices: continue
            
            geom_type_str = geometry.geom_type.replace('Multi', '').replace('String', '').lower()  
            lon_main = all_vertices[0][0] if all_vertices else 0
            
            # NEW: Pass filename to get_transformer
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

def apply_to_word_name_split(row, source_type):
    name = str(row['Name']).strip()
    original_remark = row.get('Remark', '')
    name_lower = name.lower()
    has_underscore = '_' in name
    #remark_to_return = '' if has_underscore else original_remark
    # အခု:
    remark_to_return = original_remark  # underscore ပါရင်လည်း remark ထားခဲ့မယ်  
   
    # Priority 1: Protected Prefix Split - WORKING VERSION
    if has_underscore:
        for p_name in sorted(PROTECTED_SPLIT_NAMES_WORD, key=len, reverse=True):
            p_lower = p_name.lower()
            
            # Simple approach: Remove all spaces and check
            name_no_spaces = re.sub(r'\s+', '', name_lower)
            p_no_spaces = re.sub(r'\s+', '', p_lower)
            
            if name_no_spaces.startswith(p_no_spaces + '_'):
                # Find where protected name ends in original
                # We need to extract remaining text correctly
                
                # Find position considering spaces
                pattern = r'^' + re.escape(p_name).replace('_', r'\s*_\s*') + r'(?=_|\s|$)'
                match = re.match(pattern, name, re.IGNORECASE)
                
                if match:
                    remaining_start = match.end()
                    remaining = name[remaining_start:].strip()
                    
                    # Remove leading underscore if present
                    if remaining.startswith('_'):
                        remaining = remaining[1:].strip()
                    
                    if remaining:
                        object_part = remaining
                        if 'del' in object_part.lower():
                            idx = object_part.lower().find('del')
                            object_part = object_part[:idx].strip()
                        return p_name, object_part, remark_to_return
                
    # --- Priority 2: Generic Underscore Splitting ---    
    if has_underscore:
        is_protected = any(p.lower() in name_lower for p in PROTECTED_SPLIT_NAMES_WORD)
        if not is_protected:
            parts = name.split('_', 1)
        
            if len(parts) == 2:
                object_part = parts[1].strip()
                object_lower = object_part.lower()
                
                # Check if object contains bua with del/delete
                bua_del_patterns = [
                    r'del(?:ete)?\s+bua',
                    r'bua\s+del(?:ete)?',
                    r'deleted\s+bua',
                    r'bua\s+deleted'
                ]
                
                has_bua_del = any(re.search(pattern, object_lower) for pattern in bua_del_patterns)
                
                # Find del/delete in object
                del_match = re.search(r'\b(del|delete)\b', object_part, re.IGNORECASE)
                if del_match:
                    if has_bua_del:
                        # For bua with del patterns: keep everything in object
                        return parts[0].strip(), object_part, remark_to_return
                    else:
                        # For non-bua del cases: move del and everything after to remark
                        del_start = del_match.start()
                        moved_text = object_part[del_start:].strip()
                        
                        # Update object (remove del and after)
                        object_part = object_part[:del_start].strip()
                        
                        # Set moved text to remark (THIS WAS MISSING!)
                        if moved_text:
                            remark_to_return = moved_text
                
                return parts[0].strip(), object_part, remark_to_return
            
    # --- Priority 3: Change Keywords ---
    change_result = extract_change_keywords(row)
    if change_result[0] != name:
        return change_result

    # --- Priority 4: Road Keyword Splitting ---
    matched_keyword = None
    keyword_position = -1    

    for keyword in sorted(ROAD_KEYWORDS, key=len, reverse=True):
        k_lower = keyword.lower()
        pos = name_lower.find(k_lower)  # "sa mon secondary road" မှာ "secondary road" ရှာတယ်
        
        if pos == -1:
            continue
        
        # IMPROVED: Space ပါသည်ဖြစ်စေ မပါသည်ဖြစ်စေ keyword ရှာမယ်
        # "secondary road" နဲ့ "secondary road(" ကိုပါ ရှာမယ်
        pattern = r'\b' + re.escape(keyword) + r'\b'
        match = re.search(pattern, name, re.IGNORECASE)
        
        if match:
            matched_keyword = match.group()
            keyword_position = match.start()
            break

    # Current approach ကိုပဲ ဆက်သုံးပြီး suffix မှာ space မပါရင်လည်း parentheses ရှာမယ်
    if matched_keyword and keyword_position != -1:
        base_name_segment = name[:keyword_position + len(matched_keyword)].strip()
        suffix_full = name[keyword_position + len(matched_keyword):].strip()   
              
        if suffix_full.strip():            
            # --- 1. Repair/Number Pattern/parentheses ---       
            pattern = r'^\s*(?:(?:\(?(repair|check|alignment|dual\s+lane|under\s+construction)\)?)\s*(?:\d+\s*)?\s*)+'
            repair_match = re.match(pattern, suffix_full, re.IGNORECASE)

            if repair_match:
                full_match_text = repair_match.group(0).strip()
                remaining_after_repair = suffix_full[repair_match.end():].strip()
                new_name = f"{base_name_segment} {full_match_text}".strip()                
                
                if remaining_after_repair.startswith('('):
                    paren_match = re.search(r'\(\s*(.*?)\s*\)', remaining_after_repair)
                    if paren_match:
                        after_paren_text = remaining_after_repair[paren_match.end():].strip()                        
                        
                        if not after_paren_text:
                            new_object = paren_match.group(1).strip()                            
                            return new_name, new_object, remark_to_return
                
                if remaining_after_repair:                    
                    return new_name, remaining_after_repair, remark_to_return
                else:                    
                    return new_name, '', remark_to_return
            
        if suffix_full.startswith('(') or '(' in suffix_full:
            paren_match = re.search(r'\(\s*(.*?)\s*\)', suffix_full)
            if paren_match:
                new_name = base_name_segment
                new_object = paren_match.group(1).strip()
                return new_name, new_object, remark_to_return            
            
            else:
                new_name = base_name_segment
                new_object = suffix_full.strip()                
                return new_name, new_object, remark_to_return                        
        
    return name, '', remark_to_return

def has_change_pattern(name):
        """Check if name matches change pattern with allowed keywords"""
        if not isinstance(name, str):
            return False
        
        name_lower = name.lower().strip()       
        
        # Check for " to " or " from " with allowed keywords
        if ' to ' in name_lower or ' from ' in name_lower:
            # Check if any allowed keyword is in the name
            for keyword in ALLOWED_KEYWORDS:
                if keyword in name_lower:
                    return True
        
        return False

def extract_change_keywords(row):
    name = str(row['Name']).strip()
    original_remark = row.get('Remark', '')    
        
    # Multiple patterns to match:    # 1. "Change X to Y ..."    # 2. "Change X from Y ..."  
    # - NEW    # 3. "X to Y ..."    # 4. "X from Y ..."  - NEW    
    patterns = [
        (r'^Change\s+(\w+)\s+(to|from)\s+(\w+)(?:\s+(\d+))?\s*(?:\((.*?)\))?\s*$', True),  # with "Change"
        (r'^(\w+)\s+(to|from)\s+(\w+)(?:\s+(\d+))?\s*(?:\((.*?)\))?\s*$', False)  # without "Change"
    ]
    
    for pattern_str, has_change_prefix in patterns:
        match = re.match(pattern_str, name, re.IGNORECASE)
        if match:
            if has_change_prefix:
                from_keyword = match.group(1)
                direction = match.group(2).lower()  # "to" or "from"
                to_keyword = match.group(3)
                number = match.group(4)
                description = match.group(5)
            else:
                from_keyword = match.group(1)
                direction = match.group(2).lower()  # "to" or "from"
                to_keyword = match.group(3)
                number = match.group(4)
                description = match.group(5)
            
            # Check if keywords are allowed
            if (from_keyword.lower() in ALLOWED_KEYWORDS and 
                to_keyword.lower() in ALLOWED_KEYWORDS):
                
                # Build name based on pattern
                base_name = ""
                if has_change_prefix:
                    base_name = f"Change {from_keyword} {direction} {to_keyword}"
                else:
                    base_name = f"{from_keyword} {direction} {to_keyword}"
                
                # Add number if present
                if number:
                    new_name = f"{base_name} {number}"
                else:
                    new_name = base_name
                
                new_object = description if description else ""
                
                return new_name, new_object, original_remark
    
    return name, '', original_remark
    
def should_move_te_zu(name):
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

def move_te_zu_to_object(df):
    mask = df['Name'].apply(should_move_te_zu)
    
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

def export_to_excel(all_rows, output_path, proj_choice):
    """DataFrame ကို Excel Sheets များအဖြစ် ထုတ်သည်။"""
    # Column headers
    if proj_choice in ["Custom_UTM", "WGS84_UTM"]:
        coord_cols = ["Easting", "Northing"]
    elif proj_choice == "WGS84_LatLon":
        coord_cols = ["Longitude", "Latitude"]
    else:
        coord_cols = ["X_Coord", "Y_Coord"]

    # Original headers without EastingMM, NorthingMM
    other_headers = ["No.", "Name", coord_cols[0], coord_cols[1], "Remark", "Source"]
    
    # Word headers with EastingMM, NorthingMM for To Word and aa sheets only
    word_headers = ["No.", "Name", "Object", coord_cols[0], coord_cols[1], "EastingMM", "NorthingMM", "Remark", "Source"]
        
    df_all_raw = pd.DataFrame(all_rows, columns=other_headers)
    
    # UPDATED: Helper function to check if name has only protected underscores
    def has_only_protected_underscore(name):
        name_str = str(name)
        
        # Step 1: Remove all text within parentheses (brackets) including the parentheses
        name_without_brackets = re.sub(r'\([^)]*\)', '', name_str).strip()
        
        # Step 2: Remove all protected name underscores from consideration
        temp_name = name_without_brackets
        for protected_name in PROTECTED_SPLIT_NAMES_WORD:
            protected_lower = protected_name.lower()
            temp_name_lower = temp_name.lower()
             
            # Remove the protected name part from temp name for underscore counting
            if temp_name_lower.startswith(protected_lower):
                temp_name = temp_name[len(protected_name):].strip()
                break
        
        # Step 3: Count remaining underscores after removing protected names and bracket content
        remaining_underscores = temp_name.count('_')
        
        # If no remaining underscores after removing protected names and bracket content, neglect it
        if remaining_underscores == 0:
            return True
        else:
            return False       
    

    # Helper function for new masking logic
    def has_road_keyword_and_suffix(name):
        name_lower = str(name).strip().lower()
        if not name_lower:
            return False
            
        for keyword in ROAD_KEYWORDS:
            keyword_lower = keyword.lower()
            
            # Find all occurrences
            pos = name_lower.find(keyword_lower)
            while pos != -1:
                # Check character before (if exists)
                before_ok = (pos == 0 or not name_lower[pos-1].isalnum() or name_lower[pos-1] == ' ')
                # Check character after
                after_pos = pos + len(keyword_lower)
                after_ok = (after_pos >= len(name_lower) or not name_lower[after_pos].isalnum() or name_lower[after_pos] == ' ')
                
                if before_ok and after_ok:
                    # Check if there's meaningful text after
                    remaining = name_lower[after_pos:].strip()
                    if remaining:
                        return True
                
                # Find next occurrence
                pos = name_lower.find(keyword_lower, pos + 1)
            
        return False

    # 1. Prepare To Word Sheet Logic
    new_sheet_rows_list = []
    grouped = df_all_raw.groupby('No.')

    # In export_to_excel function, where To Word sheet rows are selected:
    
    for feature_no, group in grouped:
        if group.empty: continue
            
        source_type = group['Source'].iloc[0].lower()
        feature_name = group['Name'].iloc[0].lower()
        
        # Check for road keywords FIRST (higher priority)
        has_road_keyword = any(keyword in feature_name for keyword in ROAD_KEYWORDS)
        has_polygon_keyword = any(keyword in feature_name for keyword in POLYGON_KEYWORDS)
        
        # Priority: Road > Polygon > Default
        if source_type == 'point':
            is_single_row_feature = True
            is_start_end_feature = False
        elif source_type == 'polygon':
            is_single_row_feature = True  # Polygon is always single
            is_start_end_feature = False
        elif source_type == 'line':
            if has_road_keyword:
                # Road keywords get start-end rows (2 rows)
                is_single_row_feature = False
                is_start_end_feature = True
            elif has_polygon_keyword:
                # Polygon keywords get single row
                is_single_row_feature = True
                is_start_end_feature = False
            else:
                # Default: start-end rows
                is_single_row_feature = False
                is_start_end_feature = True

        if is_single_row_feature:
            new_sheet_rows_list.append(group.iloc[0].to_dict())
        elif is_start_end_feature:
            if len(group) == 1:
                new_sheet_rows_list.append(group.iloc[0].to_dict())
            elif len(group) > 1:
                new_sheet_rows_list.append(group.iloc[0].to_dict())
                end_row = group.iloc[-1].to_dict()
                end_row['No.'] = ''
                end_row['Remark'] = ''  
                new_sheet_rows_list.append(end_row)

    df_word_filtered = pd.DataFrame(new_sheet_rows_list, columns=other_headers)
    df_word_filtered.insert(2, 'Object', '')
    
    # Add EastingMM and NorthingMM for To Word sheet
    df_word_filtered["EastingMM"] = df_word_filtered.apply(
        lambda row: round_coordinate_for_phrase(str(row[coord_cols[0]]), "easting"), 
        axis=1
    )
    df_word_filtered["NorthingMM"] = df_word_filtered.apply(
        lambda row: round_coordinate_for_phrase(str(row[coord_cols[1]]), "northing"), 
        axis=1
    )
    
    # Reorder columns for To Word sheet
    df_word_filtered = df_word_filtered[["No.", "Name", "Object", coord_cols[0], coord_cols[1], "EastingMM", "NorthingMM", "Remark", "Source"]]
    
    # Move te zu to object
    df_word_filtered = move_te_zu_to_object(df_word_filtered)
    
    # Apply Name Splitting and Masking Logic for 'To Word'
    protected_names_lower = [name.lower() for name in PROTECTED_SPLIT_NAMES_WORD]
    
    has_underscore = df_word_filtered['Name'].str.contains('_', na=False)
    has_road_suffix = df_word_filtered['Name'].apply(has_road_keyword_and_suffix)
    has_change = df_word_filtered['Name'].apply(has_change_pattern) 
    mask_to_process = (has_underscore | has_road_suffix | has_change) & ~df_word_filtered['Name'].str.lower().isin(protected_names_lower)  # UPDATED & not_te_zu_word  # UPDATED

    if mask_to_process.any():
        split_results = df_word_filtered.loc[mask_to_process].apply(
            lambda row: apply_to_word_name_split(row, row['Source']), 
            axis=1, 
            result_type='expand'
        )
        split_results.columns = ['Name_New', 'Object_New', 'Remark_New']
        df_word_filtered.loc[mask_to_process, 'Name'] = split_results['Name_New']
        df_word_filtered.loc[mask_to_process, 'Object'] = split_results['Object_New']
        df_word_filtered.loc[mask_to_process, 'Remark'] = split_results['Remark_New']

    # Re-index 'No.' sequentially for 'To Word' 
    if not df_word_filtered.empty:
        current_no = 1
        temp_no_list = []
        for index, row in df_word_filtered.iterrows():
            if row['No.'] != '':
                temp_no_list.append(current_no)
                current_no += 1
            else:
                temp_no_list.append('')
        df_word_filtered['No.'] = temp_no_list
    
    # 2. Prepare Excel Sheet (without EastingMM, NorthingMM) - UPDATED
    #not_te_zu = ~df_all_raw['Name'].str.contains(r'te\s?zu', flags=re.IGNORECASE, regex=True, na=False)
    has_underscore_raw = df_all_raw['Name'].str.contains('_', na=False)
    has_only_protected_underscore_col = df_all_raw['Name'].apply(has_only_protected_underscore)     
         
    # NEW: Check if feature is line or polygon
    is_line_or_polygon = df_all_raw['Source'].str.lower().isin(['line', 'polygon'])

    # Excel sheet အတွက် filter
    keep_final_excel = (
        # 1. Line/Polygon အကုန် (point မဟုတ်ရင်)
        is_line_or_polygon |
        
        # 2. Point မှာ underscore ပါရင် (te zu မဟုတ်ရင်၊ protected underscore only မဟုတ်ရင်)
        (~is_line_or_polygon & has_underscore_raw & ~has_only_protected_underscore_col)
    )

    df_excel_modified = df_all_raw[keep_final_excel].copy()
    
    # Apply re-indexing/sequential numbering for 'Excel' sheet
    df_excel_modified['No.'] = df_excel_modified['No.'].astype(object)
    mask_excel = df_excel_modified['No.'].ne(df_excel_modified['No.'].shift()).fillna(True)
    df_excel_modified.loc[~mask_excel, ['No.', 'Name']] = ''
    
    current_no = 1
    temp_no_list = []
    for index, row in df_excel_modified.iterrows():
        if row['No.'] != '':
            temp_no_list.append(current_no)
            current_no += 1
        else:
            temp_no_list.append('')
    df_excel_modified['No.'] = temp_no_list

    # 3. Prepare 'aa' Sheet (with EastingMM, NorthingMM)
    df_aa_final = df_all_raw[keep_final_excel].copy()
    df_aa_final.insert(2, 'Object', '')
    
    # Add EastingMM and NorthingMM for aa sheet
    df_aa_final["EastingMM"] = df_aa_final.apply(
        lambda row: round_coordinate_for_phrase(str(row[coord_cols[0]]), "easting"), 
        axis=1
    )
    df_aa_final["NorthingMM"] = df_aa_final.apply(
        lambda row: round_coordinate_for_phrase(str(row[coord_cols[1]]), "northing"), 
        axis=1
    )
    
    # Reorder columns for aa sheet
    df_aa_final = df_aa_final[["No.", "Name", "Object", coord_cols[0], coord_cols[1], "EastingMM", "NorthingMM", "Remark", "Source"]]
    
    # Move te zu to object
    df_aa_final = move_te_zu_to_object(df_aa_final)
    
    # Apply Name Splitting and Masking Logic for 'aa' Sheet
    has_underscore_aa = df_aa_final['Name'].str.contains('_', na=False)
    has_road_suffix_aa = df_aa_final['Name'].apply(has_road_keyword_and_suffix)  
    has_change_aa = df_aa_final['Name'].apply(has_change_pattern)  # NEW

    #mask_to_process_aa = (has_underscore_aa | has_road_suffix_aa) & ~df_aa_final['Name'].str.lower().isin(protected_names_lower)
    mask_to_process_aa = (has_underscore_aa | has_road_suffix_aa | has_change_aa) & ~df_aa_final['Name'].str.lower().isin(protected_names_lower)  # UPDATED

    if mask_to_process_aa.any():
        split_results_aa = df_aa_final.loc[mask_to_process_aa].apply(
            lambda row: apply_to_word_name_split(row, row['Source']), 
            axis=1, 
            result_type='expand'
        )
        split_results_aa.columns = ['Name_New', 'Object_New', 'Remark_New']
        df_aa_final.loc[mask_to_process_aa, 'Name'] = split_results_aa['Name_New']
        df_aa_final.loc[mask_to_process_aa, 'Object'] = split_results_aa['Object_New']
        df_aa_final.loc[mask_to_process_aa, 'Remark'] = split_results_aa['Remark_New']
        
    # Apply re-indexing/sequential numbering for 'aa' sheet
    df_aa_final['No.'] = df_aa_final['No.'].astype(object)
    mask_aa_blank = df_aa_final['No.'].ne(df_aa_final['No.'].shift()).fillna(True)
    df_aa_final.loc[~mask_aa_blank, ['No.', 'Name']] = ''
    
    current_no_aa = 1
    temp_no_list_aa = []
    for index, row in df_aa_final.iterrows():
        if row['No.'] != '':
            temp_no_list_aa.append(current_no_aa)
            current_no_aa += 1
        else:
            temp_no_list_aa.append('')
    df_aa_final['No.'] = temp_no_list_aa

    # 4. Prepare Raw Data Sheet (Original Indexing without EastingMM, NorthingMM)
    df_raw_modified = df_all_raw.copy()
    df_raw_modified['No.'] = df_raw_modified['No.'].astype(object)  
    mask_raw = df_raw_modified['No.'].ne(df_raw_modified['No.'].shift()).fillna(True)
    df_raw_modified.loc[~mask_raw, ['No.', 'Name']] = ''
    
    # 5. Prepare Pts Sheet (without EastingMM, NorthingMM) - UPDATED    
    df_points_raw = df_all_raw[df_all_raw['Source'].str.lower() == 'point'].copy()
    has_underscore_pts = df_points_raw['Name'].str.contains('_', na=False)    
    has_only_protected_underscore_pts = df_points_raw['Name'].apply(has_only_protected_underscore)      
    
    # Pts sheet အတွက် filter
    filter_mask_pts = (
        # Point မှာ underscore ပါရင် (te zu မဟုတ်ရင်၊ protected underscore only မဟုတ်ရင်)
        has_underscore_pts & ~has_only_protected_underscore_pts
    )
    df_points = df_points_raw[filter_mask_pts].copy()
    
    # Re-index 'No.' sequentially for the filtered 'Pts' sheet
    df_points['No.'] = range(1, len(df_points) + 1)
    
    # Export to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_word_filtered.to_excel(writer, sheet_name='To Word', index=False, header=word_headers)
        df_aa_final.to_excel(writer, sheet_name='aa', index=False, header=word_headers)
        df_excel_modified.to_excel(writer, sheet_name='Excel', index=False, header=other_headers)
        df_raw_modified.to_excel(writer, sheet_name='Raw Data', index=False, header=other_headers)
        df_points.to_excel(writer, sheet_name='Pts', index=False, header=other_headers)
        
        # Apply Styling
        def apply_styling(ws):
            for row in ws.iter_rows():
                for cell in row:
                    cell.font = Font(name="Pyidaungsu", size=13)  
                    if cell.row == 1:
                        cell.alignment = Alignment(horizontal="center")
                    else:
                        cell.alignment = Alignment(horizontal="left")

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

        for sheet_name in writer.sheets:
            apply_styling(writer.sheets[sheet_name])

class ProcessingThread(QThread):
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)    
    
    def __init__(self, input_file, output_file, proj_choice, zone_mode, manual_zone):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.proj_choice = proj_choice
        self.zone_mode = zone_mode
        self.manual_zone = manual_zone
    
    def run(self):
        try:
            self.message.emit("ဒေတာဖိုင်ဖတ်နေသည်...")
            rows = process_vector(self.input_file, self.proj_choice, self.zone_mode, self.manual_zone)
            self.progress.emit(30)
            
            if not rows:
                self.finished.emit(False, "ရွေးချယ်ထားသော ဖိုင်မှ Point, Line, သို့မဟုတ် Polygon Data များကို ရှာမတွေ့ပါ။")
                return
            
            self.message.emit("Excel ဖိုင်ထုတ်နေသည်...")
            export_to_excel(rows, self.output_file, self.proj_choice)
            self.progress.emit(100)
            
            self.finished.emit(True, f"အချက်အလက်များ အောင်မြင်စွာ ပြုပြင်ပြီး Excel ဖိုင် (၅)ခု ထုတ်ပြီးပါပြီ။\nသိမ်းဆည်းထားသော လမ်းကြောင်း: {self.output_file}")
            
        except Exception as e:
            self.finished.emit(False, f"Data ပြုပြင်နေစဉ် အမှားရှိခဲ့သည်:\n{str(e)}")

class GeoProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Geospatial Data Processor Tool")
        self.setGeometry(100, 100, 600, 500)
        
        # Set dark theme
        self.set_dark_theme()
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        # Title
        title = QLabel("Geospatial Data Processor Tool")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18pt; font-weight: bold; color: #ffffff; margin: 10px;")
        layout.addWidget(title)
        
        # Input File Section
        input_group = QGroupBox("Input Data ဖိုင်")
        input_layout = QHBoxLayout(input_group)
        
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setPlaceholderText("Input data file path...")
        input_layout.addWidget(self.input_file_edit)
        
        self.browse_input_btn = QPushButton("ရွေးချယ်ရန်")
        self.browse_input_btn.clicked.connect(self.browse_input_file)
        input_layout.addWidget(self.browse_input_btn)
        
        layout.addWidget(input_group)
        
        # Output File Section
        output_group = QGroupBox("Output Excel ဖိုင်")
        output_layout = QHBoxLayout(output_group)
        
        self.output_file_edit = QLineEdit()
        self.output_file_edit.setPlaceholderText("Output Excel file path...")
        output_layout.addWidget(self.output_file_edit)
        
        self.browse_output_btn = QPushButton("သိမ်းရန်")
        self.browse_output_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(self.browse_output_btn)
        
        layout.addWidget(output_group)
        
        # Projection Section
        proj_group = QGroupBox("Coordinate စနစ်")
        proj_layout = QHBoxLayout(proj_group)
        
        self.proj_combo = QComboBox()
        self.proj_combo.addItem("MYANMAR Datum UTM (Easting/Northing)", "Custom_UTM")
        self.proj_combo.addItem("WGS84 UTM (Easting/Northing)", "WGS84_UTM")
        self.proj_combo.addItem("WGS84 (Lon/Lat)", "WGS84_LatLon")
        proj_layout.addWidget(self.proj_combo)
        
        layout.addWidget(proj_group)
        
        # NEW: Zone Selection Section
        zone_group = QGroupBox("Zone Selection")
        zone_layout = QVBoxLayout(zone_group)
        
        # Zone mode selection
        zone_mode_layout = QHBoxLayout()
        self.zone_auto_radio = QRadioButton("Auto")
        self.zone_manual_46_radio = QRadioButton("Manual: 46")
        self.zone_manual_47_radio = QRadioButton("Manual: 47")
        
        self.zone_auto_radio.setChecked(True)  # Default to Auto
        
        zone_mode_layout.addWidget(self.zone_auto_radio)
        zone_mode_layout.addWidget(self.zone_manual_46_radio)
        zone_mode_layout.addWidget(self.zone_manual_47_radio)
        zone_layout.addLayout(zone_mode_layout)
        
        # Threshold display (read-only)
        threshold_layout = QHBoxLayout()
        self.threshold_label = QLabel("MMD: 95.9968784133°")
        self.threshold_label.setStyleSheet("color: green; font-weight: bold;")
        threshold_layout.addWidget(QLabel("Threshold:"))
        threshold_layout.addWidget(self.threshold_label)
        threshold_layout.addStretch()
        zone_layout.addLayout(threshold_layout)
        
        layout.addWidget(zone_group)
        
        # Connect signals for zone selection
        self.zone_auto_radio.toggled.connect(self.on_zone_mode_changed)
        self.zone_manual_46_radio.toggled.connect(self.on_zone_mode_changed)  # NEW
        self.zone_manual_47_radio.toggled.connect(self.on_zone_mode_changed)  # NEW
        self.proj_combo.currentTextChanged.connect(self.on_projection_changed)
        
        # Progress Section
        progress_group = QGroupBox("လုပ်ဆောင်ချက်")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setReadOnly(True)
        progress_layout.addWidget(self.log_text)
        
        layout.addWidget(progress_group)
        
        # Process Button
        self.process_btn = QPushButton("စတင်ပြုပြင်မည်")
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setStyleSheet("QPushButton { background-color: #ff9800; color: white; font-weight: bold; padding: 10px; }")
        layout.addWidget(self.process_btn)

    def on_zone_mode_changed(self, checked):
        """Update threshold display based on zone mode"""
        if checked:  # Only update when the radio button is checked (not unchecked)
            if self.zone_manual_46_radio.isChecked():
                self.threshold_label.setText("Zone 46")
                self.threshold_label.setStyleSheet("color: orange; font-weight: bold;")
            elif self.zone_manual_47_radio.isChecked():
                self.threshold_label.setText("Zone 47") 
                self.threshold_label.setStyleSheet("color: orange; font-weight: bold;")
            else:  # Auto mode
                if self.proj_combo.currentData() == "Custom_UTM":
                    self.threshold_label.setText("MMD: 95.9968784133°")
                    self.threshold_label.setStyleSheet("color: green; font-weight: bold;")
                else:
                    self.threshold_label.setText("WGS: 96.0°")
                    self.threshold_label.setStyleSheet("color: green; font-weight: bold;")

    def on_projection_changed(self):
        """Update threshold display when projection changes"""
        # Only update if in auto mode
        if self.zone_auto_radio.isChecked():
            if self.proj_combo.currentData() == "Custom_UTM":
                self.threshold_label.setText("MMD: 95.9968784133°")
            else:
                self.threshold_label.setText("WGS: 96.0°")
        
    def set_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #555555;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #ff9800;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QPushButton {
                background-color: #00796b;
                color: white;
                padding: 5px 10px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #004d40;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 3px;
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QProgressBar {
                border: 1px solid #555555;
                border-radius: 3px;
                text-align: center;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #ff9800;
            }
            QTextEdit {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 3px;
            }
        """)
    
    def browse_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Input Data ဖိုင်ရွေးချယ်ရန်",
            "",
            "Geospatial Files (*.kmz *.kml *.shp *.geojson *.gpkg);;All Files (*.*)"
        )
        if file_path:
            self.input_file_edit.setText(file_path)
            dir_name = os.path.dirname(file_path)
            base_name = os.path.basename(file_path)
            file_name_without_ext = os.path.splitext(base_name)[0]
            default_output = os.path.join(dir_name, f"{file_name_without_ext}.xlsx")
            self.output_file_edit.setText(default_output)
    
    def browse_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Output Excel ဖိုင်သိမ်းဆည်းရန်",
            self.output_file_edit.text(),
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file_path:
            self.output_file_edit.setText(file_path)
    
    def process_data(self):
        input_file = self.input_file_edit.text()
        output_file = self.output_file_edit.text()
        proj_choice = self.proj_combo.currentData()
        
        if not input_file or not os.path.exists(input_file):
            QMessageBox.critical(self, "အမှား", "ကျေးဇူးပြု၍ Input Data ဖိုင်လမ်းကြောင်းကို မှန်ကန်စွာ ရွေးချယ်ပါ။")
            return
        
        if not output_file or not output_file.lower().endswith(('.xlsx', '.xls')):
            QMessageBox.critical(self, "အမှား", "ကျေးဇူးပြု၍ Output ဖိုင်အမည်ကို .xlsx (သို့) .xls ဖြင့် အဆုံးသတ်ပါ။")
            return
        
        if not GEOSPATIAL_LIBS_READY:
            QMessageBox.critical(self, "Dependency Error", 
                            "Geopandas, pyproj, သို့မဟုတ် openpyxl ကို ရှာမတွေ့ပါ။\n'pip install geopandas openpyxl pandas pyproj' ဖြင့် install လုပ်ရန် လိုအပ်ပါသည်။")
            return
        
        # Check if manual zone is selected
        if self.zone_manual_46_radio.isChecked() or self.zone_manual_47_radio.isChecked():
            zone_name = "46" if self.zone_manual_46_radio.isChecked() else "47"
            
            # Show warning
            warning_msg = QMessageBox()
            warning_msg.setIcon(QMessageBox.Warning)
            warning_msg.setWindowTitle("Manual Zone Selected")
            warning_msg.setText(f"<b>Manual Zone {zone_name} ရွေးထားပါတယ်</b>")
            warning_msg.setInformativeText(
                f"သတိပြုရန်: Manual Zone {zone_name} ကို ရွေးထားပါတယ်။\n\n"
                f"ဒေတာထဲမှာရှိတဲ့ longitude တန်ဖိုးတွေကို မစစ်ဆေးတော့ပါ။\n"
                f"မှားယွင်းတဲ့ zone သုံးမိရင် coordinates တွေ မှားနိုင်ပါတယ်။\n\n"
                f"ဆက်လက်လုပ်ဆောင်ပါမည်လား?"
            )
            warning_msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            warning_msg.setDefaultButton(QMessageBox.No)
            
            reply = warning_msg.exec_()
            if reply != QMessageBox.Yes:
                return
        
        # Disable process button during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        # Get zone parameters from UI - ဒီအပိုင်း အရေးကြီးပါတယ်!
        if self.zone_auto_radio.isChecked():
            zone_mode = "auto"
            manual_zone = None
        elif self.zone_manual_46_radio.isChecked():
            zone_mode = "manual" 
            manual_zone = 46
        else:  # manual 47
            zone_mode = "manual"
            manual_zone = 47
        
        print(f"Debug: zone_mode = {zone_mode}, manual_zone = {manual_zone}")  # Debug အတွက်
        
        # Start processing thread with zone parameters
        self.thread = ProcessingThread(input_file, output_file, proj_choice, zone_mode, manual_zone)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.message.connect(self.log_text.append)
        self.thread.finished.connect(self.processing_finished)
        self.thread.start()
    
    def processing_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "အောင်မြင်ပါသည်", message)
        else:
            QMessageBox.critical(self, "လုပ်ဆောင်ချက် အမှား", message)

def main():
    app = QApplication(sys.argv)
    
    # Cross-platform font handling
    if sys.platform == "win32":
        font_family = "Pyidaungsu"
    elif sys.platform == "linux":
        font_family = "Noto Sans Myanmar"
    else:
        font_family = "Sans Serif"
    
    # Global application font
    app_font = QFont(font_family, 10)
    app.setFont(app_font)
    
    # Global stylesheet with dynamic font
    app.setStyleSheet(f"""
        * {{
            font-family: '{font_family}';
        }}
        QMainWindow {{
            background-color: #2b2b2b;
            color: #ffffff;
        }}
        QGroupBox {{
            font-weight: bold;
            border: 2px solid #555555;
            border-radius: 5px;
            margin-top: 1ex;
            padding-top: 10px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
            color: #ff9800;
        }}
        QLineEdit {{
            padding: 5px;
            border: 1px solid #555555;
            border-radius: 3px;
            background-color: #3c3c3c;
            color: #ffffff;
        }}
        QPushButton {{
            background-color: #00796b;
            color: white;
            padding: 5px 10px;
            border: none;
            border-radius: 3px;
        }}
        QPushButton:hover {{
            background-color: #004d40;
        }}
        QComboBox {{
            padding: 5px;
            border: 1px solid #555555;
            border-radius: 3px;
            background-color: #3c3c3c;
            color: #ffffff;
        }}
        QProgressBar {{
            border: 1px solid #555555;
            border-radius: 3px;
            text-align: center;
            color: white;
        }}
        QProgressBar::chunk {{
            background-color: #ff9800;
        }}
        QTextEdit {{
            background-color: #3c3c3c;
            color: #ffffff;
            border: 1px solid #555555;
            border-radius: 3px;
        }}
    """)
    
    window = GeoProcessorApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()