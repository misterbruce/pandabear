from datetime import datetime, timedelta
import colorsys
import json
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
import dateparser
import pandas as pd

# --- Output column order, forced explicitly so it can never drift or reorder ---
OUTPUT_COLUMNS = [
    "Member",
    "Work Email",
    "Group",
    "Start Date",
    "Start Time",
    "End Date",
    "End Time",
    "Theme Color",
    "Custom Label",
    "Unpaid Break (minutes)",
    "Notes",
    "Shared",
]

# --- Microsoft Teams Shifts "Theme Color" numbered labels (per Microsoft's ---
# --- Shifts Excel-import template, column H). Source: ---
# --- https://support.microsoft.com/en-us/teams/shifts/import-a-schedule-from-excel-to-shifts ---
# NOTE: Microsoft's own doc table lists "8." as "Blue" again (not "DarkBlue"),
# which looks like a documentation typo given the 9-12 pattern (DarkGreen,
# DarkPurple, DarkPink, DarkYellow). We're using "DarkBlue" for #8 to match
# that pattern, but you should sanity-check this one value against the
# dropdown in your real Teams "Import schedule" Excel template before
# trusting it for a real import.
TEAMS_COLOR_LABELS = {
    "White": "1. White",
    "Blue": "2. Blue",
    "Green": "3. Green",
    "Purple": "4. Purple",
    "Pink": "5. Pink",
    "Yellow": "6. Yellow",
    "Grey": "7. Grey",
    "Dark Blue": "8. DarkBlue",
    "Dark Green": "9. DarkGreen",
    "Dark Purple": "10. DarkPurple",
    "Dark Pink": "11. DarkPink",
    "Dark Yellow": "12. DarkYellow",
}


def get_time_by_index(col_index, start_col=2):
    minutes_offset = (col_index - start_col) * 15
    base_time = datetime.strptime("07:00", "%H:%M")
    actual_time = base_time + timedelta(minutes=minutes_offset)
    return actual_time.strftime("%H:%M")


# --- Reference swatches for each Teams shift-color category. These are the ---
# --- exact hex codes the original script was hand-tuned against; keeping ---
# --- them as anchors (rather than throwing them out) means colors that ---
# --- exactly match still behave exactly as before. NEW: any color that ---
# --- doesn't exactly match now gets classified to whichever anchor it's ---
# --- CLOSEST to, instead of silently defaulting to White. ---
_COLOR_ANCHORS_HEX = [
    ("FFFFFF", "White"),
    ("808080", "Grey"), ("D9D9D9", "Grey"), ("A6A6A6", "Grey"),
    ("C0C0C0", "Grey"), ("ECECEC", "Grey"), ("F2F2F2", "Grey"), ("7F7F7F", "Grey"),
    ("CC99FF", "Purple"), ("CCCCFF", "Purple"),
    ("800080", "Dark Purple"), ("7030A0", "Dark Purple"),
    ("FFFF00", "Yellow"), ("FFE699", "Yellow"),
    ("B3B300", "Dark Yellow"),
    ("DDEBF7", "Blue"), ("8FAADC", "Blue"), ("4472C4", "Blue"), ("00B0F0", "Blue"),
    ("2F5597", "Dark Blue"), ("0000FF", "Dark Blue"),
    ("00FF00", "Green"), ("E2EFDA", "Green"), ("C6E0B4", "Green"), ("A9D08E", "Green"),
    ("375623", "Dark Green"),
    ("FFC0CB", "Pink"),
    ("FF1493", "Dark Pink"), ("FF69B4", "Dark Pink"),
]


def _hex_to_rgb(hex6):
    hex6 = hex6[-6:]
    return tuple(int(hex6[i:i + 2], 16) for i in (0, 2, 4))


_COLOR_ANCHORS_RGB = [(_hex_to_rgb(h), label) for h, label in _COLOR_ANCHORS_HEX]


def is_blank_or_white(hex6, threshold=25):
    """
    Strict check for 'this cell has no meaningful color' -- no fill at all,
    or a color extremely close to pure white. Deliberately NOT the same as
    'nearest anchor is White': a pale-but-intentional shift color (e.g. a
    light tinted yellow) should still count as a real, working shift, not
    get swallowed as blank. Only near-exact white/no-fill counts as blank.
    """
    if hex6 is None:
        return True
    try:
        r, g, b = _hex_to_rgb(hex6)
    except Exception:
        return True
    dist = ((255 - r) ** 2 + (255 - g) ** 2 + (255 - b) ** 2) ** 0.5
    return dist <= threshold


def classify_color(hex6):
    """
    Classifies ANY 6-digit hex color to the nearest known Teams shift-color
    category by Euclidean RGB distance, instead of requiring an exact
    substring match. This means a color that's a slightly different shade
    of blue/green/purple/etc. than our hardcoded anchors still gets
    correctly classified, rather than silently falling through to White
    (which was the root cause of colors "always coming out White").
    """
    if hex6 is None:
        return "White"
    try:
        r, g, b = _hex_to_rgb(hex6)
    except Exception:
        return "White"
    best_label, best_dist = "White", float("inf")
    for (ar, ag, ab), label in _COLOR_ANCHORS_RGB:
        dist = (r - ar) ** 2 + (g - ag) ** 2 + (b - ab) ** 2
        if dist < best_dist:
            best_dist = dist
            best_label = label
    return best_label


def to_teams_theme_label(color_name):
    """Map a plain color name (e.g. 'Blue') to the numbered Teams import
    label (e.g. '2. Blue'). Falls back to the plain name if unrecognized,
    rather than silently dropping it."""
    return TEAMS_COLOR_LABELS.get(color_name, color_name)


# --- OOXML theme resolution -------------------------------------------------
# Excel's default "Office" theme, used only as a last-resort fallback if a
# workbook's own theme XML can't be read for any reason.
_DEFAULT_THEME_HEX = [
    "FFFFFF", "000000", "E7E6E6", "44546A",
    "4472C4", "ED7D31", "A5A5A5", "FFC000",
    "5B9BD5", "70AD47", "0563C1", "954F72",
]

_DRAWINGML_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def extract_theme_colors(workbook):
    """
    Reads the workbook's own embedded theme (xl/theme/theme1.xml, exposed by
    openpyxl as workbook.loaded_theme) and returns the 12 real theme colors
    in the order Excel uses for cell-style theme indices 0-11.

    This matters because a cell filled by clicking a swatch in Excel's
    "Theme Colors" row stores only a theme index (0-11) + an optional tint,
    NOT an RGB value. Every company's workbook can have a different custom
    theme, so those 12 colors have to be read from the actual file rather
    than assumed.
    """
    try:
        theme_xml = workbook.loaded_theme
        if not theme_xml:
            raise ValueError("no embedded theme")
        root = ET.fromstring(theme_xml)
        clr_scheme = root.find(".//a:clrScheme", _DRAWINGML_NS)
        color_map = {}
        for child in clr_scheme:
            tag = child.tag.split("}")[-1]  # dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
            srgb = child.find("a:srgbClr", _DRAWINGML_NS)
            sysclr = child.find("a:sysClr", _DRAWINGML_NS)
            if srgb is not None:
                color_map[tag] = srgb.get("val")
            elif sysclr is not None:
                color_map[tag] = sysclr.get("lastClr")
        # NOTE: Excel's cell-style theme indices swap dk1/lt1 and dk2/lt2
        # relative to the order they appear in the theme XML itself. This is
        # a documented OOXML quirk, not a typo.
        ordered_tags = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                         "accent3", "accent4", "accent5", "accent6",
                         "hlink", "folHlink"]
        colors = [color_map.get(t) for t in ordered_tags]
        if all(colors):
            return [c.upper() for c in colors]
    except Exception:
        pass
    return list(_DEFAULT_THEME_HEX)


def apply_tint(hex6, tint):
    """
    Applies an OOXML 'tint' value to a base theme RGB color -- this is how
    Excel produces the lighter/darker variants you see in the "Theme Colors"
    swatch grid (e.g. "Accent 1, Lighter 40%") from a single base color.
    """
    if not tint:
        return hex6
    r, g, b = (c / 255.0 for c in _hex_to_rgb(hex6))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return "{:02X}{:02X}{:02X}".format(
        int(round(r2 * 255)), int(round(g2 * 255)), int(round(b2 * 255))
    )


def resolve_fill_hex(fill, theme_colors):
    """
    Resolves ANY cell fill -- plain RGB, theme-based (index + tint), or the
    legacy indexed palette used by older files -- down to a plain 6-digit
    hex string. Returns None for "no fill" (blank cell). This is the piece
    that was missing before: previously only plain RGB fills were read, so
    theme-colored and indexed-colored cells always fell back to a default
    that resolved to White.
    """
    if fill is None or fill.fill_type is None or fill.fill_type == "none":
        return None
    start_color = fill.start_color
    if start_color is None:
        return None
    try:
        if start_color.type == "rgb":
            rgb = start_color.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                hex6 = rgb[-6:].upper()
                if hex6 != "000000" or rgb.upper() not in ("00000000",):
                    return hex6
            return None
        if start_color.type == "theme" and start_color.theme is not None:
            idx = start_color.theme
            if 0 <= idx < len(theme_colors):
                base_hex = theme_colors[idx]
                tint = getattr(start_color, "tint", 0) or 0
                return apply_tint(base_hex, tint)
            return None
        if start_color.type == "indexed" and start_color.indexed is not None:
            idx = start_color.indexed
            if 0 <= idx < len(COLOR_INDEX):
                val = COLOR_INDEX[idx]
                if val and len(val) >= 6:
                    return val[-6:].upper()
            return None
    except Exception:
        pass
    return None


def run_conversion(excel_path, json_path, output_path, sheet_name):
    """
    Encapsulated state-machine code engine. 
    Loads employees dynamically from the JSON database keys.
    Processes the exact sheet name targeted from the UI input.
    """
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            email_lookup_raw = json.load(f)
            email_lookup = {k.strip().lower(): v for k, v in email_lookup_raw.items()}
            allowed_employees_lower = list(email_lookup.keys())
    except (FileNotFoundError, json.JSONDecodeError):
        email_lookup = {}
        allowed_employees_lower = []

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Target sheet '{sheet_name}' was not found inside the uploaded workbook. Available sheets: {', '.join(wb.sheetnames)}")
        
    ws = wb[sheet_name]

    # Read this specific workbook's real theme palette once, up front, so
    # theme-colored fills (index + tint) can be resolved to real RGB below.
    theme_colors = extract_theme_colors(wb)

    start_date = []
    stop_date = []
    member = []       
    # shifts_per_day will now store a list of dicts for each employee to track time alongside column indices:
    # {"times": ['08:00 AM', '01:00 PM'], "col_idx": 6}
    shifts_per_day = [] 
    emails_per_day = []
    missing_matches = set()

    has_date_started = False

    for row_idx in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row_idx, column=1).value
        if cell_value is None or pd.isna(cell_value):
            continue
            
        excel_name_str = str(cell_value).strip()
        excel_name_lower = excel_name_str.lower()
        
        if excel_name_lower == "employee names":
            continue

        # If this text matches a known employee name, always treat it as an
        # employee row -- even if dateparser would otherwise mis-parse a name
        # like "May" or "Jan" as a date. This only helps disambiguation for
        # names already in the JSON database; names NOT yet in the database
        # that happen to look like a date/month could still be misclassified.
        is_known_employee = excel_name_lower in allowed_employees_lower

        is_date_parseable = False
        if not is_known_employee:
            try:
                if dateparser.parse(excel_name_str) is not None:
                    is_date_parseable = True
            except Exception:
                pass

        if not is_date_parseable and has_date_started:
            if excel_name_lower not in allowed_employees_lower:
                missing_matches.add(excel_name_str)
                
            member[-1].append(excel_name_str)
            
            safe_email_local_part = excel_name_lower.replace(" ", ".")
            employee_email = email_lookup.get(excel_name_lower, f"{safe_email_local_part}@missing_json_match.com")
            emails_per_day[-1].append(employee_email)
            
            START_COL = 2 
            END_COL = 41 
            
            employee_shifts = []
            in_shift = False
            shift_start_time = None
            shift_start_col = None
            # Tracks whether purple appeared inside the CURRENT shift block only
            # (reset every time a new shift block starts), not the whole row.
            current_shift_has_purple = False
            
            for col_idx in range(START_COL, END_COL + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                fill = cell.fill
                resolved_hex = resolve_fill_hex(fill, theme_colors)

                if is_blank_or_white(resolved_hex):
                    is_working = False
                    is_purple = False
                else:
                    category = classify_color(resolved_hex)
                    is_working = category != "Grey"
                    is_purple = category in ("Purple", "Dark Purple")
                
                if is_working and not in_shift:
                    in_shift = True
                    shift_start_time = get_time_by_index(col_idx, START_COL)
                    shift_start_col = col_idx  # Keep track of where this specific shift begins
                    current_shift_has_purple = is_purple
                elif is_working and in_shift:
                    if is_purple:
                        current_shift_has_purple = True
                elif not is_working and in_shift:
                    in_shift = False
                    shift_end_time = get_time_by_index(col_idx, START_COL) 
                    employee_shifts.append({
                        "times": [shift_start_time, shift_end_time],
                        "sample_col": shift_start_col,
                        "has_purple": current_shift_has_purple,
                    })
            
            if in_shift:
                shift_end_time = "17:00"
                employee_shifts.append({
                    "times": [shift_start_time, shift_end_time],
                    "sample_col": shift_start_col,
                    "has_purple": current_shift_has_purple,
                })
                
            # Only extend the LAST shift's end time to 6pm if purple showed up
            # within that last shift block itself (not anywhere else in the row).
            if employee_shifts and employee_shifts[-1]["has_purple"]:
                employee_shifts[-1]["times"][1] = "18:00"
                
            shifts_per_day[-1].append(employee_shifts)
            continue
            
        try:
            parsed_date = dateparser.parse(str(cell_value))
            if parsed_date is not None:
                # Manual (not strftime) formatting so month/day are never
                # zero-padded -- e.g. "9/21/2026", not "09/21/2026" -- and
                # this works identically on Windows and Linux (strftime's
                # no-zero-pad flag differs by platform: %-m on Linux/Mac vs
                # %#m on Windows).
                formatted_date = f"{parsed_date.month}/{parsed_date.day}/{parsed_date.year}"
                start_date.append(formatted_date)
                stop_date.append(formatted_date)
                
                member.append([])
                emails_per_day.append([])  
                shifts_per_day.append([])
                has_date_started = True
        except Exception:
            pass

    flat_rows = []
    for d_idx, day_val in enumerate(start_date):
        for m_idx, emp_name in enumerate(member[d_idx]):
            emp_email = emails_per_day[d_idx][m_idx]
            emp_shifts = shifts_per_day[d_idx][m_idx]
            
            target_row_idx = None
            for r in range(1, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val and str(val).strip().lower() == emp_name.lower():
                    target_row_idx = r
                    break
                    
            if not emp_shifts:
                continue

            for shift_info in emp_shifts:
                times_pair = shift_info["times"]
                active_col = shift_info["sample_col"]
                
                resolved_color_text = "White"
                if target_row_idx:
                    # DYNAMIC SAMPLING: Sample the cell exactly where this specific shift block is painted!
                    sample_cell = ws.cell(row=target_row_idx, column=active_col)
                    resolved_hex = resolve_fill_hex(sample_cell.fill, theme_colors)
                    resolved_color_text = classify_color(resolved_hex)

                flat_rows.append({
                    "Member": emp_name,
                    "Work Email": emp_email,
                    "Group": "Call Center",
                    "Start Date": day_val,
                    "Start Time": times_pair[0],
                    "End Date": day_val,
                    "End Time": times_pair[1],
                    "Theme Color": to_teams_theme_label(resolved_color_text),
                    "Custom Label": "",
                    "Unpaid Break (minutes)": "",
                    "Notes": "",
                    "Shared": "1. Shared",
                })

    if not flat_rows:
        return False, list(missing_matches)

    # Explicit columns= forces the exact header set AND order every time,
    # regardless of dict key order / pandas version behavior.
    export_df = pd.DataFrame(flat_rows, columns=OUTPUT_COLUMNS)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        export_df.to_excel(writer, sheet_name="Shifts", index=False)

    return True, list(missing_matches)