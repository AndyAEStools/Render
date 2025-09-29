import pandas as pd
import xml.etree.ElementTree as ET
import os
import re

short_to_long = {
    'N': 'North', 'NE': 'Northeast', 'E': 'East', 'SE': 'Southeast',
    'S': 'South', 'SW': 'Southwest', 'W': 'West', 'NW': 'Northwest'
}
direction_to_angle = {k: v for k, v in zip(short_to_long.values(), [0, 45, 90, 135, 180, 225, 270, 315])}
angle_to_direction = {v: k for k, v in direction_to_angle.items()}

def convert_compass(direction):
    return short_to_long.get(direction, direction)

def normalize_orientation(raw):
    if pd.isna(raw):
        return ""
    val = str(raw).strip().replace(" ", "").capitalize()
    return convert_compass(val)

def normalize_roof_pitch(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip().lower()
    if val_str in ["horizontal", "vertical"]:
        return val_str.capitalize()
    try:
        pitch = int(float(val))
        return f"_{pitch}" if pitch in [30, 45, 60] else None
    except:
        return None

def mirror_orientation_full(original_angle, dwelling_angle):
    local = (original_angle - dwelling_angle) % 360
    mirrored_local = (360 - local) % 360
    mirrored_global = (mirrored_local + dwelling_angle) % 360
    return mirrored_global

def mirror_orientation_name(current_name, dwelling_angle):
    if current_name not in direction_to_angle:
        return current_name
    mirrored_angle = mirror_orientation_full(direction_to_angle[current_name], dwelling_angle)
    return angle_to_direction.get(mirrored_angle, current_name)

def process_xmls(excel_file, xml_folder, output_folder):
    df = pd.read_excel(excel_file, header=1)
    df.columns = df.columns.str.strip()
    df = df[df["XML Filename"].notna() & df["Dwelling Orientation"].notna()]
    df["XML Filename"] = df["XML Filename"].astype(str).str.strip()
    df["Dwelling Orientation"] = df["Dwelling Orientation"].astype(str).str.strip()
    df = df[(df["XML Filename"] != "") & (df["Dwelling Orientation"].str.capitalize().isin(short_to_long.keys()) | df["Dwelling Orientation"].isin(short_to_long.values()))]

    if df["AES Reference"].duplicated().any():
        duplicate_refs = df[df["AES Reference"].duplicated(keep=False)]["AES Reference"].tolist()
        raise ValueError(f"Duplicate AES References detected: {set(duplicate_refs)}")

    for index, row in df.iterrows():
        file_name = os.path.join(xml_folder, row["XML Filename"].strip() + ("" if row["XML Filename"].strip().lower().endswith(".xml") else ".xml"))
        if not os.path.exists(file_name):
            continue

        tree = ET.parse(file_name)
        root = tree.getroot()
        assessment_element = root.find("Assessment")

        connotation_val = row["Connotation"] if pd.notna(row["Connotation"]) else None
        if connotation_val and assessment_element is not None:
            prop_type_elem = assessment_element.find("PropertyType2")
            if prop_type_elem is not None:
                val = str(connotation_val).strip().upper()
                if val == "END":
                    prop_type_elem.text = "EndTerrace"
                elif val == "SEMI":
                    prop_type_elem.text = "SemiDetached"

        if pd.notna(row["Sheltered Sides"]) and assessment_element is not None:
            try:
                sheltered_sides_elem = assessment_element.find("ShelteredSides")
                if sheltered_sides_elem is not None:
                    sheltered_sides_elem.text = str(int(float(row["Sheltered Sides"])))
            except:
                pass

        if assessment_element is not None:
            if pd.notna(row["Plot Number"]):
                assess_ref_elem = assessment_element.find("Reference")
                if assess_ref_elem is not None:
                    assess_ref_elem.text = str(row["Plot Number"]).strip()

            orientation_elem = assessment_element.find("DwellingOrientation")
            if orientation_elem is not None and orientation_elem.text:
                original = orientation_elem.text.strip()
                new_orientation = normalize_orientation(row["Dwelling Orientation"])
                orientation_elem.text = new_orientation

                if original in direction_to_angle and new_orientation in direction_to_angle:
                    rotation_angle = (direction_to_angle[new_orientation] - direction_to_angle[original]) % 360

                    for opening in root.findall(".//Openings/Opening"):
                        orient_elem = opening.find("Orientation")
                        if orient_elem is not None and orient_elem.text in direction_to_angle:
                            rotated = (direction_to_angle[orient_elem.text] + rotation_angle) % 360
                            orient_elem.text = angle_to_direction.get(rotated, orient_elem.text)

                    for pv_unit in root.findall(".//PhotovoltaicUnits/PhotovoltaicUnit"):
                        pv_orient_elem = pv_unit.find("Orientation")
                        if pv_orient_elem is not None:
                            pv_orient_elem.text = normalize_orientation(row["Roof Orientation (PV orientation)"])

                        elevation_elem = pv_unit.find("Elevation")
                        elevation_val = normalize_roof_pitch(row["Roof Pitch (PV pitch)"])
                        if elevation_elem is not None and elevation_val:
                            elevation_elem.text = elevation_val

                        if elevation_elem is None or not elevation_val:
                            pass

                    if str(row["AS/OP"]).strip().upper() == "OP":
                        for opening in root.findall(".//Openings/Opening"):
                            orient_elem = opening.find("Orientation")
                            if orient_elem is not None and orient_elem.text in direction_to_angle:
                                original_orient = orient_elem.text
                                mirrored_orient = mirror_orientation_name(original_orient, direction_to_angle[new_orientation])
                                orient_elem.text = mirrored_orient

        plot_element = root.find("Plot")
        if plot_element is not None:
            if pd.notna(row["AES Reference"]):
                ref_elem = plot_element.find("Reference")
                if ref_elem is not None:
                    ref_elem.text = str(row["AES Reference"]).strip()

            type_ref_elem = plot_element.find("TypeReference")
            house_name_elem = plot_element.find("HouseName")
            house_number_elem = plot_element.find("HouseNumber")

            final_type_ref = None
            if type_ref_elem is not None:
                original_text = type_ref_elem.text or ""
                final_type_ref = original_text

                if str(row["AS/OP"]).strip().upper() == "OP":
                    if re.search(r"\(.*?\)|\bAS\b", original_text):
                        final_type_ref = re.sub(r"\(.*?\)|\bAS\b", "(OP)", original_text)
                    else:
                        final_type_ref = original_text.strip() + " (OP)"
                type_ref_elem.text = final_type_ref

            if house_number_elem is not None and final_type_ref:
                house_number_elem.text = final_type_ref

            if house_name_elem is not None and pd.notna(row["Plot Number"]):
                house_name_elem.text = str(row["Plot Number"]).strip()

        output_file = os.path.join(output_folder, f"{row['AES Reference']}.xml")
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
