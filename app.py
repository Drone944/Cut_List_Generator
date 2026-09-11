import streamlit as st
import pandas as pd
from math import ceil

st.set_page_config(page_title="Wardrobe Cutting Calculator", page_icon="🪚", layout="wide")

# -----------------------------------------------------------------------------
# Excel-compatible helpers
# -----------------------------------------------------------------------------

def excel_round(value: float) -> int:
    """Excel ROUND(value, 0) for the positive dimensions used here."""
    return int(value + 0.5)


def excel_ceiling_half(value: float) -> float:
    """Excel CEILING(value, 0.5)."""
    return ceil(value / 0.5) * 0.5


def optional_number_input(label: str, default=None, key=None, min_value=0, max_value=100):
    """Return None for a blank optional Excel cell, otherwise the entered number."""
    text = st.text_input(label, "" if default is None else str(default), key=key)
    text = text.strip()
    if text == "":
        return None
    try:
        value = int(text)
        if value < min_value or value > max_value:
            st.warning(f"{label}: enter a value from {min_value} to {max_value}.")
            return None
        return value
    except ValueError:
        st.warning(f"{label}: enter a whole number or leave it blank.")
        return None


def format_edge_band(thickness, material):
    """Format an edge-band specification like the Excel sheet: 22*.8(BSW)."""
    return f"22*{thickness}({material})"


def get_edge_bands(excel_row, cabinet_material, exposed_color, door_colour, glossy_or_matt):
    """Return Edge L1/L2/W1/W2 using the exact edge-band pattern from Excel.

    Excel rows 18 onward use columns I:L for edge information. Blank Excel
    cells are represented as an empty string here.
    """
    normal_edge = format_edge_band(".8", cabinet_material)
    exposed_edge = format_edge_band(".8", exposed_color)

    # Carcass side legs: all four edges use the exposed colour.
    if excel_row in (19, 20):
        return exposed_edge, exposed_edge, exposed_edge, exposed_edge

    # Top, bottom and centre vertical pieces: first edge uses exposed colour;
    # remaining edges use cabinet laminate.
    if 21 <= excel_row <= 26:
        return exposed_edge, normal_edge, normal_edge, normal_edge

    # Shelves, double shelves and skirting.
    if 27 <= excel_row <= 36:
        if excel_row == 36:
            return exposed_edge, exposed_edge, exposed_edge, exposed_edge
        return normal_edge, normal_edge, normal_edge, normal_edge

    # Single- and double-door drawer components. Back panels have no edge band.
    if 40 <= excel_row <= 93:
        # 45, 51, 57, 63, 69, 75, 81, 87, 93 are back-panel rows.
        if (excel_row - 45) % 6 == 0 and excel_row >= 45:
            return "", "", "", ""
        return normal_edge, normal_edge, normal_edge, normal_edge

    # Door: Excel repeats the same edge specification on all four sides.
    if excel_row == 94:
        door_edge_thickness = "1.3" if glossy_or_matt == 1 else "2"
        door_edge = format_edge_band(door_edge_thickness, door_colour)
        return door_edge, door_edge, door_edge, door_edge

    # Pillar: L1, W1 and W2 are edged; L2 is blank.
    if excel_row == 95:
        door_edge_thickness = "1.3" if glossy_or_matt == 1 else "2"
        door_edge = format_edge_band(door_edge_thickness, door_colour)
        return door_edge, "", door_edge, door_edge

    # Skirting support pieces copy the shelf edge-band pattern.
    if excel_row == 96:
        return normal_edge, normal_edge, normal_edge, normal_edge

    # Dynamic rows used by the standalone app. Their numeric row is no longer
    # tied to a fixed Excel row, so the caller may encode the part category in
    # the row range. This fallback keeps ordinary generated pieces edged.
    if excel_row >= 1000:
        profile = excel_row // 1000
        if profile == 1:      # shelf / structural piece
            return normal_edge, normal_edge, normal_edge, normal_edge
        if profile == 2:      # back panel
            return "", "", "", ""
        if profile == 3:      # drawer component
            sub = excel_row % 1000
            if sub == 5:
                return "", "", "", ""
            return normal_edge, normal_edge, normal_edge, normal_edge
        if profile == 4:      # door
            door_edge_thickness = "1.3" if glossy_or_matt == 1 else "2"
            door_edge = format_edge_band(door_edge_thickness, door_colour)
            return door_edge, door_edge, door_edge, door_edge
        if profile == 5:      # pillar/support
            door_edge_thickness = "1.3" if glossy_or_matt == 1 else "2"
            door_edge = format_edge_band(door_edge_thickness, door_colour)
            return door_edge, "", door_edge, door_edge

    return "", "", "", ""


def get_carcass_material(
    excel_row,
    cabinet_material,
    exposed_color,
    exposed_side,
    material_type,
    door_colour,
    panel_thickness=18,
):
    """Return the material for carcass rows 19-35 using the WD CUTTING formulas.

    The Excel workbook does not use one material for both side legs:
    - Right side leg uses the exposed colour when the right side is exposed.
    - Left side leg uses the exposed colour when the left side is exposed.
    - With BOTH exposed, both side legs use the exposed colour.
    - With NONE exposed, both use the cabinet laminate.

    For material type 3, the workbook builds an MDF-labelled material string.
    """
    if excel_row == 19:  # Right side leg
        exposed = exposed_side in (1, 3)
    elif excel_row == 20:  # Left side leg
        exposed = exposed_side in (2, 3)
    else:
        exposed = False

    if material_type == 1:
        if excel_row in (19, 20) and exposed:
            return str(exposed_color)
        return cabinet_material

    if material_type == 3:
        if excel_row == 19:
            # Excel row 19: exposed -> LEFT(B6,3)&H4, then "mdf".
            material = (
                f"{str(cabinet_material)[:3]}{exposed_color}"
                if exposed else str(cabinet_material)
            )
            return f"{material}mdf"

        if excel_row == 20:
            # Excel row 20: exposed -> LEFT(B6,3)&H4; otherwise E6&B6.
            if exposed:
                return f"{str(cabinet_material)[:3]}{exposed_color}"
            return f"{panel_thickness}{cabinet_material}"

        return f"{cabinet_material}mdf"

    # This branch is not normally used by the current workbook, but keeps
    # the function safe if another material type is entered.
    return cabinet_material


def add_cut(
    cuts, excel_row, part_name, height, width, quantity, material,
    cabinet_material, exposed_color, door_colour, glossy_or_matt
):
    edge_l1, edge_l2, edge_w1, edge_w2 = get_edge_bands(
        excel_row, cabinet_material, exposed_color, door_colour, glossy_or_matt
    )
    cuts.append({
        "Excel Row": excel_row,
        "Part": part_name,
        "Height (mm)": height,
        "Width (mm)": width,
        "Qty": quantity,
        "Material": material,
        "Edge L1": edge_l1,
        "Edge L2": edge_l2,
        "Edge W1": edge_w1,
        "Edge W2": edge_w2,
    })


# -----------------------------------------------------------------------------
# Main calculation
# -----------------------------------------------------------------------------

def calculate_cut_list(inputs):
    # Main worksheet inputs (the names intentionally describe what they mean).
    room_name = inputs["room_name"]
    cabinet_height = inputs["cabinet_height"]
    cabinet_width = inputs["cabinet_width"]
    cabinet_depth = inputs["cabinet_depth"]
    wardrobe_quantity = inputs["wardrobe_quantity"]
    number_of_doors = inputs["number_of_doors"]
    exposed_color = inputs["exposed_color"]
    exposed_side = inputs["exposed_side"]
    skirting_type = inputs["skirting_type"]
    shelf_gap_from_door = inputs["shelf_gap_from_door"]
    back_panel_fitting = inputs["back_panel_fitting"]

    cabinet_laminate = inputs["cabinet_laminate"]
    material_type = inputs["material_type"]
    back_panel_material_type = inputs["back_panel_material_type"]
    panel_thickness = inputs["panel_thickness"]
    back_panel_thickness = inputs["back_panel_thickness"]
    back_panel_groove_distance = inputs["back_panel_groove_distance"]
    skirting_height = inputs["skirting_height"]
    door_colour = inputs["door_colour"]
    glossy_or_matt = inputs["glossy_or_matt"]

    single_shelf_counts = inputs["single_shelf_counts"]
    single_drawer_counts = inputs["single_drawer_counts"]
    single_drawer_heights = inputs["single_drawer_heights"]
    single_drawer_pack_overrides = inputs["single_drawer_pack_overrides"]
    drawer_depth_onset_from_door = inputs["drawer_depth_onset_from_door"]

    double_shelf_counts = inputs["double_shelf_counts"]
    double_drawer_counts = inputs["double_drawer_counts"]
    double_drawer_heights = inputs["double_drawer_heights"]
    double_drawer_pack_overrides = inputs["double_drawer_pack_overrides"]

    cuts = []

    add_cut_with_edges = add_cut

    def add_cut_row(excel_row, part_name, height, width, quantity, material):
        add_cut_with_edges(
            cuts, excel_row, part_name, height, width, quantity, material,
            cabinet_laminate, exposed_color, door_colour, glossy_or_matt
        )

    def add_dynamic_cut_row(profile, row_number, part_name, height, width, quantity, material):
        # Keep a unique internal row id while retaining the same visible part
        # information. The profile controls edge-band behavior.
        internal_row = profile * 1000 + row_number
        add_cut_with_edges(
            cuts, internal_row, part_name, height, width, quantity, material,
            cabinet_laminate, exposed_color, door_colour, glossy_or_matt
        )

    # Keep the cabinet laminate separate from the material displayed on each
    # cut. The side-leg material depends on which side of the wardrobe is
    # exposed.
    cabinet_material = cabinet_laminate
    back_material = f"{back_panel_thickness}mm{cabinet_laminate}"

    right_side_material = get_carcass_material(
        19,
        cabinet_material,
        exposed_color,
        exposed_side,
        material_type,
        door_colour,
        panel_thickness,
    )
    left_side_material = get_carcass_material(
        20,
        cabinet_material,
        exposed_color,
        exposed_side,
        material_type,
        door_colour,
        panel_thickness,
    )

    # -------------------------------------------------------------------------
    # Rows 19-26: carcass
    # -------------------------------------------------------------------------
    right_side_depth = (
        cabinet_depth
        if exposed_side in (1, 3)
        else cabinet_depth - back_panel_thickness if back_panel_fitting == 4 else cabinet_depth
    )
    left_side_depth = (
        cabinet_depth
        if exposed_side in (2, 3)
        else cabinet_depth - back_panel_thickness if back_panel_fitting == 4 else cabinet_depth
    )

    add_cut_row(
        19,
        f"{room_name} WB Side Leg R",
        cabinet_height,
        right_side_depth,
        wardrobe_quantity,
        right_side_material,
    )
    add_cut_row(
        20,
        f"{room_name} WB Side Leg L",
        cabinet_height,
        left_side_depth,
        wardrobe_quantity,
        left_side_material,
    )

    top_depth = cabinet_depth - back_panel_thickness if back_panel_groove_distance == 0 else cabinet_depth
    add_cut_row(21, f"{room_name} WB Top", cabinet_width - panel_thickness * 2,
            top_depth, wardrobe_quantity, cabinet_material)

    bottom_depth = cabinet_depth - back_panel_groove_distance - back_panel_thickness
    add_cut_row(22, f"{room_name} WB Bottom", cabinet_width - panel_thickness * 2,
            bottom_depth, wardrobe_quantity, cabinet_material)

    add_cut_row(23, f"{room_name} WB CENTRE VERTICAL",
            cabinet_height - skirting_height - panel_thickness * 2,
            bottom_depth, max(number_of_doors - 1, 0) * wardrobe_quantity, cabinet_material)

    centre_part_height = 2085 - 100 - 36 - 1018
    centre_part_depth = bottom_depth - shelf_gap_from_door
    add_cut_row(24, f"{room_name} WB CENTRE VERTICAL-PART ",
            centre_part_height, centre_part_depth, 1, cabinet_material)
    add_cut_row(25, f"{room_name} WB CENTRE VERTICAL-PART ",
            732, centre_part_depth, 0, cabinet_material)
    add_cut_row(26, f"{room_name} WB CENTRE VERTICAL-PART ",
            963, centre_part_depth, 0, cabinet_material)

    # -------------------------------------------------------------------------
    # Dynamic shelves
    # -------------------------------------------------------------------------
    # The workbook gives formulas for up to 5 doors. The pattern is:
    #   - 1 door: the whole opening is one shelf bay.
    #   - 2+ doors: first and last bays use W/N - 1.5*T.
    #   - interior bays use W/N - T.
    # This reproduces the workbook values for N=1..5 and extends the same
    # construction rule to any number of doors.
    shelf_depth = bottom_depth - shelf_gap_from_door

    if number_of_doors == 1:
        shelf_widths = [cabinet_width - panel_thickness * 2]
    else:
        edge_shelf_width = excel_round(
            cabinet_width / number_of_doors
            - (panel_thickness + panel_thickness / 2)
        )
        inner_shelf_width = excel_round(
            cabinet_width / number_of_doors - panel_thickness
        )
        shelf_widths = [edge_shelf_width]
        shelf_widths.extend([inner_shelf_width] * (number_of_doors - 2))
        shelf_widths.append(edge_shelf_width)

    # Individual-door shelves: exactly one entry for every door.
    for index in range(number_of_doors):
        shelf_count = single_shelf_counts[index]
        add_dynamic_cut_row(1,
            27 + index,
            f"{room_name} DOOR {index + 1} WB shelf",
            shelf_widths[index],
            shelf_depth,
            shelf_count * wardrobe_quantity,
            cabinet_material,
        )

    # -------------------------------------------------------------------------
    # Dynamic double-door shelves
    # -------------------------------------------------------------------------
    double_shelf_widths = []
    for index in range(number_of_doors - 1):
        double_shelf_widths.append(
            shelf_widths[index] + panel_thickness + shelf_widths[index + 1]
        )

    for index in range(number_of_doors - 1):
        add_dynamic_cut_row(1,
            27 + number_of_doors + index,
            f"{room_name} DOUBLE SHELF {index + 1}&{index + 2}",
            double_shelf_widths[index],
            shelf_depth,
            double_shelf_counts[index] * wardrobe_quantity,
            cabinet_material,
        )

    # -------------------------------------------------------------------------
    # Skirting
    # -------------------------------------------------------------------------
    skirting_width = (
        cabinet_width - panel_thickness * 2
        if skirting_type > 1
        else cabinet_width + 160
    )
    skirting_row = 27 + number_of_doors + max(number_of_doors - 1, 0)
    add_dynamic_cut_row(1,
        skirting_row,
        f"{room_name} SKIRTING",
        skirting_width,
        skirting_height,
        wardrobe_quantity,
        door_colour,
    )

    # -------------------------------------------------------------------------
    # Back panels
    # -------------------------------------------------------------------------
    # The original workbook has a hand-tuned 3-panel layout for <=5 doors.
    # Keep that exact behavior for those cases. For larger wardrobes, use a
    # scalable rule: cover the cabinet with rear panels aligned to the double
    # shelf spans, splitting the remaining width into panels as needed.
    back_panel_height = (
        cabinet_height - skirting_height
        if back_panel_groove_distance == 0
        else cabinet_height - skirting_height - 10
    )

    back_panel_start_row = skirting_row + 1

    if number_of_doors <= 5:
        # These branches reproduce the existing worksheet logic.
        if number_of_doors < 3:
            back_panel_1_width = cabinet_width - 20
        elif number_of_doors < 4:
            back_panel_1_width = (
                shelf_widths[0]
                + shelf_widths[1]
                + panel_thickness
                + panel_thickness
            )
        elif double_shelf_counts[1] * wardrobe_quantity > 0:
            back_panel_1_width = double_shelf_widths[1] + panel_thickness
        else:
            back_panel_1_width = double_shelf_widths[0] + panel_thickness

        if back_panel_fitting == 4 and exposed_side in (1, 4):
            back_panel_1_width += 9

        add_dynamic_cut_row(2,
            back_panel_start_row,
            f"{room_name} BACK PANEL 1",
            back_panel_height,
            back_panel_1_width,
            wardrobe_quantity,
            back_material,
        )

        if cabinet_width - 20 < back_panel_1_width:
            back_panel_2_base_width = 0
        elif number_of_doors == 3:
            # Match the WD CUTTING worksheet exactly. For a 3-door cabinet
            # BACK PANEL 2 uses the first single-door shelf width + panel
            # thickness. With 1500 mm width and 18 mm panels: 473 + 18 =
            # 491 mm, then the exposed-side adjustment (+9 for Left/Right)
            # produces the worksheet value of 500 mm.
            back_panel_2_base_width = shelf_widths[0] + panel_thickness
        elif number_of_doors < 5:
            if len(double_shelf_widths) > 1 and double_shelf_counts[1] * wardrobe_quantity > 0:
                back_panel_2_base_width = shelf_widths[0] + panel_thickness
            elif len(double_shelf_widths) > 2 and double_shelf_counts[2] * wardrobe_quantity > 0:
                back_panel_2_base_width = double_shelf_widths[2] + 18
            else:
                back_panel_2_base_width = cabinet_width - back_panel_1_width - 20
        else:
            back_panel_2_base_width = double_shelf_widths[2] + panel_thickness

        back_panel_2_width = back_panel_2_base_width

        # Match the worksheet exactly:
        #   exposed side = Both/None (code 4 in the original sheet) gets +20
        #   exposed side = Left or Right gets +9
        # These adjustments are independent of the back-panel fitting option.
        if exposed_side == 4 and back_panel_fitting == 4:
            back_panel_2_width += 20
        elif exposed_side in (1, 2):
            back_panel_2_width += 9

        back_panel_2_quantity = 2 if (
            back_panel_2_width * 2 + back_panel_1_width <= cabinet_width - 20
        ) else 1

        add_dynamic_cut_row(2,
            back_panel_start_row + 1,
            f"{room_name} BACK PANEL 2",
            back_panel_height,
            back_panel_2_width,
            back_panel_2_quantity * wardrobe_quantity,
            back_material,
        )

        back_panel_3_width = cabinet_width - (
            back_panel_1_width
            + back_panel_2_width * back_panel_2_quantity
            + 20
        )
        back_panel_3_quantity = 1 if back_panel_3_width > 50 else 0

        add_dynamic_cut_row(2,
            back_panel_start_row + 2,
            f"{room_name} BACK PANEL 3",
            back_panel_height,
            back_panel_3_width,
            back_panel_3_quantity * wardrobe_quantity,
            back_material,
        )
    else:
        # For >5 doors there is no source Excel formula. Use the scalable
        # cabinet rule rather than silently truncating to five doors.
        remaining_width = cabinet_width - 20
        panel_number = 1
        back_panel_widths = []

        # Group adjacent bays into panels while keeping panel widths sensible.
        # The target is roughly two bays per rear panel.
        index = 0
        while index < number_of_doors and remaining_width > 50:
            if index < number_of_doors - 1:
                candidate = double_shelf_widths[index] + panel_thickness
                bays_used = 2
            else:
                candidate = shelf_widths[index] + panel_thickness
                bays_used = 1

            if candidate <= remaining_width:
                back_panel_widths.append(candidate)
                remaining_width -= candidate
                index += bays_used
            else:
                back_panel_widths.append(max(0, remaining_width))
                remaining_width = 0
                index = number_of_doors

        for panel_index, panel_width in enumerate(back_panel_widths, start=1):
            if back_panel_fitting == 4:
                if panel_index == 1 and exposed_side in (1, 4):
                    panel_width += 9
                elif panel_index > 1 and exposed_side == 4:
                    panel_width += 20
                elif panel_index > 1 and exposed_side in (1, 2):
                    panel_width += 9

            add_dynamic_cut_row(2,
                back_panel_start_row + panel_index - 1,
                f"{room_name} BACK PANEL {panel_index}",
                back_panel_height,
                panel_width,
                wardrobe_quantity,
                back_material,
            )

    # -------------------------------------------------------------------------
    # Dynamic single-door drawers
    # -------------------------------------------------------------------------
    single_drawer_start_row = back_panel_start_row + (3 if number_of_doors <= 5 else len(back_panel_widths))

    for index in range(number_of_doors):
        drawer_count = single_drawer_counts[index]
        drawer_height = single_drawer_heights[index]
        drawer_pack_override = single_drawer_pack_overrides[index]
        drawer_depth = (
            cabinet_depth
            - drawer_depth_onset_from_door
            - back_panel_groove_distance
            - back_panel_thickness
            if drawer_count > 0 else 0
        )

        drawer_pack_quantity = (
            2 * wardrobe_quantity * drawer_count
            if drawer_pack_override is None
            else drawer_pack_override * wardrobe_quantity * drawer_count
        ) if drawer_count > 0 else 0

        side_height = drawer_depth - 20
        side_width = drawer_height - 30
        side_quantity = drawer_count * 2 * wardrobe_quantity
        row_base = single_drawer_start_row + index * 6
        door_label = f"DOOR {index + 1}"

        add_dynamic_cut_row(3,
            row_base,
            f"{room_name} WB{door_label} DRAWER PACK",
            drawer_depth,
            drawer_height,
            drawer_pack_quantity,
            cabinet_material,
        )
        add_dynamic_cut_row(3,
            row_base + 1,
            f"{room_name} WB{door_label} DR SIDE",
            side_height,
            side_width,
            side_quantity,
            cabinet_material,
        )

        if drawer_count > 0:
            reference_shelf_width = shelf_widths[index]
            drawer_back_height = (
                reference_shelf_width
                - 26
                - side_quantity * panel_thickness / (wardrobe_quantity * drawer_count)
                - drawer_pack_quantity * panel_thickness / (wardrobe_quantity * drawer_count)
            )
            drawer_back_quantity = drawer_count * wardrobe_quantity

            add_dynamic_cut_row(3,
                row_base + 2,
                f"{room_name} WB{door_label} DR BACK",
                drawer_back_height,
                side_width,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 3,
                f"{room_name} WB{door_label} DR IN FRONT",
                drawer_back_height,
                50,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 4,
                f"{room_name} WB{door_label} DR FRONT",
                drawer_back_height + 2 * panel_thickness + 26 - 4,
                drawer_height - 2,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 5,
                f"{room_name} WB{door_label} BACK PANEL",
                side_height - 10,
                drawer_back_height + panel_thickness,
                drawer_back_quantity,
                back_material,
            )

    # -------------------------------------------------------------------------
    # Dynamic double-door drawers
    # -------------------------------------------------------------------------
    double_drawer_start_row = single_drawer_start_row + number_of_doors * 6

    for index in range(number_of_doors - 1):
        drawer_count = double_drawer_counts[index]
        drawer_height = double_drawer_heights[index]
        drawer_pack_override = double_drawer_pack_overrides[index]
        drawer_depth = (
            cabinet_depth
            - drawer_depth_onset_from_door
            - back_panel_groove_distance
            - back_panel_thickness
            if drawer_count > 0 else 0
        )

        drawer_pack_quantity = (
            4 * wardrobe_quantity * drawer_count
            if drawer_pack_override is None
            else drawer_pack_override * wardrobe_quantity * drawer_count
        ) if drawer_count > 0 else 0

        side_height = drawer_depth - 20
        side_width = drawer_height - 30
        side_quantity = drawer_count * 2 * wardrobe_quantity
        row_base = double_drawer_start_row + index * 6
        pair_label = f"DOORS {index + 1}&{index + 2}"

        add_dynamic_cut_row(3,
            row_base,
            f"{room_name} WB{pair_label} DRAWER PACK",
            drawer_depth,
            drawer_height,
            drawer_pack_quantity,
            cabinet_material,
        )
        add_dynamic_cut_row(3,
            row_base + 1,
            f"{room_name} WB{pair_label} DR SIDE",
            side_height,
            side_width,
            side_quantity,
            cabinet_material,
        )

        if drawer_count > 0:
            reference_shelf_width = double_shelf_widths[index]
            drawer_back_height = (
                reference_shelf_width
                - 26
                - side_quantity * panel_thickness / (wardrobe_quantity * drawer_count)
                - drawer_pack_quantity * panel_thickness / (wardrobe_quantity * drawer_count)
            )
            drawer_back_quantity = side_quantity / 2

            add_dynamic_cut_row(3,
                row_base + 2,
                f"{room_name} WB{pair_label} DR BACK",
                drawer_back_height,
                side_width,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 3,
                f"{room_name} WB{pair_label} DR IN FRONT",
                drawer_back_height,
                80,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 4,
                f"{room_name} WB{pair_label} DR FRONT",
                drawer_back_height + 2 * panel_thickness + 26 - 4,
                drawer_height - 2,
                drawer_back_quantity,
                cabinet_material,
            )
            add_dynamic_cut_row(3,
                row_base + 5,
                f"{room_name} WB{pair_label} BACK PANEL",
                side_height - 10,
                drawer_back_height + panel_thickness,
                drawer_back_quantity,
                back_material,
            )

    # -------------------------------------------------------------------------
    # Doors and supports
    # -------------------------------------------------------------------------
    door_height = cabinet_height - skirting_height
    door_width = excel_ceiling_half(cabinet_width / number_of_doors) - (
        excel_ceiling_half((number_of_doors - 1) * 2 / number_of_doors)
        if exposed_side == 3 else
        excel_ceiling_half((number_of_doors + 1) * 2 / number_of_doors)
        if exposed_side == 4 else
        excel_ceiling_half(number_of_doors * 2 / number_of_doors)
    )

    door_row = double_drawer_start_row + max(number_of_doors - 1, 0) * 6
    add_dynamic_cut_row(4,
        door_row,
        f"{room_name} DOOR",
        door_height,
        door_width,
        number_of_doors,
        door_colour,
    )

    pillar_quantity = (
        wardrobe_quantity if exposed_side in (1, 2)
        else 0 if exposed_side == 3
        else 2 * wardrobe_quantity
    )
    add_dynamic_cut_row(5,
        door_row + 1,
        f"{room_name} PILLAR PIECE",
        door_height + skirting_height,
        80,
        pillar_quantity,
        door_colour,
    )

    support_quantity = max(number_of_doors - 1, 0)
    add_dynamic_cut_row(5,
        door_row + 2,
        "SKIRTING SUPPORT PIECES",
        cabinet_depth - 20,
        skirting_height,
        support_quantity,
        door_colour,
    )

    cut_list = pd.DataFrame(cuts)
    cut_list["Qty"] = pd.to_numeric(cut_list["Qty"], errors="coerce").fillna(0)
    cut_list = cut_list[cut_list["Qty"] > 0].copy()

    # -------------------------------------------------------------------------
    # Hardware
    # -------------------------------------------------------------------------
    hardware_category = inputs["hardware_category"]
    category_names = {1: "Normal", 2: "Soft", 3: "SS", 4: "SS soft"}
    hardware_category_name = category_names[hardware_category]
    screw_category = "Plated" if hardware_category in (1, 2) else "SS"

    hinge_0_crank_quantity = number_of_doors * 2 if number_of_doors < 3 else 4
    hinge_8_crank_quantity = number_of_doors * 2 - hinge_0_crank_quantity

    first_drawer_depth = (
        cabinet_depth - drawer_depth_onset_from_door - back_panel_groove_distance - back_panel_thickness
        if single_drawer_counts and single_drawer_counts[0] > 0 else 0
    )
    slider_length = (
        500 if first_drawer_depth > 500 else
        450 if first_drawer_depth > 450 else
        400 if first_drawer_depth > 400 else
        350 if first_drawer_depth > 350 else
        300 if first_drawer_depth > 300 else
        250 if first_drawer_depth > 250 else
        200
    )

    # Minifix follows the workbook's rule of 4 per structural/shelf piece.
    # For dynamic door counts, include all generated structural pieces.
    structural_quantities = [
        wardrobe_quantity,  # top
        wardrobe_quantity,  # bottom
        max(number_of_doors - 1, 0) * wardrobe_quantity,  # verticals
        wardrobe_quantity,  # centre vertical part
        wardrobe_quantity,  # individual shelves are added below
    ]
    minifix_quantity = sum(structural_quantities) * 4
    minifix_quantity += sum(single_shelf_counts) * wardrobe_quantity * 4
    minifix_quantity += sum(double_shelf_counts) * wardrobe_quantity * 4

    total_single_drawers = sum(single_drawer_counts)
    total_double_drawers = sum(double_drawer_counts)
    total_drawers = total_single_drawers + total_double_drawers

    rac_quantity = pillar_quantity * 3 + support_quantity * 2
    silicon_quantity = round((cabinet_height * cabinet_width) * 0.000010764 * 0.005, 2)
    cotton_quantity = round((cabinet_height * cabinet_width) * 0.000010764 * 0.002, 2)

    hardware = pd.DataFrame([
        [103, "Hinges 0 cranck", hardware_category_name, hinge_0_crank_quantity],
        [104, "Hinges 8 cranck", hardware_category_name, hinge_8_crank_quantity],
        [105, f"Slider {slider_length}", hardware_category_name, total_drawers],
        [106, "Hanger rod", "", None],
        [107, "Hanger road clip", "", None],
        [108, "Plinth adjustable leg", "", None],
        [109, "Minifix set", "", minifix_quantity],
        [110, "Knob", "", total_drawers],
        [111, "Big Handle", "", number_of_doors],
        [112, "Lock for drawer", "", total_drawers],
        [113, "Lock for Door", "", None],
        [114, '2" screw', screw_category, None],
        [115, '1.1/4 " screw', screw_category, minifix_quantity],
        [116, "5/8 screw", screw_category,
         hinge_0_crank_quantity * 12 + hinge_8_crank_quantity * 12 + total_drawers * 16 + rac_quantity * 4],
        [117, '1 " Screw', screw_category, None],
        [118, "fisher plug", "", None],
        [119, "Rac 2535", screw_category, rac_quantity],
        [120, "Silicon White", "", silicon_quantity],
        [121, "Thinner", "", silicon_quantity],
        [122, "Cotton waste (IN KG)", "", cotton_quantity],
    ], columns=["Excel Row", "Hardware", "Measurement / Category", "Quantity"])

    return cut_list, hardware

    # -------------------------------------------------------------------------
    # Hardware rows 103-122
    # -------------------------------------------------------------------------
    hardware_category = inputs["hardware_category"]
    category_names = {1: "Normal", 2: "Soft", 3: "SS", 4: "SS soft"}
    hardware_category_name = category_names[hardware_category]
    screw_category = "Plated" if hardware_category in (1, 2) else "SS"

    hinge_0_crank_quantity = number_of_doors * 2 if number_of_doors < 3 else 4
    hinge_8_crank_quantity = number_of_doors * 2 - hinge_0_crank_quantity

    first_drawer_depth = (
        cabinet_depth - drawer_depth_onset_from_door - back_panel_groove_distance - back_panel_thickness
        if single_drawer_counts[0] > 0 else 0
    )
    slider_length = (
        500 if first_drawer_depth > 500 else
        450 if first_drawer_depth > 450 else
        400 if first_drawer_depth > 400 else
        350 if first_drawer_depth > 350 else
        300 if first_drawer_depth > 300 else
        250 if first_drawer_depth > 250 else
        200
    )

    # Excel D109 sums E21:E35, excluding E25.
    quantities_by_excel_row = {
        row["Excel Row"]: row["Qty"] for row in cuts
    }
    minifix_base_rows = [21, 22, 23, 24, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35]
    minifix_quantity = sum(quantities_by_excel_row.get(row, 0) for row in minifix_base_rows) * 4

    total_single_drawers = sum(single_drawer_counts)
    total_double_drawers = sum(double_drawer_counts)
    total_drawers = total_single_drawers + total_double_drawers

    rac_quantity = pillar_quantity * 3 + support_quantity * 2
    silicon_quantity = round((cabinet_height * cabinet_width) * 0.000010764 * 0.005, 2)
    cotton_quantity = round((cabinet_height * cabinet_width) * 0.000010764 * 0.002, 2)

    hardware = pd.DataFrame([
        [103, "Hinges 0 cranck", hardware_category_name, hinge_0_crank_quantity],
        [104, "Hinges 8 cranck", hardware_category_name, hinge_8_crank_quantity],
        [105, f"Slider {slider_length}", hardware_category_name, total_drawers],
        [106, "Hanger rod", "", None],
        [107, "Hanger road clip", "", None],
        [108, "Plinth adjustable leg", "", None],
        [109, "Minifix set", "", minifix_quantity],
        [110, "Knob", "", total_drawers],
        [111, "Big Handle", "", number_of_doors],
        [112, "Lock for drawer", "", total_drawers],
        [113, "Lock for Door", "", None],
        [114, '2" screw', screw_category, None],
        [115, '1.1/4 " screw', screw_category, minifix_quantity],
        [116, "5/8 screw", screw_category,
         hinge_0_crank_quantity * 12 + hinge_8_crank_quantity * 12 + total_drawers * 16 + rac_quantity * 4],
        [117, '1 " Screw', screw_category, None],
        [118, "fisher plug", "", None],
        [119, "Rac 2535", screw_category, rac_quantity],
        [120, "Silicon White", "", silicon_quantity],
        [121, "Thinner", "", silicon_quantity],
        [122, "Cotton waste (IN KG)", "", cotton_quantity],
    ], columns=["Excel Row", "Hardware", "Measurement / Category", "Quantity"])

    return cut_list, hardware


# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------

st.markdown("""
<style>
    .block-container {
        max-width: 1500px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }

    .hero {
        padding: 1.8rem 2rem;
        border-radius: 20px;
        background: linear-gradient(135deg, rgba(255,255,255,.06), rgba(255,255,255,.015));
        border: 1px solid rgba(255,255,255,.10);
        margin-bottom: 1.5rem;
    }

    .hero h1 { margin: 0; font-size: 2.5rem; }
    .hero p { margin: .45rem 0 0; opacity: .7; font-size: 1rem; }

    .section-title {
        font-size: 1.55rem;
        font-weight: 700;
        margin: 1.7rem 0 .8rem;
    }

    .section-subtitle {
        opacity: .65;
        margin-top: -.45rem;
        margin-bottom: 1rem;
    }

    .summary-card {
        padding: 1rem 1.1rem;
        border-radius: 14px;
        border: 1px solid rgba(255,255,255,.09);
        background: rgba(255,255,255,.035);
    }

    .summary-label { opacity: .6; font-size: .82rem; }
    .summary-value { font-size: 1.35rem; font-weight: 700; margin-top: .2rem; }

    .door-badge {
        display: inline-block;
        padding: .28rem .65rem;
        border-radius: 999px;
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.10);
        font-size: .78rem;
        margin-bottom: .6rem;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 16px;
    }

    .result-title {
        font-size: 1.65rem;
        font-weight: 750;
        margin-top: 1.8rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>🪚 Wardrobe Cutting Calculator</h1>
    <p>Design the wardrobe, configure every door section, and generate the cutting list and hardware requirements.</p>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1. Wardrobe dimensions
# -----------------------------------------------------------------------------
st.markdown('<div class="section-title">1. Wardrobe details</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Start with the overall cabinet dimensions and quantity.</div>', unsafe_allow_html=True)

with st.container(border=True):
    row1 = st.columns(4)
    with row1[0]:
        room_name = st.text_input("Room", "GF1")
    with row1[1]:
        cabinet_height = st.number_input("Height (mm)", min_value=1, max_value=10000, value=2100)
    with row1[2]:
        cabinet_width = st.number_input("Width (mm)", min_value=1, max_value=10000, value=1500)
    with row1[3]:
        cabinet_depth = st.number_input("Depth (mm)", min_value=1, max_value=3000, value=560)

    row2 = st.columns(4)
    with row2[0]:
        wardrobe_quantity = st.number_input("Quantity", min_value=1, max_value=1000, value=1)
    with row2[1]:
        number_of_doors = st.number_input(
            "Number of doors",
            min_value=1,
            max_value=30,
            value=3,
            step=1,
            help="The entire door configuration below adapts automatically to this number.",
        )
    with row2[2]:
        skirting_height = st.number_input("Skirting height (mm)", min_value=0, max_value=500, value=100)
    with row2[3]:
        shelf_gap_from_door = st.number_input("Shelf gap from door (mm)", min_value=0, max_value=100, value=10)

number_of_doors = int(number_of_doors)

# Live summary
summary_cols = st.columns(5)
summary_values = [
    ("Doors", number_of_doors),
    ("Individual sections", number_of_doors),
    ("Double-door sections", max(number_of_doors - 1, 0)),
    ("Wardrobes", wardrobe_quantity),
    ("Overall size", f"{cabinet_width} × {cabinet_height} × {cabinet_depth} mm"),
]
for col, (label, value) in zip(summary_cols, summary_values):
    with col:
        st.markdown(
            f'<div class="summary-card"><div class="summary-label">{label}</div>'
            f'<div class="summary-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

# -----------------------------------------------------------------------------
# 2. Construction / material settings
# -----------------------------------------------------------------------------
st.markdown('<div class="section-title">2. Construction & materials</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">These settings control carcass materials, back fitting, edges, doors and hardware.</div>', unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("**Carcass & back panel**")
    row1 = st.columns(4)
    with row1[0]:
        cabinet_laminate = st.text_input("Cabinet laminate", "BSW")
    with row1[1]:
        material_type_options = {"Laminate": 1, "Paint": 2, "MDF": 3}
        material_type_label = st.selectbox("Cabinet material type", list(material_type_options.keys()))
        material_type = material_type_options[material_type_label]
    with row1[2]:
        back_panel_material_type_options = {"Laminate": 1, "Paint": 2, "MDF": 3}
        back_panel_material_type_label = st.selectbox("Back panel material type", list(back_panel_material_type_options.keys()))
        back_panel_material_type = back_panel_material_type_options[back_panel_material_type_label]
    with row1[3]:
        back_panel_fitting_options = {
            "Groove": 4,
            "Lcut": 1,
            "Inside Back Screw": 2,
            "Back Screw": 3,
        }
        back_panel_fitting_label = st.selectbox(
            "Back panel fitting",
            list(back_panel_fitting_options.keys()),
            help="Choose how the back panel is fitted to the wardrobe carcass.",
        )
        back_panel_fitting = back_panel_fitting_options[back_panel_fitting_label]

    row2 = st.columns(4)
    with row2[0]:
        panel_thickness = st.number_input("Panel thickness (mm)", min_value=1, max_value=50, value=18)
    with row2[1]:
        back_panel_thickness = st.number_input("Back panel thickness (mm)", min_value=0, max_value=30, value=8)
    with row2[2]:
        back_panel_groove_distance = st.number_input("Distance to back panel groove (mm)", min_value=0, max_value=100, value=19)
    with row2[3]:
        drawer_depth_onset_from_door = st.number_input("Drawer depth onset from door (mm)", min_value=0, max_value=300, value=50)

    st.divider()
    st.markdown("**Exterior finish**")
    row3 = st.columns(4)
    with row3[0]:
        exposed_color = st.number_input("Exposed colour", min_value=0, max_value=999999, value=21091)
    with row3[1]:
        exposed_side_options = {"Right": 1, "Left": 2, "Both": 3, "None": 4}
        exposed_side_label = st.selectbox("Exposed side", list(exposed_side_options.keys()), index=1)
        exposed_side = exposed_side_options[exposed_side_label]
    with row3[2]:
        door_colour = st.number_input("Door colour", min_value=0, max_value=999999, value=10510)
    with row3[3]:
        glossy_or_matt_options = {"Glossy": 1, "Matt": 2}
        glossy_or_matt_label = st.selectbox("Door finish", list(glossy_or_matt_options.keys()))
        glossy_or_matt = glossy_or_matt_options[glossy_or_matt_label]

    row4 = st.columns(3)
    with row4[0]:
        skirting_type_options = {"Full Length": 1, "Inside Cabinet": 2}
        skirting_type_label = st.selectbox("Skirting type", list(skirting_type_options.keys()), index=1)
        skirting_type = skirting_type_options[skirting_type_label]
    with row4[1]:
        hardware_category_options = {"Normal": 1, "Soft": 2, "SS": 3, "SS soft": 4}
        hardware_category_label = st.selectbox("Hardware category", list(hardware_category_options.keys()), index=3)
        hardware_category = hardware_category_options[hardware_category_label]
    with row4[2]:
        st.caption("Tip")
        st.info("Leave drawer pack override blank to use the default pack quantity.", icon="ℹ️")

# -----------------------------------------------------------------------------
# 3. Door configuration
# -----------------------------------------------------------------------------
st.markdown('<div class="section-title">3. Door configuration</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="section-subtitle">{number_of_doors} doors → {number_of_doors} individual sections + '
    f'{max(number_of_doors - 1, 0)} adjacent double-door sections.</div>',
    unsafe_allow_html=True,
)

single_shelf_counts = []
single_drawer_counts = []
single_drawer_heights = []
single_drawer_pack_overrides = []

def default_for(values, index, fallback):
    return values[index] if index < len(values) else fallback

# Individual doors
st.markdown("#### Individual doors")
for door_index in range(number_of_doors):
    door_number = door_index + 1
    default_shelf_count = default_for([2, 2, 2, 0, 0], door_index, 0)
    default_drawer_count = default_for([1, 0, 0, 0, 0], door_index, 0)
    default_pack = default_for([1, None, None, None, None], door_index, None)

    with st.container(border=True):
        st.markdown(f'<span class="door-badge">DOOR {door_number}</span>', unsafe_allow_html=True)
        cols = st.columns([1.15, 1.15, 1.15, 1.15])
        with cols[0]:
            shelf_count = st.number_input(
                "Shelves", min_value=0, max_value=30, value=default_shelf_count,
                step=1, key=f"door_{door_number}_shelves",
            )
        with cols[1]:
            drawer_count = st.number_input(
                "Drawers", min_value=0, max_value=20, value=default_drawer_count,
                step=1, key=f"door_{door_number}_drawers",
            )
        with cols[2]:
            drawer_height = st.number_input(
                "Drawer height (mm)", min_value=1, max_value=1000, value=200,
                step=1, disabled=drawer_count == 0, key=f"door_{door_number}_drawer_height",
            )
        with cols[3]:
            drawer_pack_override = optional_number_input(
                "Pack override (blank = default)", default_pack,
                key=f"door_{door_number}_pack", min_value=1, max_value=100,
            )

    single_shelf_counts.append(int(shelf_count))
    single_drawer_counts.append(int(drawer_count))
    single_drawer_heights.append(int(drawer_height))
    single_drawer_pack_overrides.append(drawer_pack_override)

# Double-door sections
double_shelf_counts = []
double_drawer_counts = []
double_drawer_heights = []
double_drawer_pack_overrides = []

if number_of_doors > 1:
    st.markdown("#### Double-door sections")
    for pair_index in range(number_of_doors - 1):
        left_door = pair_index + 1
        right_door = pair_index + 2
        default_double_shelf = default_for([0, 2, 0, 0], pair_index, 0)
        default_double_drawer = default_for([0, 1, 0, 0], pair_index, 0)
        default_double_pack = default_for([None, 3, None, None], pair_index, None)

        with st.container(border=True):
            st.markdown(
                f'<span class="door-badge">DOORS {left_door} & {right_door}</span>',
                unsafe_allow_html=True,
            )
            cols = st.columns([1.15, 1.15, 1.15, 1.15])
            with cols[0]:
                double_shelf_count = st.number_input(
                    "Shelves", min_value=0, max_value=30, value=default_double_shelf,
                    step=1, key=f"pair_{left_door}_{right_door}_shelves",
                )
            with cols[1]:
                double_drawer_count = st.number_input(
                    "Drawers", min_value=0, max_value=20, value=default_double_drawer,
                    step=1, key=f"pair_{left_door}_{right_door}_drawers",
                )
            with cols[2]:
                double_drawer_height = st.number_input(
                    "Drawer height (mm)", min_value=1, max_value=1000, value=200,
                    step=1, disabled=double_drawer_count == 0,
                    key=f"pair_{left_door}_{right_door}_height",
                )
            with cols[3]:
                double_pack_override = optional_number_input(
                    "Pack override (blank = default)", default_double_pack,
                    key=f"pair_{left_door}_{right_door}_pack", min_value=1, max_value=100,
                )

        double_shelf_counts.append(int(double_shelf_count))
        double_drawer_counts.append(int(double_drawer_count))
        double_drawer_heights.append(int(double_drawer_height))
        double_drawer_pack_overrides.append(double_pack_override)

# -----------------------------------------------------------------------------
# 4. Calculate
# -----------------------------------------------------------------------------
st.markdown('<div class="section-title">4. Generate cutting list</div>', unsafe_allow_html=True)

inputs = {
    "room_name": room_name,
    "cabinet_height": cabinet_height,
    "cabinet_width": cabinet_width,
    "cabinet_depth": cabinet_depth,
    "wardrobe_quantity": wardrobe_quantity,
    "number_of_doors": number_of_doors,
    "exposed_color": exposed_color,
    "exposed_side": exposed_side,
    "skirting_type": skirting_type,
    "shelf_gap_from_door": shelf_gap_from_door,
    "back_panel_fitting": back_panel_fitting,
    "cabinet_laminate": cabinet_laminate,
    "material_type": material_type,
    "back_panel_material_type": back_panel_material_type,
    "panel_thickness": panel_thickness,
    "back_panel_thickness": back_panel_thickness,
    "back_panel_groove_distance": back_panel_groove_distance,
    "skirting_height": skirting_height,
    "door_colour": door_colour,
    "glossy_or_matt": glossy_or_matt,
    "hardware_category": hardware_category,
    "single_shelf_counts": single_shelf_counts,
    "single_drawer_counts": single_drawer_counts,
    "single_drawer_heights": single_drawer_heights,
    "single_drawer_pack_overrides": single_drawer_pack_overrides,
    "drawer_depth_onset_from_door": drawer_depth_onset_from_door,
    "double_shelf_counts": double_shelf_counts,
    "double_drawer_counts": double_drawer_counts,
    "double_drawer_heights": double_drawer_heights,
    "double_drawer_pack_overrides": double_drawer_pack_overrides,
}

if st.button("🪚 Calculate cut list", type="primary", use_container_width=True):
    with st.spinner("Calculating cutting list..."):
        cut_list, hardware = calculate_cut_list(inputs)
    st.session_state["cut_list"] = cut_list
    st.session_state["hardware"] = hardware
    st.session_state["calculated_room"] = room_name

if "cut_list" in st.session_state:
    result_room = st.session_state.get("calculated_room", room_name)
    cut_list = st.session_state["cut_list"]
    hardware = st.session_state["hardware"]

    st.markdown('<div class="result-title">Cut list</div>', unsafe_allow_html=True)
    result_cols = st.columns(3)
    with result_cols[0]:
        st.metric("Cutting rows", len(cut_list))
    with result_cols[1]:
        st.metric("Total pieces", int(cut_list["Qty"].sum()) if not cut_list.empty else 0)
    with result_cols[2]:
        st.metric("Hardware items", len(hardware[hardware["Quantity"].fillna(0) > 0]))

    st.dataframe(cut_list, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download cut list CSV",
        cut_list.to_csv(index=False).encode("utf-8"),
        f"{result_room}_cut_list.csv",
        "text/csv",
        use_container_width=True,
    )

    st.markdown('<div class="result-title">Hardware</div>', unsafe_allow_html=True)
    st.dataframe(hardware, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download hardware CSV",
        hardware.to_csv(index=False).encode("utf-8"),
        f"{result_room}_hardware.csv",
        "text/csv",
        use_container_width=True,
    )
else:
    st.info("Configure the wardrobe above and click **Calculate cut list** when you're ready.", icon="🪚")
