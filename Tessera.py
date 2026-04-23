# -*- coding: utf-8 -*-
import rhinoscriptsyntax as rs
import Rhino.Geometry as rg

def save_data(key, value):
    rs.SetDocumentData("GridScript", key, str(value))

def load_data(key, default):
    val = rs.GetDocumentData("GridScript", key)
    if val is None or val == "":
        return default
    try:
        if '.' in val:
            return float(val)
        else:
            return int(val)
    except:
        return val

def ensure_layer(name, color, parent=None):
    if rs.IsLayer(name):
        return name
    else:
        if parent and rs.IsLayer(parent):
            return rs.AddLayer(name, color, parent)
        else:
            return rs.AddLayer(name, color)

def col_label(index):
    # Converts 0-based column index to Excel-style label: A, B, ..., Z, AA, AB, ...
    label = ""
    index += 1
    while index > 0:
        index -= 1
        label = chr(65 + (index % 26)) + label
        index //= 26
    return label

def get_boundary_curve():
    choice = rs.GetString(
        "Boundary input method",
        "ExistingShape",
        ["Draw3PointRectangle", "ExistingShape"]
    )
    if choice is None:
        return None

    # --- Option 1: 3-point rectangle ---
    if choice == "Draw3PointRectangle":
        rs.Command("_Rectangle _3Point", False)
        selected = rs.LastCreatedObjects()
        if not selected:
            rs.MessageBox("No rectangle was created.", 0, "Error")
            return None
        crv = selected[0]
        if not rs.IsCurveClosed(crv) or not rs.IsCurvePlanar(crv):
            rs.MessageBox("The created shape is not a valid closed planar curve.", 0, "Error")
            return None
        return crv

    # --- Option 2: pick existing shape ---
    else:
        crv = rs.GetObject(
            "Select a closed planar curve for outer boundary",
            rs.filter.curve
        )
        if crv is None:
            return None
        if not rs.IsCurveClosed(crv) or not rs.IsCurvePlanar(crv):
            rs.MessageBox("Please select a closed planar curve.", 0, "Error")
            return None
        return crv


def create_grid_from_selected_rectangle():
    group_name = "GridGroup"

    if rs.IsGroup(group_name):
        old_objs = rs.ObjectsByGroup(group_name)
        if old_objs:
            rs.DeleteObjects(old_objs)
        rs.DeleteGroup(group_name)

    # Get outer boundary
    outer_curve = get_boundary_curve()
    if outer_curve is None:
        return

    # Load persistent defaults
    border = load_data("border", 2)
    cols   = load_data("cols",   5)
    rows   = load_data("rows",   4)
    gap_x  = load_data("gap_x",  2)
    gap_y  = load_data("gap_y",  2)

    # Compute plane from the actual edges of the curve
    # Get the 4 corners of the rectangle from its segments
    geom_crv = rs.coercecurve(outer_curve)
    if not geom_crv:
        rs.MessageBox("Cannot read curve geometry.", 0, "Error")
        return

    # Explode into segments to get edge directions
    segs = geom_crv.DuplicateSegments()
    if not segs or len(segs) < 2:
        rs.MessageBox("Curve does not have enough segments.", 0, "Error")
        return

    # Use first segment to define X axis, second for Y axis
    seg0 = segs[0]
    seg1 = segs[1]

    pt0 = seg0.PointAtStart
    x_axis = seg0.PointAtEnd - pt0
    width = x_axis.Length
    x_axis.Unitize()

    y_axis = seg1.PointAtEnd - seg1.PointAtStart
    height = y_axis.Length
    y_axis.Unitize()

    # Build a reliable plane from the actual edge directions
    z_axis = rg.Vector3d.CrossProduct(x_axis, y_axis)
    plane = rg.Plane(pt0, x_axis, y_axis)

    # User inputs
    border = rs.GetReal("Border width (inward, equal on all sides)", border, minimum=0)
    if border is None:
        return
    if border * 2 >= width or border * 2 >= height:
        rs.MessageBox("Error: Border too large for the outer boundary size.", 0, "Input Error")
        return

    cols = rs.GetInteger("Number of columns", cols, minimum=1)
    if cols is None:
        return
    rows = rs.GetInteger("Number of rows", rows, minimum=1)
    if rows is None:
        return

    gap_x = rs.GetReal("Gap between columns", gap_x, minimum=0)
    if gap_x is None:
        return
    gap_y = rs.GetReal("Gap between rows", gap_y, minimum=0)
    if gap_y is None:
        return

    # Ask if user wants cell labels
    label_answer = rs.GetString("Add cell labels? (Yes/No)", "Yes", ["Yes", "No"])
    add_labels = label_answer and label_answer.lower() == "yes"

    # Save settings
    save_data("border", border)
    save_data("cols",   cols)
    save_data("rows",   rows)
    save_data("gap_x",  gap_x)
    save_data("gap_y",  gap_y)

    # Calculate cell size
    grid_w = width  - 2 * border
    grid_h = height - 2 * border

    cell_w = (grid_w - (cols - 1) * gap_x) / cols
    cell_h = (grid_h - (rows - 1) * gap_y) / rows

    if cell_w <= 0 or cell_h <= 0:
        rs.MessageBox("Error: Gaps or border too large for grid size.", 0, "Grid Creation Error")
        return

    unit_name = rs.UnitSystemName(abbreviate=True)

    msg = "Cell size: {:.3f} x {:.3f} {}\n\nDo you want to continue creating the grid?".format(
        cell_w, cell_h, unit_name
    )
    if rs.MessageBox(msg, 4 | 32, "Confirm Cell Size") != 6:
        return

    created_objs = []

    # pt0 is already the world-space origin of the grid (first corner of first segment)

    # Layers
    rect_layer  = ensure_layer("Rectangle",           (0,   0,   0))
    cell_layer  = ensure_layer("Rectangle::Cells",    (255, 0,   0), parent="Rectangle")
    if add_labels:
        label_layer = ensure_layer("Rectangle::Labels", (0,   100, 200), parent="Rectangle")

    # Auto text height: ~0.4% of the smaller cell dimension
    base_text_height = min(cell_w, cell_h) * 0.004

    # Create cells
    for i in range(cols):
        for j in range(rows):
            # Local offset from pt0 along plane axes
            offset_x = border + i * (cell_w + gap_x)
            offset_y = border + j * (cell_h + gap_y)

            # Corner of this cell in world space
            cell_origin = pt0 + plane.XAxis * offset_x + plane.YAxis * offset_y

            # Build a plane at the cell corner aligned to the boundary plane
            cell_plane = rg.Plane(cell_origin, plane.XAxis, plane.YAxis)

            rect = rs.AddRectangle(cell_plane, cell_w, cell_h)
            if rect:
                rs.ObjectLayer(rect, cell_layer)
                created_objs.append(rect)

            # Cell label
            if add_labels:
                label = col_label(i) + str(j + 1)

                text_height = base_text_height

                # Center of the cell in world space
                center_pt = cell_origin + plane.XAxis * (cell_w / 2.0) + plane.YAxis * (cell_h / 2.0)

                # Build a plane at the cell center aligned to the boundary plane
                text_plane = rg.Plane(center_pt, plane.XAxis, plane.YAxis)

                txt = rs.AddText(
                    label,
                    text_plane,
                    text_height,
                    justification=131072 + 2  # middle-center
                )
                if txt:
                    rs.ObjectLayer(txt, label_layer)
                    created_objs.append(txt)

    rs.ObjectLayer(outer_curve, rect_layer)
    created_objs.append(outer_curve)

    # Optional grouping
    group_answer = rs.GetString("Group all grid objects? (Yes/No)", "Yes", ["Yes", "No"])
    if group_answer and group_answer.lower() == "yes":
        if not rs.IsGroup(group_name):
            rs.AddGroup(group_name)
        rs.AddObjectsToGroup(created_objs, group_name)

    # Final summary
    total_cells = cols * rows
    msg = (
        "Done.\n"
        "-----------------------------\n"
        "Boundary:    {:.0f} x {:.0f} {}\n"
        "Border:      {:.0f} {}\n"
        "-----------------------------\n"
        "Columns:     {}\n"
        "Rows:        {}\n"
        "Total cells: {}\n"
        "-----------------------------\n"
        "Cell size:   {:.3f} x {:.3f} {}\n"
        "Labels:      {}"
    ).format(
        width, height, unit_name,
        border, unit_name,
        cols, rows, total_cells,
        cell_w, cell_h, unit_name,
        "Yes" if add_labels else "No"
    )
    rs.MessageBox(msg, 0, "Tessera")


create_grid_from_selected_rectangle()