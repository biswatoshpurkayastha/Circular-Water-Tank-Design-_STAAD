import sys
import math
import traceback
from collections import defaultdict

import numpy as np
from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QColor, QBrush, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QVBoxLayout, QFormLayout,
    QGridLayout, QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox,
    QComboBox, QMessageBox, QProgressBar, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QTabWidget, QAbstractItemView, QPlainTextEdit,
)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from matplotlib.collections import LineCollection, PolyCollection
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import pyvista as pv
from pyvistaqt import QtInteractor

try:
    from openstaadpy import os_analytical
except ImportError:
    os_analytical = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

APP_TITLE = "STAAD Tank Wall Explorer - Vertical Moment and Hoop Force"

# The load case box lists LC 1 to LC 20; any other load case or combination
# number can be typed into it.
LOAD_CASE_LIST_LIMIT = 20

# Both results are extracted in one pass; the Result Type selector only changes what is shown.
RESULT_MVERT = "Vertical Bending Moment (Mvert)"
RESULT_NHOOP = "Hoop Force (Nhoop)"
RESULT_TYPES = {
    RESULT_MVERT: {
        "key": "m_vertical", "symbol": "Mvert", "unit": "kN.m/m",
        "title": "Vertical Bending Moment Diagram",
        "sign_note": "Negative: inner-face tension | Positive: outer-face tension",
        "file_name": "Vertical_Bending_Moment",
        "section_title": "Bending Moment Diagram along the Tank Cross Section",
        "section_note": "Meridional moment (same as Mvert on the wall), drawn on the tension face",
        "positive_label": "Positive: outer-face tension",
        "negative_label": "Negative: inner-face tension",
    },
    RESULT_NHOOP: {
        "key": "n_hoop", "symbol": "Nhoop", "unit": "kN/m",
        "title": "Hoop Force Diagram",
        "sign_note": "Positive: hoop tension | Negative: hoop compression",
        "file_name": "Hoop_Force",
        "section_title": "Hoop Force Diagram along the Tank Cross Section",
        "section_note": "Hoop tension is drawn outside the outline, hoop compression inside",
        "positive_label": "Positive: hoop tension",
        "negative_label": "Negative: hoop compression",
    },
}
# Cross section through the tank axis: outline points closer than this are merged,
# and a change of direction larger than the kink angle starts a new part
# (wall to dome, wall to cone, cone to dome), so bins never straddle a junction.
SECTION_MIN_STEP = 0.02
SECTION_KINK_DEGREES = 20.0
VIEW_ALL = "Raw + Average + Envelope"
VIEW_ENVELOPE = "Binned Average + Envelope"
VIEW_AVERAGE = "Binned Average"
VIEW_RAW = "Raw Plate Results"
COLOR_POSITIVE = "#C62828"
COLOR_NEGATIVE = "#1565C0"
COLOR_PURPLE = "#4A148C"
COLOR_SELECTED = "#FFD600"


def angle_difference(a, b):
    a, b = float(a) % 360.0, float(b) % 360.0
    d = abs(a - b)
    return min(d, 360.0 - d)


def as_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple, np.ndarray)):
        return list(value)
    try:
        return list(value)
    except Exception:
        return []


def section_value(row, key):
    """Value of a plate on the cross section; the moment uses the face-based sign."""
    return row["section_m"] if key == "m_vertical" else row[key]


def order_meridian(points):
    """Orders (radius, elevation) points along the meridian, starting at the highest point."""
    remaining = points[np.lexsort((points[:, 0], -points[:, 1]))]
    ordered = [remaining[0]]
    remaining = remaining[1:]
    while len(remaining):
        distances = np.hypot(remaining[:, 0] - ordered[-1][0], remaining[:, 1] - ordered[-1][1])
        nearest = int(np.argmin(distances))
        if distances[nearest] >= SECTION_MIN_STEP:
            ordered.append(remaining[nearest])
        remaining = np.delete(remaining, nearest, axis=0)
    return np.asarray(ordered)


def split_meridian(points):
    """Splits the ordered outline into parts at gaps and at junctions (kinks)."""
    if len(points) < 2:
        return []
    steps = np.hypot(*np.diff(points, axis=0).T)
    gap = 2.5 * float(np.percentile(steps, 90))
    parts, current = [], [points[0]]
    for i in range(1, len(points)):
        if steps[i - 1] > gap:
            parts.append(current)
            current = [points[i]]
            continue
        current.append(points[i])
        if i < len(points) - 1 and steps[i] <= gap:
            a, b = points[i] - points[i - 1], points[i + 1] - points[i]
            turn = abs(math.degrees(math.atan2(a[0] * b[1] - a[1] * b[0], a[0] * b[0] + a[1] * b[1])))
            if turn > SECTION_KINK_DEGREES:
                parts.append(current)
                current = [points[i]]
    parts.append(current)
    return [np.asarray(part) for part in parts if len(part) >= 2]


def build_section(rows, bin_size, result_keys):
    """
    Builds the cross section through the tank axis from the plates on the cutting plane.

    Every plate is rotated about the tank axis into the cutting plane, so the section is
    described by (radius, elevation). Row "side" is +1 on the plane angle and -1 on the
    opposite side. The plate nodes give the outline; the plate results are averaged in
    bins of bin_size along the outline. Returns None when no outline can be formed.
    """
    if not rows:
        return None
    all_y = np.concatenate([row["coordinates"][:, 1] for row in rows])
    y_min, y_max = float(all_y.min()), float(all_y.max())
    inside = np.asarray([0.0, 0.5 * (y_min + y_max)])   # a point inside the tank, on its axis
    sides, r_max = [], 0.0
    for side in (1, -1):
        side_rows = [row for row in rows if row["side"] == side]
        nodes = {(round(math.hypot(x, z), 3), round(y, 3)) for row in side_rows for x, y, z in row["coordinates"]}
        if len(nodes) < 2:
            continue
        parts = split_meridian(order_meridian(np.asarray(sorted(nodes), dtype=float)))
        if not parts:
            continue
        # The mesh rarely has a node exactly on the tank axis, so an end that stops just
        # short of the axis (dome crown, centre of a slab) is carried to the axis.
        near_axis = 0.08 * max(float(part[:, 0].max()) for part in parts)
        for index, part in enumerate(parts):
            if 0.0 < part[0][0] <= near_axis:
                part = np.vstack([[0.0, part[0][1]], part])
            if 0.0 < part[-1][0] <= near_axis:
                part = np.vstack([part, [0.0, part[-1][1]]])
            parts[index] = part

        # One table of outline segments, with the outward normal of each segment.
        seg_a, seg_b, seg_normal, seg_part, seg_start, pieces, offset = [], [], [], [], [], [], 0.0
        for part_index, part in enumerate(parts):
            lengths = np.hypot(*np.diff(part, axis=0).T)
            cumulative = np.concatenate([[0.0], np.cumsum(lengths)])
            tangents = np.diff(part, axis=0) / lengths[:, None]
            normals = np.column_stack([tangents[:, 1], -tangents[:, 0]])
            middles = 0.5 * (part[:-1] + part[1:])
            normals[np.sum(normals * (middles - inside), axis=1) < 0.0] *= -1.0
            seg_a.append(part[:-1]); seg_b.append(part[1:]); seg_normal.append(normals)
            seg_part.append(np.full(len(lengths), part_index)); seg_start.append(cumulative[:-1])
            pieces.append({"points": part, "normals": normals, "cumulative": cumulative,
                           "length": float(cumulative[-1]), "offset": offset, "stations": []})
            offset += float(cumulative[-1])
            r_max = max(r_max, float(part[:, 0].max()))
        seg_a, seg_b, seg_normal = np.vstack(seg_a), np.vstack(seg_b), np.vstack(seg_normal)
        seg_part, seg_start = np.concatenate(seg_part), np.concatenate(seg_start)
        seg_vector = seg_b - seg_a
        seg_length_sq = np.sum(seg_vector ** 2, axis=1)

        # Project every plate onto the outline.
        for row in side_rows:
            point = np.asarray([row["radius"], row["elevation"]])
            t = np.clip(np.sum((point - seg_a) * seg_vector, axis=1) / seg_length_sq, 0.0, 1.0)
            feet = seg_a + t[:, None] * seg_vector
            j = int(np.argmin(np.hypot(feet[:, 0] - point[0], feet[:, 1] - point[1])))
            normal = seg_normal[j]
            centroid, local_z = row["centroid"], row["local_z"]
            normal_radial = (local_z[0] * centroid[0] + local_z[2] * centroid[2]) / row["radius"]
            # A positive plate moment is tension on the local +Z face. Refer it to the faces
            # of the tank instead, because local Z points in on some plates and out on others.
            orient = 1.0 if normal_radial * normal[0] + local_z[1] * normal[1] >= 0.0 else -1.0
            row["section_part"] = int(seg_part[j])
            row["section_local_s"] = float(seg_start[j] + t[j] * math.sqrt(seg_length_sq[j]))
            row["section_s"] = pieces[row["section_part"]]["offset"] + row["section_local_s"]
            row["section_point"], row["section_normal"] = feet[j].copy(), normal.copy()
            row["section_orient"] = orient
            row["section_m"] = row["m_vertical"] * orient

        # Bin the plates along each part of the outline.
        for part_index, piece in enumerate(pieces):
            members = [row for row in side_rows if row["section_part"] == part_index]
            bin_count = max(1, int(round(piece["length"] / bin_size)))
            groups = defaultdict(list)
            for row in members:
                groups[min(int(row["section_local_s"] / piece["length"] * bin_count), bin_count - 1)].append(row)
            for bin_index in sorted(groups):
                centre = (bin_index + 0.5) * piece["length"] / bin_count
                j = int(np.clip(np.searchsorted(piece["cumulative"], centre, side="right") - 1,
                                0, len(piece["normals"]) - 1))
                a, b = piece["points"][j], piece["points"][j + 1]
                position = a + (centre - piece["cumulative"][j]) / (piece["cumulative"][j + 1] - piece["cumulative"][j]) * (b - a)
                station = {
                    "side": side, "part": part_index, "s": piece["offset"] + centre,
                    "radius": float(position[0]), "elevation": float(position[1]),
                    "normal": piece["normals"][j].copy(), "count": len(groups[bin_index]),
                }
                for key in result_keys:
                    values = np.asarray([section_value(row, key) for row in groups[bin_index]], dtype=float)
                    station[key] = {"average": float(values.mean()), "minimum": float(values.min()),
                                    "maximum": float(values.max())}
                piece["stations"].append(station)
        sides.append({"side": side, "pieces": pieces, "rows": sorted(side_rows, key=lambda r: r["section_s"])})
    if not sides:
        return None
    return {"sides": sides, "y_min": y_min, "y_max": y_max, "r_max": r_max}


class NumericItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.data(Qt.UserRole)) < float(other.data(Qt.UserRole))
        except Exception:
            return super().__lt__(other)


def number_item(value, decimals=3):
    item = NumericItem(f"{float(value):.{decimals}f}")
    item.setData(Qt.UserRole, float(value))
    item.setTextAlignment(Qt.AlignCenter)
    return item


def integer_item(value):
    item = NumericItem(str(int(value)))
    item.setData(Qt.UserRole, int(value))
    item.setTextAlignment(Qt.AlignCenter)
    return item


class ExtractionWorker(QThread):
    progress = Signal(int, str)
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, y_min, y_max, load_case, target_angle, angle_tolerance, bin_size):
        super().__init__()
        self.y_min = float(y_min)
        self.y_max = float(y_max)
        self.load_case = int(load_case)
        self.target_angle = float(target_angle)
        self.angle_tolerance = float(angle_tolerance)
        self.bin_size = float(bin_size)

    def run(self):
        com_ready = False
        try:
            if pythoncom is not None:
                pythoncom.CoInitialize()
                com_ready = True
            if os_analytical is None:
                raise RuntimeError("openstaadpy is missing. Run: pip install openstaadpy")

            self.progress.emit(2, "Connecting to the STAAD model already open...")
            staad = os_analytical.connect()
            geo, out, prop = staad.Geometry, staad.Output, staad.Property
            if not out.AreResultsAvailable():
                raise RuntimeError("No analysis results found. Run STAAD analysis first.")

            # The number is picked without looking at the model, so confirm it exists.
            load = staad.Load
            available = [int(v) for v in as_list(load.GetPrimaryLoadCaseNumbers())]
            if int(load.GetLoadCombinationCaseCount()) > 0:
                available += [int(v) for v in as_list(load.GetLoadCombinationCaseNumbers())]
            if self.load_case not in available:
                raise RuntimeError(
                    f"Load case / combination {self.load_case} does not exist in the open STAAD model.\n"
                    f"Numbers available in this model: {', '.join(map(str, sorted(available)))}"
                )

            plate_ids = [int(v) for v in as_list(geo.GetPlateList())]
            if not plate_ids:
                raise RuntimeError("No plate elements found in the open STAAD model.")

            # Cache coordinates once. This is used for the whole-tank context mesh.
            node_cache = {}
            all_plates = []
            context_strip = []
            opposite_strip = []
            section_rows = []
            selected_rows = []
            elevation_bins = defaultdict(list)
            failed = 0
            total = len(plate_ids)
            plane_cos = math.cos(math.radians(self.target_angle))
            plane_sin = math.sin(math.radians(self.target_angle))

            for index, plate_id in enumerate(plate_ids):
                try:
                    incidence = [int(v) for v in as_list(geo.GetPlateIncidence(plate_id)) if int(v) > 0]
                    if len(incidence) < 3:
                        failed += 1
                        continue

                    coordinates = []
                    for node_id in incidence:
                        if node_id not in node_cache:
                            xyz = as_list(geo.GetNodeCoordinates(node_id))
                            if len(xyz) < 3:
                                raise ValueError(f"Invalid coordinates for node {node_id}")
                            node_cache[node_id] = np.asarray(xyz[:3], dtype=float)
                        coordinates.append(node_cache[node_id])
                    coordinates = np.asarray(coordinates, dtype=float)

                    # Save every plate before applying the BMD selection filters.
                    centroid = np.mean(coordinates, axis=0)
                    elevation = float(centroid[1])
                    plane_angle = math.degrees(math.atan2(centroid[2], centroid[0]))

                    base_plate = {
                        "plate_id": plate_id,
                        "node_ids": incidence.copy(),
                        "coordinates": coordinates.copy(),
                        "centroid": centroid.copy(),
                        "elevation": elevation,
                        "angle": plane_angle,
                    }
                    all_plates.append(base_plate)

                    # Full-height meridional strip, including top dome and bottom dome, and
                    # the strip on the opposite side of the tank. Together they are the
                    # cross section through the tank axis; the elevation diagram still
                    # obeys the Y limits and uses the plane-angle side only.
                    in_wedge = angle_difference(plane_angle, self.target_angle) <= self.angle_tolerance
                    if in_wedge:
                        side = 1
                    elif angle_difference(plane_angle, self.target_angle + 180.0) <= self.angle_tolerance:
                        side = -1
                    else:
                        # Near the tank axis the angular strip is narrower than a plate, so a plate
                        # that the cutting plane passes through also belongs to the cross section.
                        offsets = coordinates[:, 2] * plane_cos - coordinates[:, 0] * plane_sin
                        if not offsets.min() < 0.0 < offsets.max():
                            continue
                        side = 1 if centroid[0] * plane_cos + centroid[2] * plane_sin >= 0.0 else -1
                    (context_strip if side > 0 else opposite_strip).append(base_plate.copy())
                    in_strip = in_wedge and self.y_min <= elevation <= self.y_max

                    radius = math.hypot(centroid[0], centroid[2])
                    if radius <= 1.0e-9:
                        failed += 1
                        continue

                    c1, c2, c3 = coordinates[0], coordinates[1], coordinates[2]
                    edge_12, edge_13 = c2 - c1, c3 - c1
                    edge_length = np.linalg.norm(edge_12)
                    if edge_length <= 1.0e-12:
                        failed += 1
                        continue
                    local_x = edge_12 / edge_length
                    normal_raw = np.cross(edge_12, edge_13)
                    normal_length = np.linalg.norm(normal_raw)
                    if normal_length <= 1.0e-12:
                        failed += 1
                        continue
                    local_z = normal_raw / normal_length
                    local_y = np.cross(local_z, local_x)
                    local_y_length = np.linalg.norm(local_y)
                    if local_y_length <= 1.0e-12:
                        failed += 1
                        continue
                    local_y /= local_y_length

                    global_vertical = np.asarray([0.0, 1.0, 0.0])
                    vx = float(np.dot(global_vertical, local_x))
                    vy = float(np.dot(global_vertical, local_y))
                    if math.hypot(vx, vy) <= 1.0e-3:
                        # Horizontal plate (flat slab): the vertical has no component in the
                        # plate, so the meridional direction is the radial one.
                        global_vertical = np.asarray([centroid[0] / radius, 0.0, centroid[2] / radius])
                        vx = float(np.dot(global_vertical, local_x))
                        vy = float(np.dot(global_vertical, local_y))
                    theta_rad = math.atan2(vy, vx)
                    moments = as_list(out.GetAllPlateCenterMoments(plate_id, self.load_case))
                    if len(moments) < 3:
                        failed += 1
                        continue
                    mx, my, mxy = map(float, moments[:3])
                    # GetAllPlateCenterForces returns [SQX, SQY, SX, SY, SXY]; the last
                    # three are the membrane stresses in the plate local axes.
                    forces = as_list(out.GetAllPlateCenterForces(plate_id, self.load_case))
                    if len(forces) < 5:
                        failed += 1
                        continue
                    sx, sy, sxy = map(float, forces[2:5])
                    average_term = (mx + my) / 2.0
                    cos_2theta = math.cos(2.0 * theta_rad)
                    sin_2theta = math.sin(2.0 * theta_rad)
                    difference_term = ((mx - my) / 2.0) * cos_2theta
                    twisting_term = mxy * sin_2theta
                    m_vertical = average_term + difference_term + twisting_term

                    # Hoop direction: horizontal tangent to the circle around the tank
                    # axis (global Y through X = 0, Z = 0), projected into the plate.
                    global_hoop = np.asarray([-centroid[2] / radius, 0.0, centroid[0] / radius])
                    hx = float(np.dot(global_hoop, local_x))
                    hy = float(np.dot(global_hoop, local_y))
                    if math.hypot(hx, hy) <= 1.0e-9:
                        failed += 1
                        continue
                    alpha_rad = math.atan2(hy, hx)
                    thickness_values = [float(v) for v in as_list(prop.GetPlateThickness(plate_id)) if float(v) > 0.0]
                    if not thickness_values:
                        failed += 1
                        continue
                    thickness = float(np.mean(thickness_values))
                    cos_2alpha = math.cos(2.0 * alpha_rad)
                    sin_2alpha = math.sin(2.0 * alpha_rad)
                    hoop_average_term = (sx + sy) / 2.0
                    hoop_difference_term = ((sx - sy) / 2.0) * cos_2alpha
                    hoop_shear_term = sxy * sin_2alpha
                    sigma_hoop = hoop_average_term + hoop_difference_term + hoop_shear_term
                    n_hoop = sigma_hoop * thickness
                    bin_y = round(elevation / self.bin_size) * self.bin_size

                    row = {
                        "plate_id": plate_id, "load_case": self.load_case,
                        "elevation": elevation, "angle": plane_angle,
                        "theta_rad": theta_rad, "theta": math.degrees(theta_rad),
                        "vertical_component_local_x": vx,
                        "vertical_component_local_y": vy,
                        "global_vertical": global_vertical.copy(),
                        "local_x": local_x.copy(), "local_y": local_y.copy(),
                        "local_z": local_z.copy(), "mx": mx, "my": my, "mxy": mxy,
                        "cos_2theta": cos_2theta, "sin_2theta": sin_2theta,
                        "term_average": average_term,
                        "term_difference": difference_term,
                        "term_twisting": twisting_term,
                        "m_vertical": m_vertical, "bin_y": bin_y,
                        "radius": radius, "global_hoop": global_hoop.copy(),
                        "hoop_component_local_x": hx, "hoop_component_local_y": hy,
                        "alpha_rad": alpha_rad, "alpha": math.degrees(alpha_rad),
                        "sx": sx, "sy": sy, "sxy": sxy, "thickness": thickness,
                        "cos_2alpha": cos_2alpha, "sin_2alpha": sin_2alpha,
                        "hoop_term_average": hoop_average_term,
                        "hoop_term_difference": hoop_difference_term,
                        "hoop_term_shear": hoop_shear_term,
                        "sigma_hoop": sigma_hoop, "n_hoop": n_hoop,
                        "node_ids": incidence.copy(),
                        "coordinates": coordinates.copy(), "centroid": centroid.copy(),
                    }
                    row["side"] = side
                    section_rows.append(row)
                    if in_strip:
                        selected_rows.append(row)
                        elevation_bins[bin_y].append(row)
                except Exception:
                    failed += 1

                if index % 50 == 0 or index == total - 1:
                    pct = 5 + int(90.0 * (index + 1) / total)
                    self.progress.emit(
                        pct,
                        f"Read {index + 1}/{total} tank plates; selected strip: {len(selected_rows)}",
                    )

            if not selected_rows:
                raise RuntimeError(
                    "No plates matched the selected plane. Increase angular tolerance "
                    "or revise the elevation range."
                )

            selected_rows.sort(key=lambda r: (r["elevation"], r["plate_id"]))
            result_keys = [spec["key"] for spec in RESULT_TYPES.values()]
            bin_centers = [float(elevation) for elevation in sorted(elevation_bins)]
            binned_rows = {key: [] for key in result_keys}
            for elevation in sorted(elevation_bins):
                for key in result_keys:
                    vals = np.asarray([r[key] for r in elevation_bins[elevation]], dtype=float)
                    binned_rows[key].append({
                        "elevation": float(elevation), "average": float(np.mean(vals)),
                        "minimum": float(np.min(vals)), "maximum": float(np.max(vals)),
                        "standard_deviation": float(np.std(vals)), "count": len(vals),
                    })

            # Assign a stable categorical bin index to every selected plate.
            # Changing BIN_SIZE now changes the visible color bands in PyVista.
            bin_index_map = {center: index for index, center in enumerate(bin_centers)}
            for selected_row in selected_rows:
                selected_row["bin_center"] = float(selected_row["bin_y"])
                selected_row["bin_index"] = int(bin_index_map[selected_row["bin_y"]])

            statistics = {"failed_plates": failed, "total_model_plates": total}
            for key in result_keys:
                raw = np.asarray([r[key] for r in selected_rows], dtype=float)
                statistics[key] = {
                    "maximum": selected_rows[int(np.argmax(raw))],
                    "minimum": selected_rows[int(np.argmin(raw))],
                    "absolute": selected_rows[int(np.argmax(np.abs(raw)))],
                }
            result = {
                "parameters": {
                    "y_min": self.y_min, "y_max": self.y_max,
                    "load_case": self.load_case, "target_angle": self.target_angle,
                    "angle_tolerance": self.angle_tolerance, "bin_size": self.bin_size,
                },
                "rows": selected_rows,
                "bins": binned_rows,
                "bin_centers": bin_centers,
                "all_plates": all_plates,
                "context_strip": context_strip,
                "opposite_strip": opposite_strip,
                "section": build_section(section_rows, self.bin_size, result_keys),
                "statistics": statistics,
            }
            self.progress.emit(100, f"Loaded full tank: {len(all_plates)}; full-height strip: {len(context_strip)}; result plates: {len(selected_rows)}")
            self.completed.emit(result)
        except Exception:
            self.failed.emit(traceback.format_exc())
        finally:
            if com_ready and pythoncom is not None:
                pythoncom.CoUninitialize()


class BMDCanvas(FigureCanvas):
    plateSelected = Signal(int)

    def __init__(self, parent=None):
        self.figure = Figure(figsize=(9, 7), facecolor="white", constrained_layout=True)
        super().__init__(self.figure)
        self.setParent(parent)
        self.axes = self.figure.add_subplot(111)
        self.result = None
        self.spec = None
        self.raw_x = np.asarray([])
        self.raw_y = np.asarray([])
        self.raw_plate_ids = []
        self.annotation = None
        self.selected_marker = None
        self.mpl_connect("motion_notify_event", self.on_mouse_move)
        self.mpl_connect("button_press_event", self.on_mouse_click)
        self.draw_empty()

    def draw_empty(self):
        self.axes.clear()
        self.axes.text(0.5, 0.5, "Open an analyzed STAAD tank model and extract the results.",
                       transform=self.axes.transAxes, ha="center", va="center", fontsize=13)
        self.axes.axis("off")
        self.draw_idle()

    def set_result(self, result, mode, spec, show_zero=True, show_grid=True, show_values=False):
        self.result = result
        self.spec = spec
        key = spec["key"]
        self.axes.clear(); self.axes.axis("on")
        rows, bins, p = result["rows"], result["bins"][key], result["parameters"]
        self.raw_x = np.asarray([r[key] for r in rows], dtype=float)
        self.raw_y = np.asarray([r["elevation"] for r in rows], dtype=float)
        self.raw_plate_ids = [r["plate_id"] for r in rows]
        avg = np.asarray([r["average"] for r in bins])
        y = np.asarray([r["elevation"] for r in bins])
        mn = np.asarray([r["minimum"] for r in bins])
        mx = np.asarray([r["maximum"] for r in bins])

        if mode in (VIEW_ENVELOPE, VIEW_ALL):
            self.axes.fill_betweenx(y, mn, mx, color="#CE93D8", alpha=0.28, label="Raw envelope")
            self.axes.plot(mn, y, color="#AB47BC", linestyle=":", linewidth=1)
            self.axes.plot(mx, y, color="#AB47BC", linestyle=":", linewidth=1)
        if mode in (VIEW_RAW, VIEW_ALL):
            self.axes.scatter(self.raw_x, self.raw_y, c=self.raw_x, cmap="coolwarm", s=34,
                              edgecolors="#37474F", linewidths=0.45, alpha=0.82,
                              label="Individual plates", zorder=4)
        if mode in (VIEW_AVERAGE, VIEW_ENVELOPE, VIEW_ALL):
            self.axes.plot(avg, y, color=COLOR_PURPLE, linewidth=3, marker="o", markersize=6,
                           markerfacecolor="white", label="Binned average", zorder=5)
            self.axes.fill_betweenx(y, 0, avg, where=avg >= 0, color="#EF9A9A", alpha=0.22)
            self.axes.fill_betweenx(y, 0, avg, where=avg < 0, color="#90CAF9", alpha=0.23)
            if show_values:
                for value, elev in zip(avg, y):
                    self.axes.annotate(f"{value:.2f}", (value, elev), xytext=(6 if value >= 0 else -6, 0),
                                       textcoords="offset points", ha="left" if value >= 0 else "right",
                                       va="center", fontsize=8)

        maximum, minimum = result["statistics"][key]["maximum"], result["statistics"][key]["minimum"]
        self.axes.scatter([maximum[key]], [maximum["elevation"]], marker="^", s=125,
                          color=COLOR_POSITIVE, edgecolor="black", label="Maximum", zorder=8)
        self.axes.scatter([minimum[key]], [minimum["elevation"]], marker="v", s=125,
                          color=COLOR_NEGATIVE, edgecolor="black", label="Minimum", zorder=8)
        if show_zero:
            self.axes.axvline(0, color="#263238", linewidth=1.4, linestyle="--")
        if show_grid:
            self.axes.grid(True, linestyle=":", color="#B0BEC5")
        self.axes.set_ylim(p["y_min"], p["y_max"])
        self.axes.set_title(
            f"{spec['title']}\nAngle {p['target_angle']:.2f} deg +/- "
            f"{p['angle_tolerance']:.2f} deg | LC {p['load_case']}", fontweight="bold")
        self.axes.set_xlabel(f"{spec['symbol']} ({spec['unit']})\n{spec['sign_note']}", fontweight="bold")
        self.axes.set_ylabel("Elevation Y (m)", fontweight="bold")
        self.axes.legend(loc="best", fontsize=8.5)
        self.annotation = self.axes.annotate("", xy=(0, 0), xytext=(14, 14), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.45", facecolor="#FFFDE7", edgecolor="#6A1B9A"),
            arrowprops=dict(arrowstyle="->", color="#6A1B9A"), fontsize=9, zorder=20)
        self.annotation.set_visible(False)
        self.draw_idle()

    def nearest_index(self, event):
        if self.result is None or event.inaxes != self.axes or len(self.raw_x) == 0:
            return None
        pts = self.axes.transData.transform(np.column_stack([self.raw_x, self.raw_y]))
        distances = np.linalg.norm(pts - np.asarray([event.x, event.y]), axis=1)
        index = int(np.argmin(distances))
        return index if distances[index] <= 14 else None

    def on_mouse_move(self, event):
        if self.annotation is None:
            return
        index = self.nearest_index(event)
        if index is None:
            if self.annotation.get_visible():
                self.annotation.set_visible(False); self.draw_idle()
            return
        row = self.result["rows"][index]
        self.annotation.xy = (row[self.spec["key"]], row["elevation"])
        if self.spec["key"] == "n_hoop":
            self.annotation.set_text(
                f"Plate {row['plate_id']}\nY={row['elevation']:.3f} m\n"
                f"SX={row['sx']:.3f}\nSY={row['sy']:.3f}\nSXY={row['sxy']:.3f}\n"
                f"t={row['thickness']:.3f} m\nNhoop={row['n_hoop']:.3f}")
        else:
            self.annotation.set_text(
                f"Plate {row['plate_id']}\nY={row['elevation']:.3f} m\n"
                f"Mx={row['mx']:.3f}\nMy={row['my']:.3f}\nMxy={row['mxy']:.3f}\n"
                f"Mvert={row['m_vertical']:.3f}")
        self.annotation.set_visible(True); self.draw_idle()

    def on_mouse_click(self, event):
        if event.button == 1:
            index = self.nearest_index(event)
            if index is not None:
                self.plateSelected.emit(int(self.raw_plate_ids[index]))

    def select_plate(self, plate_id):
        if self.result is None:
            return
        row = next((r for r in self.result["rows"] if r["plate_id"] == int(plate_id)), None)
        if row is None:
            return
        if self.selected_marker is not None:
            try: self.selected_marker.remove()
            except Exception: pass
        self.selected_marker = self.axes.scatter([row[self.spec["key"]]], [row["elevation"]],
            s=190, facecolor="none", edgecolor=COLOR_SELECTED, linewidth=3, zorder=30)
        self.draw_idle()


class SectionCanvas(FigureCanvas):
    """Tank cross section through its axis, with the result drawn along the outline."""

    def __init__(self, parent=None):
        self.figure = Figure(figsize=(9, 7), facecolor="white", constrained_layout=True)
        super().__init__(self.figure)
        self.setParent(parent)
        self.axes = self.figure.add_subplot(111)
        self.spec = None
        self.hover_stations = []
        self.hover_points = np.empty((0, 2))
        self.annotation = None
        self.mpl_connect("motion_notify_event", self.on_mouse_move)
        self.draw_empty("Open an analyzed STAAD tank model and extract the results.")

    def draw_empty(self, message):
        self.axes.clear()
        self.axes.set_aspect("auto")
        self.axes.text(0.5, 0.5, message, transform=self.axes.transAxes, ha="center", va="center", fontsize=13)
        self.axes.axis("off")
        self.hover_stations, self.hover_points, self.annotation = [], np.empty((0, 2)), None
        self.draw_idle()

    def set_result(self, result, mode, spec, show_grid=True, show_values=False, scale_factor=1.0):
        section, key, p = result.get("section"), spec["key"], result["parameters"]
        self.spec = spec
        if not section:
            self.draw_empty("No plates were found on the cutting plane.\nIncrease the angular tolerance.")
            return
        ax = self.axes
        ax.clear(); ax.axis("on")
        ax.set_aspect("equal", adjustable="datalim")
        show_average = mode in (VIEW_AVERAGE, VIEW_ENVELOPE, VIEW_ALL)
        show_envelope = mode in (VIEW_ENVELOPE, VIEW_ALL)
        show_raw = mode in (VIEW_RAW, VIEW_ALL)

        stations = [st for side in section["sides"] for piece in side["pieces"] for st in piece["stations"]]
        peak = max([abs(st[key]["average"]) for st in stations], default=0.0)
        size = max(section["r_max"], 0.5 * (section["y_max"] - section["y_min"]), 1.0e-6)
        # The largest binned value is drawn 0.22 x tank size away from the outline.
        factor = 0.22 * size * scale_factor / peak if peak > 1.0e-12 else 0.0
        gap = 0.03 * size
        polygons, polygon_colors, ordinates, ordinate_colors = [], [], [], []
        extent = []
        self.hover_stations, hover_points = [], []
        ax.axvline(0.0, color="#90A4AE", linewidth=1.0, linestyle="-.", zorder=1)

        def colour(value):
            return COLOR_POSITIVE if value >= 0 else COLOR_NEGATIVE

        for side in section["sides"]:
            mirror = np.asarray([side["side"], 1.0])
            for piece in side["pieces"]:
                outline = piece["points"] * mirror
                ax.plot(outline[:, 0], outline[:, 1], color="#263238", linewidth=2.8,
                        solid_capstyle="round", zorder=6)
                extent.append(outline)
                if not piece["stations"]:
                    continue
                base = np.asarray([[st["radius"], st["elevation"]] for st in piece["stations"]]) * mirror
                normal = np.asarray([st["normal"] for st in piece["stations"]]) * mirror
                # A bin average stands for its whole bin, so the end bins run to the ends of the part.
                base = np.vstack([outline[:1], base, outline[-1:]])
                normal = np.vstack([piece["normals"][:1] * mirror, normal, piece["normals"][-1:] * mirror])
                series = {}
                for name in ("average", "minimum", "maximum"):
                    values = np.asarray([st[key][name] for st in piece["stations"]])
                    values = np.concatenate([values[:1], values, values[-1:]])
                    series[name] = (values, base + factor * values[:, None] * normal)
                average, tip = series["average"]
                extent.append(tip)
                if show_average:
                    for i in range(len(base) - 1):
                        if average[i] * average[i + 1] >= 0.0:
                            polygons.append([base[i], base[i + 1], tip[i + 1], tip[i]])
                            polygon_colors.append(colour(average[i] + average[i + 1]))
                        else:
                            zero = base[i] + average[i] / (average[i] - average[i + 1]) * (base[i + 1] - base[i])
                            polygons.append([base[i], zero, tip[i]]); polygon_colors.append(colour(average[i]))
                            polygons.append([zero, base[i + 1], tip[i + 1]]); polygon_colors.append(colour(average[i + 1]))
                    for i in range(1, len(base) - 1):
                        ordinates.append([base[i], tip[i]]); ordinate_colors.append(colour(average[i]))
                    ax.plot(tip[:, 0], tip[:, 1], color=COLOR_PURPLE, linewidth=2.2, zorder=5)
                    ax.plot(tip[1:-1, 0], tip[1:-1, 1], linestyle="none", marker="o", markersize=4.5,
                            markerfacecolor="white", markeredgecolor=COLOR_PURPLE, zorder=7)
                if show_envelope:
                    for name in ("minimum", "maximum"):
                        ax.plot(series[name][1][:, 0], series[name][1][:, 1], color="#AB47BC",
                                linestyle=":", linewidth=1.1, zorder=4)
                        extent.append(series[name][1])
                for index, st in enumerate(piece["stations"], start=1):
                    self.hover_stations.append(st)
                    hover_points.append(tip[index])
            if show_raw:
                rows = side["rows"]
                values = np.asarray([section_value(row, key) for row in rows])
                points = (np.asarray([row["section_point"] for row in rows])
                          + factor * values[:, None] * np.asarray([row["section_normal"] for row in rows])) * mirror
                ax.scatter(points[:, 0], points[:, 1], s=16, color="#78909C", edgecolors="#37474F",
                           linewidths=0.35, alpha=0.8, zorder=4)
                extent.append(points)

        if polygons:
            ax.add_collection(PolyCollection(polygons, facecolors=polygon_colors, edgecolors="none", alpha=0.30, zorder=2))
            ax.add_collection(LineCollection(ordinates, colors=ordinate_colors, linewidths=0.8, alpha=0.75, zorder=3))
        self.hover_points = np.asarray(hover_points) if hover_points else np.empty((0, 2))

        # Value labels: the maximum and the minimum always, every station when asked for.
        extent = np.vstack(extent)
        span = float(np.ptp(extent[:, 1])) + 6.0 * gap
        placed = []

        def label(st, bold):
            value = st[key]["average"]
            mirror = np.asarray([st["side"], 1.0])
            normal = st["normal"] * mirror
            direction = normal if value >= 0 else -normal
            anchor = np.asarray([st["radius"], st["elevation"]]) * mirror + factor * value * normal + gap * direction
            if not bold and any(abs(anchor[0] - q[0]) < 0.085 * span and abs(anchor[1] - q[1]) < 0.027 * span for q in placed):
                return
            placed.append(anchor)
            text = f"{value:.2f}"
            ax.text(anchor[0], anchor[1], "0.00" if text == "-0.00" else text, fontsize=9.5 if bold else 8,
                    fontweight="bold" if bold else "normal", color=colour(value) if bold else "#263238",
                    ha="left" if direction[0] > 0.35 else "right" if direction[0] < -0.35 else "center",
                    va="bottom" if direction[1] > 0.35 else "top" if direction[1] < -0.35 else "center", zorder=9)

        handles = [Line2D([], [], color="#263238", linewidth=2.8, label="Tank cross section")]
        if stations and show_average:
            handles += [Line2D([], [], color=COLOR_PURPLE, linewidth=2.2, marker="o", markersize=4.5,
                               markerfacecolor="white", label="Binned average"),
                        Patch(facecolor=COLOR_POSITIVE, alpha=0.30, label=spec["positive_label"]),
                        Patch(facecolor=COLOR_NEGATIVE, alpha=0.30, label=spec["negative_label"])]
            for st, marker, name in ((max(stations, key=lambda s: s[key]["average"]), "^", "Maximum"),
                                     (min(stations, key=lambda s: s[key]["average"]), "v", "Minimum")):
                value = st[key]["average"]
                mirror = np.asarray([st["side"], 1.0])
                point = (np.asarray([st["radius"], st["elevation"]]) + factor * value * st["normal"]) * mirror
                ax.scatter([point[0]], [point[1]], marker=marker, s=120, color=colour(value), edgecolor="black", zorder=8)
                handles.append(Line2D([], [], linestyle="none", marker=marker, markersize=10, markerfacecolor=colour(value),
                                      markeredgecolor="black", label=f"{name}: {value:.3f} {spec['unit']}"))
                label(st, True)
            if show_values:
                for st in stations:
                    label(st, False)
        if show_envelope:
            handles.append(Line2D([], [], color="#AB47BC", linestyle=":", linewidth=1.1, label="Minimum / maximum in bin"))
        if show_raw:
            handles.append(Line2D([], [], linestyle="none", marker="o", markersize=5, markerfacecolor="#78909C",
                                  markeredgecolor="#37474F", label="Individual plates"))

        # Equal scale on both axes; the limits follow the drawn extent.
        margin = 5.0 * gap
        ax.update_datalim(np.vstack([extent.min(axis=0) - margin, extent.max(axis=0) + margin]))
        ax.autoscale_view()
        if show_grid:
            ax.grid(True, linestyle=":", color="#B0BEC5")
        opposite = (p["target_angle"] + 180.0) % 360.0
        ax.set_title(
            f"{spec['section_title']}\nPlane {p['target_angle']:.2f} deg (right) and {opposite:.2f} deg (left), "
            f"+/- {p['angle_tolerance']:.2f} deg | LC {p['load_case']}", fontweight="bold")
        ax.set_xlabel(f"Distance from the tank axis (m)\n{spec['section_note']}", fontweight="bold")
        ax.set_ylabel("Elevation Y (m)", fontweight="bold")
        if factor > 0:
            ax.text(0.008, 0.012, f"Diagram scale: 1 m = {1.0 / factor:.3g} {spec['unit']}",
                    transform=ax.transAxes, fontsize=8.5, color="#455A64", zorder=9)
        ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.005, 1.0), fontsize=8.5)
        self.annotation = ax.annotate("", xy=(0, 0), xytext=(14, 14), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.45", facecolor="#FFFDE7", edgecolor="#6A1B9A"),
            arrowprops=dict(arrowstyle="->", color="#6A1B9A"), fontsize=9, zorder=20)
        self.annotation.set_visible(False)
        self.draw_idle()

    def on_mouse_move(self, event):
        if self.annotation is None:
            return
        index = None
        if event.inaxes == self.axes and len(self.hover_points):
            distances = np.linalg.norm(self.axes.transData.transform(self.hover_points) - np.asarray([event.x, event.y]), axis=1)
            if distances.min() <= 14:
                index = int(np.argmin(distances))
        if index is None:
            if self.annotation.get_visible():
                self.annotation.set_visible(False); self.draw_idle()
            return
        st, values = self.hover_stations[index], self.hover_stations[index][self.spec["key"]]
        self.annotation.xy = tuple(self.hover_points[index])
        self.annotation.set_text(
            f"{'Plane-angle side' if st['side'] > 0 else 'Opposite side'}\n"
            f"Distance along section = {st['s']:.2f} m\n"
            f"Radius = {st['radius']:.3f} m, Y = {st['elevation']:.3f} m\n"
            f"{self.spec['symbol']} average = {values['average']:.3f} {self.spec['unit']}\n"
            f"Minimum = {values['minimum']:.3f}, maximum = {values['maximum']:.3f}\n"
            f"Plates in bin = {st['count']}")
        self.annotation.set_visible(True); self.draw_idle()


class InteractiveBMDExplorer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1680, 960)
        self.result = None
        self.worker = None
        self.selected_plate = None
        self.setup_ui()
        self.apply_style()

    def setup_ui(self):
        splitter = QSplitter(Qt.Horizontal)
        controls = QWidget(); controls.setMaximumWidth(420)
        left = QVBoxLayout(controls)

        input_box = QGroupBox("Vertical Plane Selection")
        form = QFormLayout(input_box)
        self.y_min = QLineEdit("15.0"); self.y_max = QLineEdit("21.0")
        # Generic list LC 1 ... LC 20 (primary load case or combination). The box is
        # editable, so a number outside the list, such as 201, can be typed.
        self.load_case = QComboBox()
        self.load_case.setEditable(True)
        self.load_case.setInsertPolicy(QComboBox.NoInsert)
        self.load_case.addItems([f"LC {number}" for number in range(1, LOAD_CASE_LIST_LIMIT + 1)])
        self.load_case.setToolTip("Pick LC 1 to LC 20, or type any load case or combination number.")
        self.target_angle = QLineEdit("0.0")
        self.angle_tolerance = QLineEdit("3.0"); self.bin_size = QLineEdit("0.5")
        self.result_type = QComboBox(); self.result_type.addItems(list(RESULT_TYPES))
        form.addRow("Result Type:", self.result_type)
        form.addRow("Bottom Elevation Y (m):", self.y_min)
        form.addRow("Top Elevation Y (m):", self.y_max)
        form.addRow("Load Case / Combination:", self.load_case)
        form.addRow("Plane Angle (deg):", self.target_angle)
        form.addRow("Angular Tolerance (+/- deg):", self.angle_tolerance)
        form.addRow("Bin Size (m):", self.bin_size)
        left.addWidget(input_box)

        view_box = QGroupBox("3D Full Tank Display")
        view_form = QFormLayout(view_box)
        self.show_whole_tank = QCheckBox(); self.show_whole_tank.setChecked(True)
        self.show_tank_edges = QCheckBox(); self.show_tank_edges.setChecked(True)
        self.tank_opacity = QLineEdit("0.18")
        self.selected_strip_opacity = QLineEdit("1.00")
        self.result_color_mode = QComboBox()
        self.result_color_mode.addItems(["Elevation Bin Bands", "Result Contour"])
        self.show_bin_planes = QCheckBox(); self.show_bin_planes.setChecked(True)
        self.show_bin_labels = QCheckBox(); self.show_bin_labels.setChecked(True)
        view_form.addRow("Show Whole Tank:", self.show_whole_tank)
        view_form.addRow("Show Tank Mesh Edges:", self.show_tank_edges)
        view_form.addRow("Tank Opacity (0-1):", self.tank_opacity)
        view_form.addRow("Selected Strip Opacity:", self.selected_strip_opacity)
        view_form.addRow("Selected Plate Colors:", self.result_color_mode)
        view_form.addRow("Show Bin Boundary Planes:", self.show_bin_planes)
        view_form.addRow("Show Bin Labels:", self.show_bin_labels)
        left.addWidget(view_box)

        graph_box = QGroupBox("Diagram Display")
        graph_form = QFormLayout(graph_box)
        self.mode = QComboBox(); self.mode.addItems([VIEW_ALL, VIEW_ENVELOPE, VIEW_AVERAGE, VIEW_RAW])
        self.show_zero = QCheckBox(); self.show_zero.setChecked(True)
        self.show_grid = QCheckBox(); self.show_grid.setChecked(True)
        self.show_values = QCheckBox(); self.show_values.setChecked(False)
        self.section_scale = QLineEdit("1.0")
        graph_form.addRow("Representation:", self.mode)
        graph_form.addRow("Show Zero Line:", self.show_zero)
        graph_form.addRow("Show Grid:", self.show_grid)
        graph_form.addRow("Show Values on Diagrams:", self.show_values)
        graph_form.addRow("Cross Section Scale (x):", self.section_scale)
        left.addWidget(graph_box)

        action_box = QGroupBox("Actions")
        grid = QGridLayout(action_box)
        self.extract_button = QPushButton("EXTRACT AND DISPLAY FULL TANK")
        self.extract_button.setStyleSheet("background:#D84315;color:white;font-weight:bold;padding:9px")
        self.refresh_3d_button = QPushButton("REFRESH 3D VIEW")
        self.save_button = QPushButton("SAVE DIAGRAM")
        self.excel_button = QPushButton("EXPORT TO EXCEL")
        self.excel_button.setStyleSheet("background:#2E7D32;color:white;font-weight:bold")
        grid.addWidget(self.extract_button, 0, 0, 1, 2)
        grid.addWidget(self.refresh_3d_button, 1, 0)
        grid.addWidget(self.save_button, 1, 1)
        grid.addWidget(self.excel_button, 2, 0, 1, 2)
        left.addWidget(action_box)
        self.progress = QProgressBar(); self.status = QLabel("Ready. Keep analyzed STAAD model open.")
        self.status.setWordWrap(True); left.addWidget(self.progress); left.addWidget(self.status)

        summary = QGroupBox("Selection Summary")
        sf = QFormLayout(summary)
        self.total_count = QLabel("-"); self.selected_count = QLabel("-")
        self.active_bin_count = QLabel("-"); self.active_bin_size = QLabel("-")
        self.max_m = QLabel("-"); self.min_m = QLabel("-")
        sf.addRow("Entire Tank Plates:", self.total_count)
        self.max_label = QLabel("Maximum Mvert:"); self.min_label = QLabel("Minimum Mvert:")
        sf.addRow("Highlighted Result Plates:", self.selected_count)
        sf.addRow("Active Elevation Bins:", self.active_bin_count)
        sf.addRow("Current Bin Size:", self.active_bin_size)
        sf.addRow(self.max_label, self.max_m)
        sf.addRow(self.min_label, self.min_m)
        left.addWidget(summary); left.addStretch()

        self.tabs = QTabWidget()
        plot_tab = QWidget(); plot_layout = QVBoxLayout(plot_tab)
        self.canvas = BMDCanvas(self); plot_layout.addWidget(NavigationToolbar(self.canvas, self)); plot_layout.addWidget(self.canvas)
        self.tabs.addTab(plot_tab, "Interactive Diagram")

        # Whole cross section through the tank axis, with the result drawn along its outline.
        self.section_tab = QWidget(); section_layout = QVBoxLayout(self.section_tab)
        self.section_canvas = SectionCanvas(self)
        section_layout.addWidget(NavigationToolbar(self.section_canvas, self)); section_layout.addWidget(self.section_canvas)
        self.tabs.addTab(self.section_tab, "Tank Cross Section")

        tank_tab = QWidget(); tank_layout = QVBoxLayout(tank_tab)
        self.plotter = QtInteractor(tank_tab); self.plotter.set_background("#F8FAFC")
        tank_layout.addWidget(self.plotter.interactor)
        self.tabs.addTab(tank_tab, "Full Tank + Highlighted Plates")

        table_tab = QWidget(); table_layout = QVBoxLayout(table_tab)
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(["Elevation", "Plate", "Plane Angle", "theta", "Mx", "My", "Mxy", "Mvert", "Bin Y"])
        self.table.setSortingEnabled(True); self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection); self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); table_layout.addWidget(self.table)
        self.tabs.addTab(table_tab, "Selected Plate Table")

        calc_tab = QWidget(); calc_layout = QVBoxLayout(calc_tab)
        self.calculation = QPlainTextEdit(); self.calculation.setReadOnly(True); self.calculation.setFont(QFont("Consolas", 10))
        calc_layout.addWidget(self.calculation); self.tabs.addTab(calc_tab, "Calculation Report")

        splitter.addWidget(controls); splitter.addWidget(self.tabs)
        splitter.setStretchFactor(0, 0); splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

        self.extract_button.clicked.connect(self.extract)
        self.refresh_3d_button.clicked.connect(self.update_3d_view)
        self.save_button.clicked.connect(self.save_plot)
        self.excel_button.clicked.connect(self.export_to_excel)
        self.result_type.currentTextChanged.connect(self.result_type_changed)
        self.mode.currentTextChanged.connect(self.refresh_plot)
        self.show_zero.stateChanged.connect(self.refresh_plot)
        self.show_grid.stateChanged.connect(self.refresh_plot)
        self.show_values.stateChanged.connect(self.refresh_plot)
        self.section_scale.editingFinished.connect(self.refresh_plot)
        self.show_whole_tank.stateChanged.connect(self.update_3d_view)
        self.show_tank_edges.stateChanged.connect(self.update_3d_view)
        self.result_color_mode.currentTextChanged.connect(self.update_3d_view)
        self.show_bin_planes.stateChanged.connect(self.update_3d_view)
        self.show_bin_labels.stateChanged.connect(self.update_3d_view)
        self.canvas.plateSelected.connect(self.select_plate)
        self.table.itemSelectionChanged.connect(self.table_selection)

    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow { background:#ECEFF1; }
            QGroupBox { background:white; border:1px solid #B0BEC5; border-radius:6px;
                        margin-top:11px; padding-top:8px; font-weight:bold; }
            QGroupBox::title { subcontrol-origin:margin; left:9px; padding:0 4px; }
            QLineEdit, QComboBox { background:white; border:1px solid #B0BEC5;
                                  border-radius:4px; padding:4px; min-height:24px; }
            QPushButton { border-radius:5px; padding:5px 8px; min-height:30px; }
            QHeaderView::section { background:#4A148C; color:white; font-weight:bold; padding:6px; }
        """)

    def read_inputs(self):
        # Accepts a picked entry such as "LC 5" or a typed number such as "201".
        load_text = self.load_case.currentText().upper().replace("LC", "").strip()
        if not load_text.isdigit():
            raise ValueError("Pick LC 1 to LC 20, or type a load case or combination number such as 201.")
        values = dict(
            y_min=float(self.y_min.text()), y_max=float(self.y_max.text()),
            load_case=int(load_text), target_angle=float(self.target_angle.text()),
            angle_tolerance=float(self.angle_tolerance.text()), bin_size=float(self.bin_size.text()),
        )
        if values["y_max"] <= values["y_min"]:
            raise ValueError("Top elevation must exceed bottom elevation.")
        if values["load_case"] <= 0 or values["angle_tolerance"] <= 0 or values["bin_size"] <= 0:
            raise ValueError("Load case, tolerance and bin size must be greater than zero.")
        return values

    def current_spec(self):
        return RESULT_TYPES[self.result_type.currentText()]

    def update_summary(self):
        spec = self.current_spec()
        key, stats = spec["key"], self.result["statistics"][spec["key"]]
        self.max_label.setText(f"Maximum {spec['symbol']}:")
        self.min_label.setText(f"Minimum {spec['symbol']}:")
        self.max_m.setText(f"{stats['maximum'][key]:.3f} {spec['unit']}")
        self.min_m.setText(f"{stats['minimum'][key]:.3f} {spec['unit']}")

    def result_type_changed(self, *args):
        # Both results are already extracted, so switching needs no new STAAD query.
        if self.result is None: return
        self.update_summary(); self.populate_table(); self.refresh_plot(); self.update_3d_view()
        if self.selected_plate is not None: self.select_plate(self.selected_plate)

    def extract(self):
        try:
            values = self.read_inputs()
        except Exception as exc:
            QMessageBox.warning(self, "Input Error", str(exc)); return
        self.extract_button.setEnabled(False); self.progress.setValue(0)
        self.worker = ExtractionWorker(**values)
        self.worker.progress.connect(lambda n, s: (self.progress.setValue(n), self.status.setText(s)))
        self.worker.completed.connect(self.extraction_complete)
        self.worker.failed.connect(self.extraction_failed)
        self.worker.start()

    def extraction_complete(self, result):
        self.result = result; self.selected_plate = None; self.extract_button.setEnabled(True)
        self.progress.setValue(100); self.status.setText(
            f"Full tank {len(result['all_plates'])}; full-height strip {len(result['context_strip'])}; result plates {len(result['rows'])}.")
        self.total_count.setText(str(len(result["all_plates"])))
        self.selected_count.setText(str(len(result["rows"])))
        self.active_bin_count.setText(str(len(result["bin_centers"])))
        self.active_bin_size.setText(f"{result['parameters']['bin_size']:.3f} m")
        self.update_summary()
        self.populate_table(); self.refresh_plot(); self.update_3d_view()
        self.select_plate(result["statistics"][self.current_spec()["key"]]["absolute"]["plate_id"])

    def extraction_failed(self, error):
        self.extract_button.setEnabled(True); self.progress.setValue(0)
        QMessageBox.critical(self, "Extraction Error", error)

    def refresh_plot(self, *args):
        if self.result is None: return
        self.canvas.set_result(self.result, self.mode.currentText(), self.current_spec(), self.show_zero.isChecked(),
                               self.show_grid.isChecked(), self.show_values.isChecked())
        try:
            section_scale = float(self.section_scale.text())
        except ValueError:
            section_scale = 1.0
        self.section_canvas.set_result(self.result, self.mode.currentText(), self.current_spec(),
                                       self.show_grid.isChecked(), self.show_values.isChecked(),
                                       section_scale if section_scale > 0 else 1.0)
        if self.selected_plate is not None: self.canvas.select_plate(self.selected_plate)

    @staticmethod
    def build_mesh(plate_rows, scalar_key=None):
        node_to_index, points, faces, plate_ids, scalars = {}, [], [], [], []
        for row in plate_rows:
            local = []
            for node_id, xyz in zip(row["node_ids"], row["coordinates"]):
                if node_id not in node_to_index:
                    node_to_index[node_id] = len(points); points.append(np.asarray(xyz, dtype=float))
                local.append(node_to_index[node_id])
            if len(local) >= 3:
                local = local[:4]
                faces.extend([len(local)] + local)
                plate_ids.append(row["plate_id"])
                if scalar_key: scalars.append(row[scalar_key])
        if not faces: return None
        mesh = pv.PolyData(np.asarray(points, dtype=float), np.asarray(faces, dtype=np.int64))
        mesh.cell_data["Plate ID"] = np.asarray(plate_ids, dtype=int)
        if scalar_key: mesh.cell_data[scalar_key] = np.asarray(scalars, dtype=float)
        return mesh

    def update_3d_view(self, *args):
        if self.result is None: return
        try:
            tank_opacity = min(max(float(self.tank_opacity.text()), 0.0), 1.0)
            strip_opacity = min(max(float(self.selected_strip_opacity.text()), 0.0), 1.0)
        except ValueError:
            tank_opacity, strip_opacity = 0.18, 1.0

        self.plotter.clear()
        if self.show_whole_tank.isChecked():
            tank_mesh = self.build_mesh(self.result["all_plates"])
            if tank_mesh is not None:
                self.plotter.add_mesh(
                    tank_mesh, color="#B0BEC5", opacity=tank_opacity,
                    show_edges=self.show_tank_edges.isChecked(), edge_color="#78909C",
                    line_width=0.45, name="whole_tank", pickable=False,
                )

        # Full-height angular strip, including the domes, shown in orange, together with
        # the strip on the opposite side: the two form the cross section through the axis.
        context_mesh = self.build_mesh(self.result["context_strip"] + self.result["opposite_strip"])
        if context_mesh is not None:
            self.plotter.add_mesh(
                context_mesh, color="#FFB300", opacity=0.58, show_edges=True,
                edge_color="#E65100", line_width=1.5, name="full_height_context_strip",
                pickable=False,
            )

        # BMD result plates within the specified Y range. The user can color them
        # either by elevation bin or by transformed moment.
        if self.result_color_mode.currentText() == "Elevation Bin Bands":
            strip_mesh = self.build_mesh(self.result["rows"], "bin_index")
            if strip_mesh is not None:
                bin_count = max(len(self.result["bin_centers"]), 1)
                self.plotter.add_mesh(
                    strip_mesh, scalars="bin_index", cmap="turbo",
                    clim=[-0.5, bin_count - 0.5], n_colors=bin_count,
                    categories=True, opacity=strip_opacity, show_edges=True,
                    edge_color="#263238", line_width=1.8,
                    scalar_bar_args={
                        "title": f"Elevation Bins ({self.result['parameters']['bin_size']:.3f} m)",
                        "n_labels": min(bin_count, 12),
                    },
                    name="selected_strip", pickable=True,
                )
        else:
            spec = self.current_spec()
            strip_mesh = self.build_mesh(self.result["rows"], spec["key"])
            if strip_mesh is not None:
                limit = max(float(np.max(np.abs(strip_mesh.cell_data[spec["key"]]))), 1.0e-9)
                self.plotter.add_mesh(
                    strip_mesh, scalars=spec["key"], cmap="coolwarm", clim=[-limit, limit],
                    opacity=strip_opacity, show_edges=True, edge_color="#263238", line_width=1.8,
                    scalar_bar_args={"title": f"Selected Strip {spec['symbol']} ({spec['unit']})"},
                    name="selected_strip", pickable=True,
                )

        # Draw translucent horizontal planes at true bin boundaries. Because the
        # binning uses round(Y / BIN_SIZE), each band spans center +/- BIN_SIZE/2.
        if self.show_bin_planes.isChecked() and self.result["bin_centers"]:
            all_xyz = np.vstack([row["coordinates"] for row in self.result["all_plates"]])
            x_size = max(float(np.ptp(all_xyz[:, 0])) * 1.08, 1.0)
            z_size = max(float(np.ptp(all_xyz[:, 2])) * 1.08, 1.0)
            bin_size = self.result["parameters"]["bin_size"]
            boundaries = [self.result["bin_centers"][0] - 0.5 * bin_size]
            boundaries.extend(center + 0.5 * bin_size for center in self.result["bin_centers"])
            for boundary_index, y_boundary in enumerate(boundaries):
                plane = pv.Plane(
                    center=(0.0, float(y_boundary), 0.0),
                    direction=(0.0, 1.0, 0.0),
                    i_size=x_size,
                    j_size=z_size,
                    i_resolution=1,
                    j_resolution=1,
                )
                self.plotter.add_mesh(
                    plane, color="#00ACC1", opacity=0.075, show_edges=True,
                    edge_color="#00838F", line_width=0.7,
                    name=f"bin_boundary_{boundary_index}", pickable=False,
                )

        # Label each elevation bin at the side of the tank.
        if self.show_bin_labels.isChecked() and self.result["bin_centers"]:
            all_xyz = np.vstack([row["coordinates"] for row in self.result["all_plates"]])
            x_label = float(np.max(all_xyz[:, 0])) * 1.10
            label_points = np.asarray(
                [[x_label, center, 0.0] for center in self.result["bin_centers"]],
                dtype=float,
            )
            labels = [
                f"Bin {index + 1}: Y={center:.3f} m"
                for index, center in enumerate(self.result["bin_centers"])
            ]
            if len(label_points):
                self.plotter.add_point_labels(
                    pv.PolyData(label_points), labels, font_size=9,
                    point_color="#006064", text_color="#004D40",
                    shape_color="#E0F7FA", shape_opacity=0.78,
                    name="bin_labels", always_visible=True,
                )

        self.plotter.add_axes(); self.plotter.show_grid(); self.plotter.reset_camera()
        if self.selected_plate is not None:
            self.highlight_selected_plate(self.selected_plate)

    def populate_table(self):
        key = self.current_spec()["key"]
        if key == "n_hoop":
            headers = ["Elevation", "Plate", "Plane Angle", "alpha", "SX", "SY", "SXY", "Thickness", "Nhoop", "Bin Y"]
            fields = ["elevation", "plate_id", "angle", "alpha", "sx", "sy", "sxy", "thickness", "n_hoop", "bin_y"]
        else:
            headers = ["Elevation", "Plate", "Plane Angle", "theta", "Mx", "My", "Mxy", "Mvert", "Bin Y"]
            fields = ["elevation", "plate_id", "angle", "theta", "mx", "my", "mxy", "m_vertical", "bin_y"]
        value_column = fields.index(key)
        self.table.blockSignals(True)
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(headers)); self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(self.result["rows"]))
        for i, row in enumerate(self.result["rows"]):
            for j, field in enumerate(fields):
                item = integer_item(row[field]) if field == "plate_id" else number_item(row[field])
                item.setData(Qt.UserRole + 1, row["plate_id"]); self.table.setItem(i, j, item)
            self.table.item(i, value_column).setForeground(QBrush(QColor(
                COLOR_POSITIVE if row[key] >= 0 else COLOR_NEGATIVE)))
        self.table.setSortingEnabled(True)
        self.table.blockSignals(False)

    def selected_row(self, plate_id):
        return next((r for r in self.result["rows"] if r["plate_id"] == int(plate_id)), None)

    def highlight_selected_plate(self, plate_id):
        row = self.selected_row(plate_id)
        if row is None: return
        mesh = self.build_mesh([row])
        if mesh is not None:
            self.plotter.add_mesh(mesh, color=COLOR_SELECTED, opacity=0.95, show_edges=True,
                                  edge_color="black", line_width=4, name="active_plate")
            centroid = np.asarray([row["centroid"]])
            labels = pv.PolyData(centroid)
            spec = self.current_spec()
            self.plotter.add_point_labels(labels, [f"Plate {row['plate_id']}\n{spec['symbol']}={row[spec['key']]:.3f}"],
                                          font_size=11, text_color="black", shape_color="#FFF59D",
                                          shape_opacity=0.92, name="active_label")
            self.plotter.render()

    def select_plate(self, plate_id):
        row = self.selected_row(plate_id)
        if row is None: return
        self.selected_plate = int(plate_id); self.canvas.select_plate(self.selected_plate)
        self.highlight_selected_plate(self.selected_plate); self.update_calculation(row)
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 1)
            if item and int(item.data(Qt.UserRole + 1)) == self.selected_plate:
                self.table.blockSignals(True); self.table.selectRow(i); self.table.scrollToItem(item)
                self.table.blockSignals(False); break

    def table_selection(self):
        selected = self.table.selectedItems()
        if selected:
            plate_id = selected[0].data(Qt.UserRole + 1)
            if plate_id is not None: self.select_plate(int(plate_id))

    def update_calculation(self, row):
        if self.current_spec()["key"] == "n_hoop":
            self.calculation.setPlainText(self.hoop_calculation_report(row))
            return
        lx, ly, lz = row["local_x"], row["local_y"], row["local_z"]
        c = row["centroid"]
        check = row["term_average"] + row["term_difference"] + row["term_twisting"] - row["m_vertical"]
        report = f"""LOCAL-TO-GLOBAL PLATE MOMENT TRANSFORMATION
{'=' * 78}

SELECTED STAAD RESULT
---------------------
Plate ID              : {row['plate_id']}
Load Case             : {row['load_case']}
Plate Nodes           : {', '.join(map(str, row['node_ids']))}
Plate Centroid        : X={c[0]:.8f}, Y={c[1]:.8f}, Z={c[2]:.8f} m
Circumferential Angle : {row['angle']:.8f} deg

STEP 1 - FORM PLATE LOCAL AXES
-------------------------------
Local X = (Node 2 - Node 1) / |Node 2 - Node 1|
Local Z = ((Node 2 - Node 1) x (Node 3 - Node 1)) / |cross product|
Local Y = Local Z x Local X

Local X = [{lx[0]: .10f}, {lx[1]: .10f}, {lx[2]: .10f}]
Local Y = [{ly[0]: .10f}, {ly[1]: .10f}, {ly[2]: .10f}]
Local Z = [{lz[0]: .10f}, {lz[1]: .10f}, {lz[2]: .10f}]

STEP 2 - PROJECT GLOBAL VERTICAL INTO LOCAL XY
-----------------------------------------------
Global vertical = [0.0, 1.0, 0.0]
Vx = Global-Y dot Local-X = {row['vertical_component_local_x']:.10f}
Vy = Global-Y dot Local-Y = {row['vertical_component_local_y']:.10f}

STEP 3 - TRANSFORMATION ANGLE
-----------------------------
theta = atan2(Vy, Vx)
      = atan2({row['vertical_component_local_y']:.10f}, {row['vertical_component_local_x']:.10f})
      = {row['theta_rad']:.10f} rad
      = {row['theta']:.8f} deg
2 theta       = {2.0 * row['theta']:.8f} deg
cos(2 theta)  = {row['cos_2theta']:.10f}
sin(2 theta)  = {row['sin_2theta']:.10f}

STEP 4 - STAAD LOCAL PLATE-CENTRE MOMENTS
------------------------------------------
GetAllPlateCenterMoments({row['plate_id']}, {row['load_case']})
Mx  = {row['mx']:.10f} kN.m/m
My  = {row['my']:.10f} kN.m/m
Mxy = {row['mxy']:.10f} kN.m/m

STEP 5 - MOMENT TRANSFORMATION
------------------------------
Mvert = (Mx + My)/2 + ((Mx - My)/2) cos(2 theta) + Mxy sin(2 theta)

NUMERICAL SUBSTITUTION
Mvert = ({row['mx']:.10f} + {row['my']:.10f})/2
      + (({row['mx']:.10f} - {row['my']:.10f})/2) x {row['cos_2theta']:.10f}
      + ({row['mxy']:.10f}) x {row['sin_2theta']:.10f}

TERM BREAKDOWN
Average term    = {row['term_average']:.10f} kN.m/m
Difference term = {row['term_difference']:.10f} kN.m/m
Twisting term   = {row['term_twisting']:.10f} kN.m/m

FINAL RESULT
Mvert = {row['term_average']:.10f} + {row['term_difference']:.10f} + {row['term_twisting']:.10f}
Mvert = {row['m_vertical']:.10f} kN.m/m

NUMERICAL CHECK
Term sum - Mvert = {check:.12e}

SIGN NOTE
The inside/outside face interpretation depends on consistent plate local-Z orientation.
"""
        self.calculation.setPlainText(report)

    @staticmethod
    def hoop_calculation_report(row):
        lx, ly, lz = row["local_x"], row["local_y"], row["local_z"]
        c, gh = row["centroid"], row["global_hoop"]
        check = (row["hoop_term_average"] + row["hoop_term_difference"]
                 + row["hoop_term_shear"]) * row["thickness"] - row["n_hoop"]
        return f"""HOOP FORCE FROM PLATE MEMBRANE STRESSES
{'=' * 78}

SELECTED STAAD RESULT
---------------------
Plate ID              : {row['plate_id']}
Load Case             : {row['load_case']}
Plate Nodes           : {', '.join(map(str, row['node_ids']))}
Plate Centroid        : X={c[0]:.8f}, Y={c[1]:.8f}, Z={c[2]:.8f} m
Circumferential Angle : {row['angle']:.8f} deg
Radius from Tank Axis : {row['radius']:.8f} m

STEP 1 - FORM PLATE LOCAL AXES
-------------------------------
Local X = (Node 2 - Node 1) / |Node 2 - Node 1|
Local Z = ((Node 2 - Node 1) x (Node 3 - Node 1)) / |cross product|
Local Y = Local Z x Local X

Local X = [{lx[0]: .10f}, {lx[1]: .10f}, {lx[2]: .10f}]
Local Y = [{ly[0]: .10f}, {ly[1]: .10f}, {ly[2]: .10f}]
Local Z = [{lz[0]: .10f}, {lz[1]: .10f}, {lz[2]: .10f}]

STEP 2 - HOOP DIRECTION AT THE PLATE CENTROID
----------------------------------------------
Tank axis   = global Y through X = 0, Z = 0
Hoop direction = horizontal tangent to the circle through the centroid
Global hoop = [-Zc, 0, Xc] / R
            = [{gh[0]: .10f}, {gh[1]: .10f}, {gh[2]: .10f}]
Hx = Global hoop dot Local-X = {row['hoop_component_local_x']:.10f}
Hy = Global hoop dot Local-Y = {row['hoop_component_local_y']:.10f}

STEP 3 - TRANSFORMATION ANGLE
-----------------------------
alpha = atan2(Hy, Hx)
      = atan2({row['hoop_component_local_y']:.10f}, {row['hoop_component_local_x']:.10f})
      = {row['alpha_rad']:.10f} rad
      = {row['alpha']:.8f} deg
2 alpha       = {2.0 * row['alpha']:.8f} deg
cos(2 alpha)  = {row['cos_2alpha']:.10f}
sin(2 alpha)  = {row['sin_2alpha']:.10f}

STEP 4 - STAAD LOCAL PLATE-CENTRE MEMBRANE STRESSES
----------------------------------------------------
GetAllPlateCenterForces({row['plate_id']}, {row['load_case']})   returns [SQX, SQY, SX, SY, SXY]
SX  = {row['sx']:.10f} kN/m2
SY  = {row['sy']:.10f} kN/m2
SXY = {row['sxy']:.10f} kN/m2

STEP 5 - STRESS TRANSFORMATION
------------------------------
sigma_hoop = (SX + SY)/2 + ((SX - SY)/2) cos(2 alpha) + SXY sin(2 alpha)

NUMERICAL SUBSTITUTION
sigma_hoop = ({row['sx']:.10f} + {row['sy']:.10f})/2
           + (({row['sx']:.10f} - {row['sy']:.10f})/2) x {row['cos_2alpha']:.10f}
           + ({row['sxy']:.10f}) x {row['sin_2alpha']:.10f}

TERM BREAKDOWN
Average term    = {row['hoop_term_average']:.10f} kN/m2
Difference term = {row['hoop_term_difference']:.10f} kN/m2
Shear term      = {row['hoop_term_shear']:.10f} kN/m2

sigma_hoop = {row['sigma_hoop']:.10f} kN/m2

STEP 6 - HOOP FORCE PER METRE OF WALL
-------------------------------------
Plate thickness t = {row['thickness']:.10f} m   (Property.GetPlateThickness, mean of corner values)
Nhoop = sigma_hoop x t
      = {row['sigma_hoop']:.10f} x {row['thickness']:.10f}

FINAL RESULT
Nhoop = {row['n_hoop']:.10f} kN/m

NUMERICAL CHECK
Term sum x t - Nhoop = {check:.12e}

SIGN NOTE
Positive Nhoop is hoop tension; negative Nhoop is hoop compression.
"""

    def save_plot(self):
        if self.result is None: return
        # Saves the diagram of the tab that is open: the cross section or the elevation diagram.
        on_section = self.tabs.currentWidget() is self.section_tab
        name = self.current_spec()["file_name"] + ("_Cross_Section" if on_section else "")
        path, _ = QFileDialog.getSaveFileName(self, "Save Diagram", f"{name}.png",
                                               "PNG (*.png);;PDF (*.pdf);;SVG (*.svg)")
        if path:
            figure = self.section_canvas.figure if on_section else self.canvas.figure
            figure.savefig(path, dpi=220, facecolor="white", bbox_inches="tight")


    def export_to_excel(self):
        if self.result is None:
            QMessageBox.information(self, "No Results", "Extract results first.")
            return
        if not OPENPYXL_AVAILABLE:
            QMessageBox.critical(self, "Missing Library", "Install openpyxl: pip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Results", "STAAD_Tank_Wall_Results.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            wb = Workbook()
            fill = PatternFill("solid", fgColor="4A148C")
            white = Font(color="FFFFFF", bold=True)

            ws = wb.active
            ws.title = "Plate Results"
            headers = ["Elevation Y (m)", "Plate ID", "Plane Angle (deg)", "Theta (deg)",
                       "Mx", "My", "Mxy", "Mvert", "Bin Center Y (m)", "Bin Index",
                       "Alpha (deg)", "SX (kN/m2)", "SY (kN/m2)", "SXY (kN/m2)", "Thickness (m)",
                       "Sigma Hoop (kN/m2)", "Nhoop (kN/m)"]
            ws.append(headers)
            for cell in ws[1]: cell.fill, cell.font = fill, white
            for r in self.result["rows"]:
                ws.append([r["elevation"], r["plate_id"], r["angle"], r["theta"],
                           r["mx"], r["my"], r["mxy"], r["m_vertical"], r["bin_y"], r["bin_index"] + 1,
                           r["alpha"], r["sx"], r["sy"], r["sxy"], r["thickness"],
                           r["sigma_hoop"], r["n_hoop"]])
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            tr = wb.create_sheet("Moment Transformation")
            tr.append(["Plate ID", "Load Case", "Elevation", "Plane Angle", "Node IDs",
                       "Local Xx", "Local Xy", "Local Xz", "Local Yx", "Local Yy", "Local Yz",
                       "Local Zx", "Local Zy", "Local Zz", "Vx", "Vy", "Theta deg", "Theta rad",
                       "Mx", "My", "Mxy", "cos2theta", "sin2theta", "Average Term",
                       "Difference Term", "Twisting Term", "Mvert"])
            for cell in tr[1]: cell.fill, cell.font = fill, white
            for r in self.result["rows"]:
                tr.append([r["plate_id"], r["load_case"], r["elevation"], r["angle"],
                           ", ".join(map(str, r["node_ids"])), *r["local_x"], *r["local_y"], *r["local_z"],
                           r["vertical_component_local_x"], r["vertical_component_local_y"],
                           r["theta"], r["theta_rad"], r["mx"], r["my"], r["mxy"],
                           r["cos_2theta"], r["sin_2theta"], r["term_average"],
                           r["term_difference"], r["term_twisting"], r["m_vertical"]])
            tr.freeze_panes = "A2"
            tr.auto_filter.ref = tr.dimensions

            ht = wb.create_sheet("Hoop Transformation")
            ht.append(["Plate ID", "Load Case", "Elevation", "Plane Angle", "Radius",
                       "Hoop X", "Hoop Y", "Hoop Z", "Hx", "Hy", "Alpha deg", "Alpha rad",
                       "SX", "SY", "SXY", "cos2alpha", "sin2alpha", "Average Term",
                       "Difference Term", "Shear Term", "Sigma Hoop", "Thickness", "Nhoop"])
            for cell in ht[1]: cell.fill, cell.font = fill, white
            for r in self.result["rows"]:
                ht.append([r["plate_id"], r["load_case"], r["elevation"], r["angle"], r["radius"],
                           *r["global_hoop"], r["hoop_component_local_x"], r["hoop_component_local_y"],
                           r["alpha"], r["alpha_rad"], r["sx"], r["sy"], r["sxy"],
                           r["cos_2alpha"], r["sin_2alpha"], r["hoop_term_average"],
                           r["hoop_term_difference"], r["hoop_term_shear"], r["sigma_hoop"],
                           r["thickness"], r["n_hoop"]])
            ht.freeze_panes = "A2"
            ht.auto_filter.ref = ht.dimensions

            for sheet_name, spec in (("Binned BMD", RESULT_TYPES[RESULT_MVERT]),
                                     ("Binned Hoop Force", RESULT_TYPES[RESULT_NHOOP])):
                bs = wb.create_sheet(sheet_name)
                bs.append(["Elevation", "Average", "Minimum", "Maximum", "Std Dev", "Count"])
                for cell in bs[1]: cell.fill, cell.font = fill, white
                for r in self.result["bins"][spec["key"]]:
                    bs.append([r["elevation"], r["average"], r["minimum"], r["maximum"], r["standard_deviation"], r["count"]])
                if bs.max_row >= 3:
                    chart = ScatterChart(); chart.title = spec["title"]
                    chart.x_axis.title = f"{spec['symbol']} ({spec['unit']})"; chart.y_axis.title = "Elevation Y (m)"
                    # Result on the X axis and elevation on the Y axis, as in the on-screen diagram.
                    chart.series.append(Series(Reference(bs, min_col=1, min_row=2, max_row=bs.max_row),
                                               Reference(bs, min_col=2, min_row=2, max_row=bs.max_row),
                                               title="Binned Average"))
                    chart.height, chart.width = 14, 18
                    bs.add_chart(chart, "H2")

            # Cross section through the tank axis: binned stations and the plates behind them.
            section = self.result.get("section")
            if section:
                cs = wb.create_sheet("Cross Section Stations")
                cs.append(["Side", "Plane Angle (deg)", "Part", "Distance along Section (m)", "Radius (m)",
                           "Elevation Y (m)", "Plates in Bin", "Moment Average (kN.m/m)", "Moment Minimum",
                           "Moment Maximum", "Nhoop Average (kN/m)", "Nhoop Minimum", "Nhoop Maximum"])
                cp = wb.create_sheet("Cross Section Plates")
                cp.append(["Side", "Plate ID", "Distance along Section (m)", "Radius (m)", "Elevation Y (m)",
                           "Local Z Points", "Meridional Moment, Local-Z Sign (kN.m/m)",
                           "Meridional Moment, Outer-Face Tension Positive (kN.m/m)", "Nhoop (kN/m)"])
                for sheet in (cs, cp):
                    for cell in sheet[1]: cell.fill, cell.font = fill, white
                target_angle = self.result["parameters"]["target_angle"]
                for side in section["sides"]:
                    name = "Plane-angle side" if side["side"] > 0 else "Opposite side"
                    angle = target_angle if side["side"] > 0 else (target_angle + 180.0) % 360.0
                    for piece in side["pieces"]:
                        for st in piece["stations"]:
                            m, n = st["m_vertical"], st["n_hoop"]
                            cs.append([name, angle, st["part"] + 1, st["s"], st["radius"], st["elevation"], st["count"],
                                       m["average"], m["minimum"], m["maximum"], n["average"], n["minimum"], n["maximum"]])
                    for r in side["rows"]:
                        cp.append([name, r["plate_id"], r["section_s"], r["radius"], r["elevation"],
                                   "Outward" if r["section_orient"] > 0 else "Inward",
                                   r["m_vertical"], r["section_m"], r["n_hoop"]])
                for sheet in (cs, cp):
                    sheet.freeze_panes = "A2"
                    sheet.auto_filter.ref = sheet.dimensions

            summary = wb.create_sheet("Summary")
            p, st = self.result["parameters"], self.result["statistics"]
            summary.append(["Parameter", "Value", "Units", "Comment"])
            summary.append(["Entire Tank Plates", len(self.result["all_plates"]), "-", "Full shell context"])
            summary.append(["Full-height angular strip", len(self.result["context_strip"]), "-", "Includes domes"])
            summary.append(["Result plates", len(self.result["rows"]), "-", "Within selected Y range"])
            summary.append(["Load Case", p["load_case"], "-", ""])
            summary.append(["Plane Angle", p["target_angle"], "deg", ""])
            summary.append(["Angle Tolerance", p["angle_tolerance"], "deg", ""])
            for spec in RESULT_TYPES.values():
                key = spec["key"]
                summary.append([f"Maximum {spec['symbol']}", st[key]["maximum"][key], spec["unit"], f"Plate {st[key]['maximum']['plate_id']}"])
                summary.append([f"Minimum {spec['symbol']}", st[key]["minimum"][key], spec["unit"], f"Plate {st[key]['minimum']['plate_id']}"])
            for cell in summary[1]: cell.fill, cell.font = fill, white

            for sheet in wb.worksheets:
                for col in sheet.columns:
                    letter = get_column_letter(col[0].column)
                    sheet.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col) + 2, 12), 32)
            wb.save(path)
            QMessageBox.information(self, "Export Complete", f"Saved to:\n{path}")
        except Exception:
            QMessageBox.critical(self, "Export Error", traceback.format_exc())


def main():
    app = QApplication(sys.argv)
    window = InteractiveBMDExplorer()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
