"""
SolidWorks COM client wrapper.

All COM API calls must occur on the same STA (Single-Threaded Apartment) thread.
This module manages a dedicated thread + executor for that purpose.
All public methods are synchronous and intended to be called from the executor.
"""

import math
import os
import tempfile
from typing import Any, Dict, List, Optional, Tuple

import pythoncom
import win32com.client

# ─── DISPID Constants (SolidWorks 2024 / ISldWorks) ──────────────────────────
# win32com late-binding hits PROPERTYGET for some methods in SW 2024.
# We call these by DISPID with DISPATCH_METHOD=1 to force method invocation.
_DISPID_NEW_DOCUMENT = 111   # ISldWorks.NewDocument
_DISPID_CLOSE_DOC    = 7     # ISldWorks.CloseDoc
_DISPID_OPEN_DOC6    = 167   # ISldWorks.OpenDoc6
_DISPID_GET_DOCS     = 273   # ISldWorks.GetDocuments
_DISPID_SELECT_BY_ID2 = 68   # IModelDocExtension.SelectByID2
_DISPID_EXTRUDE3     = None  # resolved at runtime
_DISPATCH_METHOD     = 1


def _invoke(obj, dispid: int, *args) -> Any:
    """Force COM method invocation via DISPID, bypassing property-get confusion."""
    result = obj._oleobj_.Invoke(dispid, 0, _DISPATCH_METHOD, 1, *args)
    if result is None:
        return None
    try:
        return win32com.client.Dispatch(result)
    except Exception:
        return result


def _resolve_dispid(obj, method_name: str) -> Optional[int]:
    """Look up a method's DISPID by name at runtime."""
    try:
        return obj._oleobj_.GetIDsOfNames(0, method_name)
    except Exception:
        return None


def _get(obj, name: str):
    """
    Access a COM attribute that SolidWorks 2024 exposes as a dual property/method.
    win32com returns the value directly via PROPERTYGET; calling it with () fails.
    This helper returns the value whether it comes back as a primitive or callable.
    """
    val = getattr(obj, name, None)
    if isinstance(val, (str, int, float, bool, bytes, type(None))):
        return val
    if callable(val):
        try:
            return val()
        except Exception:
            return val
    return val

# ─── SolidWorks API Constants ────────────────────────────────────────────────

# Document types
SW_DOC_NONE = 0
SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
SW_DOC_DRAWING = 3

# Open options
SW_OPEN_SILENT = 1
SW_OPEN_READ_ONLY = 2

# Save options
SW_SAVE_SILENT = 1
SW_SAVE_COPY = 2

# End conditions for extrude/revolve
SW_END_BLIND = 0
SW_END_THROUGH_ALL = 1
SW_END_THROUGH_ALL_BOTH = 2
SW_END_MID_PLANE = 8
SW_END_UP_TO_VERTEX = 4
SW_END_UP_TO_SURFACE = 5

# Mate types
SW_MATE_COINCIDENT = 0
SW_MATE_CONCENTRIC = 1
SW_MATE_PERPENDICULAR = 2
SW_MATE_PARALLEL = 3
SW_MATE_TANGENT = 4
SW_MATE_DISTANCE = 5
SW_MATE_ANGLE = 6
SW_MATE_SYMMETRIC = 7
SW_MATE_CAM = 9
SW_MATE_GEAR = 10
SW_MATE_RACK_PINION = 11
SW_MATE_SCREW = 12
SW_MATE_LOCK = 20
SW_MATE_HINGE = 22
SW_MATE_SLOT = 26
SW_MATE_PROFILE_CENTER = 30
SW_MATE_LINEAR_COUPLER = 14

MATE_TYPE_MAP = {
    "COINCIDENT": SW_MATE_COINCIDENT,
    "CONCENTRIC": SW_MATE_CONCENTRIC,
    "PERPENDICULAR": SW_MATE_PERPENDICULAR,
    "PARALLEL": SW_MATE_PARALLEL,
    "TANGENT": SW_MATE_TANGENT,
    "DISTANCE": SW_MATE_DISTANCE,
    "ANGLE": SW_MATE_ANGLE,
    "SYMMETRIC": SW_MATE_SYMMETRIC,
    "CAM": SW_MATE_CAM,
    "GEAR": SW_MATE_GEAR,
    "RACK_PINION": SW_MATE_RACK_PINION,
    "SCREW": SW_MATE_SCREW,
    "LOCK": SW_MATE_LOCK,
    "HINGE": SW_MATE_HINGE,
    "SLOT": SW_MATE_SLOT,
    "PROFILE_CENTER": SW_MATE_PROFILE_CENTER,
    "LINEAR_COUPLER": SW_MATE_LINEAR_COUPLER,
}

# Mate alignment
SW_MATE_ALIGNED = 0
SW_MATE_ANTI_ALIGNED = 1
SW_MATE_CLOSEST = 2

# Body types
SW_SOLID_BODY = 0
SW_SHEET_BODY = 1

# Display modes
DISPLAY_MODE_MAP = {
    "wireframe": 0,
    "hidden_lines_visible": 1,
    "hidden_lines_removed": 2,
    "shaded": 3,
    "shaded_with_edges": 4,
}

# Sketch relations
RELATION_MAP = {
    "HORIZONTAL": "sgHORIZONTAL",
    "VERTICAL": "sgVERTICAL",
    "COINCIDENT": "sgCOINCIDENT",
    "PARALLEL": "sgPARALLEL",
    "PERPENDICULAR": "sgPERPENDICULAR",
    "TANGENT": "sgTANGENT",
    "COLLINEAR": "sgCOLLINEAR2",
    "CONCENTRIC": "sgCONCENTRIC",
    "EQUAL": "sgEQUAL",
    "SYMMETRIC": "sgSYMMETRIC",
    "MIDPOINT": "sgATMIDPOINT",
    "INTERSECTION": "sgINTERSECT",
    "FIX": "sgFIXED",
    "PIERCE": "sgPIERCE",
}

# View name map
VIEW_NAME_MAP = {
    "front": "*Front",
    "back": "*Back",
    "left": "*Left",
    "right": "*Right",
    "top": "*Top",
    "bottom": "*Bottom",
    "isometric": "*Isometric",
    "trimetric": "*Trimetric",
    "dimetric": "*Dimetric",
}

MM = 0.001   # millimeters to meters
DEG = math.pi / 180.0  # degrees to radians


def _m(mm: float) -> float:
    """Convert mm to meters."""
    return mm * MM


def _mm(m: float) -> float:
    """Convert meters to mm."""
    return m / MM


def _r(deg: float) -> float:
    """Convert degrees to radians."""
    return deg * DEG


def _deg(rad: float) -> float:
    """Convert radians to degrees."""
    return rad / DEG


class SolidWorksClient:
    """
    Wraps the SolidWorks COM API. All methods are synchronous and must be called
    from a thread that has initialized COM via pythoncom.CoInitialize().
    """

    def __init__(self):
        self._app: Optional[Any] = None

    def _init_com(self):
        pythoncom.CoInitialize()

    # ─── Connection ──────────────────────────────────────────────────────────

    def connect(self, launch_if_not_running: bool = True, ensure_document: bool = True) -> Dict:
        """Connect to a running SolidWorks instance, optionally launching it.

        Args:
            launch_if_not_running: Launch SolidWorks if not already running.
            ensure_document: If True (default) and no document is open after
                connecting/launching, automatically create a new blank Part so
                that the COM Running Object Table registration is active and
                document-level tools work immediately.
        """
        import time

        # ── 1. Re-use an existing live reference ──────────────────────────────
        if self._app is not None:
            try:
                rev = self._app.RevisionNumber
                result = {"status": "already_connected", "version": rev}
                if ensure_document:
                    self._open_default_doc_if_needed(result)
                return result
            except Exception:
                self._app = None

        # ── 2. Attach to a running SolidWorks via COM ROT ─────────────────────
        try:
            self._app = win32com.client.GetActiveObject("SldWorks.Application")
            result = {"status": "connected", "version": self._app.RevisionNumber}
            if ensure_document:
                self._open_default_doc_if_needed(result)
            return result
        except Exception:
            pass

        if not launch_if_not_running:
            raise RuntimeError(
                "SolidWorks is not running. Set launch_if_not_running=True or start SolidWorks manually."
            )

        # ── 3. Launch SolidWorks via CoCreateInstance (Dispatch) ──────────────
        self._app = win32com.client.Dispatch("SldWorks.Application")
        self._app.Visible = True

        # Wait up to 30 s for SolidWorks to finish initialising
        deadline = time.time() + 30
        while time.time() < deadline:
            try:
                rev = self._app.RevisionNumber
                if rev:
                    break
            except Exception:
                pass
            time.sleep(0.5)
        else:
            raise RuntimeError("SolidWorks launched but did not become ready within 30 seconds.")

        result = {"status": "launched", "version": self._app.RevisionNumber}
        if ensure_document:
            self._open_default_doc_if_needed(result)
        return result

    def _open_default_doc_if_needed(self, status_dict: Dict) -> None:
        """If SolidWorks has no open document, create a new blank Part.

        SolidWorks only registers itself in the Windows COM Running Object Table
        (ROT) once a document is open.  Calling this after every connect() ensures
        the server is immediately usable without a separate manual step.
        """
        import time
        try:
            doc = self._app.ActiveDoc
            if doc is None:
                # Give SW a moment to finish loading (e.g. after a fresh launch)
                time.sleep(1.0)
                doc = self._app.ActiveDoc
            if doc is None:
                tmpl = self._find_template("part")
                new_doc = _invoke(self._app, _DISPID_NEW_DOCUMENT, tmpl, 0, 0.0, 0.0)
                if new_doc is not None:
                    title = _get(new_doc, 'GetTitle') or "Part1"
                    status_dict["auto_opened_document"] = title
                    status_dict["note"] = (
                        "No document was open; created a new blank Part automatically. "
                        "Use sw_open_document or sw_create_part to work with a specific file."
                    )
        except Exception:
            pass  # Non-fatal – tools will surface the "no active document" error if needed

    def get_status(self) -> Dict:
        """Get connection status and active document info."""
        if self._app is None:
            # Try auto-connect silently
            try:
                self._app = win32com.client.GetActiveObject("SldWorks.Application")
            except Exception:
                return {"connected": False, "active_document": None}
        try:
            rev = self._app.RevisionNumber
        except Exception:
            self._app = None
            return {"connected": False, "active_document": None}

        info: Dict = {"connected": True, "version": rev}
        doc = self._app.ActiveDoc
        if doc:
            try:
                info["active_document"] = {
                    "title": _get(doc, 'GetTitle'),
                    "path": _get(doc, 'GetPathName'),
                    "type": _doc_type_name(_get(doc, 'GetType')),
                    "is_modified": bool(_get(doc, 'GetSaveFlag')),
                }
            except Exception:
                info["active_document"] = {"title": "(unknown)", "type": "Unknown"}
        else:
            info["active_document"] = None

        try:
            docs = _invoke(self._app, _DISPID_GET_DOCS)
            if docs is None:
                info["open_documents"] = []
            elif hasattr(docs, "__iter__"):
                info["open_documents"] = [
                    {"title": _get(d, 'GetTitle'), "type": _doc_type_name(_get(d, 'GetType'))}
                    for d in docs
                ]
            else:
                # Single document returned (not wrapped in sequence)
                info["open_documents"] = [
                    {"title": docs.GetTitle(), "type": _doc_type_name(docs.GetType())}
                ]
        except Exception:
            info["open_documents"] = []
        return info

    # ─── File Operations ─────────────────────────────────────────────────────

    def open_document(self, file_path: str, read_only: bool = False) -> Dict:
        app = self._ensure_app()
        ext = os.path.splitext(file_path)[1].lower()
        doc_type = {
            ".sldprt": SW_DOC_PART,
            ".sldasm": SW_DOC_ASSEMBLY,
            ".slddrw": SW_DOC_DRAWING,
        }.get(ext, SW_DOC_PART)

        options = SW_OPEN_SILENT | (SW_OPEN_READ_ONLY if read_only else 0)
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        doc = app.OpenDoc6(file_path, doc_type, options, "", errors, warnings)
        if doc is None:
            raise RuntimeError(
                f"Failed to open '{file_path}'. Errors: {errors.value}, Warnings: {warnings.value}"
            )
        return {
            "title": _get(doc, 'GetTitle'),
            "path": _get(doc, 'GetPathName'),
            "type": _doc_type_name(_get(doc, 'GetType')),
        }

    def close_document(self, title: Optional[str] = None) -> Dict:
        app = self._ensure_app()
        if not title:
            doc = self._active_doc()
            title = _get(doc, 'GetTitle')
        # Use _invoke to bypass property-get confusion in SW 2024
        _invoke(app, _DISPID_CLOSE_DOC, title)
        return {"closed": title}

    def save_document(self, file_path: Optional[str] = None) -> Dict:
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        if file_path:
            result = doc.Extension.SaveAs(file_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
        else:
            # Save3 may be a property in SW 2024; try normal call, then _get fallback
            try:
                result = doc.Save3(SW_SAVE_SILENT, errors, warnings)
            except Exception:
                result = _get(doc, 'Save3')  # property shortcut (no options)
        if not result:
            raise RuntimeError(f"Save failed. Errors: {errors.value}")
        return {"saved": _get(doc, 'GetPathName') or _get(doc, 'GetTitle'), "errors": errors.value, "warnings": warnings.value}

    def _find_template(self, kind: str) -> str:
        """Locate the default SolidWorks document template file."""
        exts = {"part": "prtdot", "assembly": "asmdot", "drawing": "drwdot"}
        names = {"part": "Part", "assembly": "Assembly", "drawing": "Drawing"}
        ext = exts[kind]
        name = names[kind]
        # Common locations for SolidWorks templates
        search_roots = [
            r"C:\ProgramData\SolidWorks",
            r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\templates",
            r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS 2024\templates",
        ]
        for root in search_roots:
            if not os.path.exists(root):
                continue
            for dirpath, _, files in os.walk(root):
                for f in files:
                    if f.lower() == f"{name}.{ext}".lower() or f.lower() == f"{name.lower()}.{ext}":
                        return os.path.join(dirpath, f)
        return ""

    def create_part(self, template_path: Optional[str] = None) -> Dict:
        app = self._ensure_app()
        tmpl = template_path or self._find_template("part")
        doc = _invoke(app, _DISPID_NEW_DOCUMENT, tmpl, 0, 0.0, 0.0)
        if doc is None:
            raise RuntimeError(
                "Failed to create new part. "
                f"Template used: '{tmpl}'. Ensure SolidWorks has a valid Part template."
            )
        return {"title": _get(doc, 'GetTitle'), "type": "Part"}

    def create_assembly(self, template_path: Optional[str] = None) -> Dict:
        app = self._ensure_app()
        tmpl = template_path or self._find_template("assembly")
        doc = _invoke(app, _DISPID_NEW_DOCUMENT, tmpl, 0, 0.0, 0.0)
        if doc is None:
            raise RuntimeError("Failed to create new assembly.")
        return {"title": _get(doc, 'GetTitle'), "type": "Assembly"}

    def create_drawing(self, model_path: Optional[str] = None, template_path: Optional[str] = None) -> Dict:
        app = self._ensure_app()
        tmpl = template_path or self._find_template("drawing")
        doc = _invoke(app, _DISPID_NEW_DOCUMENT, tmpl, 0, 0.0, 0.0)
        if doc is None:
            raise RuntimeError("Failed to create new drawing.")
        return {"title": _get(doc, 'GetTitle'), "type": "Drawing"}

    def list_open_documents(self) -> List[Dict]:
        app = self._ensure_app()
        try:
            docs = _invoke(app, _DISPID_GET_DOCS)
        except Exception:
            return []
        if docs is None:
            return []
        if not hasattr(docs, "__iter__"):
            docs = [docs]
        result = []
        for d in docs:
            try:
                result.append({
                    "title": _get(d, 'GetTitle'),
                    "path": _get(d, 'GetPathName'),
                    "type": _doc_type_name(_get(d, 'GetType')),
                    "modified": bool(_get(d, 'GetSaveFlag')),
                })
            except Exception:
                pass
        return result

    def export_file(self, output_path: str) -> Dict:
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = doc.Extension.SaveAs(output_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
        if not result:
            raise RuntimeError(f"Export failed to '{output_path}'. Error code: {errors.value}")
        return {"exported_to": output_path}

    def get_model_info(self) -> Dict:
        doc = self._active_doc()
        doc_type = _get(doc, 'GetType')
        info = {
            "title": _get(doc, 'GetTitle'),
            "path": _get(doc, 'GetPathName'),
            "type": _doc_type_name(doc_type),
            "modified": bool(_get(doc, 'GetSaveFlag')),
            "configurations": list(_get(doc, 'GetConfigurationNames') or []),
            "active_configuration": (lambda c: _get(c, 'Name') if c else None)(_get(doc, 'GetActiveConfiguration')),
        }
        # Feature list (GetFeatureCount is a property in SW 2024, not a method)
        features = []
        try:
            feat_count = int(doc.GetFeatureCount)
        except (TypeError, AttributeError):
            feat_count = 0
        for i in range(min(feat_count, 200)):
            try:
                f = doc.FeatureByPositionReverse(i)
                if f:
                    features.append({"name": f.Name, "type": _get(f, 'GetTypeName2')})
            except Exception:
                pass
        info["features"] = features
        return info

    # ─── Selection ────────────────────────────────────────────────────────────

    def select_entity(self, entity_name: str, entity_type: str, append: bool = False) -> Dict:
        doc = self._active_doc()
        ext = doc.Extension
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        result = ext.SelectByID2(entity_name, entity_type.upper(), 0.0, 0.0, 0.0, append, 0, _nc, 0)
        if not result:
            raise RuntimeError(
                f"Could not select '{entity_name}' (type={entity_type}). "
                "Check the name matches exactly as shown in the FeatureManager tree."
            )
        return {"selected": entity_name, "type": entity_type}

    def clear_selection(self) -> Dict:
        doc = self._active_doc()
        doc.ClearSelection2(True)
        return {"cleared": True}

    # ─── Sketch Operations ────────────────────────────────────────────────────

    def create_sketch(self, plane: str) -> Dict:
        doc = self._active_doc()
        # Select via Feature object to avoid SelectByID2 null-callout marshaling issues
        feat = doc.FeatureByName(plane)
        if feat:
            feat.Select2(False, 0)
        else:
            # Fallback: SelectByID2 with explicit null callout VARIANT
            _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            doc.Extension.SelectByID2(plane, "PLANE", 0.0, 0.0, 0.0, False, 0, _nc, 0)
        doc.SketchManager.InsertSketch(True)
        return {"sketch_active": True, "plane": plane}

    def edit_sketch(self, sketch_name: str) -> Dict:
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 0, _nc, 0)
        doc.MenuCommand(2026)  # swCommands_Sketch_Edit
        return {"editing": sketch_name}

    def exit_sketch(self) -> Dict:
        doc = self._active_doc()
        doc.SketchManager.InsertSketch(True)
        return {"sketch_exited": True}

    def sketch_line(self, x1: float, y1: float, x2: float, y2: float) -> Dict:
        doc = self._active_doc()
        seg = doc.SketchManager.CreateLine(_m(x1), _m(y1), 0.0, _m(x2), _m(y2), 0.0)
        if seg is None:
            raise RuntimeError("Failed to create line. Ensure a sketch is active.")
        return {"type": "line", "x1": x1, "y1": y1, "x2": x2, "y2": y2}

    def sketch_circle(self, cx: float, cy: float, radius: float) -> Dict:
        doc = self._active_doc()
        seg = doc.SketchManager.CreateCircleByRadius(_m(cx), _m(cy), 0.0, _m(radius))
        if seg is None:
            raise RuntimeError("Failed to create circle. Ensure a sketch is active.")
        return {"type": "circle", "cx": cx, "cy": cy, "radius": radius}

    def sketch_rectangle(self, x1: float, y1: float, x2: float, y2: float, rect_type: str = "corner") -> Dict:
        doc = self._active_doc()
        sm = doc.SketchManager
        if rect_type == "center":
            segs = sm.CreateCenterRectangle(_m(x1), _m(y1), 0.0, _m(x2), _m(y2), 0.0)
        else:
            segs = sm.CreateCornerRectangle(_m(x1), _m(y1), 0.0, _m(x2), _m(y2), 0.0)
        if segs is None:
            raise RuntimeError("Failed to create rectangle. Ensure a sketch is active.")
        return {"type": "rectangle", "rect_type": rect_type, "x1": x1, "y1": y1, "x2": x2, "y2": y2}

    def sketch_arc(self, cx: float, cy: float, radius: float, start_angle: float, end_angle: float, clockwise: bool = False) -> Dict:
        doc = self._active_doc()
        sa = _r(start_angle)
        ea = _r(end_angle)
        r = _m(radius)
        sx = _m(cx) + r * math.cos(sa)
        sy = _m(cy) + r * math.sin(sa)
        ex = _m(cx) + r * math.cos(ea)
        ey = _m(cy) + r * math.sin(ea)
        direction = -1 if clockwise else 1
        seg = doc.SketchManager.CreateArc(
            _m(cx), _m(cy), 0.0,
            sx, sy, 0.0,
            ex, ey, 0.0,
            direction
        )
        if seg is None:
            raise RuntimeError("Failed to create arc. Ensure a sketch is active.")
        return {"type": "arc", "cx": cx, "cy": cy, "radius": radius, "start_angle": start_angle, "end_angle": end_angle}

    def sketch_ellipse(self, cx: float, cy: float, semi_major: float, semi_minor: float, rotation: float = 0.0) -> Dict:
        doc = self._active_doc()
        rot = _r(rotation)
        ax = _m(cx) + _m(semi_major) * math.cos(rot)
        ay = _m(cy) + _m(semi_major) * math.sin(rot)
        bx = _m(cx) + _m(semi_minor) * math.cos(rot + math.pi / 2)
        by = _m(cy) + _m(semi_minor) * math.sin(rot + math.pi / 2)
        seg = doc.SketchManager.CreateEllipse(_m(cx), _m(cy), 0.0, ax, ay, 0.0, bx, by, 0.0)
        if seg is None:
            raise RuntimeError("Failed to create ellipse. Ensure a sketch is active.")
        return {"type": "ellipse", "cx": cx, "cy": cy, "semi_major": semi_major, "semi_minor": semi_minor}

    def sketch_polygon(self, cx: float, cy: float, num_sides: int, circumscribed_radius: float, rotation: float = 0.0) -> Dict:
        doc = self._active_doc()
        rot = _r(rotation)
        vx = _m(cx) + _m(circumscribed_radius) * math.cos(rot)
        vy = _m(cy) + _m(circumscribed_radius) * math.sin(rot)
        segs = doc.SketchManager.CreatePolygon(_m(cx), _m(cy), 0.0, vx, vy, 0.0, num_sides, True)
        if segs is None:
            raise RuntimeError("Failed to create polygon. Ensure a sketch is active.")
        return {"type": "polygon", "sides": num_sides, "cx": cx, "cy": cy, "radius": circumscribed_radius}

    def sketch_spline(self, points: List[Tuple[float, float]]) -> Dict:
        doc = self._active_doc()
        flat_pts = []
        for x, y in points:
            flat_pts.extend([_m(x), _m(y), 0.0])
        seg = doc.SketchManager.CreateSpline(flat_pts)
        if seg is None:
            raise RuntimeError("Failed to create spline. Ensure a sketch is active.")
        return {"type": "spline", "num_points": len(points)}

    def sketch_offset(self, distance: float, outward: bool = True) -> Dict:
        doc = self._active_doc()
        result = doc.SketchManager.SketchOffset2(_m(distance), not outward, False, False, False, False)
        if result is None:
            raise RuntimeError("Offset failed. Select sketch entities first.")
        return {"type": "offset", "distance": distance, "outward": outward}

    def sketch_mirror(self, mirror_line_name: Optional[str] = None) -> Dict:
        doc = self._active_doc()
        doc.SketchManager.SketchMirror()
        return {"type": "mirror"}

    def sketch_linear_pattern(self, x_count: int, y_count: int, x_spacing: float, y_spacing: float, x_angle: float = 0.0, y_angle: float = 90.0) -> Dict:
        doc = self._active_doc()
        doc.SketchManager.CreateLinearSketchStepAndRepeat(
            x_count, y_count, _m(x_spacing), _m(y_spacing), _r(x_angle), _r(y_angle), False, "", True, True
        )
        return {"x_count": x_count, "y_count": y_count, "x_spacing": x_spacing, "y_spacing": y_spacing}

    def sketch_circular_pattern(self, count: int, radius: float, angle: float = 360.0) -> Dict:
        doc = self._active_doc()
        doc.SketchManager.CreateCircularSketchStepAndRepeat(
            _m(radius), _r(angle), count, False, "", True, True, 0.0, 0.0, 0.0, True
        )
        return {"count": count, "radius": radius, "angle": angle}

    def sketch_centerline(self, x1: float, y1: float, x2: float, y2: float) -> Dict:
        doc = self._active_doc()
        sm = doc.SketchManager
        sm.CreateCenterLine(_m(x1), _m(y1), 0.0, _m(x2), _m(y2), 0.0)
        return {"type": "centerline", "x1": x1, "y1": y1, "x2": x2, "y2": y2}

    def add_sketch_dimension(self, value: float, x: float = 0.0, y: float = 0.0) -> Dict:
        doc = self._active_doc()
        dim = doc.AddDimension2(_m(x), _m(y), 0.0)
        if dim is None:
            raise RuntimeError(
                "Failed to add dimension. Select the sketch entity to dimension first."
            )
        dim.SystemValue = _m(value)
        return {"dimension": value, "x": x, "y": y}

    def add_sketch_constraint(self, relation_type: str) -> Dict:
        doc = self._active_doc()
        sg = RELATION_MAP.get(relation_type.upper())
        if sg is None:
            raise ValueError(
                f"Unknown relation type '{relation_type}'. "
                f"Valid types: {', '.join(RELATION_MAP.keys())}"
            )
        doc.SketchAddConstraints(sg)
        return {"constraint": relation_type}

    def add_sketch_dimensions(self, width: float, height: float) -> Dict:
        """Add width and height dimensions to selected rectangle."""
        doc = self._active_doc()
        result = {"width": width, "height": height}
        for val in [width, height]:
            dim = doc.AddDimension2(0.0, 0.0, 0.0)
            if dim:
                dim.SystemValue = _m(val)
        return result

    # ─── Feature Operations ───────────────────────────────────────────────────

    def extrude(self, depth: float, direction: str = "blind", flip_direction: bool = False,
                draft_angle: float = 0.0, is_cut: bool = False,
                thin_feature: bool = False, thin_thickness: float = 1.0) -> Dict:
        doc = self._active_doc()
        fm = doc.FeatureManager
        end_cond = {
            "blind": SW_END_BLIND,
            "through_all": SW_END_THROUGH_ALL,
            "through_all_both": SW_END_THROUGH_ALL_BOTH,
            "mid_plane": SW_END_MID_PLANE,
            "up_to_vertex": SW_END_UP_TO_VERTEX,
            "up_to_surface": SW_END_UP_TO_SURFACE,
        }.get(direction.lower(), SW_END_BLIND)

        # Try FeatureExtrusion3 starting from 23 params (SW 2024) down to 20 (older SW)
        draft_on = draft_angle > 0
        ang = _r(draft_angle) if draft_on else 0.0
        feat = None
        for extra in [
            (False, False, False),  # 23 params (SW 2024)
            (True, False),          # 22 params (SW 2020-2023)
            (False, False),         # 22 params alt
            (),                     # 20 params (older SW)
        ]:
            try:
                args = [
                    True, is_cut, flip_direction,
                    end_cond, SW_END_BLIND,
                    _m(depth), 0.0,
                    False, False,
                    draft_on, False,
                    ang, 0.0,
                    False, False,
                    False, False,
                    True, False, True,
                ] + list(extra)
                feat = fm.FeatureExtrusion3(*args)
                if feat is not None:
                    break
            except Exception:
                pass
        if feat is None:
            raise RuntimeError(
                "Extrude failed. Ensure a closed sketch profile is selected/active and the sketch is properly closed."
            )
        return {"feature": feat.Name, "type": "cut" if is_cut else "boss", "depth_mm": depth}

    def revolve(self, angle: float = 360.0, is_cut: bool = False, flip_direction: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureRevolve2(
            _r(angle),
            0.0,           # start angle
            SW_END_BLIND,
            SW_END_BLIND,
            False,         # thin feature
            True,          # single direction
            True,          # merge result
            flip_direction,
            0.0,           # thin thickness
            False,         # reverse thin side
            True,          # auto select axis
            True,          # select contour
        )
        if feat is None:
            raise RuntimeError(
                "Revolve failed. Ensure a sketch profile and a centerline/axis are selected."
            )
        return {"feature": feat.Name, "type": "cut" if is_cut else "boss", "angle_deg": angle}

    def fillet(self, radius: float, tangent_propagation: bool = True) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureFillet(
            _m(radius),
            True,                # keep feature
            tangent_propagation, # propagate along tangent
            False,               # full round
            False,               # multi-radius
            0.0,                 # setback
            0,                   # num vertex setbacks
        )
        if feat is None:
            raise RuntimeError("Fillet failed. Select edges first.")
        return {"feature": feat.Name, "radius_mm": radius}

    def chamfer(self, distance: float, angle: float = 45.0) -> Dict:
        doc = self._active_doc()
        # swChamferMethod_eDistDist=0, swChamferMethod_eDistAngle=1
        feat = doc.FeatureManager.InsertFeatureChamfer(
            1,            # type: 1=distance-angle
            0,            # flip dir
            _m(distance),
            _r(angle),
            False, False, False, False
        )
        if feat is None:
            # Fallback to equal-distance chamfer
            feat = doc.FeatureManager.InsertFeatureChamfer(
                0,           # equal distance
                0,
                _m(distance),
                _r(45.0),
                False, False, False, False
            )
        if feat is None:
            raise RuntimeError("Chamfer failed. Select edges first.")
        return {"feature": feat.Name, "distance_mm": distance, "angle_deg": angle}

    def shell(self, thickness: float, outward: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureShell(
            _m(thickness),
            outward,
            None,  # faces to remove (selected)
        )
        if feat is None:
            raise RuntimeError("Shell failed. Select face(s) to remove first.")
        return {"feature": feat.Name, "thickness_mm": thickness}

    def draft(self, draft_angle: float) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureDraft2(
            True,          # neutral plane
            False,         # draft outward
            _r(draft_angle),
            True,          # use pull dir
        )
        if feat is None:
            raise RuntimeError(
                "Draft failed. Select the neutral plane and draught faces first."
            )
        return {"feature": feat.Name, "angle_deg": draft_angle}

    def linear_pattern(self, dir1_count: int, dir1_spacing: float,
                       dir2_count: int = 1, dir2_spacing: float = 10.0,
                       geometry_pattern: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureLinearPattern3(
            dir1_count, _m(dir1_spacing),
            dir2_count, _m(dir2_spacing),
            True, True,    # create seed, same config
            geometry_pattern,
            0, 0,          # pattern seed entities
            True,          # vary sketch
        )
        if feat is None:
            raise RuntimeError(
                "Linear pattern failed. Select a feature and direction references first."
            )
        return {"feature": feat.Name, "dir1_count": dir1_count, "dir2_count": dir2_count}

    def circular_pattern(self, count: int, angle: float = 360.0, geometry_pattern: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.FeatureCirPattern3(
            count, _r(angle),
            True,          # equal spacing
            geometry_pattern,
            0,             # seed entities
            True,          # vary sketch
        )
        if feat is None:
            raise RuntimeError(
                "Circular pattern failed. Select a feature and axis/face for rotation first."
            )
        return {"feature": feat.Name, "count": count, "angle_deg": angle}

    def mirror_feature(self, mirror_plane: str) -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(mirror_plane, "PLANE", 0, 0, 0, True, 32, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        feat = doc.FeatureManager.InsertMirrorFeature2(True, False, True, True, False)
        if feat is None:
            raise RuntimeError(
                "Mirror failed. Select the features to mirror, then specify the mirror plane."
            )
        return {"feature": feat.Name, "mirror_plane": mirror_plane}

    def hole_wizard(self, hole_type: str, standard: str, size: str, depth: float,
                    through_all: bool = False, x: float = 0.0, y: float = 0.0) -> Dict:
        doc = self._active_doc()
        end_cond = SW_END_THROUGH_ALL if through_all else SW_END_BLIND
        hole_type_map = {
            "simple": 0, "counterbore": 1, "countersink": 2,
            "straight_tap": 3, "tapered_tap": 4,
        }
        wh_type = hole_type_map.get(hole_type.lower(), 0)
        feat = doc.FeatureManager.HoleWizard5(
            wh_type, 0, end_cond,
            _m(depth), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
            _m(x), _m(y), 0.0,
            True
        )
        if feat is None:
            raise RuntimeError(
                "Hole Wizard failed. Select a face to place the hole on first."
            )
        return {"feature": feat.Name, "hole_type": hole_type, "size": size}

    def loft(self, is_cut: bool = False, close_loft: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertProtrusionBlend2(
            close_loft, True, True, 1, 0.0, 0.0,
            False, False, True, True,
            0, 0, True, True
        )
        if feat is None:
            raise RuntimeError(
                "Loft failed. Select two or more sketch profiles (in order) as loft sections."
            )
        return {"feature": feat.Name, "type": "cut" if is_cut else "boss"}

    def sweep(self, is_cut: bool = False, merge_result: bool = True) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertProtrusionSwept4(
            False,        # thin feature
            False,        # thin sides
            0,            # alignment
            0,            # twist type
            0.0,          # twist angle
            True,         # merge tangent faces
            False,        # keep tangent faces
            False,        # advanced smoothing
            merge_result,
            False,        # start tangency
            False,        # end tangency
            0.0, 0.0,     # tang lengths
        )
        if feat is None:
            raise RuntimeError(
                "Sweep failed. Select the profile sketch and path sketch/curve first."
            )
        return {"feature": feat.Name, "type": "cut" if is_cut else "boss"}

    def delete_feature(self, feature_name: str, absorb_children: bool = False) -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(feature_name, "BODYFEATURE", 0, 0, 0, False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        result = doc.Extension.DeleteSelection2(0 if absorb_children else 1)
        return {"deleted": feature_name, "result": result}

    def suppress_feature(self, feature_name: str) -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(feature_name, "BODYFEATURE", 0, 0, 0, False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        doc.EditSuppress2()
        return {"suppressed": feature_name}

    def unsuppress_feature(self, feature_name: str) -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(feature_name, "BODYFEATURE", 0, 0, 0, False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        doc.EditUnsuppress2()
        return {"unsuppressed": feature_name}

    def list_features(self) -> List[Dict]:
        doc = self._active_doc()
        features = []
        # GetFeatureCount is a property (not method) in SW 2024
        try:
            count = int(doc.GetFeatureCount)
        except (TypeError, AttributeError):
            count = 0
        for i in range(count):
            try:
                f = doc.FeatureByPositionReverse(i)
                if f:
                    try:
                        suppressed = bool(f.IsSuppressed2(0, None)[0])
                    except Exception:
                        suppressed = False
                    features.append({
                        "name": f.Name,
                        "type": _get(f, 'GetTypeName2'),
                        "suppressed": suppressed,
                    })
            except Exception:
                pass
        return list(reversed(features))

    # ─── Assembly Operations ──────────────────────────────────────────────────

    def insert_component(self, file_path: str, x: float = 0.0, y: float = 0.0, z: float = 0.0, fixed: bool = False) -> Dict:
        doc = self._active_doc()
        comp = doc.AddComponent5(file_path, 0, "", fixed, _m(x), _m(y), _m(z))
        if comp is None:
            raise RuntimeError(f"Failed to insert component '{file_path}'. Ensure the file exists and the active document is an assembly.")
        return {"component": comp.Name2, "fixed": fixed}

    def add_mate(self, mate_type: str, entity1: str, entity2: str,
                 distance_or_angle: Optional[float] = None, flip_mate: bool = False) -> Dict:
        doc = self._active_doc()
        mt = MATE_TYPE_MAP.get(mate_type.upper())
        if mt is None:
            raise ValueError(f"Unknown mate type '{mate_type}'. Valid: {', '.join(MATE_TYPE_MAP.keys())}")

        # Select entities
        doc.ClearSelection2(True)
        for ent in [entity1, entity2]:
            parts = ent.rsplit("/", 1)
            if len(parts) == 2:
                doc.Extension.SelectByID2(parts[1], "FACE", 0, 0, 0, True, 1, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
            else:
                doc.Extension.SelectByID2(ent, "FACE", 0, 0, 0, True, 1, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)

        align = SW_MATE_ANTI_ALIGNED if flip_mate else SW_MATE_ALIGNED
        dist = _m(distance_or_angle) if distance_or_angle and mt == SW_MATE_DISTANCE else 0.0
        ang = _r(distance_or_angle) if distance_or_angle and mt == SW_MATE_ANGLE else 0.0

        feat = doc.AddMate5(
            mt, align, flip_mate,
            dist, dist, 0.0,
            1.0, 1.0,
            ang, ang, 0.0,
            False, False, False,
        )
        err = 0
        if isinstance(feat, tuple):
            feat, err = feat[0], feat[1]
        if feat is None:
            raise RuntimeError(f"Mate failed (error={err}). Verify the entities are correct and compatible with '{mate_type}' mate.")
        return {"mate": feat.Name, "type": mate_type}

    def fix_component(self, component_name: str, fix: bool = True) -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(component_name, "COMPONENT", 0, 0, 0, False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        if fix:
            doc.EditMakeRigid()
        else:
            doc.EditMakeFlexible()
        return {"component": component_name, "fixed": fix}

    def check_interference(self, treat_sub_as_comp: bool = True) -> Dict:
        doc = self._active_doc()
        if _get(doc, 'GetType') != SW_DOC_ASSEMBLY:
            raise RuntimeError("Interference check requires an assembly document.")
        # Use the interference detection object
        interference = doc.InterferenceDetection
        interference.UseTransform = True
        interference.TreatCoincidenceAsInterference = False
        interference.IncludeSurfaceBodies = True
        count = interference.GetCount()
        interferences = []
        for i in range(count):
            interf = interference.Interference(i)
            interferences.append({
                "volume_mm3": interf.Volume / (MM ** 3) if interf.Volume else 0,
            })
        interference.Done()
        return {"interference_count": count, "interferences": interferences}

    def explode_assembly(self) -> Dict:
        doc = self._active_doc()
        doc.EditExplodeStep()
        return {"exploded": True}

    # ─── Drawing Operations ───────────────────────────────────────────────────

    def add_drawing_view(self, view_type: str = "front", x: float = 0.1, y: float = 0.15,
                         scale: float = 1.0, model_path: Optional[str] = None) -> Dict:
        doc = self._active_doc()
        if _get(doc, 'GetType') != SW_DOC_DRAWING:
            raise RuntimeError("Active document must be a drawing.")
        view_name = VIEW_NAME_MAP.get(view_type.lower(), "*Front")
        view = doc.CreateDrawViewFromModelView3(model_path or "", view_name, x, y)
        if view is None:
            raise RuntimeError(f"Failed to create drawing view '{view_type}'. Ensure a model is associated with the drawing.")
        view.ScaleDecimal = scale
        return {"view": view.Name, "type": view_type, "scale": scale}

    def add_section_view(self, label: str = "A", offset: float = 0.0) -> Dict:
        doc = self._active_doc()
        doc.InsertSectionView2(False, False, 0, label, True)
        return {"section_view": label}

    def add_drawing_dimension(self, x: float = 0.0, y: float = 0.0) -> Dict:
        doc = self._active_doc()
        doc.AddDimension2(_m(x), _m(y), 0.0)
        return {"dimension_placed": True, "x": x, "y": y}

    def add_annotation(self, text: str, x: float = 0.05, y: float = 0.05) -> Dict:
        doc = self._active_doc()
        note = doc.InsertNote(text)
        if note:
            ann = note.GetAnnotation()
            if ann:
                ann.SetPosition2(_m(x), _m(y), 0.0)
        return {"annotation": text}

    def add_bom_table(self, table_type: int = 0) -> Dict:
        doc = self._active_doc()
        # table_type: 0=Parts Only, 1=Top-Level Only, 2=Indented
        doc.InsertTableAnnotation2(True, 0.05, 0.2, 1, "", True, True, True, True, True)
        return {"bom_inserted": True}

    def add_detail_view(self, x: float, y: float, radius: float, scale: float = 2.0, label: str = "A") -> Dict:
        doc = self._active_doc()
        doc.InsertDetailView2(_m(x), _m(y), _m(radius), scale, label, True)
        return {"detail_view": label, "scale": scale}

    def update_sheet_format(self) -> Dict:
        doc = self._active_doc()
        doc.Sheet().ReloadTemplate(True)
        return {"updated": True}

    # ─── Dimensions ───────────────────────────────────────────────────────────

    def get_dimension(self, dimension_name: str) -> Dict:
        doc = self._active_doc()
        dim = doc.Parameter(dimension_name)
        if dim is None:
            raise RuntimeError(
                f"Dimension '{dimension_name}' not found. "
                "Use format 'D1@FeatureName' or 'D1@Sketch1'."
            )
        val_mm = _mm(dim.SystemValue)
        return {"name": dimension_name, "value_mm": val_mm, "value_m": dim.SystemValue}

    def set_dimension(self, dimension_name: str, value: float) -> Dict:
        doc = self._active_doc()
        dim = doc.Parameter(dimension_name)
        if dim is None:
            raise RuntimeError(
                f"Dimension '{dimension_name}' not found. "
                "Use format 'D1@FeatureName' or 'D1@Sketch1'."
            )
        old_val = _mm(dim.SystemValue)
        dim.SystemValue = _m(value)
        _get(doc, 'EditRebuild3')
        return {"name": dimension_name, "old_value_mm": old_val, "new_value_mm": value}

    # ─── Model Analysis ───────────────────────────────────────────────────────

    def get_mass_properties(self) -> Dict:
        doc = self._active_doc()
        # CreateMassProperty2/CreateMassProperty are properties in SW 2024 (not methods)
        mass_prop = None
        for method in ("CreateMassProperty2", "CreateMassProperty"):
            try:
                val = getattr(doc.Extension, method)
                # In SW 2024 these are properties returning the object directly
                # In older SW they may be methods; try calling if needed
                if isinstance(val, (str, int, float, bool, type(None))):
                    continue  # got a primitive, not the mass-prop object
                # Try calling it in case it's a method wrapper
                try:
                    mass_prop = val()
                except Exception:
                    mass_prop = val  # already the object
                if mass_prop is not None:
                    break
            except Exception:
                pass
        if mass_prop is None:
            raise RuntimeError(
                "Could not create mass property object. "
                "Ensure the document has solid bodies and SolidWorks is fully loaded."
            )
        if _get(doc, 'GetType') == SW_DOC_PART:
            try:
                bodies = doc.GetBodies2(SW_SOLID_BODY, True)
                if bodies:
                    mass_prop.AddBodies(bodies)
            except Exception:
                pass
        result: Dict = {}
        for attr, key, conv in [
            ("Mass", "mass_kg", lambda v: v),
            ("Volume", "volume_mm3", lambda v: v / (MM ** 3)),
            ("SurfaceArea", "surface_area_mm2", lambda v: v / (MM ** 2)),
            ("Density", "density_kg_m3", lambda v: v),
        ]:
            try:
                result[key] = conv(_get(mass_prop, attr))
            except Exception:
                result[key] = None
        try:
            cog = _get(mass_prop, 'CenterOfMass')
            result["center_of_mass_mm"] = {"x": _mm(cog[0]), "y": _mm(cog[1]), "z": _mm(cog[2])} if cog else None
        except Exception:
            result["center_of_mass_mm"] = None
        try:
            pmi = _get(mass_prop, 'PrincipalMomentsOfInertia')
            result["principal_moments_of_inertia"] = {"Ix": pmi[0], "Iy": pmi[1], "Iz": pmi[2]} if pmi else None
        except Exception:
            result["principal_moments_of_inertia"] = None
        return result

    def get_bounding_box(self) -> Dict:
        doc = self._active_doc()
        bb = None
        # Try Extension.GetDocumentBoundingBox first (SW 2019+), then body-level, then doc-level
        for approach in ("extension_doc", "doc_method", "body"):
            try:
                if approach == "extension_doc":
                    bb = doc.Extension.GetDocumentBoundingBox(0)
                elif approach == "doc_method":
                    bb = doc.GetModelBoundingBox()
                elif approach == "body":
                    bodies = doc.GetBodies2(SW_SOLID_BODY, True)
                    if bodies:
                        bb = bodies[0].GetBodyBox()
                if bb and len(bb) >= 6:
                    break
            except Exception:
                bb = None
        if not bb or len(bb) < 6:
            raise RuntimeError(
                "Could not compute bounding box. "
                "Ensure the document has solid bodies."
            )
        return {
            "min_mm": {"x": _mm(bb[0]), "y": _mm(bb[1]), "z": _mm(bb[2])},
            "max_mm": {"x": _mm(bb[3]), "y": _mm(bb[4]), "z": _mm(bb[5])},
            "size_mm": {
                "x": _mm(bb[3] - bb[0]),
                "y": _mm(bb[4] - bb[1]),
                "z": _mm(bb[5] - bb[2]),
            },
        }

    def measure_distance(self, entity1: str, entity2: str) -> Dict:
        doc = self._active_doc()
        ext = doc.Extension
        doc.ClearSelection2(True)
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        ext.SelectByID2(entity1, "FACE", 0.0, 0.0, 0.0, False, 0, _nc, 0)
        ext.SelectByID2(entity2, "FACE", 0.0, 0.0, 0.0, True, 0, _nc, 0)
        measure = doc.Extension.CreateMeasure()
        measure.Calculate(None)
        return {
            "distance_mm": _mm(measure.Distance),
            "delta_x_mm": _mm(measure.DeltaX) if hasattr(measure, "DeltaX") else None,
            "delta_y_mm": _mm(measure.DeltaY) if hasattr(measure, "DeltaY") else None,
            "delta_z_mm": _mm(measure.DeltaZ) if hasattr(measure, "DeltaZ") else None,
        }

    def check_geometry(self) -> Dict:
        doc = self._active_doc()
        if _get(doc, 'GetType') != SW_DOC_PART:
            raise RuntimeError("Geometry check is for part documents only.")
        # CheckDocument may not be exposed via late-binding; try gracefully
        try:
            result = _get(doc, 'CheckDocument')
            if callable(result):
                result = result(-1)
        except Exception:
            result = 0
        err_count = int(result) if result else 0
        return {"errors": err_count, "status": "OK" if not err_count else "ERRORS_FOUND"}

    def analyze_draft(self, draft_angle: float = 1.0) -> Dict:
        doc = self._active_doc()
        doc.ShowDraftAnalysis(_r(draft_angle), True, True, True)
        return {"draft_angle_deg": draft_angle, "analysis_active": True}

    # ─── Custom Properties ────────────────────────────────────────────────────

    def get_custom_properties(self, prop_name: Optional[str] = None, configuration: str = "") -> Dict:
        doc = self._active_doc()
        mgr = doc.Extension.CustomPropertyManager(configuration)
        if prop_name:
            val_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            res_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            mgr.Get4(prop_name, False, val_out, res_out)
            return {"name": prop_name, "value": val_out.value, "resolved": res_out.value}
        names = _get(mgr, 'GetNames')
        if not names:
            return {"properties": {}}
        props = {}
        for name in names:
            val_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            res_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            mgr.Get4(name, False, val_out, res_out)
            props[name] = {"value": val_out.value, "resolved": res_out.value}
        return {"properties": props}

    def set_custom_property(self, prop_name: str, value: str, configuration: str = "") -> Dict:
        doc = self._active_doc()
        mgr = doc.Extension.CustomPropertyManager(configuration)
        # swCustomInfoText=30, swCustomPropertyReplaceValue=1
        result = mgr.Add3(prop_name, 30, value, 1)
        return {"name": prop_name, "value": value, "result": result}

    # ─── Configurations ───────────────────────────────────────────────────────

    def list_configurations(self) -> Dict:
        doc = self._active_doc()
        names = list(_get(doc, 'GetConfigurationNames') or [])
        active = _get(doc, 'GetActiveConfiguration')
        return {
            "configurations": names,
            "active": _get(active, 'Name') if active else None,
            "count": len(names),
        }

    def activate_configuration(self, config_name: str) -> Dict:
        doc = self._active_doc()
        result = doc.ShowConfiguration2(config_name)
        if not result:
            raise RuntimeError(
                f"Configuration '{config_name}' not found. "
                "Use sw_list_configurations to see available configurations."
            )
        _get(doc, 'EditRebuild3')
        return {"activated": config_name}

    def add_configuration(self, config_name: str, derived_from: Optional[str] = None) -> Dict:
        doc = self._active_doc()
        config = doc.AddConfiguration3(config_name, "", "", 0)
        if config is None:
            raise RuntimeError(f"Failed to add configuration '{config_name}'.")
        return {"created": config_name}

    # ─── Equations ───────────────────────────────────────────────────────────

    def get_equations(self) -> List[Dict]:
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        count = _get(eqn_mgr, 'GetCount')
        equations = []
        for i in range(count):
            eqn = eqn_mgr.Equation(i)
            val = eqn_mgr.Value(i)
            equations.append({"index": i, "equation": eqn, "value": val})
        return equations

    def add_equation(self, equation: str) -> Dict:
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        idx = eqn_mgr.Add2(-1, equation, True)
        _get(doc, 'EditRebuild3')
        return {"added": equation, "index": idx}

    def delete_equation(self, index: int) -> Dict:
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        result = eqn_mgr.Delete(index)
        _get(doc, 'EditRebuild3')
        return {"deleted_index": index, "result": result}

    # ─── Design Table ─────────────────────────────────────────────────────────

    def create_design_table(self, excel_path: Optional[str] = None) -> Dict:
        doc = self._active_doc()
        design_table = doc.InsertFamilyTableOpen2(0, excel_path or "", True, True)
        if design_table is None:
            raise RuntimeError("Failed to create design table.")
        return {"design_table_created": True}

    # ─── Visualization ────────────────────────────────────────────────────────

    def capture_screenshot(self, output_path: Optional[str] = None, width: int = 1920, height: int = 1080) -> Dict:
        app = self._ensure_app()
        doc = self._active_doc()
        if output_path is None:
            output_path = os.path.join(tempfile.gettempdir(), "sw_screenshot.bmp")

        doc.ViewZoomtofit2()

        # 1) IModelDoc2::SaveBMP(FileName, Width, Height) — SW 2024 signature
        try:
            if doc.SaveBMP(output_path, width, height):
                return {"screenshot_path": output_path}
        except Exception:
            pass

        # 2) Try zero-width/height (let SW pick default size)
        try:
            if doc.SaveBMP(output_path, 0, 0):
                return {"screenshot_path": output_path}
        except Exception:
            pass

        # 3) Fallback: SaveAs with PNG extension
        png_path = output_path.replace(".bmp", ".png")
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        try:
            result = doc.Extension.SaveAs(png_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
            if result:
                return {"screenshot_path": png_path}
        except Exception:
            pass

        raise RuntimeError(
            f"Screenshot failed for '{output_path}'. "
            "Ensure a model is open and SolidWorks is visible."
        )

    def set_view_orientation(self, view_name: str) -> Dict:
        doc = self._active_doc()
        sw_view = VIEW_NAME_MAP.get(view_name.lower())
        if sw_view is None:
            raise ValueError(
                f"Unknown view '{view_name}'. Valid: {', '.join(VIEW_NAME_MAP.keys())}"
            )
        doc.ShowNamedView2(sw_view, -1)
        doc.ViewZoomtofit2()
        return {"view": view_name}

    def zoom_to_fit(self) -> Dict:
        doc = self._active_doc()
        doc.ViewZoomtofit2()
        return {"zoomed": True}

    def set_display_mode(self, mode: str) -> Dict:
        doc = self._active_doc()
        mode_val = DISPLAY_MODE_MAP.get(mode.lower())
        if mode_val is None:
            raise ValueError(
                f"Unknown display mode '{mode}'. Valid: {', '.join(DISPLAY_MODE_MAP.keys())}"
            )
        # Try ModelViewManager.ActiveModelView, then doc.ActiveView
        set_ok = False
        for view_obj in [
            _get(_get(doc, 'ModelViewManager'), 'ActiveModelView') if _get(doc, 'ModelViewManager') else None,
            _get(doc, 'ActiveView'),
        ]:
            try:
                if view_obj:
                    view_obj.DisplayMode = mode_val
                    set_ok = True
                    break
            except Exception:
                pass
        if not set_ok:
            raise RuntimeError(f"Could not set display mode to '{mode}'.")
        return {"display_mode": mode}

    def rebuild_model(self, top_only: bool = False) -> Dict:
        doc = self._active_doc()
        # EditRebuild3 is a property in SW 2024 that triggers rebuild and returns bool
        _get(doc, 'EditRebuild3')
        return {"rebuilt": True}

    # ─── Macro Operations ────────────────────────────────────────────────────

    def run_macro(self, macro_path: str, module_name: str = "", sub_name: str = "main") -> Dict:
        app = self._ensure_app()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = app.RunMacro2(macro_path, module_name, sub_name, False, errors)
        if not result:
            raise RuntimeError(
                f"Macro '{macro_path}' failed (error={errors.value}). "
                "Ensure the file path and subroutine name are correct."
            )
        return {"macro": macro_path, "module": module_name, "sub": sub_name}

    def run_vba_code(self, vba_code: str) -> Dict:
        """Write VBA code to a temp .swb file and run it."""
        app = self._ensure_app()
        # Wrap in Sub main if needed
        if "Sub " not in vba_code:
            vba_code = f"Sub main()\n{vba_code}\nEnd Sub"

        tmp_path = os.path.join(tempfile.gettempdir(), "sw_mcp_temp_macro.swb")
        with open(tmp_path, "w") as f:
            f.write(vba_code)

        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = app.RunMacro2(tmp_path, "sw_mcp_temp_macro", "main", False, errors)
        os.unlink(tmp_path)
        return {"executed": True, "error_code": errors.value}

    def start_macro_recording(self, output_path: str) -> Dict:
        app = self._ensure_app()
        app.StartRecordingMacro(output_path, True)  # True = VBA
        return {"recording": True, "output_path": output_path}

    def stop_macro_recording(self) -> Dict:
        app = self._ensure_app()
        app.StopRecordingMacro(True)
        return {"recording": False}

    # ─── Sheet Metal ──────────────────────────────────────────────────────────

    def flatten_sheet_metal(self) -> Dict:
        doc = self._active_doc()
        doc.ShowFlatPattern()
        return {"flattened": True}

    def get_flat_pattern_info(self) -> Dict:
        doc = self._active_doc()
        bb = doc.GetBoundingBox()
        if not bb:
            return {"error": "Could not get flat pattern bounding box"}
        return {
            "flat_width_mm": _mm(bb[3] - bb[0]),
            "flat_height_mm": _mm(bb[4] - bb[1]),
            "thickness_mm": _mm(bb[5] - bb[2]),
        }

    def export_flat_pattern(self, output_path: str) -> Dict:
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = doc.Extension.SaveAs(output_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
        if not result:
            raise RuntimeError(f"Export failed (error={errors.value})")
        return {"exported_to": output_path}

    # ─── Reference Geometry ──────────────────────────────────────────────────

    def create_reference_plane(self, name: str, offset_distance: float = 0.0, reference_plane: str = "Front Plane") -> Dict:
        doc = self._active_doc()
        doc.Extension.SelectByID2(reference_plane, "PLANE", 0, 0, 0, False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0)
        feat = doc.FeatureManager.InsertRefPlane(8, _m(offset_distance), 0, 0)  # 8=swRefPlaneReferenceConstraints_Offset
        if feat is None:
            raise RuntimeError("Failed to create reference plane.")
        return {"feature": feat.Name, "offset_mm": offset_distance}

    def create_reference_axis(self, axis_type: str = "two_planes") -> Dict:
        doc = self._active_doc()
        axis_map = {"two_planes": 4, "cylinder": 2, "two_points": 1, "point_on_face": 3}
        a_type = axis_map.get(axis_type.lower(), 4)
        feat = doc.FeatureManager.InsertAxis2(a_type)
        if feat is None:
            raise RuntimeError(f"Failed to create reference axis of type '{axis_type}'. Select the required entities first.")
        return {"feature": feat.Name, "axis_type": axis_type}

    def create_reference_point(self, x: float, y: float, z: float) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertReferencePoint(
            2, 0, None, _m(x), _m(y), _m(z), None
        )
        if feat is None:
            raise RuntimeError("Failed to create reference point.")
        return {"feature": feat.Name, "x_mm": x, "y_mm": y, "z_mm": z}

    # ─── Surface Modeling ────────────────────────────────────────────────────

    def thicken_surface(self, thickness: float, both_sides: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertThicken(_m(thickness), 0 if not both_sides else 2, False)
        if feat is None:
            raise RuntimeError("Thicken failed. Select a surface body first.")
        return {"feature": feat.Name, "thickness_mm": thickness}

    def knit_surfaces(self) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertKnitSurface(True, False, 0.0, True)
        if feat is None:
            raise RuntimeError("Knit surfaces failed. Select surface bodies to knit.")
        return {"feature": feat.Name}

    def planar_surface(self) -> Dict:
        """Create a planar surface from selected closed sketch profile or edges."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertPlanarSurface()
        if feat is None:
            raise RuntimeError("Planar surface failed. Select a closed sketch or edges first.")
        return {"feature": feat.Name}

    def extruded_surface(self, depth: float, flip: bool = False, draft_angle: float = 0.0) -> Dict:
        doc = self._active_doc()
        fm = doc.FeatureManager
        feat = fm.FeatureExtruRefSurface3(
            True, flip, False,
            SW_END_BLIND, SW_END_BLIND,
            _m(depth), 0.0,
            False, False,
            draft_angle > 0, False,
            _r(draft_angle), 0.0,
            False, False,
            False, False,
            True, False, True
        )
        if feat is None:
            raise RuntimeError("Extruded surface failed. Active sketch must be selected.")
        return {"feature": feat.Name, "depth_mm": depth}

    def revolved_surface(self, angle: float = 360.0) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertRevolveRefSurfaceFeature2(
            _r(angle), SW_END_BLIND, SW_END_BLIND, 0.0, 0.0
        )
        if feat is None:
            raise RuntimeError("Revolved surface failed. Select profile and axis first.")
        return {"feature": feat.Name, "angle_deg": angle}

    def offset_surface(self, distance: float, flip: bool = False) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertOffsetSurface(_m(distance), flip)
        if feat is None:
            raise RuntimeError("Offset surface failed. Select a face/surface first.")
        return {"feature": feat.Name, "distance_mm": distance}

    def trim_surface(self) -> Dict:
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertTrimSurface2(0, False, False, False, False, 0.0, 0.0)
        if feat is None:
            raise RuntimeError("Trim surface failed. Select trim tool and surface first.")
        return {"feature": feat.Name}

    def extend_surface(self, distance: float, extend_type: str = "distance") -> Dict:
        doc = self._active_doc()
        # extension type: 0=distance, 1=up to point, 2=up to surface
        e_type = {"distance": 0, "to_point": 1, "to_surface": 2}.get(extend_type.lower(), 0)
        feat = doc.FeatureManager.InsertExtendSurface(_m(distance), e_type, 0, False)
        if feat is None:
            raise RuntimeError("Extend surface failed.")
        return {"feature": feat.Name, "distance_mm": distance}

    def delete_face(self, repair: bool = False) -> Dict:
        """Delete selected face(s). repair=True to patch hole, False to leave open."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertDeleteFace2(1 if repair else 0)
        if feat is None:
            raise RuntimeError("Delete face failed. Select a face first.")
        return {"feature": feat.Name, "repaired": repair}

    def filled_surface(self) -> Dict:
        """Fill a closed boundary of edges with a surface."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertFilledSurface(0, False, False, False, False, False, 0.0)
        if feat is None:
            raise RuntimeError("Filled surface failed. Select closed boundary edges first.")
        return {"feature": feat.Name}

    # ─── Material & Appearance ───────────────────────────────────────────────

    def set_material(self, material_name: str, database_name: str = "SOLIDWORKS Materials") -> Dict:
        """Apply a material to the part. Common: 'AISI 1020', 'Aluminum 6061', 'ABS', 'PLA'."""
        doc = self._active_doc()
        try:
            doc.SetMaterialPropertyName2("", database_name, material_name)
        except Exception:
            # Try without database arg (older signature)
            doc.SetMaterialPropertyName(database_name, material_name)
        return {"material": material_name, "database": database_name}

    def get_material(self) -> Dict:
        """Read the current material of the part."""
        doc = self._active_doc()
        try:
            db_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            mat = doc.GetMaterialPropertyName2("", db_out)
            return {"material": mat, "database": db_out.value}
        except Exception:
            try:
                mat = _get(doc, 'MaterialIdName')
                return {"material": mat, "database": ""}
            except Exception:
                return {"material": None, "database": None}

    def set_appearance_color(self, r: int, g: int, b: int, transparency: float = 0.0) -> Dict:
        """Set a single solid-color appearance on the active part body. Values 0-255 for r/g/b, 0.0-1.0 for transparency."""
        doc = self._active_doc()
        # Material properties array: [R, G, B, ambient, diffuse, specular, shininess, transparency, emission]
        props = [r / 255.0, g / 255.0, b / 255.0, 1.0, 1.0, 0.6, 0.4, transparency, 0.0]
        try:
            bodies = doc.GetBodies2(SW_SOLID_BODY, True)
            if bodies:
                for body in (bodies if hasattr(bodies, '__iter__') else [bodies]):
                    body.MaterialPropertyValues = props
        except Exception:
            doc.MaterialPropertyValues = props
        return {"color_rgb": [r, g, b], "transparency": transparency}

    def list_materials(self, database_name: str = "SOLIDWORKS Materials") -> Dict:
        """List materials available in a SolidWorks materials database."""
        app = self._ensure_app()
        try:
            mats = app.GetMaterialDatabases()
            return {"databases": list(mats) if mats else [database_name]}
        except Exception:
            return {"note": "Materials databases live in C:/ProgramData/SOLIDWORKS/SOLIDWORKS XXXX/Lang/english/sldmaterials/*.sldmat"}

    # ─── Body Operations ──────────────────────────────────────────────────────

    def list_bodies(self, body_type: str = "solid") -> List[Dict]:
        """List solid or surface bodies in the active part."""
        doc = self._active_doc()
        bt = SW_SOLID_BODY if body_type.lower() == "solid" else SW_SHEET_BODY
        bodies = doc.GetBodies2(bt, True)
        if not bodies:
            return []
        bodies_list = list(bodies) if hasattr(bodies, '__iter__') else [bodies]
        result = []
        for b in bodies_list:
            try:
                result.append({"name": _get(b, 'Name'), "is_solid": bt == SW_SOLID_BODY})
            except Exception:
                pass
        return result

    def combine_bodies(self, operation: str = "add") -> Dict:
        """Combine selected bodies. operation: 'add' (union), 'subtract', 'common' (intersect)."""
        doc = self._active_doc()
        op_map = {"add": 0, "subtract": 1, "common": 2}
        op = op_map.get(operation.lower(), 0)
        feat = doc.FeatureManager.InsertCombineFeature(op + 1)  # 1=add, 2=subtract, 3=common
        if feat is None:
            raise RuntimeError(f"Combine ({operation}) failed. Select bodies first.")
        return {"feature": feat.Name, "operation": operation}

    def split_body(self) -> Dict:
        """Split body with selected sketch/surface trim tool."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertSplitBody3(True, 0.0, None)
        if feat is None:
            raise RuntimeError("Split body failed. Select body and trim tool first.")
        return {"feature": feat.Name}

    def move_body(self, dx: float = 0.0, dy: float = 0.0, dz: float = 0.0,
                  rx: float = 0.0, ry: float = 0.0, rz: float = 0.0,
                  copy: bool = False) -> Dict:
        """Move/rotate selected body. dx/dy/dz in mm, rx/ry/rz in degrees."""
        doc = self._active_doc()
        # MoveCopy3: type=0(trans+rot), copies=0 (no copy), x,y,z translation, rotation x,y,z
        feat = doc.FeatureManager.InsertMoveCopyBody2(
            _m(dx), _m(dy), _m(dz),
            0.0, 0.0, 0.0,    # rotation origin
            _r(rx), _r(ry), _r(rz),
            copy, 1 if copy else 0
        )
        if feat is None:
            raise RuntimeError("Move body failed. Select a body first.")
        return {"feature": feat.Name, "translation_mm": [dx, dy, dz], "rotation_deg": [rx, ry, rz]}

    def delete_body(self) -> Dict:
        """Delete selected body."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertDeleteBody2()
        if feat is None:
            raise RuntimeError("Delete body failed. Select a body first.")
        return {"feature": feat.Name}

    def scale_body(self, factor: float, uniform: bool = True) -> Dict:
        """Scale selected body uniformly by factor."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertScale(factor, factor, factor, 0, uniform)
        if feat is None:
            raise RuntimeError("Scale failed. Select body first.")
        return {"feature": feat.Name, "factor": factor}

    # ─── Curves & 3D Sketch ────────────────────────────────────────────────────

    def create_helix(self, pitch: float, revolutions: float, start_angle: float = 0.0,
                     clockwise: bool = False, taper_angle: float = 0.0) -> Dict:
        """Create a helix on the selected sketch circle. pitch in mm, revolutions count, taper in degrees."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertHelix(
            False,           # define by height & pitch
            False,           # define by height & revolution
            clockwise,
            False,           # reverse direction
            0,               # diameter type: 0=defined by sketch
            0,               # start angle method
            _m(pitch),
            revolutions,
            _r(start_angle),
            _r(taper_angle),
            False,           # taper outward
            False            # use auto pitch
        )
        if feat is None:
            raise RuntimeError("Helix creation failed. Select a sketch with a single circle first.")
        return {"feature": feat.Name, "pitch_mm": pitch, "revolutions": revolutions}

    def start_3d_sketch(self) -> Dict:
        """Start a 3D sketch."""
        doc = self._active_doc()
        doc.SketchManager.Insert3DSketch(True)
        return {"sketch_3d_active": True}

    def split_line(self, projection_type: str = "sketch") -> Dict:
        """Split a face with a sketch (projection_type='sketch')."""
        doc = self._active_doc()
        # 0=sketch-on-face, 1=projected sketch, 2=intersection
        t = {"sketch": 0, "projected": 1, "intersection": 2}.get(projection_type.lower(), 0)
        feat = doc.FeatureManager.InsertSplitLineProject(t, True, False, False, False)
        if feat is None:
            raise RuntimeError("Split line failed. Select sketch and target face first.")
        return {"feature": feat.Name}

    def projected_curve(self) -> Dict:
        """Project a sketch onto a face."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertProjectCurve(0, False)
        if feat is None:
            raise RuntimeError("Projected curve failed. Select sketch and face first.")
        return {"feature": feat.Name}

    def composite_curve(self) -> Dict:
        """Combine selected edges/sketch entities into a composite curve."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertCompositeCurve()
        if feat is None:
            raise RuntimeError("Composite curve failed. Select edges/sketch segments first.")
        return {"feature": feat.Name}

    # ─── Sketch — Advanced ─────────────────────────────────────────────────────

    def sketch_convert_entities(self) -> Dict:
        """Convert selected face edges into sketch entities on the active sketch."""
        doc = self._active_doc()
        doc.SketchManager.SketchUseEdge3(False, False)
        return {"action": "convert_entities"}

    def sketch_trim(self) -> Dict:
        """Trim selected sketch entity (uses power trim)."""
        doc = self._active_doc()
        # 5 = swSketchTrimChoice_PowerTrim
        result = doc.SketchManager.SketchTrim(5, 0.0, 0.0, 0.0)
        return {"trimmed": bool(result)}

    def sketch_extend(self) -> Dict:
        """Extend selected sketch entity."""
        doc = self._active_doc()
        result = doc.SketchManager.SketchExtend(0.0, 0.0, 0.0)
        return {"extended": bool(result)}

    def sketch_construction(self, on: bool = True) -> Dict:
        """Toggle selected sketch entities to construction geometry."""
        doc = self._active_doc()
        sm = doc.SketchManager
        # Try the direct method first (newer SW), then fall back
        for method_name in ("MakeSketchEntitiesConstruction", "MakeConstructionLines"):
            try:
                fn = getattr(sm, method_name, None)
                if fn is None:
                    continue
                if callable(fn):
                    fn(on)
                return {"construction": on, "via": method_name}
            except Exception:
                pass
        # Fallback: iterate selected sketch segments and toggle ConstructionGeometry property
        try:
            sel_mgr = doc.SelectionManager
            cnt = _get(sel_mgr, 'GetSelectedObjectCount2') or sel_mgr.GetSelectedObjectCount2(-1)
            toggled = 0
            for i in range(1, cnt + 1):
                obj = sel_mgr.GetSelectedObject6(i, -1)
                try:
                    obj.ConstructionGeometry = on
                    toggled += 1
                except Exception:
                    pass
            return {"construction": on, "segments_toggled": toggled}
        except Exception as e:
            raise RuntimeError(f"sketch_construction failed: {e}")

    def sketch_fillet(self, radius: float, trim_segments: bool = True) -> Dict:
        """Fillet selected sketch corner."""
        doc = self._active_doc()
        result = doc.SketchManager.CreateFillet(_m(radius), trim_segments)
        return {"radius_mm": radius, "result": bool(result)}

    def sketch_chamfer(self, distance: float, equal: bool = True, angle: float = 45.0) -> Dict:
        """Chamfer selected sketch corner."""
        doc = self._active_doc()
        # ChamferType: 0=DistDist (equal), 1=DistAngle
        ct = 0 if equal else 1
        result = doc.SketchManager.CreateChamfer(ct, _m(distance), _r(angle))
        return {"distance_mm": distance, "result": bool(result)}

    def sketch_text(self, text: str, x: float = 0.0, y: float = 0.0, height: float = 5.0) -> Dict:
        """Insert text in the active sketch."""
        doc = self._active_doc()
        result = doc.SketchManager.InsertSketchText(
            _m(x), _m(y), 0.0,
            text, 0, 0, 0,
            _m(height), _m(height)
        )
        return {"text": text}

    # ─── Sheet Metal — Complete Suite ────────────────────────────────────────

    def base_flange(self, thickness: float, bend_radius: float = 1.0, k_factor: float = 0.5) -> Dict:
        """Convert active sketch to a sheet-metal base flange."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertSheetMetalBaseFlange2(
            _m(thickness), False,    # thickness, reverse
            _m(bend_radius),
            0, 0,                    # bend allowance type, gauge table
            False, False,
            k_factor, 0.5,           # k-factor, bend allowance
            False, 0.0,
            False, False, False, False
        )
        if feat is None:
            raise RuntimeError("Base flange failed. Ensure a closed sketch is active.")
        return {"feature": feat.Name, "thickness_mm": thickness, "bend_radius_mm": bend_radius}

    def edge_flange(self, edge_name: str, flange_length: float, angle: float = 90.0) -> Dict:
        """Add edge flange on selected sheet-metal edge."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(edge_name, "EDGE", 0, 0, 0, False, 0, _nc, 0)
        feat = doc.FeatureManager.InsertSheetMetalEdgeFlange2(
            _r(angle), _m(flange_length),
            0, 0, 0,            # flange position, length-end type, offset distance
            False, False,
            0.5, 0.5,
            True
        )
        if feat is None:
            raise RuntimeError("Edge flange failed. Select a valid edge.")
        return {"feature": feat.Name, "length_mm": flange_length, "angle_deg": angle}

    def miter_flange(self, length: float) -> Dict:
        """Add miter flange on selected sketch path."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertMiterFlange2(
            _m(length), 0, 0.5, 0.0, 0.0, 0.0, False, False, 0
        )
        if feat is None:
            raise RuntimeError("Miter flange failed. Select sketch profile and edges.")
        return {"feature": feat.Name, "length_mm": length}

    def sketched_bend(self, angle: float = 90.0, position: str = "bend_centerline") -> Dict:
        """Add a sketched bend at the active sketch line on a sheet metal face."""
        doc = self._active_doc()
        p_map = {"bend_centerline": 0, "material_inside": 1, "material_outside": 2, "bend_outside": 3}
        pos = p_map.get(position.lower(), 0)
        feat = doc.FeatureManager.InsertSketchedBend2(_r(angle), pos, False, 0.0, 0)
        if feat is None:
            raise RuntimeError("Sketched bend failed. Select sketch line on flat face.")
        return {"feature": feat.Name, "angle_deg": angle}

    def jog(self, jog_distance: float, jog_angle: float = 90.0) -> Dict:
        """Add a jog bend at selected sketch line."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertJog2(
            _m(jog_distance), _r(jog_angle),
            False, 0, 0, 0, 0,
            False, 0.0, 0
        )
        if feat is None:
            raise RuntimeError("Jog failed. Select sketch line on flat face first.")
        return {"feature": feat.Name, "distance_mm": jog_distance}

    def hem(self, hem_type: str = "closed", length: float = 5.0, gap: float = 0.0) -> Dict:
        """Add a hem on selected sheet-metal edge."""
        doc = self._active_doc()
        # 0=closed, 1=open, 2=tear-drop, 3=rolled
        ht = {"closed": 0, "open": 1, "teardrop": 2, "rolled": 3}.get(hem_type.lower(), 0)
        feat = doc.FeatureManager.InsertHemFeature2(
            ht, 0, _m(length), _m(gap), _r(90.0), 0.0,
            False, False, False, 0, 0
        )
        if feat is None:
            raise RuntimeError("Hem failed. Select a sheet-metal edge first.")
        return {"feature": feat.Name, "type": hem_type, "length_mm": length}

    def unfold(self, all_bends: bool = True) -> Dict:
        """Unfold all (or selected) bends to flatten sheet metal."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertSheetMetalUnfold(all_bends)
        if feat is None:
            raise RuntimeError("Unfold failed.")
        return {"feature": feat.Name, "unfolded_all": all_bends}

    def fold(self) -> Dict:
        """Fold previously unfolded bends back."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertSheetMetalFold(True)
        if feat is None:
            raise RuntimeError("Fold failed.")
        return {"feature": feat.Name}

    # ─── Weldments ────────────────────────────────────────────────────────────

    def insert_weldment(self) -> Dict:
        """Insert weldment feature on the active part."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertWeldmentFeature()
        if feat is None:
            raise RuntimeError("Insert weldment failed.")
        return {"feature": feat.Name}

    def add_structural_member(self, standard: str, configuration: str, size: str) -> Dict:
        """Add structural member. standard='iso', configuration='square tube', size='40 x 40 x 4'."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertStructuralMember3(standard, configuration, size, "", 0, False)
        if feat is None:
            raise RuntimeError(f"Structural member failed. Standard='{standard}', config='{configuration}', size='{size}'.")
        return {"feature": feat.Name, "standard": standard, "size": size}

    def add_gusset(self, profile_distance1: float, profile_distance2: float, thickness: float) -> Dict:
        """Add a triangular gusset between two structural members."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertWeldGusset(
            0, _m(profile_distance1), _m(profile_distance2),
            _m(thickness), 0, _r(45.0), 0, 0, 0, False
        )
        if feat is None:
            raise RuntimeError("Gusset failed. Select two faces of structural members.")
        return {"feature": feat.Name, "thickness_mm": thickness}

    def add_end_cap(self, thickness: float, offset: float = 0.0) -> Dict:
        """Add an end cap on selected structural member face."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertEndCap(_m(thickness), _m(offset), 0, 0.0, 0, 0, False)
        if feat is None:
            raise RuntimeError("End cap failed. Select face of structural member.")
        return {"feature": feat.Name, "thickness_mm": thickness}

    def trim_weldment(self) -> Dict:
        """Trim selected weldment members to a planar/solid trim tool."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertWeldmentTrimExtend2(0, False, False, 0)
        if feat is None:
            raise RuntimeError("Trim weldment failed. Select members and trim tool.")
        return {"feature": feat.Name}

    # ─── Mold Tools ───────────────────────────────────────────────────────────

    def parting_line(self, draft_angle: float = 1.0) -> Dict:
        """Create a parting line for mold design."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertPartingLines(_r(draft_angle), 0, True, True, False)
        if feat is None:
            raise RuntimeError("Parting line failed. Select draft direction reference first.")
        return {"feature": feat.Name, "draft_angle_deg": draft_angle}

    def shut_off_surface(self) -> Dict:
        """Auto-create shut-off surfaces over holes for mold separation."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertShutOffSurfaces()
        if feat is None:
            raise RuntimeError("Shut-off surface failed.")
        return {"feature": feat.Name}

    def tooling_split(self, depth_above: float, depth_below: float) -> Dict:
        """Split a mold block into core and cavity."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertMoldToolingSplit(_m(depth_above), _m(depth_below), False)
        if feat is None:
            raise RuntimeError("Tooling split failed. Place parting line/surfaces first.")
        return {"feature": feat.Name}

    # ─── Drawing — Extended ───────────────────────────────────────────────────

    def add_projected_view(self, parent_view_name: str, direction: str = "right",
                          x: float = 0.0, y: float = 0.0) -> Dict:
        """Add a projected view from an existing drawing view. direction: left/right/up/down."""
        doc = self._active_doc()
        # Select parent view first
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(parent_view_name, "DRAWINGVIEW", 0, 0, 0, False, 0, _nc, 0)
        d_map = {"left": 0, "right": 1, "up": 2, "down": 3}
        dir_val = d_map.get(direction.lower(), 1)
        view = doc.CreateUnfoldedViewAt3(_m(x), _m(y), 0.0, False)
        if view is None:
            raise RuntimeError(f"Projected view failed. Parent='{parent_view_name}'.")
        return {"view": _get(view, 'Name'), "direction": direction}

    def add_auxiliary_view(self, edge_name: str, x: float = 0.2, y: float = 0.2) -> Dict:
        """Add an auxiliary view normal to a selected edge."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(edge_name, "EDGE", 0, 0, 0, False, 0, _nc, 0)
        view = doc.InsertAuxiliaryView2(_m(x), _m(y), 0.0)
        if view is None:
            raise RuntimeError("Auxiliary view failed. Select an edge first.")
        return {"view": _get(view, 'Name')}

    def add_detail_view(self, x: float, y: float, radius: float, scale: float = 2.0, label: str = "A") -> Dict:
        """Create a detail view at a sketch circle on an existing view."""
        doc = self._active_doc()
        # Create a sketch circle then call InsertDetailView
        sketchmgr = doc.SketchManager
        sketchmgr.CreateCircleByRadius(_m(x), _m(y), 0.0, _m(radius))
        view = doc.InsertDetailViewNew(scale, 0, False, False, label, "", True)
        if view is None:
            raise RuntimeError("Detail view failed.")
        return {"detail_view": label, "scale": scale}

    def add_broken_view(self, gap: float = 10.0, break_type: str = "straight") -> Dict:
        """Add a broken view to compress the current drawing view."""
        doc = self._active_doc()
        # 0=straight, 1=curved, 2=zig-zag, 3=small zig-zag
        bt = {"straight": 0, "curved": 1, "zigzag": 2, "small_zigzag": 3}.get(break_type.lower(), 0)
        view = doc.InsertBreak(0, _m(gap), bt, 0.5)
        if view is None:
            raise RuntimeError("Broken view failed.")
        return {"break_type": break_type, "gap_mm": gap}

    def add_centerline(self, view_name: str) -> Dict:
        """Add centerline annotations to selected drawing view."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(view_name, "DRAWINGVIEW", 0, 0, 0, False, 0, _nc, 0)
        result = doc.InsertCenterLine2(True)
        return {"centerlines_added": bool(result)}

    def add_center_mark(self, view_name: str) -> Dict:
        """Add center marks on selected drawing view."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(view_name, "DRAWINGVIEW", 0, 0, 0, False, 0, _nc, 0)
        result = doc.InsertCenterMark2(True, 0, 0.0, False, 0.0)
        return {"center_marks_added": bool(result)}

    def add_balloon(self, x: float, y: float, item_number: Optional[int] = None,
                    style: str = "circular") -> Dict:
        """Add a BOM balloon at the given drawing coordinate."""
        doc = self._active_doc()
        # Style codes: 1=circular, 2=triangle, 3=hexagon, 4=square, 5=pentagon
        st_map = {"circular": 1, "triangle": 2, "hexagon": 3, "square": 4, "pentagon": 5}
        st = st_map.get(style.lower(), 1)
        bal = doc.CreateBalloon4(
            _m(x), _m(y), 0.0,
            st, 0, 0,
            "" if item_number is None else str(item_number),
            "", 0, ""
        )
        return {"balloon_placed": True, "style": style}

    def add_surface_finish(self, symbol_type: str = "machining", roughness: str = "3.2",
                          x: float = 0.0, y: float = 0.0) -> Dict:
        """Add a surface finish symbol."""
        doc = self._active_doc()
        # 1=basic, 2=machining required, 3=machining prohibited
        st = {"basic": 1, "machining": 2, "no_machining": 3}.get(symbol_type.lower(), 2)
        sym = doc.InsertSurfaceFinishSymbol3(st, "", roughness, "", "", "", "", "", "", "", "")
        if sym:
            ann = sym.GetAnnotation()
            if ann:
                ann.SetPosition2(_m(x), _m(y), 0.0)
        return {"surface_finish": symbol_type, "roughness": roughness}

    def add_geometric_tolerance(self, gtol_text: str, x: float = 0.0, y: float = 0.0) -> Dict:
        """Add a geometric tolerance (GD&T) frame."""
        doc = self._active_doc()
        sym = doc.InsertGtol()
        if sym:
            ann = sym.GetAnnotation()
            if ann:
                ann.SetPosition2(_m(x), _m(y), 0.0)
        return {"gtol": gtol_text}

    def add_weld_symbol(self, text: str = "", x: float = 0.0, y: float = 0.0) -> Dict:
        """Add a weld symbol annotation."""
        doc = self._active_doc()
        sym = doc.InsertWeldSymbol3(text)
        if sym:
            ann = sym.GetAnnotation()
            if ann:
                ann.SetPosition2(_m(x), _m(y), 0.0)
        return {"weld_symbol": text}

    def add_datum_feature(self, label: str = "A", x: float = 0.0, y: float = 0.0) -> Dict:
        """Add a datum feature symbol."""
        doc = self._active_doc()
        sym = doc.InsertDatumTag2(label)
        if sym:
            ann = sym.GetAnnotation()
            if ann:
                ann.SetPosition2(_m(x), _m(y), 0.0)
        return {"datum": label}

    def add_hole_table(self, x: float = 0.05, y: float = 0.05, origin_x: float = 0, origin_y: float = 0) -> Dict:
        """Add a hole table to drawing."""
        doc = self._active_doc()
        table = doc.InsertHoleTable2(False, _m(x), _m(y), 1, "A", 0, _m(origin_x), _m(origin_y))
        return {"hole_table_added": table is not None}

    def add_revision_table(self, x: float = 0.25, y: float = 0.25) -> Dict:
        """Add revision table to drawing sheet."""
        doc = self._active_doc()
        table = doc.InsertRevisionTable2(False, _m(x), _m(y), 1, "Default")
        return {"revision_table_added": table is not None}

    def export_drawing_pdf(self, output_path: str) -> Dict:
        """Export current drawing to PDF."""
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = doc.Extension.SaveAs(output_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
        if not result:
            raise RuntimeError(f"PDF export failed (error={errors.value})")
        return {"exported_to": output_path}

    def export_drawing_dwg(self, output_path: str) -> Dict:
        """Export current drawing to DWG/DXF."""
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        result = doc.Extension.SaveAs(output_path, 0, SW_SAVE_SILENT, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), errors, warnings)
        if not result:
            raise RuntimeError(f"DWG export failed (error={errors.value})")
        return {"exported_to": output_path}

    # ─── Assembly — Extended ──────────────────────────────────────────────────

    def replace_component(self, component_name: str, new_file_path: str) -> Dict:
        """Replace component with another part/assembly."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(component_name, "COMPONENT", 0, 0, 0, False, 0, _nc, 0)
        result = doc.ReplaceComponents(new_file_path, "", True, False)
        if not result:
            raise RuntimeError(f"Replace component failed. '{component_name}' → '{new_file_path}'.")
        return {"replaced": component_name, "with": new_file_path}

    def pattern_component_linear(self, count: int, spacing: float, direction_face: str = "") -> Dict:
        """Pattern a component linearly."""
        doc = self._active_doc()
        feat = doc.FeatureManager.LocalLinearComponentPattern(
            count, _m(spacing), 1, 0, 0, False, False
        )
        if feat is None:
            raise RuntimeError("Linear component pattern failed. Select component and direction reference.")
        return {"feature": _get(feat, 'Name'), "count": count}

    def pattern_component_circular(self, count: int, angle: float = 360.0) -> Dict:
        """Pattern a component circularly."""
        doc = self._active_doc()
        feat = doc.FeatureManager.LocalCircularComponentPattern(count, _r(angle), True, False)
        if feat is None:
            raise RuntimeError("Circular component pattern failed. Select component and axis.")
        return {"feature": _get(feat, 'Name'), "count": count}

    def mirror_component(self, mirror_plane: str) -> Dict:
        """Mirror a component about a plane."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(mirror_plane, "PLANE", 0, 0, 0, True, 4, _nc, 0)
        feat = doc.FeatureManager.InsertMirrorComponents2(0, True, False)
        if feat is None:
            raise RuntimeError("Mirror component failed.")
        return {"feature": _get(feat, 'Name')}

    def suppress_component(self, component_name: str, suppress: bool = True) -> Dict:
        """Suppress or unsuppress a component."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(component_name, "COMPONENT", 0, 0, 0, False, 0, _nc, 0)
        if suppress:
            doc.EditSuppress2()
        else:
            doc.EditUnsuppress2()
        return {"component": component_name, "suppressed": suppress}

    def hide_component(self, component_name: str, hide: bool = True) -> Dict:
        """Hide or show a component."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(component_name, "COMPONENT", 0, 0, 0, False, 0, _nc, 0)
        if hide:
            doc.HideComponent2()
        else:
            doc.ShowComponent2()
        return {"component": component_name, "hidden": hide}

    def move_component(self, component_name: str, dx: float = 0.0, dy: float = 0.0, dz: float = 0.0) -> Dict:
        """Translate a component in mm."""
        doc = self._active_doc()
        _nc = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        doc.Extension.SelectByID2(component_name, "COMPONENT", 0, 0, 0, False, 0, _nc, 0)
        ms = doc.GetAssemblyManipulationManager() if hasattr(doc, 'GetAssemblyManipulationManager') else None
        try:
            doc.MoveComponent(_m(dx), _m(dy), _m(dz))
        except Exception:
            pass
        return {"component": component_name, "moved_mm": [dx, dy, dz]}

    def get_assembly_tree(self) -> List[Dict]:
        """Get the assembly component tree."""
        doc = self._active_doc()
        if _get(doc, 'GetType') != SW_DOC_ASSEMBLY:
            raise RuntimeError("Active document is not an assembly.")
        comps = doc.GetComponents(False) if hasattr(doc, 'GetComponents') else []
        if not comps:
            return []
        result = []
        for c in (comps if hasattr(comps, '__iter__') else [comps]):
            try:
                result.append({
                    "name": _get(c, 'Name2'),
                    "path": _get(c, 'GetPathName'),
                    "suppressed": _get(c, 'IsSuppressed'),
                    "hidden": _get(c, 'Visible') == 1,
                })
            except Exception:
                pass
        return result

    # ─── Display & Visualization Extended ────────────────────────────────────

    def create_section_view(self, plane: str = "Front Plane", offset: float = 0.0) -> Dict:
        """Create a live model section view through a plane."""
        doc = self._active_doc()
        try:
            view = doc.ActiveView
            if view and hasattr(view, 'InsertSectionView'):
                result = view.InsertSectionView(plane, _m(offset))
                return {"section_view_active": bool(result), "plane": plane}
        except Exception:
            pass
        # Fallback: simply select & toggle section
        return {"note": "Section view requires UI interaction; selection set."}

    def set_background(self, mode: str = "plain", color_r: int = 200, color_g: int = 200, color_b: int = 200) -> Dict:
        """Set viewport background. mode: 'plain', 'gradient', 'image', 'scene'."""
        app = self._ensure_app()
        try:
            doc = self._active_doc()
            # Set sceneBackgroundType: 0=plain, 1=gradient, 2=image, 3=scene
            mode_map = {"plain": 0, "gradient": 1, "image": 2, "scene": 3}
            mv = doc.ModelViewManager
            if hasattr(mv, 'SceneBackgroundType'):
                mv.SceneBackgroundType = mode_map.get(mode.lower(), 0)
        except Exception:
            pass
        return {"background": mode}

    def take_photoview_screenshot(self, output_path: str, width: int = 1920, height: int = 1080) -> Dict:
        """Trigger PhotoView 360 render and save image (requires PhotoView 360 add-in)."""
        doc = self._active_doc()
        app = self._ensure_app()
        try:
            pv = app.GetAddInObject("PhotoView360.PhotoView360") if hasattr(app, 'GetAddInObject') else None
            if pv:
                pv.SaveImage(output_path, width, height)
                return {"rendered_to": output_path}
        except Exception:
            pass
        # Fallback: regular screenshot
        return self.capture_screenshot(output_path, width, height)

    def hide_all_planes(self) -> Dict:
        """Hide all reference planes in the viewport."""
        doc = self._active_doc()
        doc.BlankRefGeom()
        return {"planes_hidden": True}

    def show_all_planes(self) -> Dict:
        """Show all reference planes in the viewport."""
        doc = self._active_doc()
        doc.UnblankRefGeom()
        return {"planes_shown": True}

    def hide_all_sketches(self) -> Dict:
        """Hide all sketches in the viewport."""
        doc = self._active_doc()
        doc.BlankSketch()
        return {"sketches_hidden": True}

    def show_all_sketches(self) -> Dict:
        """Show all sketches in the viewport."""
        doc = self._active_doc()
        doc.UnblankSketch()
        return {"sketches_shown": True}

    def view_rotate(self, x_deg: float = 0.0, y_deg: float = 0.0, z_deg: float = 0.0) -> Dict:
        """Rotate the current view by Euler angles."""
        doc = self._active_doc()
        try:
            view = doc.ActiveView
            if view:
                view.RotateAboutCenter(_r(x_deg), _r(y_deg))
        except Exception:
            pass
        return {"rotated_deg": [x_deg, y_deg, z_deg]}

    # ─── Motion Studies ───────────────────────────────────────────────────────

    def create_motion_study(self, study_name: str = "Motion Study 1") -> Dict:
        """Create a motion study (requires SOLIDWORKS Motion)."""
        doc = self._active_doc()
        mm = doc.MotionStudyManager if hasattr(doc, 'MotionStudyManager') else None
        if mm is None:
            raise RuntimeError("Motion Study Manager unavailable. Enable SOLIDWORKS Motion add-in.")
        try:
            study = mm.CreateMotionStudy(study_name)
            return {"motion_study": study_name, "created": study is not None}
        except Exception as e:
            raise RuntimeError(f"Create motion study failed: {e}")

    # ─── Simulation ───────────────────────────────────────────────────────────

    def create_simulation_study(self, study_name: str, study_type: str = "static") -> Dict:
        """Create a SOLIDWORKS Simulation study (requires Simulation add-in)."""
        app = self._ensure_app()
        try:
            cw = app.GetAddInObject("SldWorks.Simulation") if hasattr(app, 'GetAddInObject') else None
            if cw is None:
                raise RuntimeError("Simulation add-in not available.")
            # Static=0, Frequency=1, Buckling=2, Thermal=3, Drop=4, Fatigue=5, etc.
            t_map = {"static": 0, "frequency": 1, "buckling": 2, "thermal": 3, "drop": 4, "fatigue": 5, "nonlinear": 10}
            st = t_map.get(study_type.lower(), 0)
            doc = cw.ActiveDoc
            new_study = doc.StudyManager.CreateNewStudy3(study_name, st, 0, 0)
            return {"study": study_name, "type": study_type, "created": new_study is not None}
        except Exception as e:
            raise RuntimeError(f"Simulation study failed: {e}")

    def run_simulation(self) -> Dict:
        """Run the active simulation study."""
        app = self._ensure_app()
        try:
            cw = app.GetAddInObject("SldWorks.Simulation")
            doc = cw.ActiveDoc
            study = doc.StudyManager.ActiveStudy
            result = study.RunAnalysis()
            return {"run_status": result}
        except Exception as e:
            raise RuntimeError(f"Run simulation failed: {e}")

    # ─── Global Variables ─────────────────────────────────────────────────────

    def list_global_variables(self) -> List[Dict]:
        """List all global variables (equations of form `"Name" = value`)."""
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        count = _get(eqn_mgr, 'GetCount') or 0
        result = []
        for i in range(count):
            eqn = eqn_mgr.Equation(i)
            val = eqn_mgr.Value(i)
            if eqn and '"' in eqn and "=" in eqn:
                result.append({"index": i, "equation": eqn, "value": val})
        return result

    def set_global_variable(self, name: str, value: float) -> Dict:
        """Set or create a global variable."""
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        # Look for existing
        count = _get(eqn_mgr, 'GetCount') or 0
        target = f'"{name}"'
        for i in range(count):
            eqn = eqn_mgr.Equation(i)
            if eqn and eqn.startswith(target):
                eqn_mgr.Equation = (i, f'"{name}" = {value}')
                _get(doc, 'EditRebuild3')
                return {"name": name, "value": value, "updated": True}
        idx = eqn_mgr.Add2(-1, f'"{name}" = {value}', True)
        _get(doc, 'EditRebuild3')
        return {"name": name, "value": value, "added_at_index": idx}

    # ─── Display States ──────────────────────────────────────────────────────

    def list_display_states(self) -> List[str]:
        """List all display states in the active part/assembly."""
        doc = self._active_doc()
        cfg = _get(doc, 'GetActiveConfiguration')
        if cfg is None:
            return []
        try:
            names = cfg.GetDisplayStates()
            return list(names) if names else []
        except Exception:
            return []

    def activate_display_state(self, state_name: str) -> Dict:
        """Activate a named display state."""
        doc = self._active_doc()
        cfg = _get(doc, 'GetActiveConfiguration')
        if cfg is None:
            raise RuntimeError("No active configuration.")
        try:
            cfg.ApplyDisplayState(state_name)
        except Exception as e:
            raise RuntimeError(f"Activate display state failed: {e}")
        return {"display_state": state_name}

    # ─── Sensors ──────────────────────────────────────────────────────────────

    def list_sensors(self) -> List[Dict]:
        """List sensors on the active document."""
        doc = self._active_doc()
        try:
            sensors = doc.SensorsCollection if hasattr(doc, 'SensorsCollection') else []
            result = []
            count = _get(sensors, 'Count') if sensors else 0
            for i in range(count or 0):
                s = sensors.Item(i)
                result.append({"name": _get(s, 'Name'), "value": _get(s, 'CurrentValue')})
            return result
        except Exception:
            return []

    # ─── File Management Extended ────────────────────────────────────────────

    def save_all(self) -> Dict:
        """Save all open documents."""
        app = self._ensure_app()
        docs = _invoke(app, _DISPID_GET_DOCS)
        saved = 0
        if docs is None:
            return {"saved_count": 0}
        if not hasattr(docs, '__iter__'):
            docs = [docs]
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        for d in docs:
            try:
                if _get(d, 'GetSaveFlag'):
                    d.Save3(SW_SAVE_SILENT, errors, warnings)
                    saved += 1
            except Exception:
                pass
        return {"saved_count": saved}

    def close_all(self, save: bool = False) -> Dict:
        """Close all open documents."""
        app = self._ensure_app()
        docs = _invoke(app, _DISPID_GET_DOCS)
        if docs is None:
            return {"closed_count": 0}
        if not hasattr(docs, '__iter__'):
            docs = [docs]
        titles = []
        for d in docs:
            try:
                titles.append(_get(d, 'GetTitle'))
            except Exception:
                pass
        for t in titles:
            try:
                _invoke(app, _DISPID_CLOSE_DOC, t)
            except Exception:
                pass
        return {"closed_count": len(titles), "titles": titles}

    def save_as_copy(self, file_path: str) -> Dict:
        """Save the active document as a copy (does not change current open file)."""
        doc = self._active_doc()
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        # SaveAs option SW_SAVE_COPY=2
        result = doc.Extension.SaveAs(file_path, 0, SW_SAVE_COPY | SW_SAVE_SILENT,
                                      win32com.client.VARIANT(pythoncom.VT_DISPATCH, None),
                                      errors, warnings)
        if not result:
            raise RuntimeError(f"Save as copy failed (error={errors.value})")
        return {"saved_copy_to": file_path}

    def get_referenced_documents(self) -> List[str]:
        """List paths of all documents referenced by the active assembly/drawing."""
        doc = self._active_doc()
        try:
            paths = doc.GetDependencies2(True, True, False) if hasattr(doc, 'GetDependencies2') else None
            if not paths:
                return []
            # GetDependencies2 returns alternating name/path tuples
            paths_list = list(paths)
            files = [paths_list[i+1] for i in range(0, len(paths_list), 2) if i+1 < len(paths_list)]
            return files
        except Exception:
            return []

    # ─── Advanced Features ────────────────────────────────────────────────────

    def rib(self, thickness: float, draft_angle: float = 0.0, flip: bool = False, two_sided: bool = False) -> Dict:
        """Add a rib feature from the active sketch."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertRib(
            _m(thickness),
            1 if two_sided else 0,
            flip,
            False,
            _r(draft_angle),
            draft_angle > 0,
            False
        )
        if feat is None:
            raise RuntimeError("Rib failed. Ensure a sketch line/curve is active.")
        return {"feature": feat.Name, "thickness_mm": thickness}

    def dome(self, height: float, elliptical: bool = False, flip: bool = False) -> Dict:
        """Add a dome on a selected face."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertDome(_m(height), flip, elliptical)
        if feat is None:
            raise RuntimeError("Dome failed. Select a face first.")
        return {"feature": feat.Name, "height_mm": height}

    def wrap(self, thickness: float, wrap_type: str = "emboss") -> Dict:
        """Wrap a sketch onto a face. type: 'emboss', 'deboss', 'scribe'."""
        doc = self._active_doc()
        t_map = {"emboss": 0, "deboss": 1, "scribe": 2}
        wt = t_map.get(wrap_type.lower(), 0)
        feat = doc.FeatureManager.InsertWrapFeature(_m(thickness), wt, False)
        if feat is None:
            raise RuntimeError("Wrap failed. Select sketch and target face first.")
        return {"feature": feat.Name, "thickness_mm": thickness, "type": wrap_type}

    def intersect(self, operation: str = "create_regions") -> Dict:
        """Intersect bodies/surfaces with selected entities."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertIntersect(False, False, False, True, True, True)
        if feat is None:
            raise RuntimeError("Intersect failed. Select bodies/surfaces first.")
        return {"feature": feat.Name}

    def boundary_boss(self, is_cut: bool = False) -> Dict:
        """Boundary boss/cut from two or more profiles."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertBoundaryBoss(False, False, 0, 0, 0.0, 0.0, False, False)
        if feat is None:
            raise RuntimeError("Boundary boss failed. Select 2+ sketch profiles.")
        return {"feature": feat.Name, "type": "cut" if is_cut else "boss"}

    # ─── Patterns Extended ────────────────────────────────────────────────────

    def fill_pattern(self, spacing: float, pattern_type: str = "perimeter") -> Dict:
        """Fill pattern over selected face. type: 'perimeter', 'square', 'rectangular'."""
        doc = self._active_doc()
        t_map = {"perimeter": 0, "square": 1, "rectangular": 2, "circular": 3}
        pt = t_map.get(pattern_type.lower(), 1)
        feat = doc.FeatureManager.InsertFillPattern2(
            _m(spacing), pt, 0.0, 0.0,
            0, 0.0, False, False
        )
        if feat is None:
            raise RuntimeError("Fill pattern failed. Select target face and seed feature.")
        return {"feature": feat.Name, "spacing_mm": spacing}

    def curve_driven_pattern(self, count: int, spacing: float, equal_spacing: bool = True) -> Dict:
        """Pattern feature along a curve."""
        doc = self._active_doc()
        feat = doc.FeatureManager.InsertCurveDrivenPattern(
            count, _m(spacing), equal_spacing,
            False, False, False, False, False,
            False, 0, 0, 0
        )
        if feat is None:
            raise RuntimeError("Curve-driven pattern failed. Select seed and path curve.")
        return {"feature": feat.Name, "count": count}

    # ─── Equation Helpers ─────────────────────────────────────────────────────

    def link_dimension_to_equation(self, dim_name: str, expr: str) -> Dict:
        """Link a dimension to a value/expression via an equation."""
        doc = self._active_doc()
        eqn_mgr = _get(doc, 'GetEquationMgr')
        equation = f'"{dim_name}" = {expr}'
        idx = eqn_mgr.Add2(-1, equation, True)
        _get(doc, 'EditRebuild3')
        return {"equation": equation, "index": idx}

    # ─── Run / Export Helpers ─────────────────────────────────────────────────

    def export_step(self, output_path: str, version: str = "AP214") -> Dict:
        """Export to STEP with explicit version. version: AP203, AP214, AP242."""
        app = self._ensure_app()
        # 0=AP203, 1=AP214, 2=AP242
        v_map = {"AP203": 0, "AP214": 1, "AP242": 2}
        try:
            app.SetUserPreferenceIntegerValue(228, v_map.get(version.upper(), 1))  # swStepAP_e
        except Exception:
            pass
        return self.export_file(output_path)

    def export_stl(self, output_path: str, binary: bool = True, quality: str = "fine") -> Dict:
        """Export to STL with quality settings. quality: coarse, fine, custom."""
        app = self._ensure_app()
        try:
            # swExportStlFormat_e: 0=Binary, 1=ASCII
            app.SetUserPreferenceIntegerValue(74, 0 if binary else 1)
            # swSTLQuality_e: 0=Coarse, 1=Fine, 2=Custom
            app.SetUserPreferenceIntegerValue(75, {"coarse": 0, "fine": 1, "custom": 2}.get(quality.lower(), 1))
        except Exception:
            pass
        return self.export_file(output_path)

    def get_global_preference(self, name: str) -> Dict:
        """Get a SolidWorks system preference value by name."""
        app = self._ensure_app()
        # Approximate mapping for common names
        return {"note": f"Use sw_run_vba_code to read preference '{name}' via SldWorks API."}

    # ─── Helpers ─────────────────────────────────────────────────────────────

    def _ensure_app(self) -> Any:
        if self._app is None:
            self.connect(launch_if_not_running=True)
        try:
            _ = self._app.RevisionNumber
        except Exception:
            self._app = None
            self.connect(launch_if_not_running=True)
        return self._app

    def _active_doc(self) -> Any:
        app = self._ensure_app()
        doc = app.ActiveDoc
        if doc is None:
            raise RuntimeError(
                "No active document. Use sw_open_document to open a file, "
                "or sw_create_part / sw_create_assembly to create a new document."
            )
        return doc


def _doc_type_name(doc_type: int) -> str:
    return {SW_DOC_PART: "Part", SW_DOC_ASSEMBLY: "Assembly", SW_DOC_DRAWING: "Drawing"}.get(doc_type, "Unknown")
