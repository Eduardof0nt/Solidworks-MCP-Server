#!/usr/bin/env python3
"""
SolidWorks MCP Server

Exposes SolidWorks automation tools via MCP (stdio transport).
Connects to a running SolidWorks instance via COM.

Usage:
    python server.py

Requirements:
    pip install mcp pydantic pywin32
    SolidWorks must be installed on this machine.
"""

import asyncio
import functools
import json
import os
import sys
import concurrent.futures
from typing import Any, Dict, List, Optional, Tuple

# ─── Single-Instance Lock ────────────────────────────────────────────────────
# Multiple Python MCP server processes would each try to grab SolidWorks via
# GetActiveObject("SldWorks.Application"), causing COM contention and tool
# call timeouts. We use a Windows named mutex (auto-released on process exit,
# no stale lock files) to ensure exactly one instance is ever alive.
#
# Set environment variable SW_MCP_ALLOW_MULTIPLE=1 to disable this check
# (useful for debugging / dev — not recommended in production).

def _acquire_single_instance_lock():
    """Acquire a Windows named mutex. Exit if another instance holds it."""
    if os.environ.get("SW_MCP_ALLOW_MULTIPLE") == "1":
        return None
    if sys.platform != "win32":
        return None  # only enforce on Windows (where SolidWorks runs)
    try:
        import win32event
        import win32api
        import winerror
    except ImportError:
        # pywin32 missing — silently skip lock (will fail later anyway when
        # the COM imports try to load)
        return None

    # Use "Local\" prefix (not "Global\") so the lock is per-user-session;
    # this avoids issues with multi-user / RDP scenarios.
    mutex_name = "Local\\SolidWorksMCPServer_SingleInstance_v1"
    mutex = win32event.CreateMutex(None, False, mutex_name)
    last_err = win32api.GetLastError()
    if last_err == winerror.ERROR_ALREADY_EXISTS:
        sys.stderr.write(
            "ERROR: another SolidWorks MCP server instance is already running.\n"
            "  Only one server can run at a time (SolidWorks COM is single-instance).\n"
            "  If you are sure no other instance is alive, restart the host (Claude\n"
            "  Desktop / Claude Code) to clear orphaned subprocesses, or set the\n"
            "  environment variable SW_MCP_ALLOW_MULTIPLE=1 to bypass this check.\n"
        )
        sys.stderr.flush()
        sys.exit(1)
    # Keep handle alive for the lifetime of the process; mutex is auto-released
    # by the kernel when the process exits (clean or crash).
    return mutex


_single_instance_mutex = _acquire_single_instance_lock()

import pythoncom
from pydantic import BaseModel, Field, field_validator, ConfigDict
from mcp.server.fastmcp import FastMCP

from sw_client import SolidWorksClient

# ─── Server & COM Executor Setup ─────────────────────────────────────────────

mcp = FastMCP("solidworks_mcp")

# SolidWorks COM must run on a single STA thread
_executor = concurrent.futures.ThreadPoolExecutor(
    max_workers=1,
    initializer=pythoncom.CoInitialize,
)
_sw = SolidWorksClient()


async def _run(fn, *args, **kwargs):
    """Execute fn(*args, **kwargs) on the dedicated SolidWorks COM thread."""
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        _executor, functools.partial(fn, *args, **kwargs)
    )


def _ok(data: Any) -> str:
    return json.dumps(data, indent=2, default=str)


def _err(e: Exception) -> str:
    msg = str(e)
    # Provide extra context for common COM errors
    if "com_error" in type(e).__name__.lower() or "pythoncom" in msg.lower():
        msg = f"SolidWorks COM error — is SolidWorks running? Details: {msg}"
    return f"Error: {msg}"


# ─── Input Models ─────────────────────────────────────────────────────────────

class ConnectInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    launch_if_not_running: bool = Field(default=True, description="Launch SolidWorks if it is not already running")


class OpenDocInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: str = Field(..., description="Full path to the file (.sldprt, .sldasm, .slddrw, .step, .iges, .stl, etc.)")
    read_only: bool = Field(default=False, description="Open in read-only mode")


class CloseDocInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    title: Optional[str] = Field(default=None, description="Document title to close (leave empty to close the active document)")


class SaveDocInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: Optional[str] = Field(default=None, description="Save-as path (leave empty to save in place)")


class NewPartInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    template_path: Optional[str] = Field(default=None, description="Path to .prtdot template (leave empty for default)")


class NewAssemblyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    template_path: Optional[str] = Field(default=None, description="Path to .asmdot template (leave empty for default)")


class NewDrawingInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    model_path: Optional[str] = Field(default=None, description="Path to part/assembly for the drawing")
    template_path: Optional[str] = Field(default=None, description="Path to .drwdot template (leave empty for default)")


class ExportInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: str = Field(..., description="Output path — extension determines format: .step, .stp, .stl, .dxf, .pdf, .iges, .igs, .x_t, .3mf, .obj, .png, .bmp")


class SelectInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    entity_name: str = Field(..., description="Entity name exactly as shown in the FeatureManager tree (e.g. 'Front Plane', 'Boss-Extrude1', 'Face<1>')")
    entity_type: str = Field(default="PLANE", description="Entity type: PLANE, FACE, EDGE, VERTEX, FEATURE, BODYFEATURE, SKETCH, REFAXIS, COMPONENT, BODY")
    append: bool = Field(default=False, description="Append to existing selection (True) or replace it (False)")


class CreateSketchInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    plane: str = Field(default="Front Plane", description="Plane name to create sketch on: 'Front Plane', 'Top Plane', 'Right Plane', or any reference plane/face name")


class EditSketchInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    sketch_name: str = Field(..., description="Sketch name to edit (e.g. 'Sketch1', 'Sketch2')")


class SketchLineInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x1: float = Field(..., description="Start X in mm")
    y1: float = Field(..., description="Start Y in mm")
    x2: float = Field(..., description="End X in mm")
    y2: float = Field(..., description="End Y in mm")


class SketchCircleInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    cx: float = Field(..., description="Center X in mm")
    cy: float = Field(..., description="Center Y in mm")
    radius: float = Field(..., description="Radius in mm", gt=0)


class SketchRectInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x1: float = Field(..., description="First corner X in mm")
    y1: float = Field(..., description="First corner Y in mm")
    x2: float = Field(..., description="Opposite corner X in mm (or center X if type='center')")
    y2: float = Field(..., description="Opposite corner Y in mm (or half-width if type='center')")
    rect_type: str = Field(default="corner", description="'corner' (two diagonally opposite corners) or 'center' (center point + corner)")


class SketchArcInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    cx: float = Field(..., description="Arc center X in mm")
    cy: float = Field(..., description="Arc center Y in mm")
    radius: float = Field(..., description="Radius in mm", gt=0)
    start_angle: float = Field(..., description="Start angle in degrees (0° = +X axis, CCW positive)")
    end_angle: float = Field(..., description="End angle in degrees")
    clockwise: bool = Field(default=False, description="Draw arc clockwise")


class SketchEllipseInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    cx: float = Field(..., description="Center X in mm")
    cy: float = Field(..., description="Center Y in mm")
    semi_major: float = Field(..., description="Semi-major axis length in mm", gt=0)
    semi_minor: float = Field(..., description="Semi-minor axis length in mm", gt=0)
    rotation: float = Field(default=0.0, description="Rotation of major axis in degrees")


class SketchPolygonInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    cx: float = Field(..., description="Center X in mm")
    cy: float = Field(..., description="Center Y in mm")
    num_sides: int = Field(..., description="Number of sides (3–40)", ge=3, le=40)
    circumscribed_radius: float = Field(..., description="Circumscribed circle radius in mm", gt=0)
    rotation: float = Field(default=0.0, description="Rotation angle in degrees")


class SketchSplineInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    points: List[Tuple[float, float]] = Field(..., description="List of [x, y] control points in mm (minimum 2 points)")

    @field_validator("points")
    @classmethod
    def at_least_two(cls, v):
        if len(v) < 2:
            raise ValueError("Spline requires at least 2 points")
        return v


class SketchOffsetInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    distance: float = Field(..., description="Offset distance in mm", gt=0)
    outward: bool = Field(default=True, description="Offset outward (True) or inward (False)")


class SketchCenterlineInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x1: float = Field(..., description="Start X in mm")
    y1: float = Field(..., description="Start Y in mm")
    x2: float = Field(..., description="End X in mm")
    y2: float = Field(..., description="End Y in mm")


class AddDimInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    value: float = Field(..., description="Dimension value in mm (use degrees for angular dimensions)")
    x: float = Field(default=0.0, description="Label placement X in mm")
    y: float = Field(default=5.0, description="Label placement Y in mm")


class AddRelationInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    relation_type: str = Field(..., description="Constraint type: HORIZONTAL, VERTICAL, COINCIDENT, PARALLEL, PERPENDICULAR, TANGENT, COLLINEAR, CONCENTRIC, EQUAL, SYMMETRIC, MIDPOINT, INTERSECTION, FIX, PIERCE")


class SketchLinearPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x_count: int = Field(..., description="Number of instances in X direction", ge=1)
    y_count: int = Field(default=1, description="Number of instances in Y direction", ge=1)
    x_spacing: float = Field(..., description="Spacing in X direction in mm")
    y_spacing: float = Field(default=10.0, description="Spacing in Y direction in mm")
    x_angle: float = Field(default=0.0, description="Angle of X direction in degrees")


class SketchCircPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    count: int = Field(..., description="Number of instances", ge=2)
    radius: float = Field(..., description="Pattern radius in mm", gt=0)
    angle: float = Field(default=360.0, description="Total sweep angle in degrees", gt=0, le=360)


class ExtrudeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    depth: float = Field(..., description="Extrusion depth in mm", gt=0)
    direction: str = Field(default="blind", description="End condition: 'blind', 'through_all', 'mid_plane', 'up_to_vertex', 'up_to_surface'")
    flip_direction: bool = Field(default=False, description="Flip extrusion direction")
    draft_angle: float = Field(default=0.0, description="Draft taper angle in degrees (0 = no draft)", ge=0, le=89)
    is_cut: bool = Field(default=False, description="True = extruded cut; False = boss/base")
    thin_feature: bool = Field(default=False, description="Create as thin-walled extrusion")
    thin_thickness: float = Field(default=1.0, description="Wall thickness for thin feature in mm", gt=0)


class RevolveInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    angle: float = Field(default=360.0, description="Revolution angle in degrees", gt=0, le=360)
    is_cut: bool = Field(default=False, description="True = revolved cut; False = boss/base")
    flip_direction: bool = Field(default=False, description="Flip revolution direction")


class FilletInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    radius: float = Field(..., description="Fillet radius in mm", gt=0)
    tangent_propagation: bool = Field(default=True, description="Propagate fillet along tangent edges")


class ChamferInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    distance: float = Field(..., description="Chamfer distance in mm", gt=0)
    angle: float = Field(default=45.0, description="Chamfer angle in degrees", gt=0, lt=90)


class ShellInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., description="Shell wall thickness in mm", gt=0)
    outward: bool = Field(default=False, description="Shell outward (adds material outside) or inward (removes material)")


class DraftInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    draft_angle: float = Field(..., description="Draft angle in degrees", gt=0, le=89)


class LinearPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    dir1_count: int = Field(..., description="Number of instances in direction 1 (includes original)", ge=2)
    dir1_spacing: float = Field(..., description="Spacing in direction 1 in mm", gt=0)
    dir2_count: int = Field(default=1, description="Number of instances in direction 2 (1 = no second direction)", ge=1)
    dir2_spacing: float = Field(default=10.0, description="Spacing in direction 2 in mm")
    geometry_pattern: bool = Field(default=False, description="Use geometry pattern (faster, copies geometry only)")


class CircPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    count: int = Field(..., description="Total number of instances including original", ge=2)
    angle: float = Field(default=360.0, description="Total sweep angle in degrees", gt=0, le=360)
    geometry_pattern: bool = Field(default=False, description="Use geometry pattern")


class MirrorInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    mirror_plane: str = Field(..., description="Mirror plane name (e.g. 'Front Plane', 'Top Plane', 'Right Plane', or a reference plane feature name)")


class HoleWizardInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    hole_type: str = Field(default="simple", description="Hole type: 'simple', 'counterbore', 'countersink', 'straight_tap', 'tapered_tap'")
    standard: str = Field(default="ANSI Metric", description="Hole standard (e.g. 'ANSI Metric', 'ISO', 'ANSI Inch')")
    size: str = Field(default="M6", description="Hole size designation (e.g. 'M6', 'M8x1.0', '1/4-20')")
    depth: float = Field(..., description="Hole depth in mm", gt=0)
    through_all: bool = Field(default=False, description="Through-all depth (ignores depth value)")
    x: float = Field(default=0.0, description="Hole placement X in mm")
    y: float = Field(default=0.0, description="Hole placement Y in mm")


class LoftInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    is_cut: bool = Field(default=False, description="True = loft cut; False = loft boss")
    close_loft: bool = Field(default=False, description="Close the loft (periodic loft)")


class SweepInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    is_cut: bool = Field(default=False, description="True = sweep cut; False = sweep boss")
    merge_result: bool = Field(default=True, description="Merge result with existing bodies")


class DeleteFeatureInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    feature_name: str = Field(..., description="Feature name to delete (as shown in FeatureManager tree)")
    absorb_children: bool = Field(default=False, description="Absorb child features instead of deleting them")


class SuppressInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    feature_name: str = Field(..., description="Feature name to suppress/unsuppress")


class InsertComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: str = Field(..., description="Full path to component file (.sldprt or .sldasm)")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")
    z: float = Field(default=0.0, description="Z position in mm")
    fixed: bool = Field(default=False, description="Fix the component in place (no degrees of freedom)")


class AddMateInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    mate_type: str = Field(..., description="Mate type: COINCIDENT, CONCENTRIC, PARALLEL, PERPENDICULAR, TANGENT, DISTANCE, ANGLE, SYMMETRIC, LOCK, HINGE, GEAR, SLOT, PROFILE_CENTER, LINEAR_COUPLER")
    entity1: str = Field(..., description="First entity name (face/edge/axis/plane in assembly context, e.g. 'Face<1>@ComponentName-1')")
    entity2: str = Field(..., description="Second entity name")
    distance_or_angle: Optional[float] = Field(default=None, description="Distance (mm) for DISTANCE mate, angle (degrees) for ANGLE mate")
    flip_mate: bool = Field(default=False, description="Flip mate alignment direction")


class FixComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    component_name: str = Field(..., description="Component name as shown in assembly FeatureManager")
    fix: bool = Field(default=True, description="True = fix, False = float (remove fix)")


class CheckInterferenceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    treat_sub_as_comp: bool = Field(default=True, description="Treat subassemblies as single rigid components")


class DrawingViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    view_type: str = Field(default="front", description="View orientation: 'front', 'back', 'left', 'right', 'top', 'bottom', 'isometric', 'trimetric', 'dimetric'")
    x: float = Field(default=0.1, description="X position on sheet in meters")
    y: float = Field(default=0.15, description="Y position on sheet in meters")
    scale: float = Field(default=1.0, description="View scale (1.0 = 1:1, 0.5 = 1:2)", gt=0)
    model_path: Optional[str] = Field(default=None, description="Model path for the first view (leave empty to use active model)")


class SectionViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    label: str = Field(default="A", description="Section view label (A, B, C, ...)")
    offset: float = Field(default=0.0, description="Section offset distance in mm")


class DrawingDimInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(default=0.0, description="Label placement X in mm")
    y: float = Field(default=0.0, description="Label placement Y in mm")


class AnnotationInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    text: str = Field(..., description="Annotation text (supports \\n for new lines)")
    x: float = Field(default=50.0, description="X position in mm on sheet")
    y: float = Field(default=50.0, description="Y position in mm on sheet")


class GetDimInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    dimension_name: str = Field(..., description="Dimension name in 'DimName@FeatureName' format, e.g. 'D1@Sketch1', 'D2@Boss-Extrude1'")


class SetDimInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    dimension_name: str = Field(..., description="Dimension name in 'DimName@FeatureName' format")
    value: float = Field(..., description="New value in mm")


class GetPropInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    property_name: Optional[str] = Field(default=None, description="Property name to retrieve (leave empty for all properties)")
    configuration: str = Field(default="", description="Configuration name (empty = document-level properties)")


class SetPropInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    property_name: str = Field(..., description="Custom property name")
    value: str = Field(..., description="Property value (can include equations like '\"D1@Boss-Extrude1\"')")
    configuration: str = Field(default="", description="Configuration name (empty = document-level)")


class ActivateConfigInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    config_name: str = Field(..., description="Configuration name to activate")


class AddConfigInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    config_name: str = Field(..., description="New configuration name")
    derived_from: Optional[str] = Field(default=None, description="Existing configuration to copy from")


class EquationInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    equation: str = Field(..., description="Equation string (e.g. '\"Width\" = 100mm', '\"D1@Sketch1\" = \"Height\" / 2')")


class DeleteEqnInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    index: int = Field(..., description="Equation index (from sw_get_equations)", ge=0)


class DesignTableInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    excel_path: Optional[str] = Field(default=None, description="Path to Excel (.xlsx) file (leave empty to create embedded table)")


class ScreenshotInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: Optional[str] = Field(default=None, description="Output image path (.bmp, .png, .jpg). Leave empty to save to temp folder.")


class ViewOrientInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    view_name: str = Field(..., description="View: 'front', 'back', 'left', 'right', 'top', 'bottom', 'isometric', 'trimetric', 'dimetric'")


class DisplayModeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    mode: str = Field(..., description="Display mode: 'wireframe', 'hidden_lines_visible', 'hidden_lines_removed', 'shaded', 'shaded_with_edges'")


class RebuildInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    top_only: bool = Field(default=False, description="True = rebuild top level only (faster for assemblies)")


class RunMacroInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    macro_path: str = Field(..., description="Full path to the macro file (.swp or .swb)")
    module_name: str = Field(default="", description="VBA module name (leave empty for default)")
    sub_name: str = Field(default="main", description="Subroutine name to run")


class RunVBAInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    vba_code: str = Field(..., description="VBA code to execute. Will be auto-wrapped in Sub main()...End Sub if no Sub declaration is present.")


class RefPlaneInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    reference_plane: str = Field(default="Front Plane", description="Reference plane to offset from")
    offset_distance: float = Field(default=0.0, description="Offset distance in mm")


class RefAxisInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    axis_type: str = Field(default="two_planes", description="Axis definition type: 'two_planes', 'cylinder', 'two_points', 'point_on_face'")


class RefPointInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(..., description="X coordinate in mm")
    y: float = Field(..., description="Y coordinate in mm")
    z: float = Field(..., description="Z coordinate in mm")


class MeasureInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    entity1: str = Field(..., description="First entity name (face, edge, or vertex)")
    entity2: str = Field(..., description="Second entity name")


class ThickenInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., description="Thicken amount in mm", gt=0)
    both_sides: bool = Field(default=False, description="Thicken on both sides")


class FlatPatternExportInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: str = Field(..., description="Output path (.dxf, .dwg, or .pdf)")


class DraftAnalysisInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    draft_angle: float = Field(default=1.0, description="Reference draft angle in degrees", gt=0, le=89)


# ─── Connection Tools ─────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_connect",
    annotations={"title": "Connect to SolidWorks", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_connect(params: ConnectInput) -> str:
    """Connect to a running SolidWorks instance or launch one.

    Must be called before other sw_* tools if SolidWorks is not yet connected.
    Most tools will auto-connect if needed, but this lets you verify the connection.

    Returns version info and connection status.
    """
    try:
        result = await _run(_sw.connect, params.launch_if_not_running)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_get_status",
    annotations={"title": "Get SolidWorks Status", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_status() -> str:
    """Get SolidWorks connection status, active document info, and list of open documents.

    Returns:
        JSON with connected (bool), version, active_document info, and open_documents list.
    """
    try:
        result = await _run(_sw.get_status)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── File / Document Tools ───────────────────────────────────────────────────

@mcp.tool(
    name="sw_open_document",
    annotations={"title": "Open SolidWorks Document", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_open_document(params: OpenDocInput) -> str:
    """Open a SolidWorks file (part, assembly, drawing, or imported format like STEP/IGES/STL).

    Supports: .sldprt, .sldasm, .slddrw, .step, .stp, .iges, .igs, .stl, .x_t, .sat, .3mf, .obj

    Returns document title, path, and type.
    """
    try:
        result = await _run(_sw.open_document, params.file_path, params.read_only)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_close_document",
    annotations={"title": "Close SolidWorks Document", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False}
)
async def sw_close_document(params: CloseDocInput) -> str:
    """Close a document. Closes the active document if no title is given.

    WARNING: Does not prompt to save — save first if needed with sw_save_document.
    """
    try:
        result = await _run(_sw.close_document, params.title)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_save_document",
    annotations={"title": "Save SolidWorks Document", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_save_document(params: SaveDocInput) -> str:
    """Save the active document. Provide file_path to Save As, leave empty to save in-place.

    Returns the saved path and any warning codes.
    """
    try:
        result = await _run(_sw.save_document, params.file_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_create_part",
    annotations={"title": "Create New Part", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_part(params: NewPartInput) -> str:
    """Create a new SolidWorks part document.

    Optionally specify a template .prtdot file, otherwise uses the SolidWorks default template.
    Returns the new document title.
    """
    try:
        result = await _run(_sw.create_part, params.template_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_create_assembly",
    annotations={"title": "Create New Assembly", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_assembly(params: NewAssemblyInput) -> str:
    """Create a new SolidWorks assembly document.

    Returns the new document title.
    """
    try:
        result = await _run(_sw.create_assembly, params.template_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_create_drawing",
    annotations={"title": "Create New Drawing", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_drawing(params: NewDrawingInput) -> str:
    """Create a new SolidWorks drawing document.

    Returns the new document title.
    """
    try:
        result = await _run(_sw.create_drawing, params.model_path, params.template_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_list_open_documents",
    annotations={"title": "List Open Documents", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_list_open_documents() -> str:
    """List all documents currently open in SolidWorks.

    Returns a list with title, path, type, and modified status for each document.
    """
    try:
        result = await _run(_sw.list_open_documents)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_export_file",
    annotations={"title": "Export File", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_export_file(params: ExportInput) -> str:
    """Export the active document to another format.

    Supported output extensions: .step / .stp, .stl, .dxf, .pdf, .iges / .igs, .x_t (Parasolid),
    .3mf, .obj, .png, .bmp, .jpg, .edrw, .u3d, .vrml

    Example: output_path = "C:/output/part.step"
    """
    try:
        result = await _run(_sw.export_file, params.output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_get_model_info",
    annotations={"title": "Get Model Info", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_model_info() -> str:
    """Get comprehensive info about the active document: title, path, type, configurations, and feature tree.

    Returns document metadata plus a full list of features with names and types.
    """
    try:
        result = await _run(_sw.get_model_info)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Selection Tools ──────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_select_entity",
    annotations={"title": "Select Entity", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_select_entity(params: SelectInput) -> str:
    """Select an entity in the active document by name and type.

    Entity types: PLANE, FACE, EDGE, VERTEX, BODYFEATURE, SKETCH, REFAXIS, COMPONENT, BODY

    Common examples:
    - Select Front Plane: name='Front Plane', type='PLANE'
    - Select an extrude feature: name='Boss-Extrude1', type='BODYFEATURE'
    - Select a sketch: name='Sketch1', type='SKETCH'

    Use append=True to add to the current selection (e.g. for mates, patterns).
    """
    try:
        result = await _run(_sw.select_entity, params.entity_name, params.entity_type, params.append)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_clear_selection",
    annotations={"title": "Clear Selection", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_clear_selection() -> str:
    """Clear all current selections in the active document."""
    try:
        result = await _run(_sw.clear_selection)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_list_features",
    annotations={"title": "List Features", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_list_features() -> str:
    """List all features in the active model's FeatureManager tree.

    Returns each feature's name, type string (e.g. 'Boss-Extrude', 'Cut-Extrude', 'Fillet'), and suppression state.
    """
    try:
        result = await _run(_sw.list_features)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Sketch Tools ─────────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_create_sketch",
    annotations={"title": "Create Sketch", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_sketch(params: CreateSketchInput) -> str:
    """Create a new sketch on the specified plane or face.

    After calling this, use sw_sketch_* tools to add geometry, then sw_exit_sketch to finish.
    Standard planes: 'Front Plane', 'Top Plane', 'Right Plane'
    """
    try:
        result = await _run(_sw.create_sketch, params.plane)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_edit_sketch",
    annotations={"title": "Edit Existing Sketch", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_edit_sketch(params: EditSketchInput) -> str:
    """Open an existing sketch for editing.

    After editing, call sw_exit_sketch. Use sw_list_features to find sketch names.
    """
    try:
        result = await _run(_sw.edit_sketch, params.sketch_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_exit_sketch",
    annotations={"title": "Exit Sketch", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_exit_sketch() -> str:
    """Exit the active sketch and return to the 3D model view.

    Must be called after finishing sketch geometry before creating features.
    """
    try:
        result = await _run(_sw.exit_sketch)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_line",
    annotations={"title": "Sketch Line", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_line(params: SketchLineInput) -> str:
    """Draw a line in the active sketch. Coordinates are in mm on the sketch plane.

    A sketch must be active (use sw_create_sketch first).
    """
    try:
        result = await _run(_sw.sketch_line, params.x1, params.y1, params.x2, params.y2)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_circle",
    annotations={"title": "Sketch Circle", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_circle(params: SketchCircleInput) -> str:
    """Draw a circle in the active sketch. Coordinates and radius are in mm.

    A sketch must be active (use sw_create_sketch first).
    """
    try:
        result = await _run(_sw.sketch_circle, params.cx, params.cy, params.radius)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_rectangle",
    annotations={"title": "Sketch Rectangle", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_rectangle(params: SketchRectInput) -> str:
    """Draw a rectangle in the active sketch.

    type='corner': x1,y1 and x2,y2 are diagonally opposite corners.
    type='center': x1,y1 is the center, x2,y2 is a corner.
    Coordinates in mm. A sketch must be active.
    """
    try:
        result = await _run(_sw.sketch_rectangle, params.x1, params.y1, params.x2, params.y2, params.rect_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_arc",
    annotations={"title": "Sketch Arc", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_arc(params: SketchArcInput) -> str:
    """Draw an arc in the active sketch defined by center, radius, and start/end angles.

    Angles are in degrees (0° = +X axis, counter-clockwise positive).
    Coordinates and radius are in mm. A sketch must be active.
    """
    try:
        result = await _run(_sw.sketch_arc, params.cx, params.cy, params.radius, params.start_angle, params.end_angle, params.clockwise)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_ellipse",
    annotations={"title": "Sketch Ellipse", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_ellipse(params: SketchEllipseInput) -> str:
    """Draw an ellipse in the active sketch. Coordinates in mm.

    rotation rotates the major axis from the X-axis. A sketch must be active.
    """
    try:
        result = await _run(_sw.sketch_ellipse, params.cx, params.cy, params.semi_major, params.semi_minor, params.rotation)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_polygon",
    annotations={"title": "Sketch Polygon", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_polygon(params: SketchPolygonInput) -> str:
    """Draw a regular polygon in the active sketch. Coordinates and radius in mm.

    circumscribed_radius is the distance from center to vertex.
    rotation rotates the polygon. A sketch must be active.
    """
    try:
        result = await _run(_sw.sketch_polygon, params.cx, params.cy, params.num_sides, params.circumscribed_radius, params.rotation)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_spline",
    annotations={"title": "Sketch Spline", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_spline(params: SketchSplineInput) -> str:
    """Draw a spline through the given control points. Points are [x, y] in mm.

    Requires at least 2 points. A sketch must be active.
    Example: points=[[0,0],[10,20],[30,15],[50,0]]
    """
    try:
        result = await _run(_sw.sketch_spline, params.points)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_offset",
    annotations={"title": "Sketch Offset Entities", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_offset(params: SketchOffsetInput) -> str:
    """Offset selected sketch entities by a given distance.

    Select the entities to offset first with sw_select_entity, then call this.
    """
    try:
        result = await _run(_sw.sketch_offset, params.distance, params.outward)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_centerline",
    annotations={"title": "Sketch Centerline", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_centerline(params: SketchCenterlineInput) -> str:
    """Draw a centerline in the active sketch. Used as axis for revolve features or symmetry.

    Coordinates in mm. A sketch must be active.
    """
    try:
        result = await _run(_sw.sketch_centerline, params.x1, params.y1, params.x2, params.y2)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_linear_pattern",
    annotations={"title": "Sketch Linear Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_linear_pattern(params: SketchLinearPatternInput) -> str:
    """Create a linear step-and-repeat pattern of selected sketch entities.

    Select the sketch entities first with sw_select_entity (append=True for multiple).
    """
    try:
        result = await _run(_sw.sketch_linear_pattern, params.x_count, params.y_count, params.x_spacing, params.y_spacing, params.x_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sketch_circular_pattern",
    annotations={"title": "Sketch Circular Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sketch_circular_pattern(params: SketchCircPatternInput) -> str:
    """Create a circular pattern of selected sketch entities.

    Select the sketch entities first with sw_select_entity.
    """
    try:
        result = await _run(_sw.sketch_circular_pattern, params.count, params.radius, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_sketch_dimension",
    annotations={"title": "Add Sketch Dimension", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_sketch_dimension(params: AddDimInput) -> str:
    """Add a Smart Dimension to the selected sketch entity and set its value.

    Select the line, circle, or arc to dimension with sw_select_entity first.
    x,y sets where the dimension label appears.

    Example workflow:
    1. sw_create_sketch('Front Plane')
    2. sw_sketch_line(x1=0, y1=0, x2=50, y2=0)
    3. sw_select_entity('Line1', 'SKETCHSEGMENT')
    4. sw_add_sketch_dimension(value=50, x=25, y=-10)
    """
    try:
        result = await _run(_sw.add_sketch_dimension, params.value, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_sketch_constraint",
    annotations={"title": "Add Sketch Constraint", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_sketch_constraint(params: AddRelationInput) -> str:
    """Add a geometric constraint (relation) to selected sketch entities.

    Select entities first with sw_select_entity (use append=True for multiple).

    Available constraints: HORIZONTAL, VERTICAL, COINCIDENT, PARALLEL, PERPENDICULAR,
    TANGENT, COLLINEAR, CONCENTRIC, EQUAL, SYMMETRIC, MIDPOINT, INTERSECTION, FIX, PIERCE

    Example: make a line horizontal
    1. sw_select_entity('Line1', 'SKETCHSEGMENT')
    2. sw_add_sketch_constraint('HORIZONTAL')
    """
    try:
        result = await _run(_sw.add_sketch_constraint, params.relation_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Feature Tools ────────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_extrude",
    annotations={"title": "Extrude Boss/Cut", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_extrude(params: ExtrudeInput) -> str:
    """Extrude the active sketch profile into a 3D boss (adds material) or cut (removes material).

    The sketch must be closed and the document must be a part.
    Use sw_exit_sketch first, then the last closed sketch is used.

    Returns the new feature name.
    """
    try:
        result = await _run(_sw.extrude, params.depth, params.direction, params.flip_direction,
                            params.draft_angle, params.is_cut, params.thin_feature, params.thin_thickness)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_revolve",
    annotations={"title": "Revolve Boss/Cut", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_revolve(params: RevolveInput) -> str:
    """Revolve the active sketch profile around a centerline/axis.

    The sketch must contain a profile and a centerline (use sw_sketch_centerline).
    Use sw_exit_sketch first.

    Returns the new feature name.
    """
    try:
        result = await _run(_sw.revolve, params.angle, params.is_cut, params.flip_direction)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_fillet",
    annotations={"title": "Add Fillet", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_fillet(params: FilletInput) -> str:
    """Add a constant-radius fillet to selected edges.

    Select edges first with sw_select_entity (type='EDGE', use append=True for multiple).

    Returns the new fillet feature name.
    """
    try:
        result = await _run(_sw.fillet, params.radius, params.tangent_propagation)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_chamfer",
    annotations={"title": "Add Chamfer", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_chamfer(params: ChamferInput) -> str:
    """Add a chamfer to selected edges.

    Select edges first with sw_select_entity (type='EDGE').

    Returns the new chamfer feature name.
    """
    try:
        result = await _run(_sw.chamfer, params.distance, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_shell",
    annotations={"title": "Shell Body", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_shell(params: ShellInput) -> str:
    """Shell a solid body (hollow it out with a uniform wall thickness).

    Select the face(s) to open (remove) with sw_select_entity (type='FACE') before calling.
    Leave unselected to shell all faces inward.

    Returns the new shell feature name.
    """
    try:
        result = await _run(_sw.shell, params.thickness, params.outward)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_draft",
    annotations={"title": "Add Draft", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_draft(params: DraftInput) -> str:
    """Add draft (taper angle) to selected faces.

    Select the neutral plane (PLANE) and the faces to draft (FACE) before calling.
    Select neutral plane first, then hold Ctrl and select the draft faces.

    Returns the new draft feature name.
    """
    try:
        result = await _run(_sw.draft, params.draft_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_linear_pattern",
    annotations={"title": "Linear Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_linear_pattern(params: LinearPatternInput) -> str:
    """Create a linear pattern of a selected feature.

    Select the feature to pattern and direction edges/axes with sw_select_entity first.
    Select direction1 edge, then Ctrl+select direction2 edge (optional), then Ctrl+select features.

    Returns the new pattern feature name.
    """
    try:
        result = await _run(_sw.linear_pattern, params.dir1_count, params.dir1_spacing,
                            params.dir2_count, params.dir2_spacing, params.geometry_pattern)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_circular_pattern",
    annotations={"title": "Circular Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_circular_pattern(params: CircPatternInput) -> str:
    """Create a circular pattern of a selected feature around an axis.

    Select the rotation axis (REFAXIS or edge) and the features to pattern first.

    Returns the new pattern feature name.
    """
    try:
        result = await _run(_sw.circular_pattern, params.count, params.angle, params.geometry_pattern)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_mirror",
    annotations={"title": "Mirror Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_mirror(params: MirrorInput) -> str:
    """Mirror selected features across a plane.

    Select features to mirror with sw_select_entity first.
    mirror_plane is the plane to mirror about (e.g. 'Front Plane', 'Right Plane').

    Returns the new mirror feature name.
    """
    try:
        result = await _run(_sw.mirror_feature, params.mirror_plane)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_hole_wizard",
    annotations={"title": "Hole Wizard", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_hole_wizard(params: HoleWizardInput) -> str:
    """Create a hole using the Hole Wizard (supports standard threaded/clearance holes).

    Select the face to place the hole on first with sw_select_entity (type='FACE').

    hole_type: 'simple', 'counterbore', 'countersink', 'straight_tap', 'tapered_tap'
    standard: 'ANSI Metric', 'ANSI Inch', 'ISO'
    size: e.g. 'M6', 'M8x1.25', '1/4-20', '#10-32'

    Returns the new hole feature name.
    """
    try:
        result = await _run(_sw.hole_wizard, params.hole_type, params.standard, params.size,
                            params.depth, params.through_all, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_loft",
    annotations={"title": "Loft Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_loft(params: LoftInput) -> str:
    """Create a loft between selected sketch profiles.

    Select 2 or more sketch profiles in order with sw_select_entity (append=True).
    Optionally select guide curves before running.

    Returns the new loft feature name.
    """
    try:
        result = await _run(_sw.loft, params.is_cut, params.close_loft)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_sweep",
    annotations={"title": "Sweep Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_sweep(params: SweepInput) -> str:
    """Create a sweep of a profile sketch along a path sketch.

    Select the profile sketch first, then Ctrl+select the path sketch.

    Returns the new sweep feature name.
    """
    try:
        result = await _run(_sw.sweep, params.is_cut, params.merge_result)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_delete_feature",
    annotations={"title": "Delete Feature", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False}
)
async def sw_delete_feature(params: DeleteFeatureInput) -> str:
    """Delete a feature from the FeatureManager tree by name.

    WARNING: This is irreversible (unless you undo). Use sw_list_features to find feature names.
    absorb_children=True will keep child features (e.g. keep a sketch when deleting an extrude).
    """
    try:
        result = await _run(_sw.delete_feature, params.feature_name, params.absorb_children)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_suppress_feature",
    annotations={"title": "Suppress Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_suppress_feature(params: SuppressInput) -> str:
    """Suppress a feature (temporarily disable it without deleting).

    Use sw_unsuppress_feature to re-enable. Use sw_list_features to find feature names.
    """
    try:
        result = await _run(_sw.suppress_feature, params.feature_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_unsuppress_feature",
    annotations={"title": "Unsuppress Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_unsuppress_feature(params: SuppressInput) -> str:
    """Unsuppress (re-enable) a previously suppressed feature.

    Use sw_list_features to find feature names and their suppression state.
    """
    try:
        result = await _run(_sw.unsuppress_feature, params.feature_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Assembly Tools ───────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_insert_component",
    annotations={"title": "Insert Assembly Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_insert_component(params: InsertComponentInput) -> str:
    """Insert a part or subassembly into the active assembly.

    x, y, z set the initial placement position in mm.
    fixed=True locks the component in place.

    Returns the component name as it appears in the assembly tree.
    """
    try:
        result = await _run(_sw.insert_component, params.file_path, params.x, params.y, params.z, params.fixed)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_mate",
    annotations={"title": "Add Assembly Mate", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_mate(params: AddMateInput) -> str:
    """Add a mate constraint between two entities in an assembly.

    entity1/entity2: entity names in 'EntityName@ComponentName-Instance' format.
    For example: 'Face<1>@Bolt-1' or 'Plane1@Housing-2'

    mate_type options: COINCIDENT, CONCENTRIC, PARALLEL, PERPENDICULAR, TANGENT,
    DISTANCE, ANGLE, SYMMETRIC, LOCK, HINGE, GEAR, SLOT, PROFILE_CENTER

    For DISTANCE mates: distance_or_angle = gap in mm
    For ANGLE mates: distance_or_angle = angle in degrees

    Returns the mate feature name.
    """
    try:
        result = await _run(_sw.add_mate, params.mate_type, params.entity1, params.entity2,
                            params.distance_or_angle, params.flip_mate)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_fix_component",
    annotations={"title": "Fix/Float Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_fix_component(params: FixComponentInput) -> str:
    """Fix a component in place (remove all DOF) or float it (restore DOF).

    fix=True makes it rigid. fix=False allows mates to position it.
    """
    try:
        result = await _run(_sw.fix_component, params.component_name, params.fix)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_check_interference",
    annotations={"title": "Check Assembly Interference", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_check_interference(params: CheckInterferenceInput) -> str:
    """Check for interference (collision) between components in the active assembly.

    Returns the number of interferences found and their approximate volumes.
    Active document must be an assembly.
    """
    try:
        result = await _run(_sw.check_interference, params.treat_sub_as_comp)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_explode_assembly",
    annotations={"title": "Explode Assembly View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_explode_assembly() -> str:
    """Create or activate an exploded view of the active assembly.

    Use with sw_capture_screenshot to generate exploded view images.
    """
    try:
        result = await _run(_sw.explode_assembly)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Drawing Tools ────────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_add_drawing_view",
    annotations={"title": "Add Drawing View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_drawing_view(params: DrawingViewInput) -> str:
    """Add a model view to the active drawing document.

    view_type: 'front', 'back', 'left', 'right', 'top', 'bottom', 'isometric', 'trimetric'
    x, y: position on sheet in meters (note: meters not mm for sheet coordinates)
    scale: 1.0 = 1:1, 0.5 = 1:2, 2.0 = 2:1

    Active document must be a drawing. Use sw_create_drawing first.
    """
    try:
        result = await _run(_sw.add_drawing_view, params.view_type, params.x, params.y, params.scale, params.model_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_section_view",
    annotations={"title": "Add Section View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_section_view(params: SectionViewInput) -> str:
    """Add a section view to the drawing.

    Select a cutting line on an existing view first. label is the section identifier (A, B, C...).
    """
    try:
        result = await _run(_sw.add_section_view, params.label, params.offset)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_drawing_dimension",
    annotations={"title": "Add Drawing Dimension", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_drawing_dimension(params: DrawingDimInput) -> str:
    """Add a dimension to the active drawing view.

    Select the entities to dimension first, then specify where to place the label.
    x, y are in mm on the drawing sheet.
    """
    try:
        result = await _run(_sw.add_drawing_dimension, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_annotation",
    annotations={"title": "Add Note/Annotation", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_annotation(params: AnnotationInput) -> str:
    """Add a text note/annotation to the active drawing or model.

    text supports \\n for multi-line notes.
    x, y are position in mm.
    """
    try:
        result = await _run(_sw.add_annotation, params.text, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_bom_table",
    annotations={"title": "Add Bill of Materials Table", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_bom_table() -> str:
    """Add a Bill of Materials (BOM) table to the active drawing.

    The drawing must contain an assembly view. Places the BOM at a default position.
    """
    try:
        result = await _run(_sw.add_bom_table)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Dimensions / Parameters ──────────────────────────────────────────────────

@mcp.tool(
    name="sw_get_dimension",
    annotations={"title": "Get Dimension Value", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_dimension(params: GetDimInput) -> str:
    """Get the current value of a named dimension.

    dimension_name format: 'DimName@FeatureName' (e.g. 'D1@Sketch1', 'Width@Boss-Extrude1')
    Values are returned in mm.

    To find dimension names: use the Properties dialog in SolidWorks or sw_get_equations.
    """
    try:
        result = await _run(_sw.get_dimension, params.dimension_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_set_dimension",
    annotations={"title": "Set Dimension Value", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_set_dimension(params: SetDimInput) -> str:
    """Set a named dimension to a new value and rebuild the model.

    dimension_name format: 'DimName@FeatureName' (e.g. 'D1@Sketch1', 'Width@Boss-Extrude1')
    value is in mm.

    This is the primary way to parametrically drive model geometry.
    """
    try:
        result = await _run(_sw.set_dimension, params.dimension_name, params.value)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Model Analysis Tools ─────────────────────────────────────────────────────

@mcp.tool(
    name="sw_get_mass_properties",
    annotations={"title": "Get Mass Properties", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_mass_properties() -> str:
    """Calculate mass properties of the active part or assembly.

    Returns: mass (kg), volume (mm³), surface area (mm²), density (kg/m³),
    center of mass (mm), and principal moments of inertia.

    Requires the material to be set for accurate mass. Volume is always correct.
    """
    try:
        result = await _run(_sw.get_mass_properties)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_get_bounding_box",
    annotations={"title": "Get Bounding Box", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_bounding_box() -> str:
    """Get the axis-aligned bounding box of the active model.

    Returns min/max corners and overall dimensions (width × height × depth) in mm.
    Useful for envelope sizing and clearance checks.
    """
    try:
        result = await _run(_sw.get_bounding_box)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_measure_distance",
    annotations={"title": "Measure Distance", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_measure_distance(params: MeasureInput) -> str:
    """Measure the minimum distance between two entities (faces, edges, vertices).

    Returns distance in mm and delta X/Y/Z components.
    Entity names must match exactly as shown in the FeatureManager tree.
    """
    try:
        result = await _run(_sw.measure_distance, params.entity1, params.entity2)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_check_geometry",
    annotations={"title": "Check Geometry", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_check_geometry() -> str:
    """Check the active part for geometry errors (invalid faces, edges, etc.).

    Returns error count and status. Run after complex operations to validate the model.
    Only works on part documents.
    """
    try:
        result = await _run(_sw.check_geometry)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_analyze_draft",
    annotations={"title": "Draft Analysis", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_analyze_draft(params: DraftAnalysisInput) -> str:
    """Perform draft analysis on the active part model.

    Highlights faces with sufficient draft (green), faces requiring draft (red),
    and straddle faces (yellow). Use for injection mold design review.

    draft_angle: minimum acceptable draft angle in degrees.
    """
    try:
        result = await _run(_sw.analyze_draft, params.draft_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Custom Properties ────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_get_custom_properties",
    annotations={"title": "Get Custom Properties", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_custom_properties(params: GetPropInput) -> str:
    """Get custom properties from the active document.

    Leave property_name empty to retrieve all properties.
    configuration: config name for config-specific properties, or empty for document-level.

    Returns name/value pairs including linked (resolved) values.
    """
    try:
        result = await _run(_sw.get_custom_properties, params.property_name, params.configuration)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_set_custom_property",
    annotations={"title": "Set Custom Property", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_set_custom_property(params: SetPropInput) -> str:
    """Set or create a custom property on the active document.

    Supports text values and linked equations (e.g. value='\"D1@Boss-Extrude1\"').
    configuration: leave empty for document-level, or specify a config name.
    """
    try:
        result = await _run(_sw.set_custom_property, params.property_name, params.value, params.configuration)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Configuration Tools ──────────────────────────────────────────────────────

@mcp.tool(
    name="sw_list_configurations",
    annotations={"title": "List Configurations", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_list_configurations() -> str:
    """List all configurations in the active document and show which is active.

    Configurations are alternate states of a part/assembly (different dimensions, features, etc.).
    """
    try:
        result = await _run(_sw.list_configurations)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_activate_configuration",
    annotations={"title": "Activate Configuration", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_activate_configuration(params: ActivateConfigInput) -> str:
    """Activate (switch to) a configuration by name.

    Use sw_list_configurations to see available configuration names.
    Rebuilds the model after switching.
    """
    try:
        result = await _run(_sw.activate_configuration, params.config_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_configuration",
    annotations={"title": "Add Configuration", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_configuration(params: AddConfigInput) -> str:
    """Add a new configuration to the active document.

    The new configuration is a copy of the current (or derived_from) configuration.
    Use sw_set_dimension or suppress/unsuppress features to differentiate it.
    """
    try:
        result = await _run(_sw.add_configuration, params.config_name, params.derived_from)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Equations Tools ──────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_get_equations",
    annotations={"title": "Get Equations", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_equations() -> str:
    """List all equations and global variables in the active document.

    Returns index, equation string, and current evaluated value for each.
    Use indices with sw_delete_equation.
    """
    try:
        result = await _run(_sw.get_equations)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_add_equation",
    annotations={"title": "Add Equation", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_add_equation(params: EquationInput) -> str:
    """Add an equation or global variable to the active document.

    Equation format examples:
    - Global variable: '"Width" = 100mm'
    - Dimension link: '"D1@Sketch1" = "Width" * 2'
    - Math: '"Radius" = sqrt("Width" ^ 2 + "Height" ^ 2)'

    Rebuilds the model after adding.
    """
    try:
        result = await _run(_sw.add_equation, params.equation)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_delete_equation",
    annotations={"title": "Delete Equation", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False}
)
async def sw_delete_equation(params: DeleteEqnInput) -> str:
    """Delete an equation by its index (from sw_get_equations).

    Rebuilds the model after deletion. Affected dimensions revert to manual.
    """
    try:
        result = await _run(_sw.delete_equation, params.index)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Design Table ─────────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_create_design_table",
    annotations={"title": "Create Design Table", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_design_table(params: DesignTableInput) -> str:
    """Create a design table for driving multiple configurations from a spreadsheet.

    Provide an Excel (.xlsx) file path for a linked table, or leave empty to create an embedded one.
    SolidWorks will open Excel for editing if the table is created successfully.
    """
    try:
        result = await _run(_sw.create_design_table, params.excel_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Visualization Tools ──────────────────────────────────────────────────────

@mcp.tool(
    name="sw_capture_screenshot",
    annotations={"title": "Capture Screenshot", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_capture_screenshot(params: ScreenshotInput) -> str:
    """Capture a screenshot of the active SolidWorks model view.

    output_path: where to save the image (.bmp, .png, .jpg). Defaults to a temp file.
    Set the view with sw_set_view_orientation before capturing.

    Returns the path to the saved image file.
    """
    try:
        result = await _run(_sw.capture_screenshot, params.output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_set_view_orientation",
    annotations={"title": "Set View Orientation", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_set_view_orientation(params: ViewOrientInput) -> str:
    """Set the 3D view orientation and zoom-to-fit.

    view_name options: 'front', 'back', 'left', 'right', 'top', 'bottom',
    'isometric', 'trimetric', 'dimetric'
    """
    try:
        result = await _run(_sw.set_view_orientation, params.view_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_zoom_to_fit",
    annotations={"title": "Zoom to Fit", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_zoom_to_fit() -> str:
    """Zoom the view to fit the entire model on screen (Ctrl+Shift+F equivalent)."""
    try:
        result = await _run(_sw.zoom_to_fit)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_set_display_mode",
    annotations={"title": "Set Display Mode", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_set_display_mode(params: DisplayModeInput) -> str:
    """Set the 3D display mode for the active document.

    modes: 'wireframe', 'hidden_lines_visible', 'hidden_lines_removed', 'shaded', 'shaded_with_edges'
    """
    try:
        result = await _run(_sw.set_display_mode, params.mode)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_rebuild_model",
    annotations={"title": "Rebuild Model", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_rebuild_model(params: RebuildInput) -> str:
    """Force-rebuild the active model (Ctrl+B equivalent).

    Use after making multiple dimension or property changes to update the geometry.
    top_only=True is faster for assemblies (skips sub-component rebuilds).
    """
    try:
        result = await _run(_sw.rebuild_model, params.top_only)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Macro Tools ──────────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_run_macro",
    annotations={"title": "Run SolidWorks Macro", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_run_macro(params: RunMacroInput) -> str:
    """Run a saved SolidWorks macro file (.swp or .swb).

    macro_path: full path to the macro file
    module_name: VBA module name (leave empty for default)
    sub_name: subroutine to run (default: 'main')

    Returns success/failure status.
    """
    try:
        result = await _run(_sw.run_macro, params.macro_path, params.module_name, params.sub_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_run_vba_code",
    annotations={"title": "Run VBA Code", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_run_vba_code(params: RunVBAInput) -> str:
    """Execute arbitrary VBA code against SolidWorks.

    Writes the code to a temporary macro file and runs it.
    If no Sub declaration is present, code is auto-wrapped in Sub main().

    Full SolidWorks API access via swApp and Part/Assembly/Drawing objects.

    Example:
        vba_code = '''
        Dim swApp As Object
        Dim swDoc As Object
        Set swApp = Application.SldWorks
        Set swDoc = swApp.ActiveDoc
        MsgBox "Document: " & swDoc.GetTitle()
        '''
    """
    try:
        result = await _run(_sw.run_vba_code, params.vba_code)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_start_macro_recording",
    annotations={"title": "Start Macro Recording", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_start_macro_recording(output_path: str) -> str:
    """Start recording a SolidWorks macro to capture manual operations.

    output_path: where to save the .swb file (e.g. 'C:/macros/my_macro.swb')
    Stop recording with sw_stop_macro_recording.
    """
    try:
        result = await _run(_sw.start_macro_recording, output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_stop_macro_recording",
    annotations={"title": "Stop Macro Recording", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_stop_macro_recording() -> str:
    """Stop recording a SolidWorks macro and save the file.

    The recorded macro can be reviewed with a text editor and run later with sw_run_macro.
    """
    try:
        result = await _run(_sw.stop_macro_recording)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Reference Geometry Tools ─────────────────────────────────────────────────

@mcp.tool(
    name="sw_create_reference_plane",
    annotations={"title": "Create Reference Plane", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_reference_plane(params: RefPlaneInput) -> str:
    """Create a reference plane offset from an existing plane.

    reference_plane: base plane to offset from (e.g. 'Front Plane', 'Top Plane')
    offset_distance: offset in mm (positive = normal direction)

    Returns the new plane feature name.
    """
    try:
        result = await _run(_sw.create_reference_plane, "Plane", params.offset_distance, params.reference_plane)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_create_reference_axis",
    annotations={"title": "Create Reference Axis", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_reference_axis(params: RefAxisInput) -> str:
    """Create a reference axis from selected geometry.

    Select the required entities first with sw_select_entity.

    axis_type options:
    - 'two_planes': intersection of two selected planes
    - 'cylinder': axis of a selected cylindrical/conical face
    - 'two_points': axis through two selected points/vertices
    - 'point_on_face': normal to face at a selected point

    Returns the new axis feature name.
    """
    try:
        result = await _run(_sw.create_reference_axis, params.axis_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_create_reference_point",
    annotations={"title": "Create Reference Point", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_create_reference_point(params: RefPointInput) -> str:
    """Create a reference point at the specified XYZ coordinates (in mm).

    Returns the new point feature name.
    """
    try:
        result = await _run(_sw.create_reference_point, params.x, params.y, params.z)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Surface Modeling Tools ───────────────────────────────────────────────────

@mcp.tool(
    name="sw_thicken_surface",
    annotations={"title": "Thicken Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_thicken_surface(params: ThickenInput) -> str:
    """Thicken a selected surface body into a solid.

    Select a surface body with sw_select_entity (type='BODY') first.
    thickness: amount to thicken in mm.
    both_sides=True thickens symmetrically.

    Returns the new solid feature name.
    """
    try:
        result = await _run(_sw.thicken_surface, params.thickness, params.both_sides)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_knit_surfaces",
    annotations={"title": "Knit Surfaces", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}
)
async def sw_knit_surfaces() -> str:
    """Knit selected surface bodies into a single surface or closed solid.

    Select all surface bodies to knit with sw_select_entity (append=True for multiple).
    If the result is a closed volume, a solid body is automatically created.
    """
    try:
        result = await _run(_sw.knit_surfaces)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Sheet Metal Tools ────────────────────────────────────────────────────────

@mcp.tool(
    name="sw_flatten_sheet_metal",
    annotations={"title": "Flatten Sheet Metal Part", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_flatten_sheet_metal() -> str:
    """Toggle the flat pattern view of the active sheet metal part.

    Shows the unbent, flat state of the sheet metal part.
    Use sw_get_flat_pattern_info to get flat dimensions.
    """
    try:
        result = await _run(_sw.flatten_sheet_metal)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_get_flat_pattern_info",
    annotations={"title": "Get Flat Pattern Info", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_get_flat_pattern_info() -> str:
    """Get dimensions of the flat pattern for the active sheet metal part.

    Returns flat width, height, and material thickness in mm.
    Activate the flat pattern view first with sw_flatten_sheet_metal.
    """
    try:
        result = await _run(_sw.get_flat_pattern_info)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(
    name="sw_export_flat_pattern",
    annotations={"title": "Export Flat Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def sw_export_flat_pattern(params: FlatPatternExportInput) -> str:
    """Export the flat pattern of the active sheet metal part.

    output_path: target file path (.dxf, .dwg, or .pdf)
    The flat pattern must be active (use sw_flatten_sheet_metal first).
    """
    try:
        result = await _run(_sw.export_flat_pattern, params.output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ═══════════════════════════════════════════════════════════════════════════════
# ─── Expanded Toolset: Materials, Bodies, Curves, Sheet Metal, Weldments,
# ─── Mold, Drawing Extended, Assembly Extended, Display, Motion, Simulation,
# ─── Global Variables, Display States, Sensors, File Mgmt, Surface Extended
# ═══════════════════════════════════════════════════════════════════════════════


# ─── Material & Appearance ────────────────────────────────────────────────────

class SetMaterialInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    material_name: str = Field(..., description="Material name as it appears in SolidWorks material database (e.g. 'AISI 1020', 'Aluminum 6061', 'ABS')")
    database_name: str = Field(default="SOLIDWORKS Materials", description="Material database name")


class AppearanceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    r: int = Field(..., ge=0, le=255, description="Red (0-255)")
    g: int = Field(..., ge=0, le=255, description="Green (0-255)")
    b: int = Field(..., ge=0, le=255, description="Blue (0-255)")
    transparency: float = Field(default=0.0, ge=0.0, le=1.0, description="Transparency (0=opaque, 1=invisible)")


class MaterialListInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    database_name: str = Field(default="SOLIDWORKS Materials", description="Material database to list")


@mcp.tool(name="sw_set_material", annotations={"title": "Set Material", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_set_material(params: SetMaterialInput) -> str:
    """Apply a material to the active part. Common names: 'AISI 1020', 'Aluminum 6061-T6', 'ABS', 'PLA', 'Brass', 'Copper'."""
    try:
        result = await _run(_sw.set_material, params.material_name, params.database_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_get_material", annotations={"title": "Get Material", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_get_material() -> str:
    """Read the current material assigned to the active part."""
    try:
        result = await _run(_sw.get_material)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_set_appearance_color", annotations={"title": "Set Appearance Color", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_set_appearance_color(params: AppearanceInput) -> str:
    """Set RGB color and transparency on the active part's solid bodies."""
    try:
        result = await _run(_sw.set_appearance_color, params.r, params.g, params.b, params.transparency)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_list_materials", annotations={"title": "List Material Databases", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_list_materials(params: MaterialListInput) -> str:
    """List available material databases."""
    try:
        result = await _run(_sw.list_materials, params.database_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Body Operations ──────────────────────────────────────────────────────────

class ListBodiesInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    body_type: str = Field(default="solid", description="Body type: 'solid' or 'surface'")


class CombineBodiesInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    operation: str = Field(default="add", description="Boolean op: 'add' (union), 'subtract', or 'common' (intersect)")


class MoveBodyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, protected_namespaces=())
    dx: float = Field(default=0.0, description="X translation in mm")
    dy: float = Field(default=0.0, description="Y translation in mm")
    dz: float = Field(default=0.0, description="Z translation in mm")
    rx: float = Field(default=0.0, description="X rotation in degrees")
    ry: float = Field(default=0.0, description="Y rotation in degrees")
    rz: float = Field(default=0.0, description="Z rotation in degrees")
    make_copy: bool = Field(default=False, alias="copy", description="Copy instead of move")


class ScaleBodyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    factor: float = Field(..., gt=0, description="Uniform scale factor (e.g. 2.0 = double size)")
    uniform: bool = Field(default=True, description="Uniform scaling (same in x/y/z)")


@mcp.tool(name="sw_list_bodies", annotations={"title": "List Bodies", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_list_bodies(params: ListBodiesInput) -> str:
    """List solid or surface bodies in the active part."""
    try:
        result = await _run(_sw.list_bodies, params.body_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_combine_bodies", annotations={"title": "Combine Bodies", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_combine_bodies(params: CombineBodiesInput) -> str:
    """Combine selected bodies via boolean operation (union/subtract/intersect)."""
    try:
        result = await _run(_sw.combine_bodies, params.operation)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_split_body", annotations={"title": "Split Body", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_split_body() -> str:
    """Split the selected body using the selected trim tool (sketch/surface/plane)."""
    try:
        result = await _run(_sw.split_body)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_move_body", annotations={"title": "Move/Copy Body", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_move_body(params: MoveBodyInput) -> str:
    """Move or copy the selected body by translation (mm) and rotation (deg)."""
    try:
        result = await _run(_sw.move_body, params.dx, params.dy, params.dz, params.rx, params.ry, params.rz, params.make_copy)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_delete_body", annotations={"title": "Delete Body", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_delete_body() -> str:
    """Delete the selected body."""
    try:
        result = await _run(_sw.delete_body)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_scale_body", annotations={"title": "Scale Body", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_scale_body(params: ScaleBodyInput) -> str:
    """Uniformly scale the selected body by a factor."""
    try:
        result = await _run(_sw.scale_body, params.factor, params.uniform)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Curves & Helix ───────────────────────────────────────────────────────────

class HelixInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    pitch: float = Field(..., gt=0, description="Pitch in mm")
    revolutions: float = Field(..., gt=0, description="Number of revolutions")
    start_angle: float = Field(default=0.0, description="Start angle in degrees")
    clockwise: bool = Field(default=False, description="Clockwise winding")
    taper_angle: float = Field(default=0.0, description="Taper angle in degrees (0 = straight helix)")


class SplitLineInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    projection_type: str = Field(default="sketch", description="'sketch', 'projected', or 'intersection'")


@mcp.tool(name="sw_create_helix", annotations={"title": "Create Helix", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_create_helix(params: HelixInput) -> str:
    """Create a helix from a selected sketch circle (used as a sweep path for threads, springs)."""
    try:
        result = await _run(_sw.create_helix, params.pitch, params.revolutions, params.start_angle, params.clockwise, params.taper_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_start_3d_sketch", annotations={"title": "Start 3D Sketch", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_start_3d_sketch() -> str:
    """Start a 3D sketch (for routing pipes, wires, complex paths)."""
    try:
        result = await _run(_sw.start_3d_sketch)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_split_line", annotations={"title": "Split Line", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_split_line(params: SplitLineInput) -> str:
    """Split a face into sub-faces using a sketch (for selective face operations)."""
    try:
        result = await _run(_sw.split_line, params.projection_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_projected_curve", annotations={"title": "Projected Curve", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_projected_curve() -> str:
    """Project a sketch onto a face to create a 3D curve."""
    try:
        result = await _run(_sw.projected_curve)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_composite_curve", annotations={"title": "Composite Curve", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_composite_curve() -> str:
    """Combine selected edges and sketch entities into a single composite curve."""
    try:
        result = await _run(_sw.composite_curve)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Advanced Sketch ──────────────────────────────────────────────────────────

class SketchFilletInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    radius: float = Field(..., gt=0, description="Fillet radius in mm")
    trim_segments: bool = Field(default=True, description="Trim original segments")


class SketchChamferInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    distance: float = Field(..., gt=0, description="Chamfer distance in mm")
    equal: bool = Field(default=True, description="Equal distance chamfer (vs. distance+angle)")
    angle: float = Field(default=45.0, description="Angle in degrees (used if equal=False)")


class SketchTextInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    text: str = Field(..., description="Text string to insert")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")
    height: float = Field(default=5.0, gt=0, description="Text height in mm")


class SketchConstructionInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    on: bool = Field(default=True, description="True = make construction, False = restore solid")


@mcp.tool(name="sw_sketch_convert_entities", annotations={"title": "Convert Entities", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_convert_entities() -> str:
    """Convert selected face edges into sketch entities (in the active sketch)."""
    try:
        result = await _run(_sw.sketch_convert_entities)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_trim", annotations={"title": "Sketch Trim", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_trim() -> str:
    """Trim selected sketch entity (power trim mode)."""
    try:
        result = await _run(_sw.sketch_trim)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_extend", annotations={"title": "Sketch Extend", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_extend() -> str:
    """Extend selected sketch entity until next boundary."""
    try:
        result = await _run(_sw.sketch_extend)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_construction", annotations={"title": "Toggle Construction Geometry", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_sketch_construction(params: SketchConstructionInput) -> str:
    """Toggle selected sketch entities to construction (dashed) geometry."""
    try:
        result = await _run(_sw.sketch_construction, params.on)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_fillet", annotations={"title": "Sketch Fillet", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_fillet(params: SketchFilletInput) -> str:
    """Fillet (round) the corner between two selected sketch lines."""
    try:
        result = await _run(_sw.sketch_fillet, params.radius, params.trim_segments)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_chamfer", annotations={"title": "Sketch Chamfer", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_chamfer(params: SketchChamferInput) -> str:
    """Chamfer (bevel) the corner between two selected sketch lines."""
    try:
        result = await _run(_sw.sketch_chamfer, params.distance, params.equal, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketch_text", annotations={"title": "Sketch Text", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketch_text(params: SketchTextInput) -> str:
    """Insert text in the active sketch (use as profile for extrudes/cuts to engrave/emboss)."""
    try:
        result = await _run(_sw.sketch_text, params.text, params.x, params.y, params.height)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Surface Modeling Extended ────────────────────────────────────────────────

class ExtrudedSurfaceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    depth: float = Field(..., gt=0, description="Extrusion depth in mm")
    flip: bool = Field(default=False, description="Flip extrusion direction")
    draft_angle: float = Field(default=0.0, description="Draft angle in degrees")


class RevolvedSurfaceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    angle: float = Field(default=360.0, gt=0, description="Revolution angle in degrees")


class OffsetSurfaceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    distance: float = Field(..., gt=0, description="Offset distance in mm")
    flip: bool = Field(default=False, description="Flip offset direction")


class ExtendSurfaceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    distance: float = Field(..., gt=0, description="Extension distance in mm")
    extend_type: str = Field(default="distance", description="'distance', 'to_point', or 'to_surface'")


class DeleteFaceInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    repair: bool = Field(default=False, description="True = patch the hole, False = leave open")


@mcp.tool(name="sw_planar_surface", annotations={"title": "Planar Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_planar_surface() -> str:
    """Create a planar surface from selected closed sketch profile or edges."""
    try:
        result = await _run(_sw.planar_surface)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_extruded_surface", annotations={"title": "Extruded Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_extruded_surface(params: ExtrudedSurfaceInput) -> str:
    """Create an extruded surface from the active sketch."""
    try:
        result = await _run(_sw.extruded_surface, params.depth, params.flip, params.draft_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_revolved_surface", annotations={"title": "Revolved Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_revolved_surface(params: RevolvedSurfaceInput) -> str:
    """Create a revolved surface from the active sketch profile + axis."""
    try:
        result = await _run(_sw.revolved_surface, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_offset_surface", annotations={"title": "Offset Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_offset_surface(params: OffsetSurfaceInput) -> str:
    """Offset a selected face/surface by a distance."""
    try:
        result = await _run(_sw.offset_surface, params.distance, params.flip)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_trim_surface", annotations={"title": "Trim Surface", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_trim_surface() -> str:
    """Trim a surface using selected trim tool."""
    try:
        result = await _run(_sw.trim_surface)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_extend_surface", annotations={"title": "Extend Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_extend_surface(params: ExtendSurfaceInput) -> str:
    """Extend a surface by a distance, up to a point, or up to another surface."""
    try:
        result = await _run(_sw.extend_surface, params.distance, params.extend_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_delete_face", annotations={"title": "Delete Face", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_delete_face(params: DeleteFaceInput) -> str:
    """Delete selected face(s). Optionally patch the resulting hole."""
    try:
        result = await _run(_sw.delete_face, params.repair)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_filled_surface", annotations={"title": "Filled Surface", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_filled_surface() -> str:
    """Fill a closed boundary of edges with a new surface."""
    try:
        result = await _run(_sw.filled_surface)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Sheet Metal Complete Suite ───────────────────────────────────────────────

class BaseFlangeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., gt=0, description="Sheet thickness in mm")
    bend_radius: float = Field(default=1.0, gt=0, description="Default bend radius in mm")
    k_factor: float = Field(default=0.5, gt=0, lt=1, description="K-factor (0.3-0.5 typical)")


class EdgeFlangeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    edge_name: str = Field(..., description="Name of sheet-metal edge to flange (e.g. 'Edge<1>')")
    flange_length: float = Field(..., gt=0, description="Flange length in mm")
    angle: float = Field(default=90.0, description="Flange angle in degrees")


class MiterFlangeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    length: float = Field(..., gt=0, description="Flange length in mm")


class SketchedBendInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    angle: float = Field(default=90.0, description="Bend angle in degrees")
    position: str = Field(default="bend_centerline", description="'bend_centerline', 'material_inside', 'material_outside', 'bend_outside'")


class JogInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    jog_distance: float = Field(..., gt=0, description="Jog distance in mm")
    jog_angle: float = Field(default=90.0, description="Jog angle in degrees")


class HemInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    hem_type: str = Field(default="closed", description="'closed', 'open', 'teardrop', or 'rolled'")
    length: float = Field(default=5.0, gt=0, description="Hem length in mm")
    gap: float = Field(default=0.0, ge=0, description="Hem gap in mm (open hem)")


class UnfoldInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    all_bends: bool = Field(default=True, description="Unfold all bends (True) or selected only (False)")


@mcp.tool(name="sw_base_flange", annotations={"title": "Base Flange (Sheet Metal)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_base_flange(params: BaseFlangeInput) -> str:
    """Convert the active sketch into a sheet-metal base flange (starts a sheet metal part)."""
    try:
        result = await _run(_sw.base_flange, params.thickness, params.bend_radius, params.k_factor)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_edge_flange", annotations={"title": "Edge Flange (Sheet Metal)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_edge_flange(params: EdgeFlangeInput) -> str:
    """Add an edge flange to a selected sheet-metal edge."""
    try:
        result = await _run(_sw.edge_flange, params.edge_name, params.flange_length, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_miter_flange", annotations={"title": "Miter Flange", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_miter_flange(params: MiterFlangeInput) -> str:
    """Add a miter flange following the active sketch profile around multiple edges."""
    try:
        result = await _run(_sw.miter_flange, params.length)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_sketched_bend", annotations={"title": "Sketched Bend", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_sketched_bend(params: SketchedBendInput) -> str:
    """Add a sketched bend along an active sketch line on a flat sheet metal face."""
    try:
        result = await _run(_sw.sketched_bend, params.angle, params.position)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_jog", annotations={"title": "Jog Bend (Sheet Metal)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_jog(params: JogInput) -> str:
    """Add a jog bend at the active sketch line on a sheet metal face."""
    try:
        result = await _run(_sw.jog, params.jog_distance, params.jog_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_hem", annotations={"title": "Hem (Sheet Metal)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_hem(params: HemInput) -> str:
    """Add a hem to a selected sheet-metal edge."""
    try:
        result = await _run(_sw.hem, params.hem_type, params.length, params.gap)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_unfold", annotations={"title": "Unfold Bends", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_unfold(params: UnfoldInput) -> str:
    """Unfold sheet-metal bends to flatten (use with subsequent features then sw_fold)."""
    try:
        result = await _run(_sw.unfold, params.all_bends)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_fold", annotations={"title": "Fold Back Bends", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_fold() -> str:
    """Fold previously-unfolded bends back to formed state."""
    try:
        result = await _run(_sw.fold)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Weldments ────────────────────────────────────────────────────────────────

class StructuralMemberInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    standard: str = Field(..., description="Weldment standard (e.g. 'iso', 'ansi inch', 'ansi metric')")
    configuration: str = Field(..., description="Profile configuration (e.g. 'square tube', 'pipe', 'angle iron')")
    size: str = Field(..., description="Profile size (e.g. '40 x 40 x 4', '25 x 25 x 3')")


class GussetInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    profile_distance1: float = Field(..., gt=0, description="Distance along face 1 in mm")
    profile_distance2: float = Field(..., gt=0, description="Distance along face 2 in mm")
    thickness: float = Field(..., gt=0, description="Gusset thickness in mm")


class EndCapInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., gt=0, description="End cap thickness in mm")
    offset: float = Field(default=0.0, description="Offset from face in mm")


@mcp.tool(name="sw_insert_weldment", annotations={"title": "Insert Weldment", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_insert_weldment() -> str:
    """Insert the weldment feature into the active part (enables weldment tools)."""
    try:
        result = await _run(_sw.insert_weldment)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_structural_member", annotations={"title": "Add Structural Member", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_structural_member(params: StructuralMemberInput) -> str:
    """Add a structural member along selected sketch path(s) using a weldment profile."""
    try:
        result = await _run(_sw.add_structural_member, params.standard, params.configuration, params.size)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_gusset", annotations={"title": "Add Gusset", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_gusset(params: GussetInput) -> str:
    """Add a triangular gusset between two selected structural member faces."""
    try:
        result = await _run(_sw.add_gusset, params.profile_distance1, params.profile_distance2, params.thickness)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_end_cap", annotations={"title": "Add End Cap", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_end_cap(params: EndCapInput) -> str:
    """Add an end cap to a selected structural member end face."""
    try:
        result = await _run(_sw.add_end_cap, params.thickness, params.offset)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_trim_weldment", annotations={"title": "Trim Weldment", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_trim_weldment() -> str:
    """Trim selected weldment members to a planar or solid trim tool."""
    try:
        result = await _run(_sw.trim_weldment)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Mold Tools ───────────────────────────────────────────────────────────────

class PartingLineInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    draft_angle: float = Field(default=1.0, gt=0, description="Draft angle in degrees")


class ToolingSplitInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    depth_above: float = Field(..., gt=0, description="Depth above parting in mm")
    depth_below: float = Field(..., gt=0, description="Depth below parting in mm")


@mcp.tool(name="sw_parting_line", annotations={"title": "Parting Line (Mold)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_parting_line(params: PartingLineInput) -> str:
    """Create a parting line for mold design (used to separate core from cavity)."""
    try:
        result = await _run(_sw.parting_line, params.draft_angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_shut_off_surface", annotations={"title": "Shut-off Surface (Mold)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_shut_off_surface() -> str:
    """Automatically create shut-off surfaces over holes for mold separation."""
    try:
        result = await _run(_sw.shut_off_surface)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_tooling_split", annotations={"title": "Tooling Split (Mold)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_tooling_split(params: ToolingSplitInput) -> str:
    """Split a mold block into core and cavity using the parting line/surfaces."""
    try:
        result = await _run(_sw.tooling_split, params.depth_above, params.depth_below)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Drawing Extended ─────────────────────────────────────────────────────────

class ProjectedViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    parent_view_name: str = Field(..., description="Parent view name (e.g. 'Drawing View1')")
    direction: str = Field(default="right", description="Projection direction: 'left', 'right', 'up', 'down'")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")


class AuxiliaryViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    edge_name: str = Field(..., description="Edge name to project normal to")
    x: float = Field(default=0.2, description="X position in mm")
    y: float = Field(default=0.2, description="Y position in mm")


class DetailViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(..., description="Center X in mm (drawing space)")
    y: float = Field(..., description="Center Y in mm")
    radius: float = Field(..., gt=0, description="Detail circle radius in mm")
    scale: float = Field(default=2.0, gt=0, description="Scale factor")
    label: str = Field(default="A", description="Detail label")


class BrokenViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    gap: float = Field(default=10.0, gt=0, description="Break gap in mm")
    break_type: str = Field(default="straight", description="'straight', 'curved', 'zigzag', 'small_zigzag'")


class ViewNameInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    view_name: str = Field(..., description="Drawing view name to operate on")


class BalloonInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(..., description="X position in mm")
    y: float = Field(..., description="Y position in mm")
    item_number: Optional[int] = Field(default=None, description="Item number text (defaults to auto)")
    style: str = Field(default="circular", description="'circular', 'triangle', 'hexagon', 'square', 'pentagon'")


class SurfaceFinishInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    symbol_type: str = Field(default="machining", description="'basic', 'machining', 'no_machining'")
    roughness: str = Field(default="3.2", description="Roughness value (e.g. '3.2', '0.8', 'Ra 1.6')")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")


class GtolInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    gtol_text: str = Field(..., description="GD&T frame text (e.g. 'Position 0.1 A B C')")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")


class WeldSymbolInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    text: str = Field(default="", description="Weld text")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")


class DatumFeatureInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    label: str = Field(default="A", description="Datum label letter")
    x: float = Field(default=0.0, description="X position in mm")
    y: float = Field(default=0.0, description="Y position in mm")


class HoleTableInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(default=50.0, description="Table position X in mm")
    y: float = Field(default=50.0, description="Table position Y in mm")
    origin_x: float = Field(default=0.0, description="Hole origin X in mm")
    origin_y: float = Field(default=0.0, description="Hole origin Y in mm")


class RevisionTableInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x: float = Field(default=250.0, description="Table position X in mm")
    y: float = Field(default=250.0, description="Table position Y in mm")


@mcp.tool(name="sw_add_projected_view", annotations={"title": "Add Projected View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_projected_view(params: ProjectedViewInput) -> str:
    """Add a projected view derived from an existing drawing view."""
    try:
        result = await _run(_sw.add_projected_view, params.parent_view_name, params.direction, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_auxiliary_view", annotations={"title": "Add Auxiliary View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_auxiliary_view(params: AuxiliaryViewInput) -> str:
    """Add an auxiliary view normal to a selected edge."""
    try:
        result = await _run(_sw.add_auxiliary_view, params.edge_name, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_detail_view", annotations={"title": "Add Detail View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_detail_view(params: DetailViewInput) -> str:
    """Add a detail view of a region at the given (x,y) with the given radius."""
    try:
        result = await _run(_sw.add_detail_view, params.x, params.y, params.radius, params.scale, params.label)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_broken_view", annotations={"title": "Add Broken View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_broken_view(params: BrokenViewInput) -> str:
    """Compress the current drawing view with a break."""
    try:
        result = await _run(_sw.add_broken_view, params.gap, params.break_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_centerline", annotations={"title": "Add Centerlines", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_add_centerline(params: ViewNameInput) -> str:
    """Add automatic centerlines to a drawing view."""
    try:
        result = await _run(_sw.add_centerline, params.view_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_center_mark", annotations={"title": "Add Center Marks", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_add_center_mark(params: ViewNameInput) -> str:
    """Add automatic center marks to circular features in a drawing view."""
    try:
        result = await _run(_sw.add_center_mark, params.view_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_balloon", annotations={"title": "Add Balloon", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_balloon(params: BalloonInput) -> str:
    """Add a BOM balloon at the specified position in the drawing."""
    try:
        result = await _run(_sw.add_balloon, params.x, params.y, params.item_number, params.style)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_surface_finish", annotations={"title": "Add Surface Finish", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_surface_finish(params: SurfaceFinishInput) -> str:
    """Add a surface finish symbol to the drawing."""
    try:
        result = await _run(_sw.add_surface_finish, params.symbol_type, params.roughness, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_geometric_tolerance", annotations={"title": "Add GD&T", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_geometric_tolerance(params: GtolInput) -> str:
    """Add a geometric tolerance (GD&T) frame to the drawing."""
    try:
        result = await _run(_sw.add_geometric_tolerance, params.gtol_text, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_weld_symbol", annotations={"title": "Add Weld Symbol", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_weld_symbol(params: WeldSymbolInput) -> str:
    """Add a welding symbol annotation to the drawing."""
    try:
        result = await _run(_sw.add_weld_symbol, params.text, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_datum_feature", annotations={"title": "Add Datum Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_datum_feature(params: DatumFeatureInput) -> str:
    """Add a datum feature symbol (A, B, C...) to the drawing."""
    try:
        result = await _run(_sw.add_datum_feature, params.label, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_hole_table", annotations={"title": "Add Hole Table", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_hole_table(params: HoleTableInput) -> str:
    """Add a hole table to the drawing."""
    try:
        result = await _run(_sw.add_hole_table, params.x, params.y, params.origin_x, params.origin_y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_add_revision_table", annotations={"title": "Add Revision Table", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_add_revision_table(params: RevisionTableInput) -> str:
    """Add a revision table to the drawing sheet."""
    try:
        result = await _run(_sw.add_revision_table, params.x, params.y)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_export_drawing_pdf", annotations={"title": "Export Drawing to PDF", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_export_drawing_pdf(params: ExportInput) -> str:
    """Export the active drawing to PDF (output_path must end in .pdf)."""
    try:
        result = await _run(_sw.export_drawing_pdf, params.output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_export_drawing_dwg", annotations={"title": "Export Drawing to DWG/DXF", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_export_drawing_dwg(params: ExportInput) -> str:
    """Export the active drawing to DWG or DXF (output_path determines format)."""
    try:
        result = await _run(_sw.export_drawing_dwg, params.output_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Assembly Extended ────────────────────────────────────────────────────────

class ReplaceComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    component_name: str = Field(..., description="Component name (e.g. 'shaft-1@assembly')")
    new_file_path: str = Field(..., description="Path to the new part/assembly file")


class PatternComponentLinearInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    count: int = Field(..., ge=2, description="Number of instances")
    spacing: float = Field(..., gt=0, description="Spacing between instances in mm")
    direction_face: str = Field(default="", description="Optional direction reference face name")


class PatternComponentCircularInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    count: int = Field(..., ge=2, description="Number of instances")
    angle: float = Field(default=360.0, description="Total angle in degrees")


class MirrorComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    mirror_plane: str = Field(..., description="Plane to mirror about")


class SuppressComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    component_name: str = Field(..., description="Component name to suppress/unsuppress")
    suppress: bool = Field(default=True, description="True = suppress, False = unsuppress")


class HideComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    component_name: str = Field(..., description="Component name to hide/show")
    hide: bool = Field(default=True, description="True = hide, False = show")


class MoveComponentInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    component_name: str = Field(..., description="Component name")
    dx: float = Field(default=0.0, description="X translation in mm")
    dy: float = Field(default=0.0, description="Y translation in mm")
    dz: float = Field(default=0.0, description="Z translation in mm")


@mcp.tool(name="sw_replace_component", annotations={"title": "Replace Component", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": True})
async def sw_replace_component(params: ReplaceComponentInput) -> str:
    """Replace a component in the active assembly with a different file."""
    try:
        result = await _run(_sw.replace_component, params.component_name, params.new_file_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_pattern_component_linear", annotations={"title": "Linear Pattern Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_pattern_component_linear(params: PatternComponentLinearInput) -> str:
    """Linearly pattern a selected component (select seed component + direction reference first)."""
    try:
        result = await _run(_sw.pattern_component_linear, params.count, params.spacing, params.direction_face)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_pattern_component_circular", annotations={"title": "Circular Pattern Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_pattern_component_circular(params: PatternComponentCircularInput) -> str:
    """Circularly pattern a selected component about an axis."""
    try:
        result = await _run(_sw.pattern_component_circular, params.count, params.angle)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_mirror_component", annotations={"title": "Mirror Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_mirror_component(params: MirrorComponentInput) -> str:
    """Mirror a selected component about a plane (select component first)."""
    try:
        result = await _run(_sw.mirror_component, params.mirror_plane)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_suppress_component", annotations={"title": "Suppress/Unsuppress Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_suppress_component(params: SuppressComponentInput) -> str:
    """Suppress or unsuppress a component in the assembly."""
    try:
        result = await _run(_sw.suppress_component, params.component_name, params.suppress)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_hide_component", annotations={"title": "Hide/Show Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_hide_component(params: HideComponentInput) -> str:
    """Hide or show a component in the assembly viewport."""
    try:
        result = await _run(_sw.hide_component, params.component_name, params.hide)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_move_component", annotations={"title": "Move Component", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_move_component(params: MoveComponentInput) -> str:
    """Move a component (free-drag style translation, subject to mate constraints)."""
    try:
        result = await _run(_sw.move_component, params.component_name, params.dx, params.dy, params.dz)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_get_assembly_tree", annotations={"title": "Get Assembly Tree", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_get_assembly_tree() -> str:
    """List all components in the active assembly with their state info."""
    try:
        result = await _run(_sw.get_assembly_tree)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Display & Visualization Extended ─────────────────────────────────────────

class SectionViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    plane: str = Field(default="Front Plane", description="Cutting plane name")
    offset: float = Field(default=0.0, description="Offset from plane in mm")


class BackgroundInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    mode: str = Field(default="plain", description="'plain', 'gradient', 'image', or 'scene'")
    color_r: int = Field(default=200, ge=0, le=255)
    color_g: int = Field(default=200, ge=0, le=255)
    color_b: int = Field(default=200, ge=0, le=255)


class PhotoViewInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: str = Field(..., description="Output image path (.png, .jpg)")
    width: int = Field(default=1920, ge=100)
    height: int = Field(default=1080, ge=100)


class ViewRotateInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    x_deg: float = Field(default=0.0)
    y_deg: float = Field(default=0.0)
    z_deg: float = Field(default=0.0)


@mcp.tool(name="sw_create_section_view", annotations={"title": "Section View (Model)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_create_section_view(params: SectionViewInput) -> str:
    """Create a live model section view through a plane (visualization only)."""
    try:
        result = await _run(_sw.create_section_view, params.plane, params.offset)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_set_background", annotations={"title": "Set Viewport Background", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_set_background(params: BackgroundInput) -> str:
    """Set the viewport background style and color."""
    try:
        result = await _run(_sw.set_background, params.mode, params.color_r, params.color_g, params.color_b)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_take_photoview_screenshot", annotations={"title": "PhotoView 360 Render", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_take_photoview_screenshot(params: PhotoViewInput) -> str:
    """Render via PhotoView 360 add-in if available, else regular screenshot."""
    try:
        result = await _run(_sw.take_photoview_screenshot, params.output_path, params.width, params.height)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_hide_all_planes", annotations={"title": "Hide All Reference Planes", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_hide_all_planes() -> str:
    """Hide all reference planes in the viewport."""
    try:
        result = await _run(_sw.hide_all_planes)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_show_all_planes", annotations={"title": "Show All Reference Planes", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_show_all_planes() -> str:
    """Show all reference planes in the viewport."""
    try:
        result = await _run(_sw.show_all_planes)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_hide_all_sketches", annotations={"title": "Hide All Sketches", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_hide_all_sketches() -> str:
    """Hide all sketches in the viewport."""
    try:
        result = await _run(_sw.hide_all_sketches)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_show_all_sketches", annotations={"title": "Show All Sketches", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_show_all_sketches() -> str:
    """Show all sketches in the viewport."""
    try:
        result = await _run(_sw.show_all_sketches)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_view_rotate", annotations={"title": "Rotate View", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_view_rotate(params: ViewRotateInput) -> str:
    """Rotate the current view by Euler angles (degrees)."""
    try:
        result = await _run(_sw.view_rotate, params.x_deg, params.y_deg, params.z_deg)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Motion Study ─────────────────────────────────────────────────────────────

class MotionStudyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    study_name: str = Field(default="Motion Study 1", description="Motion study name")


@mcp.tool(name="sw_create_motion_study", annotations={"title": "Create Motion Study", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_create_motion_study(params: MotionStudyInput) -> str:
    """Create a new motion study (requires SOLIDWORKS Motion add-in)."""
    try:
        result = await _run(_sw.create_motion_study, params.study_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Simulation ───────────────────────────────────────────────────────────────

class SimulationStudyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    study_name: str = Field(..., description="Study name")
    study_type: str = Field(default="static", description="'static', 'frequency', 'buckling', 'thermal', 'drop', 'fatigue', 'nonlinear'")


@mcp.tool(name="sw_create_simulation_study", annotations={"title": "Create Simulation Study", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_create_simulation_study(params: SimulationStudyInput) -> str:
    """Create a new SOLIDWORKS Simulation study (requires Simulation add-in)."""
    try:
        result = await _run(_sw.create_simulation_study, params.study_name, params.study_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_run_simulation", annotations={"title": "Run Simulation Study", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_run_simulation() -> str:
    """Run the active SOLIDWORKS Simulation study."""
    try:
        result = await _run(_sw.run_simulation)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Global Variables ────────────────────────────────────────────────────────

class GlobalVarInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    name: str = Field(..., description="Global variable name (without quotes)")
    value: float = Field(..., description="Numeric value")


@mcp.tool(name="sw_list_global_variables", annotations={"title": "List Global Variables", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_list_global_variables() -> str:
    """List all global variables in the active model."""
    try:
        result = await _run(_sw.list_global_variables)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_set_global_variable", annotations={"title": "Set Global Variable", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_set_global_variable(params: GlobalVarInput) -> str:
    """Set or create a global variable (used in equations like 'Diameter=2*\"Radius\"')."""
    try:
        result = await _run(_sw.set_global_variable, params.name, params.value)
        return _ok(result)
    except Exception as e:
        return _err(e)


class LinkDimensionInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    dim_name: str = Field(..., description="Dimension name (e.g. 'D1@Sketch1')")
    expr: str = Field(..., description="Expression or value (e.g. '\"Width\"*2', '50mm')")


@mcp.tool(name="sw_link_dimension_to_equation", annotations={"title": "Link Dimension to Equation", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_link_dimension_to_equation(params: LinkDimensionInput) -> str:
    """Link a dimension to an expression via an equation (drives the dimension parametrically)."""
    try:
        result = await _run(_sw.link_dimension_to_equation, params.dim_name, params.expr)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Display States ──────────────────────────────────────────────────────────

class DisplayStateInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    state_name: str = Field(..., description="Display state name")


@mcp.tool(name="sw_list_display_states", annotations={"title": "List Display States", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_list_display_states() -> str:
    """List all display states in the active configuration."""
    try:
        result = await _run(_sw.list_display_states)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_activate_display_state", annotations={"title": "Activate Display State", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_activate_display_state(params: DisplayStateInput) -> str:
    """Activate a named display state."""
    try:
        result = await _run(_sw.activate_display_state, params.state_name)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Sensors ──────────────────────────────────────────────────────────────────

@mcp.tool(name="sw_list_sensors", annotations={"title": "List Sensors", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_list_sensors() -> str:
    """List sensors on the active document with current values."""
    try:
        result = await _run(_sw.list_sensors)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── File Management Extended ────────────────────────────────────────────────

class SaveAsCopyInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    file_path: str = Field(..., description="Path to save copy to")


class CloseAllInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    save: bool = Field(default=False, description="Save modified documents before closing")


@mcp.tool(name="sw_save_all", annotations={"title": "Save All Open Documents", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_save_all() -> str:
    """Save all open SolidWorks documents that have unsaved changes."""
    try:
        result = await _run(_sw.save_all)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_close_all", annotations={"title": "Close All Documents", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": True, "openWorldHint": False})
async def sw_close_all(params: CloseAllInput) -> str:
    """Close all open SolidWorks documents."""
    try:
        result = await _run(_sw.close_all, params.save)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_save_as_copy", annotations={"title": "Save As Copy", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_save_as_copy(params: SaveAsCopyInput) -> str:
    """Save the active document as a copy without changing the current open file."""
    try:
        result = await _run(_sw.save_as_copy, params.file_path)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_get_referenced_documents", annotations={"title": "Get Referenced Documents", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_get_referenced_documents() -> str:
    """List all documents (parts, subassemblies) referenced by the active assembly/drawing."""
    try:
        result = await _run(_sw.get_referenced_documents)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Advanced Features ────────────────────────────────────────────────────────

class RibInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., gt=0, description="Rib thickness in mm")
    draft_angle: float = Field(default=0.0, ge=0, description="Draft angle in degrees")
    flip: bool = Field(default=False, description="Flip rib direction")
    two_sided: bool = Field(default=False, description="Two-sided rib (centered on sketch)")


class DomeInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    height: float = Field(..., gt=0, description="Dome height in mm")
    elliptical: bool = Field(default=False, description="Elliptical dome (vs spherical)")
    flip: bool = Field(default=False, description="Flip dome direction (concave)")


class WrapInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    thickness: float = Field(..., gt=0, description="Wrap thickness in mm")
    wrap_type: str = Field(default="emboss", description="'emboss', 'deboss', or 'scribe'")


class BoundaryBossInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    is_cut: bool = Field(default=False, description="Cut instead of boss")


@mcp.tool(name="sw_rib", annotations={"title": "Rib Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_rib(params: RibInput) -> str:
    """Add a rib feature from the active sketch (used to stiffen thin walls)."""
    try:
        result = await _run(_sw.rib, params.thickness, params.draft_angle, params.flip, params.two_sided)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_dome", annotations={"title": "Dome Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_dome(params: DomeInput) -> str:
    """Add a dome on a selected planar face."""
    try:
        result = await _run(_sw.dome, params.height, params.elliptical, params.flip)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_wrap", annotations={"title": "Wrap Feature", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_wrap(params: WrapInput) -> str:
    """Wrap a sketch around a selected curved face (emboss/deboss/scribe)."""
    try:
        result = await _run(_sw.wrap, params.thickness, params.wrap_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_intersect", annotations={"title": "Intersect Bodies", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": False})
async def sw_intersect() -> str:
    """Create intersection regions from selected bodies/surfaces."""
    try:
        result = await _run(_sw.intersect)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_boundary_boss", annotations={"title": "Boundary Boss/Cut", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_boundary_boss(params: BoundaryBossInput) -> str:
    """Create a boundary boss/cut from two or more profile curves."""
    try:
        result = await _run(_sw.boundary_boss, params.is_cut)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Patterns Extended ────────────────────────────────────────────────────────

class FillPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    spacing: float = Field(..., gt=0, description="Spacing between instances in mm")
    pattern_type: str = Field(default="square", description="'perimeter', 'square', 'rectangular', or 'circular'")


class CurveDrivenPatternInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    count: int = Field(..., ge=2, description="Number of instances")
    spacing: float = Field(..., gt=0, description="Spacing in mm")
    equal_spacing: bool = Field(default=True, description="Equal spacing along curve")


@mcp.tool(name="sw_fill_pattern", annotations={"title": "Fill Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_fill_pattern(params: FillPatternInput) -> str:
    """Pattern a seed feature to fill a target face."""
    try:
        result = await _run(_sw.fill_pattern, params.spacing, params.pattern_type)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_curve_driven_pattern", annotations={"title": "Curve Driven Pattern", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False})
async def sw_curve_driven_pattern(params: CurveDrivenPatternInput) -> str:
    """Pattern a seed feature along a curve."""
    try:
        result = await _run(_sw.curve_driven_pattern, params.count, params.spacing, params.equal_spacing)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Export Extended ──────────────────────────────────────────────────────────

class StepExportInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: str = Field(..., description="Path ending in .step or .stp")
    version: str = Field(default="AP214", description="STEP version: 'AP203', 'AP214', or 'AP242'")


class StlExportInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)
    output_path: str = Field(..., description="Path ending in .stl")
    binary: bool = Field(default=True, description="Binary STL (vs ASCII)")
    quality: str = Field(default="fine", description="'coarse', 'fine', or 'custom'")


@mcp.tool(name="sw_export_step", annotations={"title": "Export STEP (with version)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_export_step(params: StepExportInput) -> str:
    """Export the active model to STEP with explicit AP version."""
    try:
        result = await _run(_sw.export_step, params.output_path, params.version)
        return _ok(result)
    except Exception as e:
        return _err(e)


@mcp.tool(name="sw_export_stl", annotations={"title": "Export STL (with quality)", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def sw_export_stl(params: StlExportInput) -> str:
    """Export the active part to STL with binary/ASCII and quality control."""
    try:
        result = await _run(_sw.export_stl, params.output_path, params.binary, params.quality)
        return _ok(result)
    except Exception as e:
        return _err(e)


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run()
