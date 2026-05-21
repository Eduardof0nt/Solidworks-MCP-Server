# SolidWorks MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![MCP](https://img.shields.io/badge/MCP-stdio-purple.svg)](https://modelcontextprotocol.io)
[![SolidWorks 2024](https://img.shields.io/badge/SolidWorks-2018%E2%80%932024-red.svg)](https://www.solidworks.com)

A comprehensive **Model Context Protocol (MCP)** server that gives Claude (or any MCP-compatible LLM client)
full control of SolidWorks through the COM API. **187 tools** spanning every major workflow — modeling,
sketching, assemblies, drawings, sheet metal, weldments, mold tools, surfaces, simulation, motion, rendering,
materials, and more.

> ✅ **Verified on SolidWorks 2024** — handles SW 2024's dual property/method API quirks (DISPID invocation,
> property-get bypass, 23-arg `FeatureExtrusion3`, VARIANT marshaling, etc.). Should also work on SW 2018+.

---

## ✨ What you can ask Claude to do

- *"Create a 100×60×40 mm aluminum part, render it, and export to STEP"*
- *"Open this assembly, check for interferences, then mass-analyze each component"*
- *"Generate a 3-view drawing of `bracket.sldprt` with dimensions, GD&T, and a BOM, then save as PDF"*
- *"Design a parametric sheet-metal enclosure with 2 mm thickness, four edge flanges, and export flat patterns"*
- *"Build a weldment frame from 40×40×4 square tube along this sketch"*
- *"Set up a static simulation, apply this fixture and 500 N load, and run the analysis"*

---

## 📋 Prerequisites

| Requirement | Notes |
|---|---|
| **Windows 10/11** | SolidWorks is Windows-only |
| **SolidWorks 2018 or newer** | Tested and developed against SolidWorks 2024 |
| **Python 3.10+** | [python.org](https://www.python.org/downloads/) — make sure "Add to PATH" is checked |
| **An MCP client** | [Claude Desktop](https://claude.ai/download) or [Claude Code](https://docs.anthropic.com/claude-code) recommended |

---

## 🚀 Installation

### 1. Clone the repository

```bash
git clone https://github.com/Eduardof0nt/Solidworks-MCP-Server.git
cd Solidworks-MCP-Server
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

If you hit permission issues installing globally, use a virtual environment or `--user`:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Verify the server starts

```bash
python server.py
```

The server will sit waiting for MCP stdio traffic — that's correct. Hit `Ctrl+C` to stop; you'll wire it
into a client in the next step.

---

## 🔌 Configuring an MCP client

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json` (create it if it doesn't exist):

```json
{
  "mcpServers": {
    "solidworks": {
      "command": "python",
      "args": ["C:\\full\\path\\to\\Solidworks-MCP-Server\\server.py"]
    }
  }
}
```

> ⚠️ Replace `C:\\full\\path\\to\\Solidworks-MCP-Server\\server.py` with the **absolute path** on your machine.
> Use double-backslashes (`\\`) in JSON.

Restart Claude Desktop. The `solidworks` server should appear in the MCP indicator (bottom-right "🔌" icon).

### Claude Code (CLI)

```bash
claude mcp add solidworks -- python "C:\full\path\to\Solidworks-MCP-Server\server.py"
```

Or edit `~/.claude.json` (or your project's `.mcp.json`) and add the same `mcpServers` block as above.

### Other MCP clients

The server speaks **stdio MCP**. Any MCP-compatible client can connect by spawning `python server.py` as
a subprocess.

---

## ▶️ First run

1. Start SolidWorks (the server can auto-launch it, but starting it manually is faster).
2. Open your MCP client (Claude Desktop, Claude Code, etc.).
3. Ask Claude: *"Connect to SolidWorks and check the status."*

Claude should call `sw_connect` and `sw_get_status`, returning the SolidWorks version and any open documents.

## Tools Reference

### Connection
| Tool | Description |
|------|-------------|
| `sw_connect` | Connect to SolidWorks (launches if needed) |
| `sw_get_status` | Connection status + active document info |

### Files & Documents
| Tool | Description |
|------|-------------|
| `sw_open_document` | Open .sldprt, .sldasm, .slddrw, .step, .stl, etc. |
| `sw_close_document` | Close a document |
| `sw_save_document` | Save / Save As |
| `sw_create_part` | New part document |
| `sw_create_assembly` | New assembly document |
| `sw_create_drawing` | New drawing document |
| `sw_list_open_documents` | List all open documents |
| `sw_export_file` | Export to STEP, STL, DXF, PDF, IGES, etc. |
| `sw_get_model_info` | Full model info + feature tree |

### Selection
| Tool | Description |
|------|-------------|
| `sw_select_entity` | Select by name and type |
| `sw_clear_selection` | Clear selection |
| `sw_list_features` | List all FeatureManager features |

### Sketch
| Tool | Description |
|------|-------------|
| `sw_create_sketch` | Start sketch on plane/face |
| `sw_edit_sketch` | Open existing sketch |
| `sw_exit_sketch` | Exit sketch mode |
| `sw_sketch_line` | Draw line |
| `sw_sketch_circle` | Draw circle |
| `sw_sketch_rectangle` | Draw rectangle |
| `sw_sketch_arc` | Draw arc |
| `sw_sketch_ellipse` | Draw ellipse |
| `sw_sketch_polygon` | Draw regular polygon |
| `sw_sketch_spline` | Draw spline |
| `sw_sketch_centerline` | Draw centerline |
| `sw_sketch_offset` | Offset selected entities |
| `sw_sketch_linear_pattern` | Linear step-and-repeat |
| `sw_sketch_circular_pattern` | Circular step-and-repeat |
| `sw_add_sketch_dimension` | Add smart dimension |
| `sw_add_sketch_constraint` | Add geometric constraint |

### Features (Part)
| Tool | Description |
|------|-------------|
| `sw_extrude` | Extrude boss or cut |
| `sw_revolve` | Revolve boss or cut |
| `sw_fillet` | Add constant fillet |
| `sw_chamfer` | Add chamfer |
| `sw_shell` | Shell body |
| `sw_draft` | Add draft angle |
| `sw_linear_pattern` | Linear feature pattern |
| `sw_circular_pattern` | Circular feature pattern |
| `sw_mirror` | Mirror feature |
| `sw_hole_wizard` | Hole Wizard (threaded/clearance holes) |
| `sw_loft` | Loft between profiles |
| `sw_sweep` | Sweep profile along path |
| `sw_delete_feature` | Delete feature |
| `sw_suppress_feature` | Suppress feature |
| `sw_unsuppress_feature` | Unsuppress feature |

### Reference Geometry
| Tool | Description |
|------|-------------|
| `sw_create_reference_plane` | Offset reference plane |
| `sw_create_reference_axis` | Reference axis |
| `sw_create_reference_point` | Reference point |

### Surface Modeling
| Tool | Description |
|------|-------------|
| `sw_thicken_surface` | Thicken surface to solid |
| `sw_knit_surfaces` | Knit surfaces together |

### Assembly
| Tool | Description |
|------|-------------|
| `sw_insert_component` | Insert part/subassembly |
| `sw_add_mate` | Add constraint mate |
| `sw_fix_component` | Fix/float component |
| `sw_check_interference` | Interference detection |
| `sw_explode_assembly` | Exploded view |

### Drawing
| Tool | Description |
|------|-------------|
| `sw_add_drawing_view` | Add model view |
| `sw_add_section_view` | Add section view |
| `sw_add_drawing_dimension` | Add dimension |
| `sw_add_annotation` | Add text note |
| `sw_add_bom_table` | Add Bill of Materials |

### Dimensions & Parameters
| Tool | Description |
|------|-------------|
| `sw_get_dimension` | Read dimension value |
| `sw_set_dimension` | Set dimension value (parametric drive) |

### Analysis
| Tool | Description |
|------|-------------|
| `sw_get_mass_properties` | Mass, volume, center of gravity |
| `sw_get_bounding_box` | Envelope dimensions |
| `sw_measure_distance` | Distance between entities |
| `sw_check_geometry` | Geometry validity check |
| `sw_analyze_draft` | Draft angle analysis |

### Properties & Configurations
| Tool | Description |
|------|-------------|
| `sw_get_custom_properties` | Read custom properties |
| `sw_set_custom_property` | Write custom property |
| `sw_list_configurations` | List configurations |
| `sw_activate_configuration` | Switch configuration |
| `sw_add_configuration` | Add new configuration |

### Equations
| Tool | Description |
|------|-------------|
| `sw_get_equations` | List all equations |
| `sw_add_equation` | Add equation/global variable |
| `sw_delete_equation` | Delete equation |
| `sw_create_design_table` | Create design table |

### Visualization
| Tool | Description |
|------|-------------|
| `sw_capture_screenshot` | Screenshot of model view |
| `sw_set_view_orientation` | Set view (front/iso/etc.) |
| `sw_zoom_to_fit` | Zoom to fit |
| `sw_set_display_mode` | Wireframe/shaded/etc. |
| `sw_rebuild_model` | Force rebuild |

### Macros & Automation
| Tool | Description |
|------|-------------|
| `sw_run_macro` | Run .swp/.swb macro file |
| `sw_run_vba_code` | Execute VBA code directly |
| `sw_start_macro_recording` | Record macro |
| `sw_stop_macro_recording` | Stop recording |

### Sheet Metal — Basic
| Tool | Description |
|------|-------------|
| `sw_flatten_sheet_metal` | Toggle flat pattern view |
| `sw_get_flat_pattern_info` | Flat pattern dimensions |
| `sw_export_flat_pattern` | Export flat pattern DXF/PDF |

---

## Extended Toolset (additional 99 tools)

### Materials & Appearance
| Tool | Description |
|------|-------------|
| `sw_set_material` | Apply material from SolidWorks database |
| `sw_get_material` | Read current material |
| `sw_set_appearance_color` | Set RGB color and transparency |
| `sw_list_materials` | List material databases |

### Body Operations
| Tool | Description |
|------|-------------|
| `sw_list_bodies` | List solid/surface bodies |
| `sw_combine_bodies` | Union / Subtract / Intersect bodies |
| `sw_split_body` | Split body with trim tool |
| `sw_move_body` | Translate/rotate (or copy) selected body |
| `sw_delete_body` | Delete selected body |
| `sw_scale_body` | Scale body uniformly |

### Curves & 3D Sketch
| Tool | Description |
|------|-------------|
| `sw_create_helix` | Helix for threads, springs |
| `sw_start_3d_sketch` | Enter 3D sketch mode |
| `sw_split_line` | Split a face with a sketch |
| `sw_projected_curve` | Project sketch onto face |
| `sw_composite_curve` | Combine edges into single curve |

### Advanced Sketch
| Tool | Description |
|------|-------------|
| `sw_sketch_convert_entities` | Convert edges → sketch entities |
| `sw_sketch_trim` | Power-trim selected entity |
| `sw_sketch_extend` | Extend selected entity |
| `sw_sketch_construction` | Toggle construction geometry |
| `sw_sketch_fillet` | Fillet corner |
| `sw_sketch_chamfer` | Chamfer corner |
| `sw_sketch_text` | Insert text for engraving |

### Surface Modeling Extended
| Tool | Description |
|------|-------------|
| `sw_planar_surface` | Planar surface from sketch/edges |
| `sw_extruded_surface` | Extruded surface |
| `sw_revolved_surface` | Revolved surface |
| `sw_offset_surface` | Offset face/surface |
| `sw_trim_surface` | Trim surface |
| `sw_extend_surface` | Extend surface |
| `sw_delete_face` | Delete face (with optional patch) |
| `sw_filled_surface` | Fill closed edge boundary |

### Sheet Metal — Complete
| Tool | Description |
|------|-------------|
| `sw_base_flange` | Convert sketch → sheet metal base |
| `sw_edge_flange` | Add edge flange |
| `sw_miter_flange` | Add miter flange |
| `sw_sketched_bend` | Sketched bend on flat face |
| `sw_jog` | Jog bend |
| `sw_hem` | Closed/open/teardrop/rolled hem |
| `sw_unfold` | Unfold bends to flatten |
| `sw_fold` | Re-fold bends |

### Weldments
| Tool | Description |
|------|-------------|
| `sw_insert_weldment` | Enable weldment feature |
| `sw_add_structural_member` | Add structural member (tube, pipe, angle) |
| `sw_add_gusset` | Add triangular gusset |
| `sw_add_end_cap` | Add end cap on member face |
| `sw_trim_weldment` | Trim members to trim tool |

### Mold Tools
| Tool | Description |
|------|-------------|
| `sw_parting_line` | Parting line for mold |
| `sw_shut_off_surface` | Auto shut-off surfaces |
| `sw_tooling_split` | Split mold block to core/cavity |

### Drawing — Extended
| Tool | Description |
|------|-------------|
| `sw_add_projected_view` | Projected view from parent |
| `sw_add_auxiliary_view` | Auxiliary view normal to edge |
| `sw_add_detail_view` | Detail view of region |
| `sw_add_broken_view` | Broken view (compressed) |
| `sw_add_centerline` | Auto centerlines |
| `sw_add_center_mark` | Auto center marks |
| `sw_add_balloon` | BOM balloon |
| `sw_add_surface_finish` | Surface finish symbol |
| `sw_add_geometric_tolerance` | GD&T frame |
| `sw_add_weld_symbol` | Welding symbol |
| `sw_add_datum_feature` | Datum feature symbol |
| `sw_add_hole_table` | Hole table |
| `sw_add_revision_table` | Revision table |
| `sw_export_drawing_pdf` | Export drawing → PDF |
| `sw_export_drawing_dwg` | Export drawing → DWG/DXF |

### Assembly — Extended
| Tool | Description |
|------|-------------|
| `sw_replace_component` | Swap component with different file |
| `sw_pattern_component_linear` | Linear component pattern |
| `sw_pattern_component_circular` | Circular component pattern |
| `sw_mirror_component` | Mirror component about plane |
| `sw_suppress_component` | Suppress/unsuppress |
| `sw_hide_component` | Hide/show component |
| `sw_move_component` | Translate component |
| `sw_get_assembly_tree` | List all components with state |

### Display & Visualization
| Tool | Description |
|------|-------------|
| `sw_create_section_view` | Live model section view |
| `sw_set_background` | Viewport background style |
| `sw_take_photoview_screenshot` | PhotoView 360 render |
| `sw_hide_all_planes` | Hide all reference planes |
| `sw_show_all_planes` | Show all reference planes |
| `sw_hide_all_sketches` | Hide all sketches |
| `sw_show_all_sketches` | Show all sketches |
| `sw_view_rotate` | Rotate view by Euler angles |

### Motion & Simulation
| Tool | Description |
|------|-------------|
| `sw_create_motion_study` | Create motion study |
| `sw_create_simulation_study` | Create simulation study (static, frequency, buckling, thermal, drop, fatigue) |
| `sw_run_simulation` | Run active simulation |

### Global Variables & Parametrics
| Tool | Description |
|------|-------------|
| `sw_list_global_variables` | List globals |
| `sw_set_global_variable` | Set/create a global variable |
| `sw_link_dimension_to_equation` | Link dimension to expression |

### Display States & Sensors
| Tool | Description |
|------|-------------|
| `sw_list_display_states` | List display states |
| `sw_activate_display_state` | Activate display state |
| `sw_list_sensors` | List sensors with values |

### File Management Extended
| Tool | Description |
|------|-------------|
| `sw_save_all` | Save all open documents |
| `sw_close_all` | Close all open documents |
| `sw_save_as_copy` | Save as copy (keeps current open) |
| `sw_get_referenced_documents` | List referenced files |

### Advanced Features
| Tool | Description |
|------|-------------|
| `sw_rib` | Rib feature from sketch |
| `sw_dome` | Dome on planar face |
| `sw_wrap` | Wrap sketch on curved face (emboss/deboss/scribe) |
| `sw_intersect` | Intersect bodies/surfaces |
| `sw_boundary_boss` | Boundary boss/cut |
| `sw_fill_pattern` | Fill face with seed pattern |
| `sw_curve_driven_pattern` | Pattern along curve |

### Export Extended
| Tool | Description |
|------|-------------|
| `sw_export_step` | STEP with AP203/AP214/AP242 |
| `sw_export_stl` | STL with binary/ASCII + quality |

---

## Example Workflows

### Create a Simple Box
```
1. sw_create_part()
2. sw_create_sketch(plane='Front Plane')
3. sw_sketch_rectangle(x1=0, y1=0, x2=100, y2=60)
4. sw_exit_sketch()
5. sw_extrude(depth=40)
6. sw_save_document(file_path='C:/parts/box.sldprt')
```

### Parametric Cylinder
```
1. sw_create_part()
2. sw_create_sketch(plane='Front Plane')
3. sw_sketch_circle(cx=0, cy=0, radius=25)
4. sw_exit_sketch()
5. sw_extrude(depth=100)
6. sw_add_equation('"Height" = 100mm')
7. sw_add_equation('"D1@Boss-Extrude1" = "Height"')
```

### Assembly with Mates
```
1. sw_create_assembly()
2. sw_insert_component(file_path='C:/parts/base.sldprt', fixed=True)
3. sw_insert_component(file_path='C:/parts/shaft.sldprt', z=50)
4. sw_add_mate(mate_type='CONCENTRIC', entity1='Face<1>@shaft-1', entity2='Face<1>@base-1')
5. sw_add_mate(mate_type='COINCIDENT', entity1='Face<2>@shaft-1', entity2='Face<2>@base-1')
```

### Sheet Metal Bracket (with bend)
```
1. sw_create_part()
2. sw_create_sketch(plane='Top Plane')
3. sw_sketch_rectangle(x1=0, y1=0, x2=200, y2=100)
4. sw_exit_sketch()
5. sw_base_flange(thickness=2.0, bend_radius=1.5)
6. sw_create_sketch(plane='<top face>')
7. sw_sketch_line(...)  # bend line
8. sw_exit_sketch()
9. sw_sketched_bend(angle=90)
10. sw_flatten_sheet_metal()
11. sw_export_flat_pattern('C:/parts/bracket_flat.dxf')
```

### Aluminum Part with Custom Color and Mass Analysis
```
1. sw_create_part()
2. sw_create_sketch('Front Plane')
3. sw_sketch_circle(0, 0, 25)
4. sw_exit_sketch()
5. sw_extrude(depth=100)
6. sw_set_material('Aluminum 6061-T6')
7. sw_set_appearance_color(r=180, g=180, b=200, transparency=0.0)
8. sw_get_mass_properties()  # returns mass, volume, COG
```

### Drawing with Multiple Views + GD&T
```
1. sw_create_drawing()
2. sw_add_drawing_view(view_type='front', x=100, y=200, scale=1.0)
3. sw_add_projected_view(parent_view_name='Drawing View1', direction='right')
4. sw_add_detail_view(x=80, y=150, radius=15, scale=2.0, label='A')
5. sw_add_centerline(view_name='Drawing View1')
6. sw_add_geometric_tolerance(gtol_text='Position 0.1 A B C', x=50, y=80)
7. sw_add_surface_finish(symbol_type='machining', roughness='3.2')
8. sw_add_bom_table()
9. sw_export_drawing_pdf('C:/drawings/part.pdf')
```

### Parametric Design with Global Variables
```
1. sw_create_part()
2. sw_set_global_variable(name='Width', value=100)
3. sw_set_global_variable(name='Height', value=60)
4. sw_create_sketch('Front Plane')
5. sw_sketch_rectangle(0, 0, 100, 60)
6. sw_exit_sketch()
7. sw_extrude(depth=40)
8. sw_link_dimension_to_equation(dim_name='D1@Sketch1', expr='"Width"')
9. sw_link_dimension_to_equation(dim_name='D2@Sketch1', expr='"Height"')
# Now changing "Width" or "Height" globals drives the geometry
```

## 📝 Notes & Conventions

- All length inputs are in **millimeters (mm)**. The server converts internally to meters for the SolidWorks COM API.
- All angles are in **degrees** (converted to radians internally).
- SolidWorks must be open (or the server will auto-launch it via `sw_connect`).
- Entity names must match exactly as shown in the SolidWorks FeatureManager tree (case-sensitive).
- The server runs as a **stdio MCP server** (single client, local only — no network exposure).
- All COM calls are routed through a single STA thread to satisfy SolidWorks' threading model.

---

## 🏗️ Architecture

```
┌──────────────────┐      stdio       ┌──────────────────┐    COM API    ┌──────────────┐
│  Claude / MCP    │ ◀──────────────▶ │   server.py      │ ◀───────────▶ │  SolidWorks  │
│     client       │   JSON-RPC       │   (FastMCP)      │   (win32com)  │   (running)  │
└──────────────────┘                  └──────────────────┘               └──────────────┘
                                              │
                                              ▼
                                      ┌──────────────┐
                                      │ sw_client.py │
                                      │ COM wrapper  │
                                      │ (STA thread) │
                                      └──────────────┘
```

- **`server.py`** — FastMCP stdio server. Defines all 187 `sw_*` tools with Pydantic input validation.
- **`sw_client.py`** — Synchronous `SolidWorksClient` class wrapping SolidWorks COM. Includes:
  - `_get(obj, name)` helper for SW 2024's dual property/method getters
  - `_invoke(obj, dispid, *args)` for DISPID-based method dispatch (bypasses win32com's property-get confusion)
  - `VARIANT(VT_DISPATCH, None)` callout helpers for SelectByID2-style calls
- All COM operations are queued onto a single-threaded `ThreadPoolExecutor` initialized with `pythoncom.CoInitialize`.

---

## 🛠️ Troubleshooting

### "SolidWorks is not running" / `sw_connect` fails
- Make sure SolidWorks itself is installed and licensed.
- Try launching SolidWorks manually first, then call `sw_connect`.
- If COM isn't registered: open SolidWorks once as Administrator, then close.

### `pywintypes.com_error: (-2147352571, 'Type mismatch.', ...)`
- An argument has the wrong type. Most common cause: passing `None` where COM expects an interface pointer.
  This server already wraps null callouts with `VARIANT(VT_DISPATCH, None)` — if you see this in a new
  tool, copy that pattern.

### `'str' object is not callable` or `'int' object is not callable`
- A SolidWorks getter is being exposed as a property in your SW version. Use `_get(obj, 'MethodName')`
  instead of `obj.MethodName()`. See `sw_client.py` for examples.

### `FeatureExtrusion3 — Parameter not optional`
- Your SolidWorks version expects a different parameter count. `sw_extrude` already falls back through
  23/22/20-param signatures.

### Claude doesn't see the `solidworks` server
- Confirm the path in `claude_desktop_config.json` is absolute and uses `\\` escaping.
- Confirm `python` is on PATH (try `python --version` from a fresh terminal).
- Check Claude Desktop logs: `%APPDATA%\Claude\logs\mcp*.log`.

### "SolidWorks COM error" on first tool call
- The first call after launch can be slow (SolidWorks initializing). Retry after a few seconds.
- If persistent, restart SolidWorks and your MCP client.

### `ERROR: another SolidWorks MCP server instance is already running`
The server enforces a **single-instance lock** (Windows named mutex `Local\SolidWorksMCPServer_SingleInstance_v1`) because multiple processes fighting over the same SolidWorks COM connection causes every tool call to hang.
- If both Claude Desktop and Claude Code try to launch the server, only the first succeeds — the second client will see no `solidworks` tools.
- To fix: fully close whichever client you don't need (system tray → Quit for Claude Desktop, or close all CLI sessions). The mutex auto-releases when the holding process exits.
- To bypass for debugging only, set `SW_MCP_ALLOW_MULTIPLE=1` in the spawning environment.

### SolidWorks GUI never appears (processes alive but no window)
If `Get-Process SLDWORKS` shows running processes with `MainWindowHandle: 0`, those are zombies from an interrupted launch — typically after a force-kill (`Stop-Process -Force`) during startup. They will never finish initializing and are not registered in the COM Running Object Table.
- Fix: `Get-Process SLDWORKS,sldworks_fs | Stop-Process -Force` to kill them, then launch SolidWorks fresh from the Start menu and **wait for the main window** before calling any `sw_*` tool.

### `pip install` permission denied
- Use `pip install --user -r requirements.txt` or create a virtualenv.

---

## 🤝 Contributing

Contributions welcome — there's plenty of surface area in SolidWorks' COM API. Some ideas:

- Add tools for missing categories (SimulationXpress, Toolbox, PDM, Route, Electrical)
- Improve error messages with SW-specific context
- Add unit tests for tools that don't require a live SolidWorks instance
- Document additional SolidWorks-version-specific quirks you find

To contribute:

1. Fork the repo and create a feature branch.
2. Follow the existing tool pattern: client method in `sw_client.py` + Pydantic input model + `@mcp.tool`
   wrapper in `server.py`.
3. For methods you've actually run against SolidWorks, note the SW version in the docstring.
4. Open a pull request with a description of the change and which SW version you tested on.

If you find a SW version quirk that needs handling (different DISPID, different param count, property vs.
method), please document it in the PR — that's the most valuable kind of contribution for this project.

---

## ⚠️ Limitations

- **Windows only.** SolidWorks COM is not available on macOS/Linux.
- **Single MCP client at a time** (stdio transport).
- **Long-running operations block.** Simulations and large rebuilds will tie up the COM thread.
- **SolidWorks UI must remain interactive** for some operations (file dialogs, error prompts).
  This server uses `swSaveSilent` and equivalent flags to suppress dialogs where the API allows.
- **Add-in dependent tools** (`sw_create_simulation_study`, `sw_take_photoview_screenshot`) require the
  corresponding add-in to be installed and loaded.

---

## 📄 License

Released under the [MIT License](LICENSE). SolidWorks® is a registered trademark of Dassault Systèmes
SolidWorks Corporation; this project is not affiliated with or endorsed by Dassault Systèmes.

---

## 🔗 Related Projects

- [Model Context Protocol](https://modelcontextprotocol.io) — the open protocol this server implements.
- [SolidWorks API documentation](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks_namespace.html)
- [pywin32](https://github.com/mhammond/pywin32) — Python COM bindings used under the hood.
