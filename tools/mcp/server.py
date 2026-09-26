"""MCP tools for the CodexVBE bridge. Run with ``uv run --project tools/mcp server.py``."""

from __future__ import annotations

import ctypes
import json
from pathlib import Path
from typing import Any, Literal

from mcp.server import MCPServer


mcp = MCPServer(
    "CodexVBE",
    instructions=(
        "Control an already open VBE add-in in Excel or SOLIDWORKS. "
        "Call status and list_projects for the intended host PID first. "
        "Read a module or form before edits, then pass its returned hash/version. "
        "A successful debug command only confirms invocation; inspect debug_state afterward."
    ),
)

MAX_RESPONSE = 10 * 1024 * 1024


def _call(host_process_id: int, command: str, **arguments: Any) -> Any:
    if host_process_id <= 0:
        raise ValueError("host_process_id must be a positive Windows process ID")
    pipe_name = rf"\\.\pipe\CodexVBE.{host_process_id}"
    if not ctypes.windll.kernel32.WaitNamedPipeW(pipe_name, 3000):
        raise ConnectionError(f"CodexVBE is unavailable in host PID {host_process_id}; open its VBE and load the add-in")
    request = {"Command": command, **arguments}
    with open(pipe_name, "r+b", buffering=0) as pipe:
        pipe.write((json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8"))
        raw = pipe.readline(MAX_RESPONSE + 1)
    if not raw or len(raw) > MAX_RESPONSE:
        raise RuntimeError("CodexVBE returned an empty or oversized response")
    response = json.loads(raw)
    if not response.get("Ok"):
        raise RuntimeError(response.get("Error") or "CodexVBE failed without an error message")
    return response.get("Data")


@mcp.tool()
def vbe_status(host_process_id: int) -> dict:
    """Check that the add-in bridge is loaded in the specified Excel or SOLIDWORKS PID."""
    return _call(host_process_id, "status")


@mcp.tool()
def vbe_list_projects(host_process_id: int) -> list:
    """List open VBA projects and their design/run/break modes in one VBE host."""
    return _call(host_process_id, "list_projects")


@mcp.tool()
def vbe_list_modules(host_process_id: int, project: str) -> list:
    """List code components in an exact, unambiguous VBA project name."""
    return _call(host_process_id, "list_modules", Project=project)


@mcp.tool()
def vbe_read_module(host_process_id: int, project: str, module: str) -> dict:
    """Read live module code and its SHA-256 token for guarded edits."""
    return _call(host_process_id, "read_module", Project=project, Module=module)


@mcp.tool()
def vbe_replace_lines(
    host_process_id: int, project: str, module: str, start_line: int,
    count: int, text: str, expected_sha256: str,
) -> dict:
    """Replace a line range in design mode only if the module still has the read SHA-256."""
    return _call(host_process_id, "replace_lines", Project=project, Module=module,
                 StartLine=start_line, Count=count, Text=text, ExpectedSha256=expected_sha256)


@mcp.tool()
def vbe_debug_state(host_process_id: int, project: str) -> dict:
    """Read project mode and the active code-pane selection after debug actions."""
    return _call(host_process_id, "debug_state", Project=project)


@mcp.tool()
def vbe_list_commands(host_process_id: int, query: str = "") -> list:
    """Find native VBE command-bar controls and their current IDs, captions and enabled state."""
    return _call(host_process_id, "list_commands", Query=query)


@mcp.tool()
def vbe_select_code(
    host_process_id: int, project: str, module: str, start_line: int, expected_sha256: str,
) -> dict:
    """Select a code line through VBIDE; requires the current module SHA-256."""
    return _call(host_process_id, "select_code", Project=project, Module=module,
                 StartLine=start_line, ExpectedSha256=expected_sha256)


@mcp.tool()
def vbe_invoke_debug(
    host_process_id: int, project: str, module: str, start_line: int,
    expected_sha256: str, expected_mode: int,
    action: Literal["toggle_breakpoint", "run", "continue", "step_into", "step_over"],
    control_id: int, control_caption: str,
) -> dict:
    """Invoke a validated native VBE debug command; read debug_state to verify its effect."""
    return _call(host_process_id, "invoke_debug", Project=project, Module=module,
                 StartLine=start_line, ExpectedSha256=expected_sha256, ExpectedMode=expected_mode,
                 Action=action, ControlId=control_id, ControlCaption=control_caption)


@mcp.tool()
def vbe_list_forms(host_process_id: int, project: str) -> list:
    """List UserForms and whether their designers are open."""
    return _call(host_process_id, "list_forms", Project=project)


@mcp.tool()
def vbe_form_state(host_process_id: int, project: str, form: str) -> dict:
    """Read UserForm and control names, captions, geometry, fonts and edit version."""
    return _call(host_process_id, "form_state", Project=project, Form=form)


@mcp.tool()
def vbe_form_properties(host_process_id: int, project: str, form: str) -> list:
    """Enumerate scalar VBIDE properties of a UserForm component."""
    return _call(host_process_id, "form_properties", Project=project, Form=form)


@mcp.tool()
def vbe_create_form(host_process_id: int, project: str, form: str) -> dict:
    """Create and visibly open a named UserForm in a project in design mode."""
    return _call(host_process_id, "create_form", Project=project, Form=form)


@mcp.tool()
def vbe_open_form(host_process_id: int, project: str, form: str) -> dict:
    """Visibly open an existing UserForm designer and return its state."""
    return _call(host_process_id, "open_form", Project=project, Form=form)


@mcp.tool()
def vbe_add_form_control(
    host_process_id: int, project: str, form: str, control: str, control_type: str,
    left: float, top: float, width: float, height: float, expected_form_version: str,
    caption: str | None = None,
) -> dict:
    """Add a built-in MSForms control to a visible UserForm, guarded by its version."""
    args = dict(Project=project, Form=form, Control=control, ControlType=control_type,
                Left=left, Top=top, Width=width, Height=height,
                ExpectedFormVersion=expected_form_version)
    if caption is not None:
        args["Caption"] = caption
    return _call(host_process_id, "add_form_control", **args)


@mcp.tool()
def vbe_set_form_control_geometry(
    host_process_id: int, project: str, form: str, control: str,
    left: float, top: float, width: float, height: float, expected_form_version: str,
) -> dict:
    """Set a UserForm control's position and size in points, guarded by form version."""
    return _call(host_process_id, "set_form_control_geometry", Project=project, Form=form,
                 Control=control, Left=left, Top=top, Width=width, Height=height,
                 ExpectedFormVersion=expected_form_version)


@mcp.tool()
def vbe_rename_form_control(
    host_process_id: int, project: str, form: str, control: str,
    new_name: str, expected_form_version: str,
) -> dict:
    """Rename a UserForm control in design mode, guarded by form version."""
    return _call(host_process_id, "rename_form_control", Project=project, Form=form,
                 Control=control, NewName=new_name, ExpectedFormVersion=expected_form_version)


@mcp.tool()
def vbe_set_form_control_caption(
    host_process_id: int, project: str, form: str, control: str,
    caption: str, expected_form_version: str,
) -> dict:
    """Set a UserForm control's visible Caption, guarded by form version."""
    return _call(host_process_id, "set_form_control_caption", Project=project, Form=form,
                 Control=control, Caption=caption, ExpectedFormVersion=expected_form_version)


@mcp.tool()
def vbe_set_form_control_font(
    host_process_id: int, project: str, form: str, control: str,
    font_name: str, font_size: float, font_bold: bool, expected_form_version: str,
) -> dict:
    """Set a UserForm control's font name, point size and bold flag, guarded by version."""
    return _call(host_process_id, "set_form_control_font", Project=project, Form=form,
                 Control=control, FontName=font_name, FontSize=font_size,
                 FontBold=font_bold, ExpectedFormVersion=expected_form_version)


if __name__ == "__main__":
    mcp.run()
