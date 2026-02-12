# pptx-mcp-live

MCP server for Microsoft PowerPoint with live COM automation on Windows.

![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)
![Tools: 58](https://img.shields.io/badge/tools-58-orange)

32 cross-platform tools (python-pptx) work everywhere. 26 Windows-only live tools use COM automation to work with presentations **while they're open in PowerPoint** — picture effects, shape formatting, transitions, PDF/image export, and more that python-pptx simply cannot do.

## What's New vs the Original

This project builds on [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (32 tools) and adds 26 Windows COM tools:

| Category | New Tools | What They Do |
|----------|-----------|--------------|
| Read (COM) | `ppt_live_get_info`, `ppt_live_get_text`, `ppt_live_get_slide_info`, `ppt_live_find_text`, `ppt_live_get_notes`, `ppt_live_list_comments`, `ppt_live_list_open` | Read from open presentations |
| Edit (COM) | `ppt_live_insert_text`, `ppt_live_replace_text`, `ppt_live_add_slide`, `ppt_live_delete_slide`, `ppt_live_set_notes`, `ppt_live_add_image`, `ppt_live_add_comment`, `ppt_live_delete_comment` | Edit open presentations |
| Format (COM) | `ppt_live_format_picture`, `ppt_live_format_shape`, `ppt_live_arrange_shapes`, `ppt_live_set_transition`, `ppt_live_set_header_footer`, `ppt_live_manage_sections`, `ppt_live_modify_table` | Effects and formatting python-pptx can't do |
| Export (COM) | `ppt_live_export_pdf`, `ppt_live_export_slide_image`, `ppt_live_export_shape_image` | PDF and image export |
| Screen capture | `ppt_screen_capture` | Screenshot of the PowerPoint window |

### COM-exclusive capabilities

These features are **only possible via COM** — python-pptx has no support:

- Picture formatting (brightness, contrast, crop, transparency)
- Shape effects (shadow, glow, reflection, soft edge, 3D)
- Slide transitions and timing
- Presentation sections
- Table modification (add/delete rows/columns, merge cells)
- PDF and image export
- Arrange shapes (group, align, distribute, z-order)
- Headers, footers, and slide numbers

## Tool List

### Cross-Platform Tools (32)

These work on Windows, macOS, and Linux using python-pptx.

<details>
<summary>Presentation Management (7)</summary>

| Tool | Description |
|------|-------------|
| `create_presentation` | Create a new presentation |
| `create_presentation_from_template` | Create from a template file |
| `open_presentation` | Open an existing presentation |
| `save_presentation` | Save to file |
| `get_presentation_info` | Get presentation metadata |
| `get_template_file_info` | Get template layouts and properties |
| `set_core_properties` | Set document properties |

</details>

<details>
<summary>Content (8)</summary>

| Tool | Description |
|------|-------------|
| `add_slide` | Add a new slide with optional background |
| `get_slide_info` | Get slide shape information |
| `extract_slide_text` | Extract text from a slide |
| `extract_presentation_text` | Extract text from all slides |
| `populate_placeholder` | Fill a placeholder with text |
| `add_bullet_points` | Add bullet points to a placeholder |
| `manage_text` | Unified text management (add, format, validate) |
| `manage_image` | Unified image management (add, enhance) |

</details>

<details>
<summary>Structural Elements (4)</summary>

| Tool | Description |
|------|-------------|
| `add_table` | Add a table with formatting options |
| `format_table_cell` | Format a specific table cell |
| `add_shape` | Add an auto shape with options |
| `add_chart` | Add a chart with comprehensive formatting |

</details>

<details>
<summary>Professional Design (3)</summary>

| Tool | Description |
|------|-------------|
| `apply_professional_design` | Themes, slides, and visual enhancements |
| `apply_picture_effects` | Apply multiple picture effects |
| `manage_fonts` | Font analysis, optimization, recommendations |

</details>

<details>
<summary>Templates (7)</summary>

| Tool | Description |
|------|-------------|
| `list_slide_templates` | List available layout templates |
| `apply_slide_template` | Apply a template to an existing slide |
| `create_slide_from_template` | Create slide from a template |
| `create_presentation_from_templates` | Create presentation from template sequence |
| `get_template_info` | Get template details |
| `auto_generate_presentation` | Auto-generate from topic and preferences |
| `optimize_slide_text` | Optimize text for readability |

</details>

<details>
<summary>Advanced Features (5)</summary>

| Tool | Description |
|------|-------------|
| `manage_hyperlinks` | Add, remove, list, update hyperlinks |
| `update_chart_data` | Replace chart data |
| `add_connector` | Add connector lines between shapes |
| `manage_slide_masters` | Access and manage slide masters |
| `manage_slide_transitions` | Basic transition support |

</details>

<details>
<summary>Utility (3)</summary>

| Tool | Description |
|------|-------------|
| `list_presentations` | List loaded presentations |
| `switch_presentation` | Switch active presentation |
| `get_server_info` | Server information |

</details>

### Windows Live Tools (26)

These require Windows with Microsoft PowerPoint installed. They operate on **currently open presentations** via COM automation.

| Tool | Description |
|------|-------------|
| **Read** | |
| `ppt_live_get_info` | Presentation metadata (name, slides, dimensions) |
| `ppt_live_get_text` | Extract all text and speaker notes |
| `ppt_live_get_slide_info` | Detailed shape info for a slide |
| `ppt_live_find_text` | Search text across all slides |
| `ppt_live_get_notes` | Get speaker notes |
| `ppt_live_list_comments` | List all comments |
| `ppt_live_list_open` | List open presentations |
| **Edit** | |
| `ppt_live_insert_text` | Insert text into a shape |
| `ppt_live_replace_text` | Find and replace text |
| `ppt_live_add_slide` | Add a new slide |
| `ppt_live_delete_slide` | Delete a slide |
| `ppt_live_set_notes` | Set speaker notes |
| `ppt_live_add_image` | Add an image |
| `ppt_live_add_comment` | Add a comment |
| `ppt_live_delete_comment` | Delete a comment |
| **Format** | |
| `ppt_live_format_picture` | Brightness, contrast, crop, transparency |
| `ppt_live_format_shape` | Shadow, glow, reflection, 3D, fill, line |
| `ppt_live_arrange_shapes` | Group, align, distribute, z-order, flip |
| `ppt_live_set_transition` | Slide transitions and timing |
| `ppt_live_set_header_footer` | Slide numbers, date, footer |
| `ppt_live_manage_sections` | List, add, rename, delete sections |
| `ppt_live_modify_table` | Add/delete rows/columns, merge cells |
| **Export** | |
| `ppt_live_export_pdf` | Export to PDF |
| `ppt_live_export_slide_image` | Export slide as PNG/JPG |
| `ppt_live_export_shape_image` | Export a shape as image |
| **Screen Capture** | |
| `ppt_screen_capture` | Screenshot of the PowerPoint window |

## Quick Start

### Claude Code (`.mcp.json`)

```json
{
  "mcpServers": {
    "pptx": {
      "command": "python",
      "args": ["path/to/ppt_mcp_server.py"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "pptx": {
      "command": "uvx",
      "args": ["--from", "pptx-mcp-live", "ppt_mcp_server"]
    }
  }
}
```

### From source

```bash
git clone https://github.com/ykarapazar/pptx-mcp-live.git
cd pptx-mcp-live
pip install -r requirements.txt
python ppt_mcp_server.py
```

## Requirements

- Python 3.10+
- `python-pptx`, `mcp[cli]`, `Pillow`, `fonttools` (see `requirements.txt`)
- **Windows Live tools only:** Windows 10/11 + Microsoft PowerPoint + `pywin32`

## Acknowledgments

Built on top of [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (MIT License).

Additional libraries: [python-pptx](https://python-pptx.readthedocs.io/), [pywin32](https://github.com/mhammond/pywin32).

## License

MIT License — see [LICENSE](LICENSE) for details.
