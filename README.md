# Markdown to Word (Dify Plugin)

Convert Markdown text into a styled Microsoft Word (.docx) document directly inside Dify. This tool accepts Markdown and presentation options (font, line spacing, margins, theme, etc.), generates a .docx file, and returns it to the workflow as a downloadable file along with JSON metadata.

## Key Features
- Markdown → Word (.docx) conversion
- Configurable typography (font family/size) and layout (line spacing, margins, paper size)
- Optional page numbers
- Theme hook for future style templates
- Returns both a file (docx) and a JSON message with useful metadata

## Requirements
- Python 3.12+
- Install dependencies: `pip install -r requirements.txt`
- A running Dify instance for plugin debugging or installation

## Getting Started (Debugging)
1. Copy `.env.example` to `.env` and fill in the required values (remote install URL and key from your Dify server).
2. Install dependencies: `pip install -r requirements.txt`.
3. Start the plugin in debug mode: `python -m main`.
4. In Dify → Plugins, use the debugging connection to test the tool in a workflow.

## Tool Parameters
All parameters are defined in `tools/markdown_to_word.yaml`.
- markdown_text (string, required): The Markdown content to convert.
- templates (string, optional): Theme/template name for styling. Default: `default`.
- font_family (string, optional): Default `"微软雅黑"` (Microsoft YaHei). You can use "SimSun", "Times New Roman", etc.
- font_size (number, optional): Font size for body text. Default: `12`.
- line_spacing (number, optional): Line spacing, e.g. `1.0`, `1.5`, `2.0`. Default: `1.5`.
- page_margins (number, optional): Page margins in centimeters (all sides). Default: `2.5`.
- paper_size (string, optional): Paper size, e.g. `A4`, `A3`, `Letter`. Default: `A4`.
- output_file (string, optional): Output file name, e.g. `output.docx`. Default: `output.docx`.
- add_page_numbers (boolean, optional): Whether to add page numbers. Default: `true`.

## Outputs
The tool streams two messages:
1) File (BLOB): The generated Word document (.docx). It appears in the workflow result under `files` and can be downloaded.
2) JSON: A summary payload, including `result`, `output_file`, `file_size`, and the settings used.

MIME type for the returned file: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.

## Example Usage
- Minimal input: provide `markdown_text` with your Markdown content.
- Optional: tune `font_family`, `font_size`, `line_spacing`, `page_margins`, `paper_size`, `templates`, `add_page_numbers`, `output_file`.
- See example Markdown in `examples/basic/simple_document.md`.

Expected result:
- `files` contains one `.docx` file
- JSON message indicates success and includes file size and settings

## Notes & Troubleshooting
- If your Dify result shows an empty `files` array, ensure your environment uses a version of the plugin that returns a file BLOB (this project already does).
- Very large Markdown inputs may take longer to convert.
- Customize themes via `src/templates/themes` and integrate them through the `templates` parameter.

## Privacy
See `PRIVACY.md` for privacy notes.



