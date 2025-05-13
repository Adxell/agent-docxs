from typing import Optional, Dict, Any, List

from pathlib import Path
from mcp.server.fastmcp import FastMCP

from word_document_editor import WordDocumentEditor


mcp = FastMCP("Office Document Server") 
docx_editor_instance = WordDocumentEditor() 

async def run_sync_tool(func, *args, **kwargs):
    return func(*args, **kwargs)


@mcp.tool()
async def create_docx_document(filename: str = "new_document.docx") -> Dict[str, Any]:
    """
    Creates a new, blank Word document (.docx) in memory.
    Args:
        filename (str, optional): Default filename for saving. Stored in 'documents_fastmcp'.
    Returns:
        Dict: Status message and current filename.
    """
    try:
        message = await run_sync_tool(docx_editor_instance.create_document, filename=filename)
        return {"status": "success", "message": message, "current_filename": docx_editor_instance.current_filename}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def load_docx_document(filename: str) -> Dict[str, Any]:
    """
    Loads an existing Word document (.docx) from 'documents_fastmcp' into memory.
    Args:
        filename (str): Name of the Word document file to load.
    Returns:
        Dict: Status message and current filename.
    """
    try:
        message = await run_sync_tool(docx_editor_instance.load_document, filename=filename)
        return {"status": "success", "message": message, "current_filename": docx_editor_instance.current_filename}
    except FileNotFoundError as e:
        return {"status": "error", "message": str(e)}
    except Exception as e:
        return {"status": "error", "message": f"An unexpected error occurred: {str(e)}"}

@mcp.tool()
async def save_docx_document(filename: Optional[str] = None) -> Dict[str, Any]:
    """
    Saves the current in-memory Word document (.docx) to 'documents_fastmcp'.
    Args:
        filename (str, optional): New filename. If None, uses current filename.
    Returns:
        Dict: Status message and saved filename.
    """
    try:
        message = await run_sync_tool(docx_editor_instance.save_document, filename=filename)
        return {"status": "success", "message": message, "saved_filename": docx_editor_instance.current_filename}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_paragraph(text: str, style: Optional[str] = None) -> Dict[str, Any]:
    """
    Adds a paragraph to the Word document.
    Args:
        text (str): The text content of the paragraph.
        style (str, optional): The style to apply (e.g., 'Normal', 'BodyText', 'Heading1', 'ListBullet').
    Returns:
        Dict: Status, message, and details of the added paragraph.
    """
    try:
        para_info = await run_sync_tool(docx_editor_instance.add_paragraph, text=text, style=style)
        return {"status": "success", "message": "Paragraph added to Word document.", "data": para_info}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_heading(text: str, level: int = 1) -> Dict[str, Any]:
    """
    Adds a heading to the Word document.
    Args:
        text (str): The text of the heading.
        level (int, optional): Heading level (0 for Title, 1-9 for Headings). Defaults to 1.
    Returns:
        Dict: Status, message, and details of the added heading.
    """
    try:
        heading_info = await run_sync_tool(docx_editor_instance.add_heading, text=text, level=level)
        return {"status": "success", "message": "Heading added to Word document.", "data": heading_info}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_styled_text_paragraph(
    text_runs: List[Dict[str, Any]], 
    paragraph_style: Optional[str] = None
) -> Dict[str, Any]:
    """
    Adds a paragraph to Word doc with multiple styled text runs.
    Each run in text_runs is a dict: {"text": "str", "bold": bool, "italic": bool, 
                                      "font_size_pt": int, "font_name": "str", "font_color_rgb": [R,G,B]}
    Args:
        text_runs (List[Dict[str, Any]]): List of text run dictionaries.
        paragraph_style (str, optional): Style for the entire paragraph.
    Returns:
        Dict: Status, message, and the combined text of the added paragraph.
    """
    try:
        # Basic validation for text_runs structure
        if not isinstance(text_runs, list) or not all(isinstance(run, dict) and "text" in run for run in text_runs):
            return {"status": "error", "message": "text_runs must be a list of dictionaries, each with a 'text' key."}
        
        added_text = await run_sync_tool(
            docx_editor_instance.add_styled_text_to_paragraph, 
            text_runs=text_runs, 
            paragraph_style=paragraph_style
        )
        return {"status": "success", "message": "Styled text paragraph added to Word document.", "data": {"full_text": added_text}}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_table(rows: int, cols: int, data_list: Optional[List[List[str]]] = None, style: Optional[str] = 'TableGrid') -> Dict[str, Any]:
    """
    Adds a table to the Word document.
    Args:
        rows (int): Number of rows.
        cols (int): Number of columns.
        data_list (List[List[str]], optional): Data to populate table cells (list of rows).
        style (str, optional): Table style (e.g., 'TableGrid', 'LightShading-Accent1'). Defaults to 'TableGrid'.
    Returns:
        Dict: Status, message, and details of the added table.
    """
    try:
        table_info = await run_sync_tool(docx_editor_instance.add_table, rows=rows, cols=cols, data_list=data_list, style=style)
        return {"status": "success", "message": "Table added to Word document.", "data": table_info}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_picture(image_path: str, width_inch: Optional[float] = None, height_inch: Optional[float] = None) -> Dict[str, Any]:
    """
    Adds a picture to the Word document from a server-accessible local path.
    Args:
        image_path (str): Server-local path to the image file.
        width_inch (float, optional): Desired width of the picture in inches.
        height_inch (float, optional): Desired height of the picture in inches.
    Returns:
        Dict: Status and message.
    """
    # SECURITY WARNING: Ensure image_path is validated or restricted to prevent access to arbitrary files.
    # For this example, we assume the LLM provides safe paths or paths are pre-validated.
    try:
        message = await run_sync_tool(docx_editor_instance.add_picture, image_path=image_path, width_inch=width_inch, height_inch=height_inch)
        return {"status": "success", "message": message}
    except FileNotFoundError:
        return {"status": "error", "message": f"Image file not found at path: {image_path}. Ensure the path is correct and accessible by the server."}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@mcp.tool()
async def add_docx_page_break() -> Dict[str, Any]:
    """
    Adds a manual page break to the Word document.
    Returns:
        Dict: Status and message.
    """
    try:
        message = await run_sync_tool(docx_editor_instance.add_page_break)
        return {"status": "success", "message": message}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    mcp.run("stdio")
