#!/usr/bin/env python3
"""
Google Workspace MCP Server (v12 - Shared Drive Support)

🎉 98 TOOLS TOTAL - COMPREHENSIVE GOOGLE WORKSPACE AUTOMATION 🎉

CHANGELOG from v11:
- Added Shared Drive (Team Drive) support to all Drive API calls
- Added supportsAllDrives=True to all files() and permissions() method calls
- Added includeItemsFromAllDrives=True to files().list() calls
- Added corpora='allDrives' to search_drive() for cross-drive search
- Fixes 404 errors when accessing files in Shared Drives

PREVIOUS (v10):
- Added Google Tasks API as the 7th Google Workspace service
- Added tasks scope: https://www.googleapis.com/auth/tasks
- NOTE: Delete token.json and re-authenticate to pick up the new scope
- Added 12 new Google Tasks tools:

  TASK LIST MANAGEMENT (3 tools):
  * tasks_list_tasklists - List all task lists
  * tasks_create_tasklist - Create a new task list
  * tasks_delete_tasklist - Delete a task list

  TASK CRUD (5 tools):
  * tasks_list_tasks - List tasks in a task list
  * tasks_get_task - Get a specific task
  * tasks_create_task - Create a task (with optional subtask support)
  * tasks_update_task - Update a task
  * tasks_delete_task - Delete a task

  TASK ACTIONS (4 tools):
  * tasks_complete_task - Mark a task as completed
  * tasks_move_task - Reorder/reparent a task
  * tasks_clear_completed - Clear completed tasks from a list
  * tasks_search_tasks - Search tasks across all lists by keyword

PREVIOUS (v10):
- Added 21 new Advanced Docs tools for tables, formatting, and document structure

PREVIOUS (v9):
- Added 13 new Advanced Sheets tools for formatting and data management:
  * sheets_batch_update - Execute multiple batch update requests (advanced)
  * sheets_rename_sheet - Rename a sheet/tab
  * sheets_format_cells - Format cells (bold, colors, borders, fonts, alignment)
  * sheets_set_column_width - Set column widths
  * sheets_set_row_height - Set row heights
  * sheets_freeze_rows_columns - Freeze rows and/or columns
  * sheets_merge_cells - Merge a range of cells
  * sheets_unmerge_cells - Unmerge a range of cells
  * sheets_add_filter - Add filter view to a sheet
  * sheets_data_validation - Add data validation (dropdowns, number ranges)
  * sheets_conditional_formatting - Add conditional formatting rules
  * sheets_named_range - Create named ranges
  * sheets_auto_resize - Auto-resize columns/rows to fit content

PREVIOUS (v8):
- Added 4 new Calendar tools:
  * calendar_event_update - Update an existing calendar event
  * calendar_event_delete - Delete a calendar event
  * calendar_quick_add - Create events using natural language
  * calendar_list_calendars - List all calendars the user has access to

PREVIOUS (v7):
- Added 4 new Slides tools:
  * slides_delete_slide - Delete a slide from a presentation
  * slides_replace_text - Find and replace text across all slides
  * slides_duplicate_slide - Duplicate an existing slide
  * slides_get_details - Get detailed slide information including object IDs

PREVIOUS (v6):
- Added 4 new Docs tools:
  * docs_replace_text - Find and replace all occurrences of text
  * docs_insert_text - Insert text at a specific index position
  * docs_delete_content - Delete content between two index positions
  * docs_get_structure - Get document structure with character indexes

PREVIOUS (v5):
- Added 5 new Sheets tools:
  * sheets_create - Create new Google Spreadsheet
  * sheets_get_metadata - Get spreadsheet metadata (title, sheets, dimensions)
  * sheets_clear - Clear a range of cells
  * sheets_add_sheet - Add a new sheet/tab to existing spreadsheet
  * sheets_delete_sheet - Delete a sheet/tab from spreadsheet

PREVIOUS (v4):
- Added Gmail scopes: gmail.compose, gmail.modify, gmail.labels
- Added 6 new Gmail tools:
  * gmail_draft_create - Create draft emails
  * gmail_draft_list - List all drafts
  * gmail_draft_send - Send an existing draft
  * gmail_reply - Reply to email threads
  * gmail_labels_list - List all labels
  * gmail_message_modify - Add/remove labels, archive, mark read/unread

TOOL COUNT (98 tools):
  - Docs: 29 tools (7 basic + 22 advanced)
  - Sheets: 21 tools (8 basic + 13 advanced)
  - Drive: 13 tools
  - Calendar: 7 tools
  - Gmail: 9 tools
  - Slides: 7 tools
  - Tasks: 12 tools

INSTALLATION:
1. Replace server.py with this file
2. Delete token.json and re-authenticate (new Tasks scope)
3. Run: python3 server.py
4. Restart Claude Code
"""

import os
import io
import json
import logging
import base64
from pathlib import Path
from typing import Any, Optional
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Scopes required for Google APIs
# v4: Added gmail.compose, gmail.modify, gmail.labels
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets', 
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',   # NEW: for drafts
    'https://www.googleapis.com/auth/gmail.modify',    # NEW: for labels/modify
    'https://www.googleapis.com/auth/gmail.labels',    # NEW: for label management
    'https://www.googleapis.com/auth/presentations',
]

# Paths for credentials
SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_PATH = SCRIPT_DIR / "credentials.json"
TOKEN_PATH = SCRIPT_DIR / "token.json"


class GoogleWorkspaceClient:
    """Client for interacting with Google Workspace APIs."""
    
    def __init__(self):
        self.creds: Optional[Credentials] = None
        self._docs_service = None
        self._sheets_service = None
        self._drive_service = None
        self._calendar_service = None
        self._gmail_service = None
        self._slides_service = None

    def authenticate(self) -> bool:
        """Authenticate with Google APIs using OAuth2."""
        try:
            if TOKEN_PATH.exists():
                self.creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
            
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                else:
                    if not CREDENTIALS_PATH.exists():
                        logger.error(f"Credentials file not found at {CREDENTIALS_PATH}")
                        return False
                    
                    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
                    self.creds = flow.run_local_server(port=0)
                
                with open(TOKEN_PATH, 'w') as token:
                    token.write(self.creds.to_json())
            
            return True
        except Exception as e:
            logger.error(f"Authentication failed: {e}")
            return False
    
    @property
    def docs_service(self):
        if not self._docs_service:
            self._docs_service = build('docs', 'v1', credentials=self.creds)
        return self._docs_service
    
    @property
    def sheets_service(self):
        if not self._sheets_service:
            self._sheets_service = build('sheets', 'v4', credentials=self.creds)
        return self._sheets_service
    
    @property
    def drive_service(self):
        if not self._drive_service:
            self._drive_service = build('drive', 'v3', credentials=self.creds)
        return self._drive_service
    
    @property
    def calendar_service(self):
        if not self._calendar_service:
            self._calendar_service = build('calendar', 'v3', credentials=self.creds)
        return self._calendar_service
    
    @property
    def gmail_service(self):
        if not self._gmail_service:
            self._gmail_service = build('gmail', 'v1', credentials=self.creds)
        return self._gmail_service
    
    @property
    def slides_service(self):
        if not self._slides_service:
            self._slides_service = build('slides', 'v1', credentials=self.creds)
        return self._slides_service

google_client = GoogleWorkspaceClient()
server = Server("google-workspace-mcp")


@server.list_tools()
async def list_tools() -> list[Tool]:
    """List available tools."""
    return [
        # ==================== GOOGLE DOCS ====================
        Tool(
            name="google_docs_create",
            description="Create a new Google Doc with optional content",
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string", "description": "Title of the document"},
                    "content": {"type": "string", "description": "Initial content (plain text)"},
                    "folder_id": {"type": "string", "description": "Optional folder ID to create in"}
                },
                "required": ["title"]
            }
        ),
        Tool(
            name="google_docs_read",
            description="Read the content of a Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"}
                },
                "required": ["document_id"]
            }
        ),
        Tool(
            name="google_docs_append",
            description="Append content to an existing Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "content": {"type": "string", "description": "Content to append"}
                },
                "required": ["document_id", "content"]
            }
        ),
        Tool(
            name="docs_replace_text",
            description="Find and replace all occurrences of text in a Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "find_text": {"type": "string", "description": "Text to find"},
                    "replace_text": {"type": "string", "description": "Text to replace with"},
                    "match_case": {"type": "boolean", "description": "Case-sensitive match (default false)"}
                },
                "required": ["document_id", "find_text", "replace_text"]
            }
        ),
        Tool(
            name="docs_insert_text",
            description="Insert text at a specific index position in a Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "text": {"type": "string", "description": "Text to insert"},
                    "index": {"type": "integer", "description": "Position to insert at (1 = beginning of document)"}
                },
                "required": ["document_id", "text", "index"]
            }
        ),
        Tool(
            name="docs_delete_content",
            description="Delete content between two index positions in a Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start position (inclusive)"},
                    "end_index": {"type": "integer", "description": "End position (exclusive)"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        Tool(
            name="docs_get_structure",
            description="Get document structure with character indexes - useful for planning edits",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"}
                },
                "required": ["document_id"]
            }
        ),
        # NEW in v10: Advanced Docs tools - Table Operations
        Tool(
            name="docs_insert_table",
            description="Insert a table at a specific index in a Google Doc",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "index": {"type": "integer", "description": "Position to insert table (use docs_get_structure to find indexes)"},
                    "rows": {"type": "integer", "description": "Number of rows"},
                    "columns": {"type": "integer", "description": "Number of columns"}
                },
                "required": ["document_id", "index", "rows", "columns"]
            }
        ),
        Tool(
            name="docs_insert_table_row",
            description="Insert row(s) in an existing table",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_index": {"type": "integer", "description": "Row index to insert at (0-based)"},
                    "insert_below": {"type": "boolean", "description": "Insert below the specified row (default true)"}
                },
                "required": ["document_id", "table_start_index", "row_index"]
            }
        ),
        Tool(
            name="docs_insert_table_column",
            description="Insert column(s) in an existing table",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "column_index": {"type": "integer", "description": "Column index to insert at (0-based)"},
                    "insert_right": {"type": "boolean", "description": "Insert to the right of the specified column (default true)"}
                },
                "required": ["document_id", "table_start_index", "column_index"]
            }
        ),
        Tool(
            name="docs_delete_table_row",
            description="Delete row(s) from a table",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_index": {"type": "integer", "description": "Row index to delete (0-based)"}
                },
                "required": ["document_id", "table_start_index", "row_index"]
            }
        ),
        Tool(
            name="docs_delete_table_column",
            description="Delete column(s) from a table",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "column_index": {"type": "integer", "description": "Column index to delete (0-based)"}
                },
                "required": ["document_id", "table_start_index", "column_index"]
            }
        ),
        Tool(
            name="docs_write_table_cell",
            description="Write text to a specific table cell",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_index": {"type": "integer", "description": "Row index (0-based)"},
                    "column_index": {"type": "integer", "description": "Column index (0-based)"},
                    "text": {"type": "string", "description": "Text to write to the cell"},
                    "replace_existing": {"type": "boolean", "description": "Replace existing cell content (default true)"}
                },
                "required": ["document_id", "table_start_index", "row_index", "column_index", "text"]
            }
        ),
        Tool(
            name="docs_write_table_bulk",
            description="Write text to multiple table cells in a single API call. Much more efficient than multiple docs_write_table_cell calls. Use this when populating tables.",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "cells": {
                        "type": "array",
                        "description": "Array of cell data to write",
                        "items": {
                            "type": "object",
                            "properties": {
                                "row": {"type": "integer", "description": "Row index (0-based)"},
                                "column": {"type": "integer", "description": "Column index (0-based)"},
                                "text": {"type": "string", "description": "Text to write"}
                            },
                            "required": ["row", "column", "text"]
                        }
                    },
                    "replace_existing": {"type": "boolean", "description": "Replace existing cell content (default true)"}
                },
                "required": ["document_id", "table_start_index", "cells"]
            }
        ),
        Tool(
            name="docs_merge_table_cells",
            description="Merge table cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_start": {"type": "integer", "description": "Starting row index (0-based)"},
                    "row_end": {"type": "integer", "description": "Ending row index (exclusive)"},
                    "column_start": {"type": "integer", "description": "Starting column index (0-based)"},
                    "column_end": {"type": "integer", "description": "Ending column index (exclusive)"}
                },
                "required": ["document_id", "table_start_index", "row_start", "row_end", "column_start", "column_end"]
            }
        ),
        Tool(
            name="docs_unmerge_table_cells",
            description="Unmerge table cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_start": {"type": "integer", "description": "Starting row index (0-based)"},
                    "row_end": {"type": "integer", "description": "Ending row index (exclusive)"},
                    "column_start": {"type": "integer", "description": "Starting column index (0-based)"},
                    "column_end": {"type": "integer", "description": "Ending column index (exclusive)"}
                },
                "required": ["document_id", "table_start_index", "row_start", "row_end", "column_start", "column_end"]
            }
        ),
        # NEW in v10: Advanced Docs tools - Table Formatting
        Tool(
            name="docs_format_table_cell",
            description="Format table cell (background color, borders, padding). Colors use RGB 0-1 scale.",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_start": {"type": "integer", "description": "Starting row index (0-based)"},
                    "row_end": {"type": "integer", "description": "Ending row index (exclusive)"},
                    "column_start": {"type": "integer", "description": "Starting column index (0-based)"},
                    "column_end": {"type": "integer", "description": "Ending column index (exclusive)"},
                    "background_color": {"type": "object", "description": "Background color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "border_color": {"type": "object", "description": "Border color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "border_width": {"type": "number", "description": "Border width in points"},
                    "padding_top": {"type": "number", "description": "Top padding in points"},
                    "padding_bottom": {"type": "number", "description": "Bottom padding in points"},
                    "padding_left": {"type": "number", "description": "Left padding in points"},
                    "padding_right": {"type": "number", "description": "Right padding in points"},
                    "vertical_alignment": {"type": "string", "description": "Vertical alignment: 'TOP', 'MIDDLE', 'BOTTOM'"}
                },
                "required": ["document_id", "table_start_index", "row_start", "row_end", "column_start", "column_end"]
            }
        ),
        Tool(
            name="docs_set_table_column_width",
            description="Set table column widths",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "column_index": {"type": "integer", "description": "Column index (0-based)"},
                    "width": {"type": "number", "description": "Width in points (72 points = 1 inch)"}
                },
                "required": ["document_id", "table_start_index", "column_index", "width"]
            }
        ),
        Tool(
            name="docs_set_table_row_height",
            description="Set minimum table row height",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "table_start_index": {"type": "integer", "description": "Start index of the table (from docs_get_structure)"},
                    "row_index": {"type": "integer", "description": "Row index (0-based)"},
                    "min_height": {"type": "number", "description": "Minimum height in points (72 points = 1 inch)"}
                },
                "required": ["document_id", "table_start_index", "row_index", "min_height"]
            }
        ),
        # NEW in v10: Advanced Docs tools - Text Formatting
        Tool(
            name="docs_format_text",
            description="Format text range (bold, italic, underline, color, font, size). Colors use RGB 0-1 scale.",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of text range"},
                    "end_index": {"type": "integer", "description": "End index of text range"},
                    "bold": {"type": "boolean", "description": "Make text bold"},
                    "italic": {"type": "boolean", "description": "Make text italic"},
                    "underline": {"type": "boolean", "description": "Underline text"},
                    "strikethrough": {"type": "boolean", "description": "Strikethrough text"},
                    "font_size": {"type": "number", "description": "Font size in points"},
                    "font_family": {"type": "string", "description": "Font family (e.g., 'Arial', 'Times New Roman')"},
                    "foreground_color": {"type": "object", "description": "Text color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "background_color": {"type": "object", "description": "Highlight/background color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "link_url": {"type": "string", "description": "URL to link the text to"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        Tool(
            name="docs_format_paragraph",
            description="Format paragraph (alignment, spacing, indentation)",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of paragraph range"},
                    "end_index": {"type": "integer", "description": "End index of paragraph range"},
                    "alignment": {"type": "string", "description": "Alignment: 'START', 'CENTER', 'END', 'JUSTIFIED'"},
                    "line_spacing": {"type": "number", "description": "Line spacing (1.0 = single, 1.5 = 1.5 lines, 2.0 = double)"},
                    "space_above": {"type": "number", "description": "Space above paragraph in points"},
                    "space_below": {"type": "number", "description": "Space below paragraph in points"},
                    "indent_first_line": {"type": "number", "description": "First line indent in points"},
                    "indent_start": {"type": "number", "description": "Left/start indent in points"},
                    "indent_end": {"type": "number", "description": "Right/end indent in points"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        Tool(
            name="docs_create_bullet_list",
            description="Create bulleted list from paragraphs",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of paragraph range"},
                    "end_index": {"type": "integer", "description": "End index of paragraph range"},
                    "bullet_preset": {"type": "string", "description": "Bullet style: 'BULLET_DISC_CIRCLE_SQUARE', 'BULLET_DIAMONDX_ARROW3D_SQUARE', 'BULLET_CHECKBOX', 'BULLET_ARROW_DIAMOND_DISC', 'BULLET_STAR_CIRCLE_SQUARE', 'BULLET_ARROW3D_CIRCLE_SQUARE', 'BULLET_LEFTTRIANGLE_DIAMOND_DISC' (default: BULLET_DISC_CIRCLE_SQUARE)"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        Tool(
            name="docs_create_numbered_list",
            description="Create numbered list from paragraphs",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of paragraph range"},
                    "end_index": {"type": "integer", "description": "End index of paragraph range"},
                    "number_preset": {"type": "string", "description": "Number style: 'NUMBERED_DECIMAL_ALPHA_ROMAN', 'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS', 'NUMBERED_DECIMAL_NESTED', 'NUMBERED_UPPERALPHA_ALPHA_ROMAN', 'NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL', 'NUMBERED_ZERODECIMAL_ALPHA_ROMAN' (default: NUMBERED_DECIMAL_ALPHA_ROMAN)"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        Tool(
            name="docs_remove_bullets",
            description="Remove bullets/numbering from paragraphs",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of paragraph range"},
                    "end_index": {"type": "integer", "description": "End index of paragraph range"}
                },
                "required": ["document_id", "start_index", "end_index"]
            }
        ),
        # NEW in v10: Advanced Docs tools - Document Structure
        Tool(
            name="docs_insert_page_break",
            description="Insert a page break at a specific index",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "index": {"type": "integer", "description": "Position to insert page break"}
                },
                "required": ["document_id", "index"]
            }
        ),
        Tool(
            name="docs_insert_section_break",
            description="Insert a section break at a specific index",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "index": {"type": "integer", "description": "Position to insert section break"},
                    "section_type": {"type": "string", "description": "Section type: 'NEXT_PAGE', 'CONTINUOUS' (default: NEXT_PAGE)"}
                },
                "required": ["document_id", "index"]
            }
        ),
        Tool(
            name="docs_insert_horizontal_rule",
            description="Insert a horizontal rule/line at a specific index",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "index": {"type": "integer", "description": "Position to insert horizontal rule"}
                },
                "required": ["document_id", "index"]
            }
        ),
        Tool(
            name="docs_apply_heading_style",
            description="Apply heading style (H1, H2, etc.) to paragraphs",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "start_index": {"type": "integer", "description": "Start index of paragraph range"},
                    "end_index": {"type": "integer", "description": "End index of paragraph range"},
                    "heading_level": {"type": "string", "description": "Heading level: 'TITLE', 'SUBTITLE', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6', 'NORMAL_TEXT'"}
                },
                "required": ["document_id", "start_index", "end_index", "heading_level"]
            }
        ),
        # NEW in v10: Advanced Docs tools - Batch Operations
        Tool(
            name="docs_batch_update",
            description="Execute multiple document requests in one call. Advanced tool for complex operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "document_id": {"type": "string", "description": "The document ID"},
                    "requests": {"type": "array", "items": {"type": "object"}, "description": "Array of request objects (see Google Docs API batchUpdate docs)"}
                },
                "required": ["document_id", "requests"]
            }
        ),

        # ==================== GOOGLE SHEETS ====================
        Tool(
            name="google_sheets_read",
            description="Read data from a Google Sheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "range": {"type": "string", "description": "A1 notation range (e.g., 'Sheet1!A1:D10')"}
                },
                "required": ["spreadsheet_id", "range"]
            }
        ),
        Tool(
            name="google_sheets_write",
            description="Write data to a Google Sheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "range": {"type": "string", "description": "A1 notation range"},
                    "values": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}, "description": "2D array of values"}
                },
                "required": ["spreadsheet_id", "range", "values"]
            }
        ),
        Tool(
            name="google_sheets_append",
            description="Append rows to a Google Sheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "range": {"type": "string", "description": "Range to append to (e.g., 'Sheet1!A:D')"},
                    "values": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}, "description": "2D array of rows"}
                },
                "required": ["spreadsheet_id", "range", "values"]
            }
        ),
        Tool(
            name="sheets_create",
            description="Create a new Google Spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string", "description": "Spreadsheet title"},
                    "sheet_titles": {"type": "array", "items": {"type": "string"}, "description": "Optional list of sheet/tab names to create"},
                    "folder_id": {"type": "string", "description": "Optional folder ID to create in"}
                },
                "required": ["title"]
            }
        ),
        Tool(
            name="sheets_get_metadata",
            description="Get spreadsheet metadata (title, sheets, dimensions)",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"}
                },
                "required": ["spreadsheet_id"]
            }
        ),
        Tool(
            name="sheets_clear",
            description="Clear a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "range": {"type": "string", "description": "A1 notation range to clear (e.g., 'Sheet1!A1:D10')"}
                },
                "required": ["spreadsheet_id", "range"]
            }
        ),
        Tool(
            name="sheets_add_sheet",
            description="Add a new sheet/tab to existing spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "title": {"type": "string", "description": "Name for the new sheet"}
                },
                "required": ["spreadsheet_id", "title"]
            }
        ),
        Tool(
            name="sheets_delete_sheet",
            description="Delete a sheet/tab from spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (not name) - get from sheets_get_metadata"}
                },
                "required": ["spreadsheet_id", "sheet_id"]
            }
        ),
        # NEW in v9: Advanced Sheets tools
        Tool(
            name="sheets_batch_update",
            description="Execute multiple batch update requests. Advanced tool for complex operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "requests": {"type": "array", "items": {"type": "object"}, "description": "Array of request objects (see Google Sheets API batchUpdate docs)"}
                },
                "required": ["spreadsheet_id", "requests"]
            }
        ),
        Tool(
            name="sheets_rename_sheet",
            description="Rename a sheet/tab",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "new_title": {"type": "string", "description": "New name for the sheet"}
                },
                "required": ["spreadsheet_id", "sheet_id", "new_title"]
            }
        ),
        Tool(
            name="sheets_format_cells",
            description="Format cells (bold, colors, borders, fonts, alignment). Colors use RGB 0-1 scale.",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "range": {"type": "string", "description": "A1 notation range (e.g., 'A1:D1' or 'Sheet1!A1:D1')"},
                    "bold": {"type": "boolean", "description": "Make text bold"},
                    "italic": {"type": "boolean", "description": "Make text italic"},
                    "font_size": {"type": "integer", "description": "Font size in points"},
                    "font_family": {"type": "string", "description": "Font family (e.g., 'Arial', 'Times New Roman')"},
                    "text_color": {"type": "object", "description": "Text color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "background_color": {"type": "object", "description": "Background color {red: 0-1, green: 0-1, blue: 0-1}"},
                    "horizontal_alignment": {"type": "string", "description": "Alignment: 'LEFT', 'CENTER', 'RIGHT'"},
                    "vertical_alignment": {"type": "string", "description": "Alignment: 'TOP', 'MIDDLE', 'BOTTOM'"},
                    "wrap_strategy": {"type": "string", "description": "Text wrapping: 'OVERFLOW_CELL', 'CLIP', 'WRAP'"},
                    "borders": {"type": "object", "description": "Border config {top, bottom, left, right} each with {style, color}. Styles: 'SOLID', 'DASHED', 'DOTTED', 'DOUBLE'"}
                },
                "required": ["spreadsheet_id", "sheet_id", "range"]
            }
        ),
        Tool(
            name="sheets_set_column_width",
            description="Set the width of columns",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based, A=0, B=1, etc.)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"},
                    "width": {"type": "integer", "description": "Width in pixels"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_column", "end_column", "width"]
            }
        ),
        Tool(
            name="sheets_set_row_height",
            description="Set the height of rows",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "height": {"type": "integer", "description": "Height in pixels"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "end_row", "height"]
            }
        ),
        Tool(
            name="sheets_freeze_rows_columns",
            description="Freeze rows and/or columns",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "frozen_rows": {"type": "integer", "description": "Number of rows to freeze (0 to unfreeze)"},
                    "frozen_columns": {"type": "integer", "description": "Number of columns to freeze (0 to unfreeze)"}
                },
                "required": ["spreadsheet_id", "sheet_id"]
            }
        ),
        Tool(
            name="sheets_merge_cells",
            description="Merge a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"},
                    "merge_type": {"type": "string", "description": "Merge type: 'MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS' (default: MERGE_ALL)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "end_row", "start_column", "end_column"]
            }
        ),
        Tool(
            name="sheets_unmerge_cells",
            description="Unmerge a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "end_row", "start_column", "end_column"]
            }
        ),
        Tool(
            name="sheets_add_filter",
            description="Add a filter view to a sheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based, usually 0 for header row)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive, omit to include all rows)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "start_column", "end_column"]
            }
        ),
        Tool(
            name="sheets_data_validation",
            description="Add data validation (dropdowns, number ranges, etc.)",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"},
                    "validation_type": {"type": "string", "description": "Type: 'ONE_OF_LIST', 'NUMBER_BETWEEN', 'NUMBER_GREATER', 'NUMBER_LESS', 'DATE_VALID', 'TEXT_CONTAINS', 'CUSTOM_FORMULA'"},
                    "values": {"type": "array", "items": {"type": "string"}, "description": "For ONE_OF_LIST: list of allowed values"},
                    "min_value": {"type": "number", "description": "For NUMBER_BETWEEN/GREATER: minimum value"},
                    "max_value": {"type": "number", "description": "For NUMBER_BETWEEN/LESS: maximum value"},
                    "custom_formula": {"type": "string", "description": "For CUSTOM_FORMULA: formula like '=A1>0'"},
                    "show_dropdown": {"type": "boolean", "description": "Show dropdown arrow for list (default true)"},
                    "strict": {"type": "boolean", "description": "Reject invalid input (default true)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "end_row", "start_column", "end_column", "validation_type"]
            }
        ),
        Tool(
            name="sheets_conditional_formatting",
            description="Add conditional formatting rules",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"},
                    "rule_type": {"type": "string", "description": "Type: 'TEXT_CONTAINS', 'TEXT_EQ', 'NUMBER_GREATER', 'NUMBER_LESS', 'NUMBER_BETWEEN', 'CUSTOM_FORMULA'"},
                    "values": {"type": "array", "items": {"type": "string"}, "description": "Values for comparison (1-2 values depending on rule type)"},
                    "custom_formula": {"type": "string", "description": "For CUSTOM_FORMULA: formula like '=$A1>100'"},
                    "background_color": {"type": "object", "description": "Background color when rule matches {red: 0-1, green: 0-1, blue: 0-1}"},
                    "text_color": {"type": "object", "description": "Text color when rule matches {red: 0-1, green: 0-1, blue: 0-1}"},
                    "bold": {"type": "boolean", "description": "Make text bold when rule matches"}
                },
                "required": ["spreadsheet_id", "sheet_id", "start_row", "end_row", "start_column", "end_column", "rule_type"]
            }
        ),
        Tool(
            name="sheets_named_range",
            description="Create or manage named ranges",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "name": {"type": "string", "description": "Name for the range (must be unique, no spaces)"},
                    "start_row": {"type": "integer", "description": "Start row index (0-based)"},
                    "end_row": {"type": "integer", "description": "End row index (exclusive)"},
                    "start_column": {"type": "integer", "description": "Start column index (0-based)"},
                    "end_column": {"type": "integer", "description": "End column index (exclusive)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "name", "start_row", "end_row", "start_column", "end_column"]
            }
        ),
        Tool(
            name="sheets_auto_resize",
            description="Auto-resize columns or rows to fit content",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {"type": "string", "description": "The spreadsheet ID"},
                    "sheet_id": {"type": "integer", "description": "The sheet ID (get from sheets_get_metadata)"},
                    "dimension": {"type": "string", "description": "Dimension to resize: 'COLUMNS' or 'ROWS'"},
                    "start_index": {"type": "integer", "description": "Start index (0-based)"},
                    "end_index": {"type": "integer", "description": "End index (exclusive)"}
                },
                "required": ["spreadsheet_id", "sheet_id", "dimension", "start_index", "end_index"]
            }
        ),

        # ==================== GOOGLE DRIVE ====================
        Tool(
            name="google_drive_list",
            description="List files in a folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_id": {"type": "string", "description": "Folder ID (use 'root' for root folder)"},
                    "query": {"type": "string", "description": "Optional name filter"}
                },
                "required": []
            }
        ),
        Tool(
            name="google_drive_search",
            description="Search for files across all of Drive",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Search query (searches name and content)"},
                    "file_type": {"type": "string", "description": "Optional: 'document', 'spreadsheet', 'presentation', 'folder', 'pdf'"},
                    "max_results": {"type": "integer", "description": "Maximum results (default 20)"}
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="google_drive_get",
            description="Get detailed metadata for a file",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "The file ID"}
                },
                "required": ["file_id"]
            }
        ),
        Tool(
            name="google_drive_create_folder",
            description="Create a new folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Folder name"},
                    "parent_id": {"type": "string", "description": "Optional parent folder ID"}
                },
                "required": ["name"]
            }
        ),
        Tool(
            name="google_drive_copy",
            description="Copy a file (great for templates)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "Source file ID to copy"},
                    "new_name": {"type": "string", "description": "Name for the copy"},
                    "folder_id": {"type": "string", "description": "Optional destination folder ID"}
                },
                "required": ["file_id", "new_name"]
            }
        ),
        Tool(
            name="google_drive_move",
            description="Move a file to a different folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to move"},
                    "new_parent_id": {"type": "string", "description": "Destination folder ID"}
                },
                "required": ["file_id", "new_parent_id"]
            }
        ),
        Tool(
            name="google_drive_rename",
            description="Rename a file or folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to rename"},
                    "new_name": {"type": "string", "description": "New name"}
                },
                "required": ["file_id", "new_name"]
            }
        ),
        Tool(
            name="google_drive_delete",
            description="Delete a file or folder (moves to trash)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to delete"}
                },
                "required": ["file_id"]
            }
        ),
        Tool(
            name="google_drive_share",
            description="Share a file with specific people",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to share"},
                    "email": {"type": "string", "description": "Email address to share with"},
                    "role": {"type": "string", "description": "Permission: 'reader', 'commenter', or 'writer'"},
                    "notify": {"type": "boolean", "description": "Send notification email (default true)"}
                },
                "required": ["file_id", "email", "role"]
            }
        ),
        Tool(
            name="google_drive_permissions",
            description="List who has access to a file",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to check permissions"}
                },
                "required": ["file_id"]
            }
        ),
        Tool(
            name="google_drive_export",
            description="Export a Google Doc/Sheet/Slides to PDF, DOCX, XLSX, etc.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "Google Doc/Sheet/Slides ID"},
                    "export_format": {"type": "string", "description": "'pdf', 'docx', 'xlsx', 'pptx', 'txt', 'csv'"},
                    "output_path": {"type": "string", "description": "Local path to save the exported file"}
                },
                "required": ["file_id", "export_format", "output_path"]
            }
        ),
        Tool(
            name="google_drive_upload",
            description="Upload a local file to Google Drive",
            inputSchema={
                "type": "object",
                "properties": {
                    "local_path": {"type": "string", "description": "Path to local file"},
                    "name": {"type": "string", "description": "Name for the file in Drive (optional)"},
                    "folder_id": {"type": "string", "description": "Optional destination folder ID"},
                    "convert": {"type": "boolean", "description": "Convert to Google format"}
                },
                "required": ["local_path"]
            }
        ),
        Tool(
            name="google_drive_download",
            description="Download a file from Google Drive",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "File ID to download"},
                    "output_path": {"type": "string", "description": "Local path to save the file"}
                },
                "required": ["file_id", "output_path"]
            }
        ),
        
        # ==================== GOOGLE CALENDAR ====================
        Tool(
            name="google_calendar_list",
            description="List upcoming calendar events",
            inputSchema={
                "type": "object",
                "properties": {
                    "max_results": {"type": "integer", "description": "Maximum events (default 10)"},
                    "days_ahead": {"type": "integer", "description": "Days to look ahead (default 7)"}
                },
                "required": []
            }
        ),
        Tool(
            name="google_calendar_get",
            description="Get details of a specific event",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_id": {"type": "string", "description": "The event ID"}
                },
                "required": ["event_id"]
            }
        ),
        Tool(
            name="google_calendar_create",
            description="Create a new calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "summary": {"type": "string", "description": "Event title"},
                    "description": {"type": "string", "description": "Event description"},
                    "start_time": {"type": "string", "description": "Start time (ISO format)"},
                    "end_time": {"type": "string", "description": "End time (ISO format)"},
                    "attendees": {"type": "array", "items": {"type": "string"}, "description": "Attendee emails"},
                    "recurrence": {"type": "array", "items": {"type": "string"}, "description": "Recurrence rules in RRULE format (e.g. ['RRULE:FREQ=WEEKLY;BYDAY=MO'])"}
                },
                "required": ["summary", "start_time", "end_time"]
            }
        ),
        Tool(
            name="calendar_event_update",
            description="Update an existing calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_id": {"type": "string", "description": "The event ID to update"},
                    "summary": {"type": "string", "description": "New event title"},
                    "description": {"type": "string", "description": "New event description"},
                    "start_time": {"type": "string", "description": "New start time (ISO format)"},
                    "end_time": {"type": "string", "description": "New end time (ISO format)"},
                    "location": {"type": "string", "description": "Event location"},
                    "attendees": {"type": "array", "items": {"type": "string"}, "description": "New list of attendee emails"},
                    "recurrence": {"type": "array", "items": {"type": "string"}, "description": "Recurrence rules in RRULE format (e.g. ['RRULE:FREQ=WEEKLY;BYDAY=MO'])"}
                },
                "required": ["event_id"]
            }
        ),
        Tool(
            name="calendar_event_delete",
            description="Delete a calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "event_id": {"type": "string", "description": "The event ID to delete"}
                },
                "required": ["event_id"]
            }
        ),
        Tool(
            name="calendar_quick_add",
            description="Create a calendar event using natural language (e.g., 'Lunch with John tomorrow at noon')",
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "Natural language description of the event"}
                },
                "required": ["text"]
            }
        ),
        Tool(
            name="calendar_list_calendars",
            description="List all calendars the user has access to",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),

        # ==================== GMAIL (ENHANCED in v4) ====================
        Tool(
            name="gmail_search",
            description="Search emails using Gmail query syntax",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Gmail search query (e.g., 'from:someone@example.com', 'is:unread', 'subject:meeting')"},
                    "max_results": {"type": "integer", "description": "Maximum emails (default 10)"}
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="gmail_read",
            description="Read a specific email by ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "The email message ID"}
                },
                "required": ["message_id"]
            }
        ),
        Tool(
            name="gmail_send",
            description="Send a new email",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email"},
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body (plain text)"},
                    "cc": {"type": "string", "description": "CC recipients (comma-separated)"},
                    "bcc": {"type": "string", "description": "BCC recipients (comma-separated)"}
                },
                "required": ["to", "subject", "body"]
            }
        ),
        # NEW in v4: Draft tools
        Tool(
            name="gmail_draft_create",
            description="Create a draft email (saved but not sent)",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email"},
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body (plain text)"},
                    "cc": {"type": "string", "description": "CC recipients (comma-separated)"},
                    "bcc": {"type": "string", "description": "BCC recipients (comma-separated)"}
                },
                "required": ["to", "subject", "body"]
            }
        ),
        Tool(
            name="gmail_draft_list",
            description="List all draft emails",
            inputSchema={
                "type": "object",
                "properties": {
                    "max_results": {"type": "integer", "description": "Maximum drafts to return (default 10)"}
                },
                "required": []
            }
        ),
        Tool(
            name="gmail_draft_send",
            description="Send an existing draft",
            inputSchema={
                "type": "object",
                "properties": {
                    "draft_id": {"type": "string", "description": "The draft ID to send"}
                },
                "required": ["draft_id"]
            }
        ),
        # NEW in v4: Reply tool
        Tool(
            name="gmail_reply",
            description="Reply to an existing email thread",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "The message ID to reply to"},
                    "body": {"type": "string", "description": "Reply body (plain text)"},
                    "reply_all": {"type": "boolean", "description": "Reply to all recipients (default false)"}
                },
                "required": ["message_id", "body"]
            }
        ),
        # NEW in v4: Labels tools
        Tool(
            name="gmail_labels_list",
            description="List all Gmail labels",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="gmail_message_modify",
            description="Modify a message: add/remove labels, archive, mark read/unread",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "The message ID to modify"},
                    "add_labels": {"type": "array", "items": {"type": "string"}, "description": "Label IDs to add (e.g., ['STARRED', 'IMPORTANT'])"},
                    "remove_labels": {"type": "array", "items": {"type": "string"}, "description": "Label IDs to remove (e.g., ['UNREAD', 'INBOX'])"}
                },
                "required": ["message_id"]
            }
        ),
        
        # ==================== GOOGLE SLIDES ====================
        Tool(
            name="google_slides_create",
            description="Create a new presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string", "description": "Presentation title"},
                    "folder_id": {"type": "string", "description": "Optional folder ID"}
                },
                "required": ["title"]
            }
        ),
        Tool(
            name="google_slides_add_slide",
            description="Add a slide to a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "title": {"type": "string", "description": "Slide title"},
                    "body": {"type": "string", "description": "Slide body text"}
                },
                "required": ["presentation_id"]
            }
        ),
        Tool(
            name="google_slides_read",
            description="Read slides from a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"}
                },
                "required": ["presentation_id"]
            }
        ),
        Tool(
            name="slides_delete_slide",
            description="Delete a slide from a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_id": {"type": "string", "description": "The slide object ID to delete (get from google_slides_read)"}
                },
                "required": ["presentation_id", "slide_id"]
            }
        ),
        Tool(
            name="slides_replace_text",
            description="Find and replace text across all slides in a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "find_text": {"type": "string", "description": "Text to find"},
                    "replace_text": {"type": "string", "description": "Text to replace with"},
                    "match_case": {"type": "boolean", "description": "Case-sensitive match (default false)"}
                },
                "required": ["presentation_id", "find_text", "replace_text"]
            }
        ),
        Tool(
            name="slides_duplicate_slide",
            description="Duplicate a slide in a presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"},
                    "slide_id": {"type": "string", "description": "The slide object ID to duplicate"}
                },
                "required": ["presentation_id", "slide_id"]
            }
        ),
        Tool(
            name="slides_get_details",
            description="Get detailed information about slides including object IDs for editing",
            inputSchema={
                "type": "object",
                "properties": {
                    "presentation_id": {"type": "string", "description": "Presentation ID"}
                },
                "required": ["presentation_id"]
            }
        ),
    ]



@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    
    if not google_client.creds:
        return [TextContent(type="text", text="Not authenticated. Please restart the server.")]
    
    try:
        # Google Docs
        if name == "google_docs_create":
            return await create_doc(arguments)
        elif name == "google_docs_read":
            return await read_doc(arguments)
        elif name == "google_docs_append":
            return await append_doc(arguments)
        elif name == "docs_replace_text":
            return await replace_text_doc(arguments)
        elif name == "docs_insert_text":
            return await insert_text_doc(arguments)
        elif name == "docs_delete_content":
            return await delete_content_doc(arguments)
        elif name == "docs_get_structure":
            return await get_doc_structure(arguments)
        # NEW in v10: Advanced Docs tools
        elif name == "docs_insert_table":
            return await insert_table(arguments)
        elif name == "docs_insert_table_row":
            return await insert_table_row(arguments)
        elif name == "docs_insert_table_column":
            return await insert_table_column(arguments)
        elif name == "docs_delete_table_row":
            return await delete_table_row(arguments)
        elif name == "docs_delete_table_column":
            return await delete_table_column(arguments)
        elif name == "docs_write_table_cell":
            return await write_table_cell(arguments)
        elif name == "docs_write_table_bulk":
            return await write_table_bulk(arguments)
        elif name == "docs_merge_table_cells":
            return await merge_table_cells(arguments)
        elif name == "docs_unmerge_table_cells":
            return await unmerge_table_cells(arguments)
        elif name == "docs_format_table_cell":
            return await format_table_cell(arguments)
        elif name == "docs_set_table_column_width":
            return await set_table_column_width(arguments)
        elif name == "docs_set_table_row_height":
            return await set_table_row_height(arguments)
        elif name == "docs_format_text":
            return await format_text(arguments)
        elif name == "docs_format_paragraph":
            return await format_paragraph(arguments)
        elif name == "docs_create_bullet_list":
            return await create_bullet_list(arguments)
        elif name == "docs_create_numbered_list":
            return await create_numbered_list(arguments)
        elif name == "docs_remove_bullets":
            return await remove_bullets(arguments)
        elif name == "docs_insert_page_break":
            return await insert_page_break(arguments)
        elif name == "docs_insert_section_break":
            return await insert_section_break(arguments)
        elif name == "docs_insert_horizontal_rule":
            return await insert_horizontal_rule(arguments)
        elif name == "docs_apply_heading_style":
            return await apply_heading_style(arguments)
        elif name == "docs_batch_update":
            return await docs_batch_update(arguments)

        # Google Sheets
        elif name == "google_sheets_read":
            return await read_sheet(arguments)
        elif name == "google_sheets_write":
            return await write_sheet(arguments)
        elif name == "google_sheets_append":
            return await append_sheet(arguments)
        elif name == "sheets_create":
            return await create_sheet(arguments)
        elif name == "sheets_get_metadata":
            return await get_sheet_metadata(arguments)
        elif name == "sheets_clear":
            return await clear_sheet(arguments)
        elif name == "sheets_add_sheet":
            return await add_sheet_tab(arguments)
        elif name == "sheets_delete_sheet":
            return await delete_sheet_tab(arguments)
        # NEW in v9: Advanced Sheets tools
        elif name == "sheets_batch_update":
            return await sheets_batch_update(arguments)
        elif name == "sheets_rename_sheet":
            return await rename_sheet(arguments)
        elif name == "sheets_format_cells":
            return await format_cells(arguments)
        elif name == "sheets_set_column_width":
            return await set_column_width(arguments)
        elif name == "sheets_set_row_height":
            return await set_row_height(arguments)
        elif name == "sheets_freeze_rows_columns":
            return await freeze_rows_columns(arguments)
        elif name == "sheets_merge_cells":
            return await merge_cells(arguments)
        elif name == "sheets_unmerge_cells":
            return await unmerge_cells(arguments)
        elif name == "sheets_add_filter":
            return await add_filter(arguments)
        elif name == "sheets_data_validation":
            return await data_validation(arguments)
        elif name == "sheets_conditional_formatting":
            return await conditional_formatting(arguments)
        elif name == "sheets_named_range":
            return await named_range(arguments)
        elif name == "sheets_auto_resize":
            return await auto_resize(arguments)

        # Google Drive
        elif name == "google_drive_list":
            return await list_drive(arguments)
        elif name == "google_drive_search":
            return await search_drive(arguments)
        elif name == "google_drive_get":
            return await get_drive_file(arguments)
        elif name == "google_drive_create_folder":
            return await create_folder(arguments)
        elif name == "google_drive_copy":
            return await copy_file(arguments)
        elif name == "google_drive_move":
            return await move_file(arguments)
        elif name == "google_drive_rename":
            return await rename_file(arguments)
        elif name == "google_drive_delete":
            return await delete_file(arguments)
        elif name == "google_drive_share":
            return await share_file(arguments)
        elif name == "google_drive_permissions":
            return await list_permissions(arguments)
        elif name == "google_drive_export":
            return await export_file(arguments)
        elif name == "google_drive_upload":
            return await upload_file(arguments)
        elif name == "google_drive_download":
            return await download_file(arguments)
        
        # Google Calendar
        elif name == "google_calendar_list":
            return await list_calendar_events(arguments)
        elif name == "google_calendar_get":
            return await get_calendar_event(arguments)
        elif name == "google_calendar_create":
            return await create_calendar_event(arguments)
        elif name == "calendar_event_update":
            return await update_calendar_event(arguments)
        elif name == "calendar_event_delete":
            return await delete_calendar_event(arguments)
        elif name == "calendar_quick_add":
            return await quick_add_event(arguments)
        elif name == "calendar_list_calendars":
            return await list_calendars(arguments)

        # Gmail (including new v4 tools)
        elif name == "gmail_search":
            return await search_gmail(arguments)
        elif name == "gmail_read":
            return await read_gmail(arguments)
        elif name == "gmail_send":
            return await send_gmail(arguments)
        elif name == "gmail_draft_create":
            return await create_gmail_draft(arguments)
        elif name == "gmail_draft_list":
            return await list_gmail_drafts(arguments)
        elif name == "gmail_draft_send":
            return await send_gmail_draft(arguments)
        elif name == "gmail_reply":
            return await reply_gmail(arguments)
        elif name == "gmail_labels_list":
            return await list_gmail_labels(arguments)
        elif name == "gmail_message_modify":
            return await modify_gmail_message(arguments)
        
        # Google Slides
        elif name == "google_slides_create":
            return await create_slides(arguments)
        elif name == "google_slides_add_slide":
            return await add_slide(arguments)
        elif name == "google_slides_read":
            return await read_slides(arguments)
        elif name == "slides_delete_slide":
            return await delete_slide(arguments)
        elif name == "slides_replace_text":
            return await replace_text_slides(arguments)
        elif name == "slides_duplicate_slide":
            return await duplicate_slide(arguments)
        elif name == "slides_get_details":
            return await get_slides_details(arguments)


        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]
    
    except HttpError as e:
        return [TextContent(type="text", text=f"Google API error: {e}")]
    except Exception as e:
        return [TextContent(type="text", text=f"Error: {e}")]


# ==================== GOOGLE DOCS HANDLERS ====================

async def create_doc(args: dict) -> list[TextContent]:
    doc = google_client.docs_service.documents().create(body={"title": args["title"]}).execute()
    doc_id = doc.get("documentId")
    
    if args.get("content"):
        requests = [{"insertText": {"location": {"index": 1}, "text": args["content"]}}]
        google_client.docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
    
    if args.get("folder_id"):
        google_client.drive_service.files().update(
            fileId=doc_id, addParents=args["folder_id"], removeParents='root', fields='id, parents',
            supportsAllDrives=True
        ).execute()

    return [TextContent(type="text", text=f"Created document '{args['title']}'\nID: {doc_id}\nURL: https://docs.google.com/document/d/{doc_id}/edit")]


async def read_doc(args: dict) -> list[TextContent]:
    doc = google_client.docs_service.documents().get(documentId=args["document_id"]).execute()
    
    content = []
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            for text_element in element["paragraph"].get("elements", []):
                if "textRun" in text_element:
                    content.append(text_element["textRun"].get("content", ""))
    
    return [TextContent(type="text", text=f"Title: {doc.get('title', 'Untitled')}\n\n{''.join(content)}")]


async def append_doc(args: dict) -> list[TextContent]:
    doc = google_client.docs_service.documents().get(documentId=args["document_id"]).execute()
    end_index = doc["body"]["content"][-1]["endIndex"] - 1

    requests = [{"insertText": {"location": {"index": end_index}, "text": args["content"]}}]
    google_client.docs_service.documents().batchUpdate(documentId=args["document_id"], body={"requests": requests}).execute()

    return [TextContent(type="text", text=f"Appended content to document {args['document_id']}")]


async def replace_text_doc(args: dict) -> list[TextContent]:
    requests = [{
        'replaceAllText': {
            'containsText': {
                'text': args['find_text'],
                'matchCase': args.get('match_case', False)
            },
            'replaceText': args['replace_text']
        }
    }]

    result = google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    occurrences = result.get('replies', [{}])[0].get('replaceAllText', {}).get('occurrencesChanged', 0)
    return [TextContent(type="text", text=f"Replaced {occurrences} occurrences of '{args['find_text']}' with '{args['replace_text']}'")]


async def insert_text_doc(args: dict) -> list[TextContent]:
    requests = [{
        'insertText': {
            'location': {'index': args['index']},
            'text': args['text']
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Inserted text at index {args['index']}")]


async def delete_content_doc(args: dict) -> list[TextContent]:
    requests = [{
        'deleteContentRange': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Deleted content from index {args['start_index']} to {args['end_index']}")]


async def get_doc_structure(args: dict) -> list[TextContent]:
    doc = google_client.docs_service.documents().get(documentId=args['document_id']).execute()

    output = [
        f"Title: {doc.get('title', 'Untitled')}",
        f"Document ID: {doc.get('documentId')}",
        "",
        "Content Structure:",
        "=================="
    ]

    for element in doc.get('body', {}).get('content', []):
        start_idx = element.get('startIndex', 0)
        end_idx = element.get('endIndex', 0)

        if 'paragraph' in element:
            # Extract text from paragraph
            para_text = ""
            for elem in element['paragraph'].get('elements', []):
                if 'textRun' in elem:
                    para_text += elem['textRun'].get('content', '')

            # Truncate long paragraphs
            preview = para_text.strip()[:50]
            if len(para_text.strip()) > 50:
                preview += "..."

            style = element['paragraph'].get('paragraphStyle', {}).get('namedStyleType', 'NORMAL_TEXT')
            output.append(f"[{start_idx}-{end_idx}] {style}: {repr(preview)}")

        elif 'table' in element:
            rows = element['table'].get('rows', 0)
            cols = element['table'].get('columns', 0)
            output.append(f"[{start_idx}-{end_idx}] TABLE: {rows}x{cols}")

        elif 'sectionBreak' in element:
            output.append(f"[{start_idx}-{end_idx}] SECTION_BREAK")

    return [TextContent(type="text", text="\n".join(output))]


# NEW in v10: Advanced Docs Handlers - Table Operations

def _get_table_info(doc: dict, table_start_index: int) -> dict:
    """Helper to get table information from a document given the table's start index."""
    for element in doc.get('body', {}).get('content', []):
        if element.get('startIndex') == table_start_index and 'table' in element:
            return element['table']
    return None


def _get_cell_content_indexes(doc: dict, table_start_index: int, row_index: int, column_index: int) -> tuple:
    """Helper to get the start and end indexes of a table cell's content."""
    for element in doc.get('body', {}).get('content', []):
        if element.get('startIndex') == table_start_index and 'table' in element:
            table = element['table']
            if row_index < len(table.get('tableRows', [])):
                row = table['tableRows'][row_index]
                if column_index < len(row.get('tableCells', [])):
                    cell = row['tableCells'][column_index]
                    # Find the content start and end within the cell
                    cell_content = cell.get('content', [])
                    if cell_content:
                        # Get the first paragraph's start index and last element's end index
                        start = cell_content[0].get('startIndex', 0)
                        end = cell_content[-1].get('endIndex', start + 1)
                        return (start, end)
    return (None, None)


async def insert_table(args: dict) -> list[TextContent]:
    """Insert a table at a specific index."""
    requests = [{
        'insertTable': {
            'rows': args['rows'],
            'columns': args['columns'],
            'location': {'index': args['index']}
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Inserted {args['rows']}x{args['columns']} table at index {args['index']}")]


async def insert_table_row(args: dict) -> list[TextContent]:
    """Insert row(s) in an existing table."""
    insert_below = args.get('insert_below', True)

    requests = [{
        'insertTableRow': {
            'tableCellLocation': {
                'tableStartLocation': {'index': args['table_start_index']},
                'rowIndex': args['row_index'],
                'columnIndex': 0
            },
            'insertBelow': insert_below
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    position = "below" if insert_below else "above"
    return [TextContent(type="text", text=f"Inserted row {position} row {args['row_index']}")]


async def insert_table_column(args: dict) -> list[TextContent]:
    """Insert column(s) in an existing table."""
    insert_right = args.get('insert_right', True)

    requests = [{
        'insertTableColumn': {
            'tableCellLocation': {
                'tableStartLocation': {'index': args['table_start_index']},
                'rowIndex': 0,
                'columnIndex': args['column_index']
            },
            'insertRight': insert_right
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    position = "right of" if insert_right else "left of"
    return [TextContent(type="text", text=f"Inserted column {position} column {args['column_index']}")]


async def delete_table_row(args: dict) -> list[TextContent]:
    """Delete row(s) from a table."""
    requests = [{
        'deleteTableRow': {
            'tableCellLocation': {
                'tableStartLocation': {'index': args['table_start_index']},
                'rowIndex': args['row_index'],
                'columnIndex': 0
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Deleted row {args['row_index']}")]


async def delete_table_column(args: dict) -> list[TextContent]:
    """Delete column(s) from a table."""
    requests = [{
        'deleteTableColumn': {
            'tableCellLocation': {
                'tableStartLocation': {'index': args['table_start_index']},
                'rowIndex': 0,
                'columnIndex': args['column_index']
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Deleted column {args['column_index']}")]


async def write_table_cell(args: dict) -> list[TextContent]:
    """Write text to a specific table cell."""
    # First, get the document to find the cell's content indexes
    doc = google_client.docs_service.documents().get(documentId=args['document_id']).execute()

    start_idx, end_idx = _get_cell_content_indexes(
        doc,
        args['table_start_index'],
        args['row_index'],
        args['column_index']
    )

    if start_idx is None:
        return [TextContent(type="text", text="Error: Could not find the specified table cell")]

    requests = []
    replace_existing = args.get('replace_existing', True)

    if replace_existing and end_idx > start_idx + 1:
        # Delete existing content (but keep the newline at the end)
        requests.append({
            'deleteContentRange': {
                'range': {
                    'startIndex': start_idx,
                    'endIndex': end_idx - 1  # Keep the trailing newline
                }
            }
        })
        # Insert new text at the start
        requests.append({
            'insertText': {
                'location': {'index': start_idx},
                'text': args['text']
            }
        })
    else:
        # Just insert at the start of the cell
        requests.append({
            'insertText': {
                'location': {'index': start_idx},
                'text': args['text']
            }
        })

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Wrote text to cell ({args['row_index']}, {args['column_index']})")]


async def write_table_bulk(args: dict) -> list[TextContent]:
    """Write text to multiple table cells in a single API call."""
    # Get the document once to find all cell indexes
    doc = google_client.docs_service.documents().get(documentId=args['document_id']).execute()
    
    replace_existing = args.get('replace_existing', True)
    cells = args['cells']
    table_start_index = args['table_start_index']
    
    # Collect all cell indexes first
    cell_data = []
    for cell in cells:
        start_idx, end_idx = _get_cell_content_indexes(
            doc,
            table_start_index,
            cell['row'],
            cell['column']
        )
        if start_idx is not None:
            cell_data.append({
                'row': cell['row'],
                'column': cell['column'],
                'text': cell['text'],
                'start_idx': start_idx,
                'end_idx': end_idx
            })
    
    if not cell_data:
        return [TextContent(type="text", text="Error: Could not find any of the specified table cells")]
    
    # Sort by start_idx in descending order to avoid index shifting issues
    # When we modify content, indexes after the modification point shift
    # By processing from highest index to lowest, we avoid this problem
    cell_data.sort(key=lambda x: x['start_idx'], reverse=True)
    
    requests = []
    for cell in cell_data:
        start_idx = cell['start_idx']
        end_idx = cell['end_idx']
        
        if replace_existing and end_idx > start_idx + 1:
            # Delete existing content (but keep the newline at the end)
            requests.append({
                'deleteContentRange': {
                    'range': {
                        'startIndex': start_idx,
                        'endIndex': end_idx - 1  # Keep the trailing newline
                    }
                }
            })
        # Insert new text at the start
        requests.append({
            'insertText': {
                'location': {'index': start_idx},
                'text': cell['text']
            }
        })
    
    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()
    
    return [TextContent(type="text", text=f"Wrote text to {len(cell_data)} cells in a single batch operation")]


async def merge_table_cells(args: dict) -> list[TextContent]:
    """Merge table cells."""
    requests = [{
        'mergeTableCells': {
            'tableRange': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': args['table_start_index']},
                    'rowIndex': args['row_start'],
                    'columnIndex': args['column_start']
                },
                'rowSpan': args['row_end'] - args['row_start'],
                'columnSpan': args['column_end'] - args['column_start']
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Merged cells from ({args['row_start']}, {args['column_start']}) to ({args['row_end']-1}, {args['column_end']-1})")]


async def unmerge_table_cells(args: dict) -> list[TextContent]:
    """Unmerge table cells."""
    requests = [{
        'unmergeTableCells': {
            'tableRange': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': args['table_start_index']},
                    'rowIndex': args['row_start'],
                    'columnIndex': args['column_start']
                },
                'rowSpan': args['row_end'] - args['row_start'],
                'columnSpan': args['column_end'] - args['column_start']
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Unmerged cells in range ({args['row_start']}, {args['column_start']}) to ({args['row_end']-1}, {args['column_end']-1})")]


# NEW in v10: Advanced Docs Handlers - Table Formatting

async def format_table_cell(args: dict) -> list[TextContent]:
    """Format table cell (background color, borders, padding)."""
    table_cell_style = {}
    fields = []

    # Background color
    if args.get('background_color'):
        table_cell_style['backgroundColor'] = {'color': {'rgbColor': args['background_color']}}
        fields.append('backgroundColor')

    # Border configuration
    if args.get('border_color') or args.get('border_width'):
        border_style = {}
        if args.get('border_color'):
            border_style['color'] = {'color': {'rgbColor': args['border_color']}}
        if args.get('border_width'):
            border_style['width'] = {'magnitude': args['border_width'], 'unit': 'PT'}
        border_style['dashStyle'] = 'SOLID'

        # Apply to all borders
        for side in ['borderTop', 'borderBottom', 'borderLeft', 'borderRight']:
            table_cell_style[side] = border_style
            fields.append(side)

    # Padding
    for padding_name, padding_field in [
        ('padding_top', 'paddingTop'),
        ('padding_bottom', 'paddingBottom'),
        ('padding_left', 'paddingLeft'),
        ('padding_right', 'paddingRight')
    ]:
        if args.get(padding_name):
            table_cell_style[padding_field] = {'magnitude': args[padding_name], 'unit': 'PT'}
            fields.append(padding_field)

    # Vertical alignment
    if args.get('vertical_alignment'):
        table_cell_style['contentAlignment'] = args['vertical_alignment']
        fields.append('contentAlignment')

    if not table_cell_style:
        return [TextContent(type="text", text="No formatting options specified")]

    requests = [{
        'updateTableCellStyle': {
            'tableRange': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': args['table_start_index']},
                    'rowIndex': args['row_start'],
                    'columnIndex': args['column_start']
                },
                'rowSpan': args['row_end'] - args['row_start'],
                'columnSpan': args['column_end'] - args['column_start']
            },
            'tableCellStyle': table_cell_style,
            'fields': ','.join(fields)
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Formatted cells from ({args['row_start']}, {args['column_start']}) to ({args['row_end']-1}, {args['column_end']-1})")]


async def set_table_column_width(args: dict) -> list[TextContent]:
    """Set table column widths."""
    requests = [{
        'updateTableColumnProperties': {
            'tableStartLocation': {'index': args['table_start_index']},
            'columnIndices': [args['column_index']],
            'tableColumnProperties': {
                'widthType': 'FIXED_WIDTH',
                'width': {'magnitude': args['width'], 'unit': 'PT'}
            },
            'fields': 'widthType,width'
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Set column {args['column_index']} width to {args['width']} points")]


async def set_table_row_height(args: dict) -> list[TextContent]:
    """Set minimum table row height."""
    requests = [{
        'updateTableRowStyle': {
            'tableStartLocation': {'index': args['table_start_index']},
            'rowIndices': [args['row_index']],
            'tableRowStyle': {
                'minRowHeight': {'magnitude': args['min_height'], 'unit': 'PT'}
            },
            'fields': 'minRowHeight'
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Set row {args['row_index']} minimum height to {args['min_height']} points")]


# NEW in v10: Advanced Docs Handlers - Text Formatting

async def format_text(args: dict) -> list[TextContent]:
    """Format text range (bold, italic, underline, color, font, size)."""
    text_style = {}
    fields = []

    # Basic text styles
    if args.get('bold') is not None:
        text_style['bold'] = args['bold']
        fields.append('bold')
    if args.get('italic') is not None:
        text_style['italic'] = args['italic']
        fields.append('italic')
    if args.get('underline') is not None:
        text_style['underline'] = args['underline']
        fields.append('underline')
    if args.get('strikethrough') is not None:
        text_style['strikethrough'] = args['strikethrough']
        fields.append('strikethrough')

    # Font settings
    if args.get('font_size'):
        text_style['fontSize'] = {'magnitude': args['font_size'], 'unit': 'PT'}
        fields.append('fontSize')
    if args.get('font_family'):
        text_style['weightedFontFamily'] = {'fontFamily': args['font_family']}
        fields.append('weightedFontFamily')

    # Colors
    if args.get('foreground_color'):
        text_style['foregroundColor'] = {'color': {'rgbColor': args['foreground_color']}}
        fields.append('foregroundColor')
    if args.get('background_color'):
        text_style['backgroundColor'] = {'color': {'rgbColor': args['background_color']}}
        fields.append('backgroundColor')

    # Link
    if args.get('link_url'):
        text_style['link'] = {'url': args['link_url']}
        fields.append('link')

    if not text_style:
        return [TextContent(type="text", text="No formatting options specified")]

    requests = [{
        'updateTextStyle': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            },
            'textStyle': text_style,
            'fields': ','.join(fields)
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Formatted text from index {args['start_index']} to {args['end_index']}")]


async def format_paragraph(args: dict) -> list[TextContent]:
    """Format paragraph (alignment, spacing, indentation)."""
    paragraph_style = {}
    fields = []

    # Alignment
    if args.get('alignment'):
        paragraph_style['alignment'] = args['alignment']
        fields.append('alignment')

    # Line spacing (convert to percentage - 100 = single, 150 = 1.5, 200 = double)
    if args.get('line_spacing'):
        paragraph_style['lineSpacing'] = args['line_spacing'] * 100
        fields.append('lineSpacing')

    # Spacing above and below
    if args.get('space_above'):
        paragraph_style['spaceAbove'] = {'magnitude': args['space_above'], 'unit': 'PT'}
        fields.append('spaceAbove')
    if args.get('space_below'):
        paragraph_style['spaceBelow'] = {'magnitude': args['space_below'], 'unit': 'PT'}
        fields.append('spaceBelow')

    # Indentation
    if args.get('indent_first_line'):
        paragraph_style['indentFirstLine'] = {'magnitude': args['indent_first_line'], 'unit': 'PT'}
        fields.append('indentFirstLine')
    if args.get('indent_start'):
        paragraph_style['indentStart'] = {'magnitude': args['indent_start'], 'unit': 'PT'}
        fields.append('indentStart')
    if args.get('indent_end'):
        paragraph_style['indentEnd'] = {'magnitude': args['indent_end'], 'unit': 'PT'}
        fields.append('indentEnd')

    if not paragraph_style:
        return [TextContent(type="text", text="No formatting options specified")]

    requests = [{
        'updateParagraphStyle': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            },
            'paragraphStyle': paragraph_style,
            'fields': ','.join(fields)
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Formatted paragraph(s) from index {args['start_index']} to {args['end_index']}")]


async def create_bullet_list(args: dict) -> list[TextContent]:
    """Create bulleted list from paragraphs."""
    bullet_preset = args.get('bullet_preset', 'BULLET_DISC_CIRCLE_SQUARE')

    requests = [{
        'createParagraphBullets': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            },
            'bulletPreset': bullet_preset
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Created bullet list from index {args['start_index']} to {args['end_index']}")]


async def create_numbered_list(args: dict) -> list[TextContent]:
    """Create numbered list from paragraphs."""
    number_preset = args.get('number_preset', 'NUMBERED_DECIMAL_ALPHA_ROMAN')

    requests = [{
        'createParagraphBullets': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            },
            'bulletPreset': number_preset
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Created numbered list from index {args['start_index']} to {args['end_index']}")]


async def remove_bullets(args: dict) -> list[TextContent]:
    """Remove bullets/numbering from paragraphs."""
    requests = [{
        'deleteParagraphBullets': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            }
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Removed bullets/numbering from index {args['start_index']} to {args['end_index']}")]


# NEW in v10: Advanced Docs Handlers - Document Structure

async def insert_page_break(args: dict) -> list[TextContent]:
    """Insert a page break at a specific index."""
    requests = [{
        'insertPageBreak': {
            'location': {'index': args['index']}
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Inserted page break at index {args['index']}")]


async def insert_section_break(args: dict) -> list[TextContent]:
    """Insert a section break at a specific index."""
    section_type = args.get('section_type', 'NEXT_PAGE')

    requests = [{
        'insertSectionBreak': {
            'location': {'index': args['index']},
            'sectionType': section_type
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Inserted {section_type} section break at index {args['index']}")]


async def insert_horizontal_rule(args: dict) -> list[TextContent]:
    """Insert a horizontal rule/line at a specific index.

    Note: Google Docs doesn't have a direct 'insertHorizontalRule' API.
    We achieve this by inserting a paragraph with a bottom border.
    """
    # Insert a new paragraph with a horizontal rule style
    # We do this by inserting text and then formatting the paragraph with a bottom border
    requests = [
        # First insert a newline to create a paragraph
        {
            'insertText': {
                'location': {'index': args['index']},
                'text': '\n'
            }
        },
        # Then apply a bottom border to simulate a horizontal rule
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': args['index'],
                    'endIndex': args['index'] + 1
                },
                'paragraphStyle': {
                    'borderBottom': {
                        'color': {'color': {'rgbColor': {'red': 0, 'green': 0, 'blue': 0}}},
                        'width': {'magnitude': 1, 'unit': 'PT'},
                        'padding': {'magnitude': 1, 'unit': 'PT'},
                        'dashStyle': 'SOLID'
                    },
                    'spaceBelow': {'magnitude': 6, 'unit': 'PT'}
                },
                'fields': 'borderBottom,spaceBelow'
            }
        }
    ]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Inserted horizontal rule at index {args['index']}")]


async def apply_heading_style(args: dict) -> list[TextContent]:
    """Apply heading style (H1, H2, etc.) to paragraphs."""
    requests = [{
        'updateParagraphStyle': {
            'range': {
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            },
            'paragraphStyle': {
                'namedStyleType': args['heading_level']
            },
            'fields': 'namedStyleType'
        }
    }]

    google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Applied {args['heading_level']} style from index {args['start_index']} to {args['end_index']}")]


# NEW in v10: Advanced Docs Handlers - Batch Operations

async def docs_batch_update(args: dict) -> list[TextContent]:
    """Execute multiple document requests in one call."""
    result = google_client.docs_service.documents().batchUpdate(
        documentId=args['document_id'],
        body={'requests': args['requests']}
    ).execute()

    replies = result.get('replies', [])
    return [TextContent(type="text", text=f"Batch update completed. {len(replies)} operations executed.")]


# ==================== GOOGLE SHEETS HANDLERS ====================

async def read_sheet(args: dict) -> list[TextContent]:
    result = google_client.sheets_service.spreadsheets().values().get(
        spreadsheetId=args["spreadsheet_id"], range=args["range"]
    ).execute()
    
    values = result.get("values", [])
    if not values:
        return [TextContent(type="text", text="No data found.")]
    
    output = []
    for row in values:
        output.append(" | ".join(str(cell) for cell in row))
    
    return [TextContent(type="text", text="\n".join(output))]


async def write_sheet(args: dict) -> list[TextContent]:
    body = {"values": args["values"]}
    result = google_client.sheets_service.spreadsheets().values().update(
        spreadsheetId=args["spreadsheet_id"], range=args["range"],
        valueInputOption="USER_ENTERED", body=body
    ).execute()
    
    return [TextContent(type="text", text=f"Updated {result.get('updatedCells', 0)} cells")]


async def append_sheet(args: dict) -> list[TextContent]:
    body = {"values": args["values"]}
    result = google_client.sheets_service.spreadsheets().values().append(
        spreadsheetId=args["spreadsheet_id"], range=args["range"],
        valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body
    ).execute()

    return [TextContent(type="text", text=f"Appended {len(args['values'])} rows")]


async def create_sheet(args: dict) -> list[TextContent]:
    body = {'properties': {'title': args['title']}}

    if args.get('sheet_titles'):
        body['sheets'] = [{'properties': {'title': t}} for t in args['sheet_titles']]

    spreadsheet = google_client.sheets_service.spreadsheets().create(body=body).execute()
    ss_id = spreadsheet['spreadsheetId']

    if args.get('folder_id'):
        google_client.drive_service.files().update(
            fileId=ss_id,
            addParents=args['folder_id'],
            removeParents='root',
            fields='id, parents',
            supportsAllDrives=True
        ).execute()

    return [TextContent(type="text", text=f"Created spreadsheet '{args['title']}'\nID: {ss_id}\nURL: https://docs.google.com/spreadsheets/d/{ss_id}/edit")]


async def get_sheet_metadata(args: dict) -> list[TextContent]:
    spreadsheet = google_client.sheets_service.spreadsheets().get(
        spreadsheetId=args['spreadsheet_id']
    ).execute()

    output = [
        f"Title: {spreadsheet.get('properties', {}).get('title', 'Untitled')}",
        f"ID: {spreadsheet['spreadsheetId']}",
        f"URL: {spreadsheet.get('spreadsheetUrl', 'N/A')}",
        "",
        "Sheets:"
    ]

    for sheet in spreadsheet.get('sheets', []):
        props = sheet.get('properties', {})
        grid = props.get('gridProperties', {})
        output.append(f"  - {props.get('title', 'Untitled')} (ID: {props.get('sheetId')}, Rows: {grid.get('rowCount', 'N/A')}, Cols: {grid.get('columnCount', 'N/A')})")

    return [TextContent(type="text", text="\n".join(output))]


async def clear_sheet(args: dict) -> list[TextContent]:
    google_client.sheets_service.spreadsheets().values().clear(
        spreadsheetId=args['spreadsheet_id'],
        range=args['range']
    ).execute()

    return [TextContent(type="text", text=f"Cleared range {args['range']}")]


async def add_sheet_tab(args: dict) -> list[TextContent]:
    request = {
        'addSheet': {
            'properties': {
                'title': args['title']
            }
        }
    }

    result = google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    new_sheet = result.get('replies', [{}])[0].get('addSheet', {}).get('properties', {})

    return [TextContent(type="text", text=f"Added sheet '{new_sheet.get('title')}'\nSheet ID: {new_sheet.get('sheetId')}")]


async def delete_sheet_tab(args: dict) -> list[TextContent]:
    request = {
        'deleteSheet': {
            'sheetId': args['sheet_id']
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Deleted sheet with ID {args['sheet_id']}")]


# NEW in v9: Advanced Sheets Handlers

def _parse_a1_range(range_str: str, sheet_id: int) -> dict:
    """Parse A1 notation (e.g., 'A1:D10' or 'Sheet1!A1:D10') into GridRange dict."""
    import re

    # Remove sheet name if present
    if '!' in range_str:
        range_str = range_str.split('!')[1]

    # Match pattern like A1:D10 or A:D or 1:10
    match = re.match(r'^([A-Z]*)(\d*):?([A-Z]*)(\d*)$', range_str.upper())
    if not match:
        raise ValueError(f"Invalid A1 notation: {range_str}")

    start_col, start_row, end_col, end_row = match.groups()

    grid_range = {'sheetId': sheet_id}

    # Convert column letters to 0-based index
    def col_to_index(col):
        if not col:
            return None
        result = 0
        for char in col:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1

    if start_col:
        grid_range['startColumnIndex'] = col_to_index(start_col)
    if end_col:
        grid_range['endColumnIndex'] = col_to_index(end_col) + 1
    elif start_col:
        grid_range['endColumnIndex'] = col_to_index(start_col) + 1

    if start_row:
        grid_range['startRowIndex'] = int(start_row) - 1
    if end_row:
        grid_range['endRowIndex'] = int(end_row)
    elif start_row:
        grid_range['endRowIndex'] = int(start_row)

    return grid_range


async def sheets_batch_update(args: dict) -> list[TextContent]:
    """Execute multiple batch update requests. Advanced tool for complex operations."""
    result = google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': args['requests']}
    ).execute()

    replies = result.get('replies', [])
    return [TextContent(type="text", text=f"Batch update completed. {len(replies)} operations executed.")]


async def rename_sheet(args: dict) -> list[TextContent]:
    """Rename a sheet/tab."""
    request = {
        'updateSheetProperties': {
            'properties': {
                'sheetId': args['sheet_id'],
                'title': args['new_title']
            },
            'fields': 'title'
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Renamed sheet {args['sheet_id']} to '{args['new_title']}'")]


async def format_cells(args: dict) -> list[TextContent]:
    """Format cells with various styling options."""
    grid_range = _parse_a1_range(args['range'], args['sheet_id'])

    # Build the cell format
    cell_format = {}

    # Text format (bold, italic, font size, font family)
    text_format = {}
    if args.get('bold') is not None:
        text_format['bold'] = args['bold']
    if args.get('italic') is not None:
        text_format['italic'] = args['italic']
    if args.get('font_size'):
        text_format['fontSize'] = args['font_size']
    if args.get('font_family'):
        text_format['fontFamily'] = args['font_family']
    if args.get('text_color'):
        text_format['foregroundColor'] = args['text_color']

    if text_format:
        cell_format['textFormat'] = text_format

    # Background color
    if args.get('background_color'):
        cell_format['backgroundColor'] = args['background_color']

    # Alignment
    if args.get('horizontal_alignment'):
        cell_format['horizontalAlignment'] = args['horizontal_alignment']
    if args.get('vertical_alignment'):
        cell_format['verticalAlignment'] = args['vertical_alignment']

    # Wrap strategy
    if args.get('wrap_strategy'):
        cell_format['wrapStrategy'] = args['wrap_strategy']

    # Borders
    if args.get('borders'):
        borders = {}
        for side in ['top', 'bottom', 'left', 'right']:
            if args['borders'].get(side):
                border_config = args['borders'][side]
                borders[side] = {
                    'style': border_config.get('style', 'SOLID'),
                    'color': border_config.get('color', {'red': 0, 'green': 0, 'blue': 0})
                }
        if borders:
            cell_format['borders'] = borders

    # Build fields string for what we're updating
    fields = []
    if text_format:
        fields.append('userEnteredFormat.textFormat')
    if args.get('background_color'):
        fields.append('userEnteredFormat.backgroundColor')
    if args.get('horizontal_alignment'):
        fields.append('userEnteredFormat.horizontalAlignment')
    if args.get('vertical_alignment'):
        fields.append('userEnteredFormat.verticalAlignment')
    if args.get('wrap_strategy'):
        fields.append('userEnteredFormat.wrapStrategy')
    if args.get('borders'):
        fields.append('userEnteredFormat.borders')

    request = {
        'repeatCell': {
            'range': grid_range,
            'cell': {
                'userEnteredFormat': cell_format
            },
            'fields': ','.join(fields) if fields else 'userEnteredFormat'
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Formatted cells in range {args['range']}")]


async def set_column_width(args: dict) -> list[TextContent]:
    """Set the width of columns."""
    request = {
        'updateDimensionProperties': {
            'range': {
                'sheetId': args['sheet_id'],
                'dimension': 'COLUMNS',
                'startIndex': args['start_column'],
                'endIndex': args['end_column']
            },
            'properties': {
                'pixelSize': args['width']
            },
            'fields': 'pixelSize'
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Set column width to {args['width']}px for columns {args['start_column']} to {args['end_column']-1}")]


async def set_row_height(args: dict) -> list[TextContent]:
    """Set the height of rows."""
    request = {
        'updateDimensionProperties': {
            'range': {
                'sheetId': args['sheet_id'],
                'dimension': 'ROWS',
                'startIndex': args['start_row'],
                'endIndex': args['end_row']
            },
            'properties': {
                'pixelSize': args['height']
            },
            'fields': 'pixelSize'
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Set row height to {args['height']}px for rows {args['start_row']} to {args['end_row']-1}")]


async def freeze_rows_columns(args: dict) -> list[TextContent]:
    """Freeze rows and/or columns."""
    grid_properties = {}

    if args.get('frozen_rows') is not None:
        grid_properties['frozenRowCount'] = args['frozen_rows']
    if args.get('frozen_columns') is not None:
        grid_properties['frozenColumnCount'] = args['frozen_columns']

    if not grid_properties:
        return [TextContent(type="text", text="No freeze parameters specified. Provide frozen_rows and/or frozen_columns.")]

    fields = []
    if 'frozenRowCount' in grid_properties:
        fields.append('gridProperties.frozenRowCount')
    if 'frozenColumnCount' in grid_properties:
        fields.append('gridProperties.frozenColumnCount')

    request = {
        'updateSheetProperties': {
            'properties': {
                'sheetId': args['sheet_id'],
                'gridProperties': grid_properties
            },
            'fields': ','.join(fields)
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    output = []
    if args.get('frozen_rows') is not None:
        output.append(f"Frozen {args['frozen_rows']} rows")
    if args.get('frozen_columns') is not None:
        output.append(f"Frozen {args['frozen_columns']} columns")

    return [TextContent(type="text", text=', '.join(output))]


async def merge_cells(args: dict) -> list[TextContent]:
    """Merge a range of cells."""
    merge_type = args.get('merge_type', 'MERGE_ALL')

    request = {
        'mergeCells': {
            'range': {
                'sheetId': args['sheet_id'],
                'startRowIndex': args['start_row'],
                'endRowIndex': args['end_row'],
                'startColumnIndex': args['start_column'],
                'endColumnIndex': args['end_column']
            },
            'mergeType': merge_type
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Merged cells from ({args['start_row']}, {args['start_column']}) to ({args['end_row']-1}, {args['end_column']-1}) using {merge_type}")]


async def unmerge_cells(args: dict) -> list[TextContent]:
    """Unmerge a range of cells."""
    request = {
        'unmergeCells': {
            'range': {
                'sheetId': args['sheet_id'],
                'startRowIndex': args['start_row'],
                'endRowIndex': args['end_row'],
                'startColumnIndex': args['start_column'],
                'endColumnIndex': args['end_column']
            }
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Unmerged cells in range ({args['start_row']}, {args['start_column']}) to ({args['end_row']-1}, {args['end_column']-1})")]


async def add_filter(args: dict) -> list[TextContent]:
    """Add a filter view to a sheet."""
    grid_range = {
        'sheetId': args['sheet_id'],
        'startRowIndex': args['start_row'],
        'startColumnIndex': args['start_column'],
        'endColumnIndex': args['end_column']
    }

    if args.get('end_row'):
        grid_range['endRowIndex'] = args['end_row']

    request = {
        'setBasicFilter': {
            'filter': {
                'range': grid_range
            }
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Added filter to sheet {args['sheet_id']}")]


async def data_validation(args: dict) -> list[TextContent]:
    """Add data validation rules (dropdowns, number ranges, etc.)."""
    validation_type = args['validation_type']

    # Build the condition based on validation type
    condition = {'type': validation_type}

    if validation_type == 'ONE_OF_LIST':
        if not args.get('values'):
            return [TextContent(type="text", text="ONE_OF_LIST requires 'values' array")]
        condition['values'] = [{'userEnteredValue': v} for v in args['values']]

    elif validation_type == 'NUMBER_BETWEEN':
        if args.get('min_value') is None or args.get('max_value') is None:
            return [TextContent(type="text", text="NUMBER_BETWEEN requires 'min_value' and 'max_value'")]
        condition['values'] = [
            {'userEnteredValue': str(args['min_value'])},
            {'userEnteredValue': str(args['max_value'])}
        ]

    elif validation_type == 'NUMBER_GREATER':
        if args.get('min_value') is None:
            return [TextContent(type="text", text="NUMBER_GREATER requires 'min_value'")]
        condition['values'] = [{'userEnteredValue': str(args['min_value'])}]

    elif validation_type == 'NUMBER_LESS':
        if args.get('max_value') is None:
            return [TextContent(type="text", text="NUMBER_LESS requires 'max_value'")]
        condition['values'] = [{'userEnteredValue': str(args['max_value'])}]

    elif validation_type == 'CUSTOM_FORMULA':
        if not args.get('custom_formula'):
            return [TextContent(type="text", text="CUSTOM_FORMULA requires 'custom_formula'")]
        condition['values'] = [{'userEnteredValue': args['custom_formula']}]

    elif validation_type == 'TEXT_CONTAINS':
        if not args.get('values'):
            return [TextContent(type="text", text="TEXT_CONTAINS requires 'values' array with text to match")]
        condition['values'] = [{'userEnteredValue': args['values'][0]}]

    # Build the validation rule
    rule = {
        'condition': condition,
        'strict': args.get('strict', True),
        'showCustomUi': args.get('show_dropdown', True)
    }

    request = {
        'setDataValidation': {
            'range': {
                'sheetId': args['sheet_id'],
                'startRowIndex': args['start_row'],
                'endRowIndex': args['end_row'],
                'startColumnIndex': args['start_column'],
                'endColumnIndex': args['end_column']
            },
            'rule': rule
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Added {validation_type} validation to range")]


async def conditional_formatting(args: dict) -> list[TextContent]:
    """Add conditional formatting rules."""
    rule_type = args['rule_type']

    # Build the boolean condition
    condition = {'type': rule_type}

    if rule_type == 'CUSTOM_FORMULA':
        if not args.get('custom_formula'):
            return [TextContent(type="text", text="CUSTOM_FORMULA requires 'custom_formula'")]
        condition['values'] = [{'userEnteredValue': args['custom_formula']}]
    elif args.get('values'):
        condition['values'] = [{'userEnteredValue': str(v)} for v in args['values']]

    # Build the format
    format_style = {}
    if args.get('background_color'):
        format_style['backgroundColor'] = args['background_color']
    if args.get('text_color'):
        format_style['textFormat'] = {'foregroundColor': args['text_color']}
        if args.get('bold'):
            format_style['textFormat']['bold'] = True
    elif args.get('bold'):
        format_style['textFormat'] = {'bold': True}

    request = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [{
                    'sheetId': args['sheet_id'],
                    'startRowIndex': args['start_row'],
                    'endRowIndex': args['end_row'],
                    'startColumnIndex': args['start_column'],
                    'endColumnIndex': args['end_column']
                }],
                'booleanRule': {
                    'condition': condition,
                    'format': format_style
                }
            },
            'index': 0  # Add at the beginning of the rules list
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Added conditional formatting rule ({rule_type}) to range")]


async def named_range(args: dict) -> list[TextContent]:
    """Create a named range."""
    request = {
        'addNamedRange': {
            'namedRange': {
                'name': args['name'],
                'range': {
                    'sheetId': args['sheet_id'],
                    'startRowIndex': args['start_row'],
                    'endRowIndex': args['end_row'],
                    'startColumnIndex': args['start_column'],
                    'endColumnIndex': args['end_column']
                }
            }
        }
    }

    result = google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    named_range_id = result.get('replies', [{}])[0].get('addNamedRange', {}).get('namedRange', {}).get('namedRangeId', 'N/A')

    return [TextContent(type="text", text=f"Created named range '{args['name']}'\nNamed Range ID: {named_range_id}")]


async def auto_resize(args: dict) -> list[TextContent]:
    """Auto-resize columns or rows to fit content."""
    request = {
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': args['sheet_id'],
                'dimension': args['dimension'],
                'startIndex': args['start_index'],
                'endIndex': args['end_index']
            }
        }
    }

    google_client.sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=args['spreadsheet_id'],
        body={'requests': [request]}
    ).execute()

    return [TextContent(type="text", text=f"Auto-resized {args['dimension'].lower()} {args['start_index']} to {args['end_index']-1}")]


# ==================== GOOGLE DRIVE HANDLERS ====================

async def list_drive(args: dict) -> list[TextContent]:
    folder_id = args.get("folder_id", "root")
    query = f"'{folder_id}' in parents and trashed = false"
    
    if args.get("query"):
        query += f" and name contains '{args['query']}'"
    
    results = google_client.drive_service.files().list(
        q=query, pageSize=50,
        fields="files(id, name, mimeType, modifiedTime, size)",
        supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    
    files = results.get("files", [])
    if not files:
        return [TextContent(type="text", text="No files found.")]
    
    output = []
    for f in files:
        file_type = "📁" if f["mimeType"] == "application/vnd.google-apps.folder" else "📄"
        output.append(f"{file_type} {f['name']} (ID: {f['id']})")
    
    return [TextContent(type="text", text="\n".join(output))]


async def search_drive(args: dict) -> list[TextContent]:
    query = f"fullText contains '{args['query']}' and trashed = false"
    
    type_map = {
        "document": "application/vnd.google-apps.document",
        "spreadsheet": "application/vnd.google-apps.spreadsheet",
        "presentation": "application/vnd.google-apps.presentation",
        "folder": "application/vnd.google-apps.folder",
        "pdf": "application/pdf"
    }
    
    if args.get("file_type") and args["file_type"] in type_map:
        query += f" and mimeType = '{type_map[args['file_type']]}'"
    
    max_results = args.get("max_results", 20)
    
    results = google_client.drive_service.files().list(
        q=query, pageSize=max_results,
        fields="files(id, name, mimeType, modifiedTime, webViewLink)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
        corpora='allDrives'
    ).execute()
    
    files = results.get("files", [])
    if not files:
        return [TextContent(type="text", text="No files found.")]
    
    output = []
    for f in files:
        output.append(f"📄 {f['name']}\n   ID: {f['id']}\n   Link: {f.get('webViewLink', 'N/A')}")
    
    return [TextContent(type="text", text="\n\n".join(output))]


async def get_drive_file(args: dict) -> list[TextContent]:
    file = google_client.drive_service.files().get(
        fileId=args["file_id"],
        fields="id, name, mimeType, size, createdTime, modifiedTime, owners, webViewLink, parents",
        supportsAllDrives=True
    ).execute()
    
    output = [
        f"Name: {file.get('name')}",
        f"ID: {file.get('id')}",
        f"Type: {file.get('mimeType')}",
        f"Size: {file.get('size', 'N/A')} bytes",
        f"Created: {file.get('createdTime')}",
        f"Modified: {file.get('modifiedTime')}",
        f"Owners: {', '.join(o.get('emailAddress', '') for o in file.get('owners', []))}",
        f"Link: {file.get('webViewLink', 'N/A')}"
    ]
    
    return [TextContent(type="text", text="\n".join(output))]


async def create_folder(args: dict) -> list[TextContent]:
    metadata = {
        "name": args["name"],
        "mimeType": "application/vnd.google-apps.folder"
    }
    
    if args.get("parent_id"):
        metadata["parents"] = [args["parent_id"]]
    
    folder = google_client.drive_service.files().create(body=metadata, fields="id, name, webViewLink", supportsAllDrives=True).execute()
    
    return [TextContent(type="text", text=f"Created folder '{args['name']}'\nID: {folder.get('id')}\nLink: {folder.get('webViewLink')}")]


async def copy_file(args: dict) -> list[TextContent]:
    body = {"name": args["new_name"]}
    if args.get("folder_id"):
        body["parents"] = [args["folder_id"]]
    
    copied = google_client.drive_service.files().copy(fileId=args["file_id"], body=body, fields="id, name, webViewLink", supportsAllDrives=True).execute()
    
    return [TextContent(type="text", text=f"Copied to '{copied.get('name')}'\nID: {copied.get('id')}\nLink: {copied.get('webViewLink')}")]


async def move_file(args: dict) -> list[TextContent]:
    file = google_client.drive_service.files().get(fileId=args["file_id"], fields="parents", supportsAllDrives=True).execute()
    previous_parents = ",".join(file.get("parents", []))

    google_client.drive_service.files().update(
        fileId=args["file_id"],
        addParents=args["new_parent_id"],
        removeParents=previous_parents,
        fields="id, parents",
        supportsAllDrives=True
    ).execute()
    
    return [TextContent(type="text", text=f"Moved file {args['file_id']} to folder {args['new_parent_id']}")]


async def rename_file(args: dict) -> list[TextContent]:
    google_client.drive_service.files().update(
        fileId=args["file_id"],
        body={"name": args["new_name"]},
        fields="id, name",
        supportsAllDrives=True
    ).execute()
    
    return [TextContent(type="text", text=f"Renamed to '{args['new_name']}'")]


async def delete_file(args: dict) -> list[TextContent]:
    google_client.drive_service.files().update(
        fileId=args["file_id"],
        body={"trashed": True},
        supportsAllDrives=True
    ).execute()
    
    return [TextContent(type="text", text=f"Moved {args['file_id']} to trash")]


async def share_file(args: dict) -> list[TextContent]:
    permission = {
        "type": "user",
        "role": args["role"],
        "emailAddress": args["email"]
    }
    
    notify = args.get("notify", True)
    
    google_client.drive_service.permissions().create(
        fileId=args["file_id"],
        body=permission,
        sendNotificationEmail=notify,
        fields="id",
        supportsAllDrives=True
    ).execute()
    
    return [TextContent(type="text", text=f"Shared with {args['email']} as {args['role']}")]


async def list_permissions(args: dict) -> list[TextContent]:
    permissions = google_client.drive_service.permissions().list(
        fileId=args["file_id"],
        fields="permissions(id, emailAddress, role, type, displayName)",
        supportsAllDrives=True
    ).execute()
    
    perms = permissions.get("permissions", [])
    if not perms:
        return [TextContent(type="text", text="No permissions found.")]
    
    output = []
    for p in perms:
        output.append(f"{p.get('displayName', p.get('emailAddress', 'Unknown'))} - {p.get('role')} ({p.get('type')})")
    
    return [TextContent(type="text", text="\n".join(output))]


async def export_file(args: dict) -> list[TextContent]:
    format_map = {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "txt": "text/plain",
        "csv": "text/csv"
    }
    
    mime_type = format_map.get(args["export_format"].lower())
    if not mime_type:
        return [TextContent(type="text", text=f"Unsupported format: {args['export_format']}")]
    
    request = google_client.drive_service.files().export_media(fileId=args["file_id"], mimeType=mime_type, supportsAllDrives=True)
    
    output_path = Path(args["output_path"])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
    
    return [TextContent(type="text", text=f"Exported to {output_path}")]


async def upload_file(args: dict) -> list[TextContent]:
    local_path = Path(args["local_path"])
    if not local_path.exists():
        return [TextContent(type="text", text=f"File not found: {local_path}")]
    
    mime_type, _ = __import__("mimetypes").guess_type(str(local_path))
    
    metadata = {"name": args.get("name", local_path.name)}
    if args.get("folder_id"):
        metadata["parents"] = [args["folder_id"]]
    
    media = MediaFileUpload(str(local_path), mimetype=mime_type)
    
    file = google_client.drive_service.files().create(
        body=metadata, media_body=media, fields="id, name, webViewLink",
        supportsAllDrives=True
    ).execute()
    
    return [TextContent(type="text", text=f"Uploaded '{file.get('name')}'\nID: {file.get('id')}\nLink: {file.get('webViewLink')}")]


async def download_file(args: dict) -> list[TextContent]:
    request = google_client.drive_service.files().get_media(fileId=args["file_id"], supportsAllDrives=True)
    
    output_path = Path(args["output_path"])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
    
    return [TextContent(type="text", text=f"Downloaded to {output_path}")]


# ==================== GOOGLE CALENDAR HANDLERS ====================

async def list_calendar_events(args: dict) -> list[TextContent]:
    max_results = args.get("max_results", 10)
    days_ahead = args.get("days_ahead", 7)
    
    now = datetime.utcnow()
    time_min = now.isoformat() + 'Z'
    time_max = (now + timedelta(days=days_ahead)).isoformat() + 'Z'
    
    events_result = google_client.calendar_service.events().list(
        calendarId='primary', timeMin=time_min, timeMax=time_max,
        maxResults=max_results, singleEvents=True, orderBy='startTime'
    ).execute()
    
    events = events_result.get('items', [])
    if not events:
        return [TextContent(type="text", text="No upcoming events found.")]
    
    output = []
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        output.append(f"📅 {start} - {event.get('summary', 'No title')}\n   ID: {event.get('id')}")
    
    return [TextContent(type="text", text="\n\n".join(output))]


async def get_calendar_event(args: dict) -> list[TextContent]:
    event = google_client.calendar_service.events().get(calendarId='primary', eventId=args["event_id"]).execute()
    
    output = [
        f"Title: {event.get('summary', 'No title')}",
        f"Start: {event['start'].get('dateTime', event['start'].get('date'))}",
        f"End: {event['end'].get('dateTime', event['end'].get('date'))}",
        f"Description: {event.get('description', 'None')}",
        f"Location: {event.get('location', 'None')}",
        f"Attendees: {', '.join(a.get('email', '') for a in event.get('attendees', []))}"
    ]
    
    return [TextContent(type="text", text="\n".join(output))]


async def create_calendar_event(args: dict) -> list[TextContent]:
    event = {
        'summary': args["summary"],
        'description': args.get("description", ""),
        'start': {'dateTime': args["start_time"], 'timeZone': 'UTC'},
        'end': {'dateTime': args["end_time"], 'timeZone': 'UTC'},
    }

    if "attendees" in args:
        event['attendees'] = [{'email': email} for email in args["attendees"]]

    if "recurrence" in args:
        event['recurrence'] = args["recurrence"]

    created = google_client.calendar_service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
    return [TextContent(type="text", text=f"Created event '{args['summary']}'\nID: {created.get('id')}\nLink: {created.get('htmlLink')}")]


async def update_calendar_event(args: dict) -> list[TextContent]:
    # Get existing event first
    event = google_client.calendar_service.events().get(
        calendarId='primary',
        eventId=args['event_id']
    ).execute()

    # Update fields if provided
    if args.get('summary'):
        event['summary'] = args['summary']
    if args.get('description'):
        event['description'] = args['description']
    if args.get('location'):
        event['location'] = args['location']
    if args.get('start_time'):
        event['start'] = {'dateTime': args['start_time'], 'timeZone': 'UTC'}
    if args.get('end_time'):
        event['end'] = {'dateTime': args['end_time'], 'timeZone': 'UTC'}
    if args.get('attendees'):
        event['attendees'] = [{'email': email} for email in args['attendees']]

    if args.get('recurrence'):
        event['recurrence'] = args['recurrence']

    updated = google_client.calendar_service.events().update(
        calendarId='primary',
        eventId=args['event_id'],
        body=event,
        sendUpdates='all'
    ).execute()

    return [TextContent(type="text", text=f"Updated event '{updated.get('summary')}'\nLink: {updated.get('htmlLink')}")]


async def delete_calendar_event(args: dict) -> list[TextContent]:
    google_client.calendar_service.events().delete(
        calendarId='primary',
        eventId=args['event_id']
    ).execute()

    return [TextContent(type="text", text=f"Deleted event {args['event_id']}")]


async def quick_add_event(args: dict) -> list[TextContent]:
    event = google_client.calendar_service.events().quickAdd(
        calendarId='primary',
        text=args['text']
    ).execute()

    start = event.get('start', {}).get('dateTime', event.get('start', {}).get('date', 'N/A'))

    return [TextContent(type="text", text=f"Created event '{event.get('summary', 'N/A')}'\nWhen: {start}\nID: {event.get('id')}\nLink: {event.get('htmlLink')}")]


async def list_calendars(args: dict) -> list[TextContent]:
    calendars = google_client.calendar_service.calendarList().list().execute()

    output = ["Your Calendars:", "==============="]

    for cal in calendars.get('items', []):
        primary = " (PRIMARY)" if cal.get('primary') else ""
        access = cal.get('accessRole', 'N/A')
        output.append(f"\n📅 {cal.get('summary', 'Untitled')}{primary}")
        output.append(f"   ID: {cal.get('id')}")
        output.append(f"   Access: {access}")

    return [TextContent(type="text", text="\n".join(output))]


# ==================== GMAIL HANDLERS (ENHANCED in v4) ====================

async def search_gmail(args: dict) -> list[TextContent]:
    query = args["query"]
    max_results = args.get("max_results", 10)
    
    results = google_client.gmail_service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
    
    messages = results.get('messages', [])
    if not messages:
        return [TextContent(type="text", text="No emails found.")]
    
    output = []
    for msg in messages:
        msg_detail = google_client.gmail_service.users().messages().get(
            userId='me', id=msg['id'], format='metadata', metadataHeaders=['From', 'Subject', 'Date']
        ).execute()
        
        headers = {h['name']: h['value'] for h in msg_detail.get('payload', {}).get('headers', [])}
        output.append(f"📧 {headers.get('Date', 'No date')}\n   From: {headers.get('From', 'Unknown')}\n   Subject: {headers.get('Subject', 'No subject')}\n   ID: {msg['id']}\n   Thread ID: {msg_detail.get('threadId')}")
    
    return [TextContent(type="text", text="\n\n".join(output))]


async def read_gmail(args: dict) -> list[TextContent]:
    msg = google_client.gmail_service.users().messages().get(userId='me', id=args["message_id"], format='full').execute()
    
    headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}
    
    body = ""
    payload = msg.get('payload', {})
    if 'body' in payload and payload['body'].get('data'):
        body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8')
    elif 'parts' in payload:
        for part in payload['parts']:
            if part.get('mimeType') == 'text/plain' and part.get('body', {}).get('data'):
                body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                break
    
    output = [
        f"From: {headers.get('From', 'Unknown')}",
        f"To: {headers.get('To', 'Unknown')}",
        f"Date: {headers.get('Date', 'Unknown')}",
        f"Subject: {headers.get('Subject', 'No subject')}",
        f"Thread ID: {msg.get('threadId')}",
        f"Message ID: {msg.get('id')}",
        f"\n--- Body ---\n{body}"
    ]
    
    return [TextContent(type="text", text="\n".join(output))]


async def send_gmail(args: dict) -> list[TextContent]:
    message = MIMEText(args["body"])
    message['to'] = args["to"]
    message['subject'] = args["subject"]
    
    if args.get("cc"):
        message['cc'] = args["cc"]
    if args.get("bcc"):
        message['bcc'] = args["bcc"]
    
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    sent = google_client.gmail_service.users().messages().send(userId='me', body={'raw': raw}).execute()
    
    return [TextContent(type="text", text=f"Email sent to {args['to']}\nMessage ID: {sent.get('id')}")]


# NEW in v4: Draft handlers

async def create_gmail_draft(args: dict) -> list[TextContent]:
    """Create a draft email."""
    message = MIMEText(args["body"])
    message['to'] = args["to"]
    message['subject'] = args["subject"]
    
    if args.get("cc"):
        message['cc'] = args["cc"]
    if args.get("bcc"):
        message['bcc'] = args["bcc"]
    
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    draft = google_client.gmail_service.users().drafts().create(
        userId='me',
        body={'message': {'raw': raw}}
    ).execute()
    
    return [TextContent(type="text", text=f"Draft created\nDraft ID: {draft.get('id')}\nTo: {args['to']}\nSubject: {args['subject']}")]


async def list_gmail_drafts(args: dict) -> list[TextContent]:
    """List all drafts."""
    max_results = args.get("max_results", 10)
    
    results = google_client.gmail_service.users().drafts().list(userId='me', maxResults=max_results).execute()
    
    drafts = results.get('drafts', [])
    if not drafts:
        return [TextContent(type="text", text="No drafts found.")]
    
    output = []
    for draft in drafts:
        draft_detail = google_client.gmail_service.users().drafts().get(userId='me', id=draft['id'], format='metadata').execute()
        msg = draft_detail.get('message', {})
        headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}
        
        output.append(f"📝 Draft ID: {draft['id']}\n   To: {headers.get('To', 'N/A')}\n   Subject: {headers.get('Subject', 'No subject')}")
    
    return [TextContent(type="text", text="\n\n".join(output))]


async def send_gmail_draft(args: dict) -> list[TextContent]:
    """Send an existing draft."""
    sent = google_client.gmail_service.users().drafts().send(
        userId='me',
        body={'id': args["draft_id"]}
    ).execute()
    
    return [TextContent(type="text", text=f"Draft sent!\nMessage ID: {sent.get('id')}")]


# NEW in v4: Reply handler

async def reply_gmail(args: dict) -> list[TextContent]:
    """Reply to an existing email thread."""
    # Get the original message to extract thread info and headers
    original = google_client.gmail_service.users().messages().get(
        userId='me', id=args["message_id"], format='metadata',
        metadataHeaders=['From', 'To', 'Cc', 'Subject', 'Message-ID', 'References']
    ).execute()
    
    headers = {h['name']: h['value'] for h in original.get('payload', {}).get('headers', [])}
    thread_id = original.get('threadId')
    
    # Determine recipients
    if args.get("reply_all"):
        # Reply to sender + all original recipients
        to_addresses = [headers.get('From', '')]
        if headers.get('To'):
            to_addresses.append(headers['To'])
        to = ', '.join(filter(None, to_addresses))
    else:
        # Reply only to sender
        to = headers.get('From', '')
    
    # Build subject with Re: prefix
    subject = headers.get('Subject', '')
    if not subject.lower().startswith('re:'):
        subject = f"Re: {subject}"
    
    # Build References header for proper threading
    message_id = headers.get('Message-ID', '')
    references = headers.get('References', '')
    if references:
        references = f"{references} {message_id}"
    else:
        references = message_id
    
    # Create the reply message
    message = MIMEText(args["body"])
    message['to'] = to
    message['subject'] = subject
    message['In-Reply-To'] = message_id
    message['References'] = references
    
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    
    sent = google_client.gmail_service.users().messages().send(
        userId='me',
        body={'raw': raw, 'threadId': thread_id}
    ).execute()
    
    return [TextContent(type="text", text=f"Reply sent!\nTo: {to}\nThread ID: {thread_id}\nMessage ID: {sent.get('id')}")]


# NEW in v4: Labels handlers

async def list_gmail_labels(args: dict) -> list[TextContent]:
    """List all Gmail labels."""
    results = google_client.gmail_service.users().labels().list(userId='me').execute()
    
    labels = results.get('labels', [])
    if not labels:
        return [TextContent(type="text", text="No labels found.")]
    
    # Separate system labels from user labels
    system_labels = []
    user_labels = []
    
    for label in labels:
        label_type = label.get('type', 'user')
        label_info = f"{label['name']} (ID: {label['id']})"
        
        if label_type == 'system':
            system_labels.append(label_info)
        else:
            user_labels.append(label_info)
    
    output = ["=== System Labels ==="]
    output.extend(system_labels)
    output.append("\n=== User Labels ===")
    output.extend(user_labels if user_labels else ["(none)"])
    
    return [TextContent(type="text", text="\n".join(output))]


async def modify_gmail_message(args: dict) -> list[TextContent]:
    """Modify a message: add/remove labels."""
    body = {}
    
    if args.get("add_labels"):
        body["addLabelIds"] = args["add_labels"]
    if args.get("remove_labels"):
        body["removeLabelIds"] = args["remove_labels"]
    
    if not body:
        return [TextContent(type="text", text="No modifications specified. Provide add_labels or remove_labels.")]
    
    modified = google_client.gmail_service.users().messages().modify(
        userId='me',
        id=args["message_id"],
        body=body
    ).execute()
    
    changes = []
    if args.get("add_labels"):
        changes.append(f"Added: {', '.join(args['add_labels'])}")
    if args.get("remove_labels"):
        changes.append(f"Removed: {', '.join(args['remove_labels'])}")
    
    return [TextContent(type="text", text=f"Message modified\nID: {modified.get('id')}\n{chr(10).join(changes)}\nCurrent labels: {', '.join(modified.get('labelIds', []))}")]


# ==================== GOOGLE SLIDES HANDLERS ====================

async def create_slides(args: dict) -> list[TextContent]:
    presentation = google_client.slides_service.presentations().create(body={"title": args["title"]}).execute()
    pres_id = presentation.get("presentationId")
    
    if args.get("folder_id"):
        google_client.drive_service.files().update(
            fileId=pres_id, addParents=args["folder_id"], removeParents='root', fields='id, parents',
            supportsAllDrives=True
        ).execute()
    
    return [TextContent(type="text", text=f"Created presentation '{args['title']}'\nID: {pres_id}\nURL: https://docs.google.com/presentation/d/{pres_id}/edit")]


async def add_slide(args: dict) -> list[TextContent]:
    presentation_id = args["presentation_id"]
    title = args.get("title", "")
    body = args.get("body", "")
    
    requests = [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]
    response = google_client.slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}
    ).execute()
    
    slide_id = response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
    
    presentation = google_client.slides_service.presentations().get(presentationId=presentation_id).execute()
    
    text_requests = []
    for slide in presentation.get('slides', []):
        if slide.get('objectId') == slide_id:
            for element in slide.get('pageElements', []):
                shape = element.get('shape', {})
                placeholder = shape.get('placeholder', {})
                
                if placeholder.get('type') == 'TITLE' and title:
                    text_requests.append({"insertText": {"objectId": element['objectId'], "text": title}})
                elif placeholder.get('type') == 'BODY' and body:
                    text_requests.append({"insertText": {"objectId": element['objectId'], "text": body}})
    
    if text_requests:
        google_client.slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": text_requests}
        ).execute()
    
    return [TextContent(type="text", text=f"Added slide to presentation {presentation_id}")]


async def read_slides(args: dict) -> list[TextContent]:
    presentation = google_client.slides_service.presentations().get(presentationId=args["presentation_id"]).execute()

    slides = presentation.get('slides', [])
    output = [f"Presentation: {presentation.get('title', 'Untitled')}", f"Slides: {len(slides)}", ""]

    for i, slide in enumerate(slides, 1):
        slide_text = []
        for element in slide.get('pageElements', []):
            shape = element.get('shape', {})
            if 'text' in shape:
                for text_element in shape['text'].get('textElements', []):
                    if 'textRun' in text_element:
                        slide_text.append(text_element['textRun'].get('content', '').strip())

        output.append(f"Slide {i}: {' | '.join(filter(None, slide_text)) or '(empty)'}")

    return [TextContent(type="text", text="\n".join(output))]


async def delete_slide(args: dict) -> list[TextContent]:
    requests = [{
        'deleteObject': {
            'objectId': args['slide_id']
        }
    }]

    google_client.slides_service.presentations().batchUpdate(
        presentationId=args['presentation_id'],
        body={'requests': requests}
    ).execute()

    return [TextContent(type="text", text=f"Deleted slide {args['slide_id']}")]


async def replace_text_slides(args: dict) -> list[TextContent]:
    requests = [{
        'replaceAllText': {
            'containsText': {
                'text': args['find_text'],
                'matchCase': args.get('match_case', False)
            },
            'replaceText': args['replace_text']
        }
    }]

    result = google_client.slides_service.presentations().batchUpdate(
        presentationId=args['presentation_id'],
        body={'requests': requests}
    ).execute()

    occurrences = result.get('replies', [{}])[0].get('replaceAllText', {}).get('occurrencesChanged', 0)
    return [TextContent(type="text", text=f"Replaced {occurrences} occurrences of '{args['find_text']}' with '{args['replace_text']}'")]


async def duplicate_slide(args: dict) -> list[TextContent]:
    requests = [{
        'duplicateObject': {
            'objectId': args['slide_id']
        }
    }]

    result = google_client.slides_service.presentations().batchUpdate(
        presentationId=args['presentation_id'],
        body={'requests': requests}
    ).execute()

    new_id = result.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId', 'N/A')
    return [TextContent(type="text", text=f"Duplicated slide\nOriginal: {args['slide_id']}\nNew slide ID: {new_id}")]


async def get_slides_details(args: dict) -> list[TextContent]:
    presentation = google_client.slides_service.presentations().get(
        presentationId=args['presentation_id']
    ).execute()

    output = [
        f"Title: {presentation.get('title', 'Untitled')}",
        f"Presentation ID: {presentation.get('presentationId')}",
        f"Total Slides: {len(presentation.get('slides', []))}",
        "",
        "Slides:",
        "======="
    ]

    for i, slide in enumerate(presentation.get('slides', []), 1):
        slide_id = slide.get('objectId')
        output.append(f"\nSlide {i}:")
        output.append(f"  ID: {slide_id}")

        # List page elements
        elements = slide.get('pageElements', [])
        output.append(f"  Elements: {len(elements)}")

        for element in elements:
            elem_id = element.get('objectId')

            if 'shape' in element:
                shape = element['shape']
                shape_type = shape.get('shapeType', 'UNKNOWN')

                # Extract text if present
                text_content = ""
                if 'text' in shape:
                    for text_elem in shape['text'].get('textElements', []):
                        if 'textRun' in text_elem:
                            text_content += text_elem['textRun'].get('content', '')

                text_preview = text_content.strip()[:40]
                if len(text_content.strip()) > 40:
                    text_preview += "..."

                if text_preview:
                    output.append(f"    - {elem_id}: {shape_type} \"{text_preview}\"")
                else:
                    output.append(f"    - {elem_id}: {shape_type}")

            elif 'image' in element:
                output.append(f"    - {elem_id}: IMAGE")

            elif 'table' in element:
                rows = element['table'].get('rows', 0)
                cols = element['table'].get('columns', 0)
                output.append(f"    - {elem_id}: TABLE {rows}x{cols}")

    return [TextContent(type="text", text="\n".join(output))]


# ==================== MAIN ====================

async def main():
    if not google_client.authenticate():
        logger.error("Failed to authenticate")
        return

    logger.info("Google Workspace MCP Server (v12 - ReisierX build, Tasks removed) starting...")
    logger.info("🎉 86 total tools across 6 Google Workspace APIs 🎉")
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
