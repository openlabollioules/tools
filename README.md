# OpenWebUI Document Generation Tools

> **Note:**
>
> * Make sure to install the requirements `requirements.txt` otherwise the docx tool will not work.

This project contains a collection of professional document generation tools designed for OpenWebUI. These tools enable AI chatbots to create, format, and deliver various types of office documents directly to users through OpenWebUI's file system.

## 🚀 Overview

The tools are built specifically for OpenWebUI and integrate seamlessly with its file management and user system. They allow users to request AI-generated documents in professional formats including:

- **PowerPoint Presentations** (PPTX)
- **Word Documents** (DOCX) 
- **Excel Spreadsheets** (XLSX)
- **Basic Text Files** (TXT, JSON, etc.)
- **Template Analysis** for PowerPoint layouts

## 📋 Available Tools

### 1. PowerPoint Generator (`generate_pptx.py`)
**Creates professional PowerPoint presentations from structured JSON data.**

- **Features:**
  - Multi-language support (French/English)
  - Confidentiality levels (Public/Internal/Confidential)
  - Professional templates with corporate branding
  - Automatic bullet point formatting
  - Chapter slides, content slides, and title slides
  - Date and author information

- **JSON Structure:**
```json
{
  "titre": "Presentation Title",
  "slides": [
    {
      "type": "titre",
      "titre": "Main Title"
    },
    {
      "type": "chapitre", 
      "titre": "Chapter Title"
    },
    {
      "type": "contenu",
      "titre": "Content Title",
      "contenu": "Content text\n* Bullet point\n    * Sub-bullet"
    }
  ]
}
```

### 2. Word Document Generator (`generate_docx.py`)
**Creates structured Word documents with professional formatting.**

- **Features:**
  - Cover page generation
  - Table of contents
  - Multiple heading levels
  - Professional styling (Calibri/Arial fonts)
  - Bibliography support
  - Page numbering
  - Bullet point formatting

- **JSON Structure:** 
```json
{
  "titre": "Document Title",
  "sous_titre": "Subtitle",
  "auteur": "Author Name",
  "date": "Date",
  "inclure_table_matieres": true,
  "sections": [
    {
      "type": "page_garde",
      "titre": "Cover Title",
      "sous_titre": "Cover Subtitle"  
    },
    {
      "type": "heading",
      "titre": "Section Title",
      "niveau": 1
    },
    {
      "type": "contenu",
      "contenu": "Section content..."
    }
  ]
}
```

### 3. Excel Generator (`generate_excel.py`)
**Creates formatted Excel spreadsheets with tables and styling.**

- **Features:**
  - Multiple worksheets
  - Professional table formatting
  - Header styling with colors
  - Automatic column width adjustment
  - Excel table creation with filters
  - Formulas for totals
  - Alternating row colors

- **JSON Structure:**
```json
{
  "titre": "Spreadsheet Title",
  "feuilles": [
    {
      "nom": "Sheet Name",
      "tableau": {
        "colonnes": ["Column 1", "Column 2", "Column 3"],
        "données": [
          ["Data 1", "Data 2", "Data 3"],
          ["Data 4", "Data 5", "Data 6"]
        ]
      }
    }
  ]
}
```

### 4. Basic File Generator (`tool_generate_basic_file.py`)
**Creates simple text-based files with custom content.**

- **Features:**
  - Support for multiple file extensions
  - Binary and text file handling
  - Base64 encoding support for binary files
  - Automatic file upload to OpenWebUI

### 5. Template Analyzer (`analyse_slides_templates.py`)
**Analyzes PowerPoint templates to identify layouts and placeholders.**

- **Features:**
  - Template structure analysis
  - Shape and placeholder identification
  - Layout mapping
  - Code generation suggestions

## 🛠️ Installation & Setup

### Prerequisites

```bash
# Python packages required
pip install python-pptx python-docx openpyxl fastapi open-webui
```

### Directory Structure

```
tools/
├── README.md                      # This file
├── generate_pptx.py              # PowerPoint generation
├── generate_docx.py              # Word document generation  
├── generate_excel.py             # Excel spreadsheet generation
├── tool_generate_basic_file.py   # Basic file creation
├── analyse_slides_templates.py   # Template analysis
└── tools_templates/              # Template and documentation
    ├── README.md                 # Development guide
    └── tools_template.py         # Base template for new tools
```

### Template Files

The PowerPoint generator requires template files in:
```
templates/
├── fr/                           # French templates
│   ├── CS-PU-template_fr.pptx   # Public template
│   ├── CS-IN-template_fr.pptx   # Internal template
│   └── CS-CO-template_fr.pptx   # Confidential template
└── en/                           # English templates
    ├── CS-PU-template_en.pptx   # Public template
    ├── CS-IN-template_en.pptx   # Internal template
    └── CS-CO-template_en.pptx   # Confidential template
```

## 🔧 OpenWebUI Integration

### Tool Structure

Each tool follows the OpenWebUI standard format:

```python
"""
title: Tool Name
author: openlab
version: 0.1
license: MIT
description: Tool description
"""

class EventEmitter:
    """Handles status updates during execution"""
    
class HelpFunctions:
    """Contains helper methods and utilities"""
    
class Tools:
    """Main class with tool methods"""
    
    async def tool_method(self, param, __request__, __event_emitter__=None, __user__=None):
        """Tool method with required OpenWebUI parameters"""
```

### Key Components

- **EventEmitter**: Provides real-time status updates to users
- **File Upload System**: Integrates with OpenWebUI's file management
- **User Context**: Access to user information and permissions
- **Request Handling**: FastAPI request object for HTTP context

## 📖 Usage Examples

### Generate a PowerPoint Presentation

```python
# In OpenWebUI chat, use the tool with JSON data:
{
  "language": "fr",
  "confidentiality": "public", 
  "json_data": {
    "titre": "My Presentation",
    "slides": [
      {"type": "titre", "titre": "Welcome"},
      {"type": "contenu", "titre": "Overview", "contenu": "* Point 1\n* Point 2"}
    ]
  }
}
```

### Generate a Word Document

```python
{
  "json_data": {
    "titre": "Report Title",
    "inclure_table_matieres": true,
    "sections": [
      {"type": "heading", "titre": "Introduction", "niveau": 1},
      {"type": "contenu", "contenu": "Introduction text..."}
    ]
  }
}
```

### Generate an Excel Spreadsheet

```python
{
  "json_data": {
    "titre": "Financial Report",
    "feuilles": [
      {
        "nom": "Summary",
        "tableau": {
          "colonnes": ["Month", "Revenue", "Expenses"],
          "données": [
            ["Jan", 10000, 8000],
            ["Feb", 12000, 9000]
          ]
        }
      }
    ]
  }
}
```

## 🎨 Features

### Professional Formatting
- Corporate templates with consistent branding
- Professional fonts (Calibri, Arial)
- Color schemes and styling
- Automatic layout management

### Multi-language Support
- French and English templates
- Localized content formatting
- Date and text formatting per locale

### File Management
- Automatic file upload to OpenWebUI
- Download links generation
- Temporary file cleanup
- File metadata tracking

### Error Handling
- Comprehensive exception handling
- User-friendly error messages
- Debug logging for troubleshooting
- Graceful fallbacks

## 🔧 Development

### Creating New Tools

1. Use `tools_templates/tools_template.py` as a base
2. Implement the required classes: `EventEmitter`, `HelpFunctions`, `Tools`
3. Follow the OpenWebUI parameter convention
4. Include comprehensive docstrings with examples
5. Add proper error handling and logging

### Code Standards

- Python 3.11+ compatibility
