# Open WebUI Tools

This directory contains tools that can be used with Open WebUI's tools feature. These tools allow you to extend Open WebUI with custom functionality.

## Creating a New Tool

To create a new tool, follow these steps:

1. Create a new Python file in the `tools` directory, e.g., `my_tool.py`.
2. Use the template from `tools_template.py` as a starting point.
3. Implement your tool functionality following the guidelines below.

## Tool File Structure

Each tool file should follow this structure:

```python
"""
title: Your Tool Name
author: Your Name
version: 0.1
license: MIT
description: A brief description of what your tool does
"""

# Imports
import os
from typing import Optional, Callable, Any
from fastapi import UploadFile, Request

# Required Open WebUI imports
from open_webui.models.users import Users
from open_webui.routers.files import upload_file

# EventEmitter class for status updates
class EventEmitter:
    # ...

# Helper classes/functions
class HelperFunctions:
    # ...

# Main Tools class
class Tools:
    def __init__(self):
        # ...
    
    async def your_tool_method(self, param1, param2, __request__: Request, 
                              __event_emitter__: Callable[[dict], Any] = None, 
                              __user__=None):
        """
        Tool method docstring with description and example usage
        """
        # Tool implementation
```

## Guidelines for Tool Development

### 1. Metadata Header

Include a docstring at the top of your file with the following metadata:

```python
"""
title: Your Tool Name
author: Your Name
version: 0.1
license: MIT
description: A brief description of what your tool does
"""
```

### 2. EventEmitter

Always use the `EventEmitter` to provide status updates during tool execution:

```python
emitter = EventEmitter(__event_emitter__)
await emitter.emit("Starting process...")
# ... do work ...
await emitter.emit(status="complete", description="Process completed", done=True)
```

### 3. Tool Methods

- All tool methods must be `async` methods in the `Tools` class
- Include the special parameters: `__request__`, `__event_emitter__`, and `__user__`
- Provide clear docstrings with parameters description and usage examples
- Handle exceptions properly

### 4. File Handling

If your tool generates files:

1. Create a temporary directory for files (`self.FILES_DIR = "./tmp"`)
2. Use the `upload_file` function to add files to Open WebUI's file system
3. Return file links using the standard format:

```python
return (
    f"<source><source_id>{doc.filename}</source_id><source_context>" 
    + str(download_link)
    + "</source_context></source>\n"
)
```

### 5. Best Practices

- Keep helper functions in a separate class or module
- Use descriptive variable names
- Add debug logging with `print("[DEBUG] message")`
- Handle all exceptions and provide meaningful error messages
- Follow Python PEP 8 style guidelines
- Use type hints for better code readability

## Example

See `generate_pptx.py` and `generate_docx.py` for complete examples of tool implementation.

## Testing Your Tool

To test your tool:

1. Place your tool file in the `tools` directory
2. Restart Open WebUI or run the reload command
3. Your tool should appear in the Tools dropdown in the chat interface
4. Test your tool with different inputs to ensure it works correctly

## Common Issues

- Make sure your tool file has the correct permissions (readable by the Open WebUI process)
- Check for syntax errors in your Python code
- Verify that all required dependencies are installed
- Make sure the `__init__` method doesn't have required parameters 