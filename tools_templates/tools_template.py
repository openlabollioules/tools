"""
title: Tool Template
author: Your Name
version: 0.1
license: MIT
description: Template for creating tools for Open WebUI
"""

import os
from typing import Optional, Callable, Any
from fastapi import UploadFile, Request

from open_webui.models.users import Users
from open_webui.routers.files import upload_file

class EventEmitter:
    """
    Handles status updates during tool execution.
    """
    def __init__(self, event_emitter: Callable[[dict], Any] = None):
        self.event_emitter = event_emitter
    
    async def emit(self, description="Unknown State", status="in_progress", done=False):
        """
        Emits a status update.
        
        Args:
            description (str): Description of the current state
            status (str): Status of the operation ('in_progress', 'complete', 'error')
            done (bool): Whether the operation is complete
        """
        if self.event_emitter:
            await self.event_emitter(
                {
                    "type": "status",
                    "data": {
                        "status": status,
                        "description": description,
                        "done": done,
                    },
                }
            )

class HelperFunctions:
    """
    Place helper functions and utilities here.
    These functions support the main tool functionality.
    """
    def __init__(self):
        pass

    def example_helper(self, param: str) -> str:
        """
        Example helper function.
        
        Args:
            param (str): Input parameter
            
        Returns:
            str: Processed result
        """
        return f"Processed: {param}"

class Tools:
    """
    Main Tools class containing all tool methods.
    """
    def __init__(self):
        # Initialize constants and helpers
        self.FILES_DIR = "./tmp"
        self.API_BASE_URL = "http://localhost:3000/api/v1/files/"
        os.makedirs(self.FILES_DIR, exist_ok=True)
        self.helper = HelperFunctions()
        
    async def example_tool(self, input_param: str, __request__: Request, 
                           __event_emitter__: Callable[[dict], Any] = None, 
                           __user__=None):
        """
        Example tool method.
        
        Args:
            input_param (str): Input parameter from the user
            __request__ (Request): FastAPI request object
            __event_emitter__ (Callable): Function to emit status updates
            __user__ (dict): User information
            
        Returns:
            str: Result of the tool execution
            
        Example:
        Here is an example input:
        
        ```json
        {
            "input_param": "example value"
        }
        ```
        """
        # Initialize event emitter
        emitter = EventEmitter(__event_emitter__)
        
        # Log input for debugging
        print(f"[DEBUG] Received input: {input_param}")
        
        # Emit initial status
        await emitter.emit(f"Starting example tool with input: {input_param}")
        
        try:
            # Perform the main tool functionality
            result = self.helper.example_helper(input_param)
            
            # Process completed successfully
            await emitter.emit(
                status="complete",
                description=f"Example tool completed successfully",
                done=True
            )
            
            return result
            
        except Exception as e:
            # Handle errors
            print(f"[ERROR] {str(e)}")
            await emitter.emit(
                status="error",
                description=f"Error: {str(e)}",
                done=True
            )
            return f"Error: {str(e)}"
    
    async def upload_file_example(self, file: UploadFile, __request__: Request, 
                                 __event_emitter__: Callable[[dict], Any] = None, 
                                 __user__=None):
        """
        Example tool for file uploads.
        
        Args:
            file (UploadFile): File to upload
            __request__ (Request): FastAPI request object
            __event_emitter__ (Callable): Function to emit status updates
            __user__ (dict): User information
            
        Returns:
            str: Download link for the uploaded file
        """
        emitter = EventEmitter(__event_emitter__)
        metadata = {"data": {"generated_by": "example_tool"}}
 
        await emitter.emit(f"Uploading file: {file.filename}")
        
        # Get the user for permissions
        user = Users.get_user_by_id(id=__user__['id'])
        
        # Upload the file to the database
        doc = upload_file(request=__request__, file=file, user=user, file_metadata=metadata, process=False)
        
        # Get the download link
        download_link = f"{self.API_BASE_URL}{doc.id}/content"
        
        await emitter.emit(
            status="complete",
            description=f"File uploaded successfully",
            done=True
        )
        
        return (
            f"<source><source_id>{doc.filename}</source_id><source_context>" 
            + str(download_link)
            + "</source_context></source>\n"
        )
