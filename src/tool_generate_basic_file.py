"""
title: Generate a file
author: openlab
version: 0.1
description: allows to generate and give a download link
"""

import os, uuid
from pydantic import BaseModel, Field
from open_webui.storage.provider import Storage
from open_webui.models.files import Files
from open_webui.models.files import FileForm

from fastapi import UploadFile


class Tools:

    def __init__(self):
        self.API_BASE_URL = "http://localhost:3000/api/v1/files/"
        self.FILES_DIR = "./tmp"
        self.TIMEOUT = 10

    def create_file(
        self,
        file_name: str,
        file_content: str,
        file_extension: str = None,
        __user__=None,
    ):
        """
        Create a file with the given name and content.
        Handles different file extensions appropriately.
        When you use this function make sure the always give to the user the link to download the file

        Args:
            file_name: Name of the file to create
            file_content: Content to write to the file
            file_extension: Optional extension to append if not already in file_name
        Returns:
            the download link of the file generated always send this to the user
        """

        # If extension is provided and not already in file_name, append it
        if file_extension and not file_name.endswith(f".{file_extension}"):
            file_name = f"{file_name}.{file_extension}"

        file_path = os.path.abspath(file_name)
        if not os.path.exists(self.FILES_DIR):
            os.makedirs(self.FILES_DIR)

        # Determine if the file is binary or text based on extension
        binary_extensions = ["pdf", "png", "jpg", "jpeg", "gif", "zip", "exe", "bin"]
        is_binary = file_extension and file_extension.lower() in binary_extensions
        file_path = os.path.join(self.FILES_DIR, file_name)
        print(f"[DEBUG] File path: {file_path}")
        # Write the file with appropriate mode
        if is_binary:
            # For binary files, content should be properly encoded
            with open(file_path, "wb") as f:
                # Assuming content might be base64 encoded for binary files
                try:
                    import base64

                    f.write(base64.b64decode(file_content))
                except:
                    # Fallback if not base64 encoded
                    f.write(file_content.encode("utf-8"))
        else:
            # For text files
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(file_content)
        print(f"[DEBUG] File created: {file_path}")

        # Upload the file to the OpenAI API
        download_link = self.get_file_download_link(file_path, __user__)
        return f"This is the dowload link for the {file_name} : {download_link}"

    # Method to download a filde using the openwebui api.
    def get_file_download_link(self, file: str, __user__=None):
        """
        get the download link of a file
        it utilise the openwebui api to upload the file and get the download link
        ARGS:
            file: the path to the file to upload
        RETURNS:
            the download link of the file
        """
        try:
            file_path = os.path.abspath(file)
            print(f"[DEBUG] File path: {file_path}")
            
            file_id = ""
            try:
                with open(file, "rb") as f:
                    files = UploadFile(file=f, filename=os.path.basename(file))
                    # files = {"file": f,"filename": os.path.basename(file),"content_type": "application/octet-stream"}
                    print(f"[DEBUG] Files: {files}")
                    # Use direct requests instead of self.post for more contro
                    response = self.upload_file(files, __user__["id"])
                    print(f"[DEBUG] Response: {response}")

                    # Parse the response
                    file_id = response.id
                    print(f"[DEBUG] File ID: {file_id}")
                    if not file_id:
                        return {"error": "No file ID returned from upload"}

            except Exception as e:
                print(f"[DEBUG] Error uploading file: {str(e)}")
                return {"error": f"Error uploading file: {str(e)}"}

            download_url = f"{self.API_BASE_URL}{file_id}/content"
            print(f"[DEBUG] Download URL: {download_url}")
            # delete the file from the local directory
            # os.remove(file_path)
            print(f"[DEBUG] File {file_path} deleted")
            return download_url
        except Exception as e:
            # os.remove(file_path)
            print(f"[DEBUG] Error in download_file_openwebui: {str(e)}")
            return {"error": f"Error in download_file_openwebui: {str(e)}"}

    def upload_file(self, file: UploadFile, user_id: str):
        """
        upload a file to the openwebui data base without the API (API doesn't work with the tools in version 0.6.5)
        ARGS:
            file: the file to upload
            user_id: the id of the user
        RETURNS:
            the file item
        """
        filename = file.filename
        id = str(uuid.uuid4())
        name = filename
        filename = f"{id}_{filename}"
        contents, file_path = Storage.upload_file(file.file, filename)

        file_item = Files.insert_new_file(
            user_id,
            FileForm(
                **{
                    "id": id,
                    "filename": name,
                    "path": file_path,
                    "meta": {
                        "name": name,
                        "content_type": file.content_type,
                        "size": len(contents),
                        "data": {"generated_by": "upload_file"},
                    },
                }
            ),
        )
        print(f"[DEBUG] File item: {file_item}")

        return file_item
