import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.user_credential import UserCredential


class SharePointClient:
    def __init__(
        self,
        site_url: str,
        client_id: str = None,
        client_secret: str = None,
        username: str = None,
        password: str = None,
    ):
        """
        Initialize SharePoint connection.
        Supports:
            - App authentication (client_id + client_secret)
            - User authentication (username + password)
        """
        self.site_url = site_url

        if client_id and client_secret:
            creds = ClientCredential(client_id, client_secret)
            self.ctx = ClientContext(site_url).with_credentials(creds)
        elif username and password:
            creds = UserCredential(username, password)
            self.ctx = ClientContext(site_url).with_credentials(creds)
        else:
            raise ValueError("Provide either app credentials or user credentials")

    # ==========================
    # DOWNLOAD FILE
    # ==========================
    def download_file(self, sharepoint_file_url: str, local_path: str):
        """
        Download a file from SharePoint to local path.

        :param sharepoint_file_url: Server-relative URL
               Example: "/sites/project/Shared Documents/folder/file.xlsx"
        :param local_path: Local file path to save
        """
        try:
            file = self.ctx.web.get_file_by_server_relative_url(
                sharepoint_file_url
            )
            with open(local_path, "wb") as local_file:
                file.download(local_file).execute_query()

            print(f"✅ File downloaded to: {local_path}")

        except Exception as e:
            print(f"❌ Error downloading file: {e}")
            raise

    # ==========================
    # UPLOAD FILE
    # ==========================
    def upload_file(self, local_file_path: str, sharepoint_folder_url: str):
        """
        Upload a file to a SharePoint folder.

        :param local_file_path: Local file path
        :param sharepoint_folder_url: Server-relative folder URL
               Example: "/sites/project/Shared Documents/folder"
        """
        try:
            with open(local_file_path, "rb") as content_file:
                file_content = content_file.read()

            file_name = os.path.basename(local_file_path)

            folder = self.ctx.web.get_folder_by_server_relative_url(
                sharepoint_folder_url
            )

            upload_file = folder.upload_file(file_name, file_content).execute_query()

            print(f"✅ Uploaded file to: {upload_file.serverRelativeUrl}")

        except Exception as e:
            print(f"❌ Error uploading file: {e}")
            raise

    # ==========================
    # LIST FILES
    # ==========================
    def list_files(self, sharepoint_folder_url: str):
        """
        List files in a SharePoint folder.
        """
        try:
            folder = self.ctx.web.get_folder_by_server_relative_url(
                sharepoint_folder_url
            )
            files = folder.files.get().execute_query()

            file_list = [f.properties["Name"] for f in files]

            print("📂 Files:")
            for f in file_list:
                print(f"- {f}")

            return file_list

        except Exception as e:
            print(f"❌ Error listing files: {e}")
            raise