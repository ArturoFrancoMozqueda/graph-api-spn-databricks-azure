import msal
import time
import requests
import urllib.parse
#import logging
import os
import re
from datetime import datetime
from typing import Tuple, Dict, List

class SharePointAccess:
    """
    Client to access SharePoint using Microsoft Graph API and MSAL.
    
    Improvements include:
    - Persistent requests.Session for connection pooling.
    - Consistent error handling via helper methods.
    - Logging integration for traceability.
    - Type annotations and detailed docstrings.
    """
    
    def __init__(self, client_id: str, tenant_id: str, client_secret: str):
        """
        Initializes the SharePoint client.
        
        Args:
            client_id (str): Cluster client ID.
            tenant_id (str): Cluster tenant ID.
            client_secret (str): Cluster secret client.
        """
        self._client_id = client_id
        self._tenant_id = tenant_id
        self._client_secret = client_secret
        self._authority = f'https://login.microsoftonline.com/{self._tenant_id}'
        
        # Initialize persistent session for connection pooling
        self._session = requests.Session()
        
        # Date format for SharePoint dates
        self.time_format = '%Y-%m-%d %H:%M:%S'
        
        # Create MSAL client and retrieve an access token
        self.__client = self.__msla_client()
        self.__access_token = self.__get_access_token()
        self._headers = {'Authorization': 'Bearer ' + self.__access_token,
                         'Content-Type': 'application/json'}

    # --- Setter Methods ---
    def set_client_id(self, new_client_id: str) -> None:
        """Set a new client ID."""
        self._client_id = new_client_id

    def set_tenant_id(self, new_tenant_id: str) -> None:
        """Set a new tenant ID."""
        self._tenant_id = new_tenant_id

    def set_client_secret(self, new_client_secret: str) -> None:
        """Set a new client secret."""
        self._client_secret = new_client_secret

    # --- MSAL Authentication Methods ---
    def __msla_client(self) -> msal.ConfidentialClientApplication:
        """
        Creates the MSAL authentication client.
        
        Returns:
            msal.ConfidentialClientApplication: The MSAL client.
        """
        client_msa = msal.ConfidentialClientApplication(
            client_id=self._client_id, 
            client_credential=self._client_secret, 
            authority=self._authority
        )
        return client_msa

    def __get_access_token(self) -> str:
        """
        Retrieves an access token using MSAL.
        
        Returns:
            str: The access token.
        """
        scope = ['https://graph.microsoft.com/.default']
        token_result = self.__client.acquire_token_silent(scope, account=None)
        if not token_result:
            token_result = self.__client.acquire_token_for_client(scopes=scope)
        if "access_token" not in token_result:
            raise Exception(f"Could not obtain access token: {token_result.get('error_description')}")
        return token_result["access_token"]

    # --- Helper Method for Direct Connection ---
    def __connect_to_site(self, url: str) -> requests.Response:
        """
        Connects to a URL using the access token.
        
        Args:
            url (str): The URL to connect to.
        
        Returns:
            requests.Response: The response object.
        
        Raises:
            Exception: If connection errors occur.
        """
        try:
            response = self._session.get(url=url, headers=self._headers)
        except requests.exceptions.ConnectionError as con_err:
            print(f"Connection error: {con_err}")
            raise Exception(f"Connection error: {con_err}")
        except requests.exceptions.Timeout as tm:
            print(f"Timeout: {tm}")
            raise Exception(f"Timeout: {tm}")
        except requests.exceptions.HTTPError as htt_err:
            print(f"HTTP Error: {htt_err}")
            raise Exception(f"HTTP Error: {htt_err}")

        if response.status_code == 400:
            raise Exception("Bad Request")
        elif response.status_code == 401:
            raise Exception("Unauthorized")
        elif response.status_code == 404:
            raise Exception(f"Not found: {response.status_code}")
        return response

    # --- SharePoint Methods ---
    def get_site_id(self, sharepoint_domain: str, sharepoint_site_name: str) -> str:
        """
        Retrieves the SharePoint site ID.
        
        Args:
            sharepoint_domain (str): The SharePoint domain.
            sharepoint_site_name (str): The site name.
        
        Returns:
            str: The site ID.
        """
        url = f'https://graph.microsoft.com/v1.0/sites/{sharepoint_domain}:{sharepoint_site_name}'
        response = self.__connect_to_site(url)
        site_id = response.json().get('id')
        if not site_id:
            raise Exception("Site ID not found in response.")
        # If site ID is in format "site,abc,xyz", split and return second part.
        return site_id.split(",", 1)[1] if "," in site_id else site_id

    def get_drive_id(self, site_id: str) -> Dict[str, str]:
        """
        Retrieves drive IDs for the given SharePoint site.
        
        Args:
            site_id (str): The site ID.
        
        Returns:
            Dict[str, str]: A dictionary mapping drive names to IDs.
        """
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
        response = self.__connect_to_site(url)
        drives = response.json().get('value', [])
        return {drive['name']: drive['id'] for drive in drives}
    
    def set_range_number_format(self, site_id: str, drive_id: str, item_id: str, 
                            raw_data_sheet_name: str, range_address: str, 
                            number_format: str) -> None:
        """
        Sets number format for a specific range in an Excel worksheet.
        
        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID (document library)
            item_id: File ID (workbook)
            raw_data_sheet_name: Worksheet name
            range_address: Excel range (e.g., "A:A")
            number_format: Excel format code (e.g., "@" for text)
        """
        format_endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{raw_data_sheet_name}')/range(address='{range_address}')/format"
        )

        print(f"Setting format '{number_format}' for range '{range_address}'...")
        
        response = self._session.patch(
            url=format_endpoint,
            headers=self._headers,  # Use the same headers that work for clear
            json={
                "numberFormat": {
                    "format": number_format
                }
            }
        )

        if response.status_code == 200:
            print("Format updated successfully.")
        else:
            raise Exception(f"Format update failed: {response.status_code}, {response.text}")

    def get_directory_list(
        self,
        sharepoint_domain: str,
        sharepoint_site_name: str,
        sub_drive_name: str,
        sharepoint_path: str
    ) -> Tuple[List[Dict], List[Dict]]:
        """
        Retrieves lists of folders and files from a SharePoint directory.
        
        Args:
            sharepoint_domain (str): The SharePoint domain.
            sharepoint_site_name (str): The site name.
            sub_drive_name (str): The drive name.
            sharepoint_path (str): The directory path.
        
        Returns:
            Tuple[List[Dict], List[Dict]]: (folders, files)
        """
        site_id = self.get_site_id(sharepoint_domain, sharepoint_site_name)
        drive_dict = self.get_drive_id(site_id)
        drive_id = drive_dict.get(sub_drive_name)
        if not drive_id:
            raise Exception(f"Drive '{sub_drive_name}' not found.")
        
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/root:/{sharepoint_path}:/children'
        response = self.__connect_to_site(url)

        folder_list, file_list = [], []
        for item in response.json().get('value', []):
            if 'folder' in item:
                folder_list.append({
                    'id': item['id'], 
                    'name': item['name'], 
                    'type': "Folder",
                    'createdDateTime': item['createdDateTime'].replace('T', ' ').replace("Z", ''),
                    'webUrl': item['webUrl']
                })
            elif 'file' in item:
                file_list.append({
                    'id': item['id'], 
                    'name': item['name'],
                    'type': "File", 
                    'createdDateTime': item['createdDateTime'].replace('T', ' ').replace("Z", ''),
                    'downloadUrl': item.get('@microsoft.graph.downloadUrl')
                })
        return folder_list, file_list

    def download_file_in_dbfs(self, file_name: str, download_url: str, dbfs_temp_folder: str = "/dbfs/") -> str:
        """
        Downloads a file from SharePoint and saves it to DBFS.
        
        Args:
            file_name (str): The file name.
            download_url (str): The URL to download the file.
            dbfs_temp_folder (str): The target folder path.
        
        Returns:
            str: The full DBFS path where the file was saved.
        """
        response = self.__connect_to_site(download_url)
        os.makedirs(dbfs_temp_folder, exist_ok=True)
        dbfs_path = os.path.join(dbfs_temp_folder, file_name)
        try:
            with open(dbfs_path, "wb") as file:
                file.write(response.content)
            print(f"File '{file_name}' saved successfully to {dbfs_temp_folder}.")
        except Exception as e:
            print(f"Error saving file '{file_name}': {e}")
            raise
        return dbfs_path

    def get_most_recent_file(
        self,
        sharepoint_domain: str,
        sharepoint_site_name: str,
        sub_drive_name: str,
        sharepoint_path: str,
        flag_download: bool = False
    ) -> Dict:
        """
        Retrieves metadata of the most recent file in a SharePoint directory.
        
        Args:
            sharepoint_domain (str): The SharePoint domain.
            sharepoint_site_name (str): The site name.
            sub_drive_name (str): The drive name.
            sharepoint_path (str): The directory path.
            flag_download (bool): If True, downloads the file.
        
        Returns:
            Dict: Metadata of the most recent file.
        """
        _, file_list = self.get_directory_list(sharepoint_domain, sharepoint_site_name, sub_drive_name, sharepoint_path)
        if not file_list:
            raise Exception("No files found in the directory.")
        
        # Convert createdDateTime to datetime objects for sorting.
        for item in file_list:
            item['parsedDateTime'] = datetime.strptime(item['createdDateTime'], self.time_format)
        file_list.sort(key=lambda x: x['parsedDateTime'])
        most_recent = file_list[-1]
        most_recent['createdDateTime'] = most_recent['parsedDateTime'].strftime(self.time_format)
        
        if flag_download:
            self.download_file_in_dbfs(file_name=most_recent['name'], download_url=most_recent['downloadUrl'])
        return most_recent

    def get_most_recent_folder(
        self,
        sharepoint_domain: str,
        sub_drive_name: str,
        sharepoint_site_name: str,
        sharepoint_path: str
    ) -> Dict:
        """
        Retrieves metadata of the most recently created folder in a SharePoint directory.
        
        Args:
            sharepoint_domain (str): The SharePoint domain.
            sub_drive_name (str): The drive name.
            sharepoint_site_name (str): The site name.
            sharepoint_path (str): The directory path.
        
        Returns:
            Dict: Metadata of the most recent folder.
        """
        folder_list, _ = self.get_directory_list(sharepoint_domain, sharepoint_site_name, sub_drive_name, sharepoint_path)
        if not folder_list:
            raise Exception("No folders found in the directory.")
        for item in folder_list:
            item['parsedDateTime'] = datetime.strptime(item['createdDateTime'], self.time_format)
        folder_list.sort(key=lambda x: x['parsedDateTime'])
        most_recent = folder_list[-1]
        most_recent['createdDateTime'] = most_recent['parsedDateTime'].strftime(self.time_format)
        return most_recent

    def delete_file(self, file_object: Dict, sharepoint_domain: str, sharepoint_site_name: str, sub_drive_name: str, sharepoint_path: str) -> None:
        """
        Deletes a file from SharePoint.
        
        Args:
            file_object (Dict): The file metadata object.
            sharepoint_domain (str): The SharePoint domain.
            sharepoint_site_name (str): The site name.
            sub_drive_name (str): The drive name.
            sharepoint_path (str): The directory path.
        """
        site_id = self.get_site_id(sharepoint_domain, sharepoint_site_name)
        drive_dict = self.get_drive_id(site_id)
        drive_id = drive_dict.get(sub_drive_name)
        if not drive_id:
            raise Exception(f"Drive '{sub_drive_name}' not found.")

        file_name = file_object.get("name")
        file_id = file_object.get("id")
        delete_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}"
        response = self._session.delete(url= delete_endpoint,
                                        headers=self._headers)
        if response.status_code in [200, 204]:
            print(f"Existing file '{file_name}' deleted successfully.")
        else:
            raise Exception(f"Failed to delete existing file: {response.text}")

    def wait_for_file(self, site_id: str, drive_id: str, folder_path: str, file_name: str, timeout: int = 60, poll_interval: int = 3) -> Dict:
        """
        Waits until a specific file is available in SharePoint.
        
        Args:
            site_id (str): The SharePoint site ID.
            drive_id (str): The drive ID.
            folder_path (str): The folder path.
            file_name (str): The file name.
            timeout (int): Maximum seconds to wait.
            poll_interval (int): Seconds between checks.
        
        Returns:
            Dict: Metadata of the file.
        """
        file_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}"
        print(f"Waiting for file '{file_name}'...")
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                response = self._session.get(url=file_endpoint,
                                            headers=self._headers)
                if response.ok:
                    data = response.json()
                    print(f"File '{file_name}' is now available.")
                    return {
                        "id": data.get("id"),
                        "name": data.get("name"),
                        "createdDateTime": data.get("createdDateTime", "").replace("T", " ").replace("Z", ""),
                        "webUrl": data.get("webUrl")
                    }
            except Exception as e:
                print(f"File not available yet: {e}")
            time.sleep(poll_interval)
        raise Exception(f"Timeout waiting for file '{file_name}' in '{folder_path}'.")

    def download_file_content(self, site_id: str, drive_id: str, folder_path: str, file_name: str) -> bytes:
        """
        Downloads the content of a file from SharePoint.
        
        Args:
            site_id (str): The SharePoint site ID.
            drive_id (str): The drive ID.
            folder_path (str): The folder path.
            file_name (str): The file name.
        
        Returns:
            bytes: The file content.
        """
        file_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}"
        print(f"Retrieving metadata for file '{file_name}'")
        response = self._session.get(url = file_endpoint, 
                                     headers=self._headers)
        if not response.ok:
            raise Exception(f"Failed to locate file '{file_name}': {response.text}")
        item_id = response.json().get("id")
        download_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}/content"
        download_resp = self._session.get(url=download_endpoint,
                                          headers=self._headers)
        if not download_resp.ok:
            raise Exception(f"Failed to download file '{file_name}': {download_resp.text}")
        return download_resp.content

    def upload_new_file(self, site_id: str, drive_id: str, folder_path: str, new_file_name: str, file_data: bytes) -> None:
        """
        Uploads a new file to SharePoint.
        
        Args:
            site_id (str): The SharePoint site ID.
            drive_id (str): The drive ID.
            folder_path (str): The target folder path.
            new_file_name (str): The new file name.
            file_data (bytes): The file content.
        """
        upload_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{new_file_name}:/content"
        print(f"Uploading file '{new_file_name}'...")
        response = self._session.put(url = upload_endpoint,
                                     headers=self._headers,
                                     data=file_data)
        if response.status_code not in [200, 201]:
            raise Exception(f"Failed to upload file '{new_file_name}': {response.text}")
        print(f"File '{new_file_name}' uploaded successfully.")

    def clear_worksheet_range(self, site_id: str, drive_id: str, item_id: str, raw_data_sheet_name: str, clear_range: str) -> None:
        """
        Clears a specific range in an Excel worksheet stored on SharePoint.
        
        Args:
            site_id (str): The SharePoint site ID.
            drive_id (str): The drive ID.
            item_id (str): The Excel file ID.
            raw_data_sheet_name (str): The worksheet name.
            clear_range (str): The cell range to clear.
        """
        clear_endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{raw_data_sheet_name}')/range(address='{clear_range}')/clear"
        )
        print(f"Clearing range '{clear_range}' in worksheet '{raw_data_sheet_name}'...")
        response = self._session.post(url = clear_endpoint,
                                      headers=self._headers)
        if response.status_code in [200, 204]:
            print("Sheet cleared successfully (keeping headers).")
        else:
            raise Exception(f"Failed to clear sheet: {response.status_code}, {response.text}")
    
    def update_range_data(self, site_id: str, drive_id: str, new_item_id: str, raw_data_sheet_name: str, 
                      update_range: str, chunk_data: list, start: int, end: int, max_retries: int = 3) -> bool:
        update_endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{new_item_id}"
            f"/workbook/worksheets('{raw_data_sheet_name}')/range(address='{update_range}')"
        )
        updated_body = {"values": chunk_data}
        
        for attempt in range(max_retries):
            response = self._session.patch(
                url=update_endpoint,
                headers=self._headers,
                json=updated_body
            )
            if response.ok:
                print(f"Successfully updated rows {start+1} to {end}.")
                return True  # Return True upon successful update
            else:
                if response.status_code == 429 or "MaxRequestDurationExceeded" in response.text:
                    retry_after = response.headers.get("Retry-After")
                    # Use the server's suggested wait time if provided; otherwise, use exponential backoff.
                    wait_time = int(retry_after) if retry_after and retry_after.isdigit() else (2 ** attempt)
                    print(f"Request timed out (attempt {attempt+1}/{max_retries}). Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    raise Exception(f"Failed to update rows {start+1} to {end}: {response.text}")
        
        raise Exception(f"Failed to update rows {start+1} to {end} after {max_retries} attempts.")


    def refresh_pivot_table(self, site_id: str, drive_id: str, item_id: str,
                            pivot_table_sheet: str, max_retries: int = 3) -> None:
        # Codificar el nombre de la hoja para manejar espacios y caracteres especiales
        encoded_sheet_name = urllib.parse.quote(pivot_table_sheet)
        
        refresh_pivot_endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{encoded_sheet_name}')/pivotTables/refreshAll"
        )
        
        retry_count = 0
        last_refresh_error = None

        while retry_count < max_retries:
            refresh_resp = self._session.post(url=refresh_pivot_endpoint, headers=self._headers)
            if refresh_resp.ok:
                print("Pivot Table refreshed successfully.")
                return
            else:
                response_text = refresh_resp.text
                last_refresh_error = response_text
                # Si el error es por tiempo excedido, reintentar
                if "MaxRequestDurationExceeded" in response_text:
                    retry_count += 1
                    wait_time = 2 ** retry_count  # backoff exponencial
                    print(f"Timeout error occurred. Retrying in {wait_time} seconds (attempt {retry_count}/{max_retries})...")
                    time.sleep(wait_time)
                else:
                    # Si es otro error, intenta obtener información de diagnóstico de las pivot tables
                    pivot_tables_endpoint = (
                        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
                        f"/workbook/worksheets('{encoded_sheet_name}')/pivotTables"
                    )
                    diagnostic_resp = self._session.get(url=pivot_tables_endpoint, headers=self._headers)
                    if diagnostic_resp.ok:
                        pivot_tables = diagnostic_resp.json()
                        print("Diagnostic: Pivot tables found:", pivot_tables)
                    else:
                        print("Diagnostic: Error retrieving pivot tables:", diagnostic_resp.status_code, diagnostic_resp.text)
                    raise Exception(f"Failed to refresh Pivot Table: {response_text}")

        # Después de alcanzar el máximo de reintentos, se intenta obtener las pivot tables para diagnóstico.
        print("Max retry attempts reached. Attempting to retrieve pivot tables for troubleshooting...")
        pivot_tables_endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{encoded_sheet_name}')/pivotTables"
        )
        diagnostic_resp = self._session.get(url=pivot_tables_endpoint, headers=self._headers)
        if diagnostic_resp.ok:
            pivot_tables = diagnostic_resp.json()
            print("Diagnostic: Pivot tables found:", pivot_tables)
        else:
            print("Diagnostic: Error retrieving pivot tables:", diagnostic_resp.status_code, diagnostic_resp.text)
        
        raise Exception(f"Max retry attempts reached. The pivot table refresh could not be completed. Last error: {last_refresh_error}")

    def refresh_individual_pivot_table(self, site_id: str, drive_id: str, item_id: str,
                                    worksheet_name: str, pivot_table_name: str,
                                    max_retries: int = 3) -> None:
        """
        Refresca de forma individual una pivot table en una worksheet usando el endpoint beta de Graph API.
        
        Args:
            site_id (str): ID del sitio de SharePoint.
            drive_id (str): ID del drive.
            item_id (str): ID del workbook.
            worksheet_name (str): Nombre de la hoja en la que se encuentra la pivot table.
            pivot_table_name (str): Nombre (o ID) de la pivot table a refrescar.
            max_retries (int): Número máximo de reintentos en caso de error.
        
        Raises:
            Exception: Si no se puede refrescar la pivot table luego de los reintentos.
        """
        # Codificar nombres para la URL
        encoded_worksheet_name = urllib.parse.quote(worksheet_name)
        encoded_pivot_name = urllib.parse.quote(pivot_table_name)
        
        # Endpoint para refrescar una pivot table específica (beta)
        endpoint = (
            f"https://graph.microsoft.com/beta/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{encoded_worksheet_name}')/pivotTables('{encoded_pivot_name}')/refresh"
        )
        
        retry_count = 0
        last_refresh_error = None

        while retry_count < max_retries:
            refresh_resp = self._session.post(url=endpoint, headers=self._headers)
            if refresh_resp.ok:
                print("Pivot Table refreshed successfully.")
                return
            else:
                response_text = refresh_resp.text
                last_refresh_error = response_text
                # Si el error es por tiempo excedido, se reintenta con backoff exponencial
                if "MaxRequestDurationExceeded" in response_text:
                    retry_count += 1
                    wait_time = 2 ** retry_count
                    print(f"Timeout error occurred. Retrying in {wait_time} seconds (attempt {retry_count}/{max_retries})...")
                    time.sleep(wait_time)
                else:
                    print("Error refreshing pivot table:", response_text)
                    raise Exception(f"Failed to refresh Pivot Table: {response_text}")
        
        raise Exception(f"Max retry attempts reached. The pivot table refresh could not be completed. Last error: {last_refresh_error}")

    def list_pivot_tables(self, site_id: str, drive_id: str, item_id: str, worksheet_name: str) -> dict:
        encoded_sheet_name = urllib.parse.quote(worksheet_name)
        endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{encoded_sheet_name}')/pivotTables"
        )
        response = self._session.get(url=endpoint, headers=self._headers)
        if response.ok:
            return response.json()
        else:
            raise Exception(f"Error listing pivot tables: {response.status_code}, {response.text}")


    def save_file_in_sharepoint(
        self,
        sharepoint_domain: str,
        sharepoint_site_name: str,
        sharepoint_path: str,
        sharepoint_sub_drive: str,
        dbfs_path: str,
        file_name: str,
        content_type: str
    ) -> None:
        """
        Uploads a file from DBFS to SharePoint.
        
        Args:
            sharepoint_domain (str): The SharePoint domain.
            sharepoint_site_name (str): The site name.
            sharepoint_path (str): The destination folder path in SharePoint.
            sharepoint_sub_drive (str): The drive name.
            dbfs_path (str): The local DBFS file path.
            file_name (str): The target file name.
            content_type (str): The MIME type of the file.
        """
        site_id = self.get_site_id(sharepoint_domain, sharepoint_site_name)
        drive_dict = self.get_drive_id(site_id)
        drive_id = drive_dict.get(sharepoint_sub_drive)
        if not drive_id:
            raise Exception(f"Drive '{sharepoint_sub_drive}' not found.")
        
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/root:/{sharepoint_path}/{file_name}:/content"
        with open(dbfs_path, 'rb') as file:
            file_content = file.read()
        headers = self._headers.copy()
        headers["Content-Type"] = content_type
        response = self._session.put(url, headers=headers, data=file_content)
        if response.status_code not in [200, 201]:
            raise Exception(f"File upload failed: {response.text}")
        print("File written successfully to SharePoint.")
        
