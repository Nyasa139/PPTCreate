
import os
import shutil  # noqa: F401
from pptx import Presentation
from pptcreator import pptcreate  # Ensure pptcreator.py is in the same directory or installed as a module
import streamlit as st
import requests
from msal import PublicClientApplication
import time
st.title("PPT Batch Processor")

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
SCOPES = ["Files.Read.All"]  # Use scope in short form, MSAL expects this style
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
count=0
if not CLIENT_ID or not TENANT_ID:
    st.error("AZURE_CLIENT_ID and AZURE_TENANT_ID must be set in environment variables.")
    st.stop()

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

fold={"Inp":"D:\MUJ SID Program\Input","OUT":"D:\MUJ SID Program\Output","arch":"D:\MUJ SID Program\Input\Archives","lay":["D:\MUJ SID Program\Templates\Layout_MUJ_V10.pptx"],"xlsx":"D:\MUJ SID Program\XLS\layout_shapes.xlsx"}
class AuthManager:
    def __init__(self, flow=None):
        self.flow = flow
        self.token = None

    def initiate_device_flow(self):
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception(f"Device flow initiation failed: {flow}")
        self.flow = flow
        return flow["message"]

    def acquire_token(self):
        if not self.flow:
            raise Exception("Device flow not initiated.")
        result = app.acquire_token_by_device_flow(self.flow)
        if "access_token" not in result:
            raise Exception(f"Authentication failed: {result.get('error_description', str(result))}")
        self.token = result["access_token"]
        return self.token


def list_and_download_files(token, fold_name, local_dir):
    headers = {"Authorization": f"Bearer {token}"}
    os.makedirs(local_dir, exist_ok=True)

    shared_url = "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe"
    response = requests.get(shared_url, headers=headers)
    response.raise_for_status()

    vsb_folder = None
    for item in response.json().get("value", []):
        if item["name"] == fold_name and "folder" in item:
            vsb_folder = item["remoteItem"]
            break

    if not vsb_folder:
        raise Exception(f"Folder '{fold_name}' not found in shared items.")

    drive_id = vsb_folder["parentReference"]["driveId"]
    item_id = vsb_folder["id"]
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    response = requests.get(children_url, headers=headers)
    response.raise_for_status()

    files = []
    for item in response.json().get("value", []):
        if "file" in item:
            download_url = item["@microsoft.graph.downloadUrl"]
            local_path = os.path.join(local_dir, item["name"])
            with requests.get(download_url, stream=True) as r:
                r.raise_for_status()
                with open(local_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
            files.append(local_path)
    return files


def main():
    # Initialize session state variables
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if "flow" not in st.session_state:
        st.session_state.flow = None

    if "token" not in st.session_state:
        st.session_state.token = None

    # Keep track of downloaded files list
    if "downloaded_files" not in st.session_state:
        st.session_state.downloaded_files = []

    auth_manager = AuthManager(flow=st.session_state.flow)

    # --- Authentication UI ---
    if not st.session_state.authenticated:
        st.write("### Step 1: Authenticate")

        if st.button("Start Authentication"):
            try:
                message = auth_manager.initiate_device_flow()
                st.session_state.device_flow_message = message
                st.session_state.flow = auth_manager.flow
                st.session_state.verification_uri = auth_manager.flow.get("verification_uri")
                st.session_state.user_code = auth_manager.flow.get("user_code")
                st.success("Device flow started! Please follow instructions below.")
            except Exception as e:
                st.error(f"Authentication initiation failed: {e}")

        if st.session_state.flow:
            st.write("Copy the code below and paste it in the auth window:")
            st.code(st.session_state.user_code, language="text")
            if "copied" in st.session_state:
                st.success("Code copied")
                time.sleep(2)
            st.link_button("Open authenticator",url=st.session_state.verification_uri)

            if st.button("Complete Authentication"):
                try:
                    auth_manager.flow = st.session_state.flow
                    token = auth_manager.acquire_token()
                    st.session_state.token = token
                    st.session_state.authenticated = True
                    st.success("Authentication successful!")
                except Exception as e:
                    st.error(f"Authentication failed: {e}")

    # --- Show inputs and processing UI once authenticated ---
    if st.session_state.authenticated:
        st.success("Authenticated! You can now provide input details.")

        # Input fields
        to_run_folder = fold["Inp"]
        st.write(fold["Inp"])
        processed = fold["arch"]
        st.write(fold["arch"])
        output_folder = fold["OUT"]
        st.write(fold["OUT"])
        excel = fold["xlsx"]
        st.write(fold["xlsx"])
        layppt = st.selectbox("Layout PPTX Path",fold["lay"])
        foldname = st.text_input("OneDrive folder name:")

        # Button to download files
        if st.button("Download files"):
            try:
                with st.spinner("Downloading files..."):
                    files = list_and_download_files(st.session_state.token, foldname, to_run_folder)
                st.success(f"Downloaded {len(files)} files.")
                st.session_state.downloaded_files = files  # Store files in session state for processing
            except Exception as e:
                st.error(f"Error downloading files: {e}")

        # Button to start processing
        if st.button("Start Processing"):
            if not st.session_state.downloaded_files:
                st.warning("No files downloaded yet. Please download files first.")
            else:
                try:
                    total_files=len(st.session_state.downloaded_files)
                    count=1
                    for file in st.session_state.downloaded_files:
                        if os.path.isfile(file):
                            with st.spinner(f"Creating PPT...{count}/{total_files}"):
                                filename = os.path.basename(file)
                                inf=st.info(f"✅ Processing file: {filename}")
                                fi = Presentation()
                                out = os.path.join(output_folder, f"OUTPUT_{filename}.pptx")
                                fi.save(out)
                                inf2=st.info(f"✅ Starting ppt creation for {filename}")
                                pptcreate(excel, layppt, out, file)
                                suc=st.success(f"✅ PPT created for {filename} and saved to Output folder")
                                count+=1
                                inf.empty()
                                inf2.empty()
                                time.sleep(5)
                                suc.empty()
                                inf3=st.info("✅ Shifting files to Archive folder")
                                time.sleep(5)
                                inf3.empty()
                                # Uncomment to move processed input files
                                shutil.move(file, os.path.join(processed, filename))
                                count+=1
                    st.success("✅ All files processed!")
                    # Clear downloaded files after processing
                    st.session_state.downloaded_files = []
                except Exception as e:
                    st.error(f"❌ Encountered exception: {e}")


if __name__ == "__main__":
    main()

