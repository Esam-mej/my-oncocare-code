import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import sys
import os
import pandas as pd
from datetime import datetime
from tkcalendar import Calendar
import bcrypt
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from collections import defaultdict
import webbrowser
from docx import Document
from docx.shared import Pt
import shutil
import socket
import threading
import time
from PIL import Image, ImageTk
import firebase_admin
from firebase_admin import credentials, firestore
from firebase_admin.exceptions import FirebaseError
from concurrent.futures import ThreadPoolExecutor
from google.cloud.exceptions import NotFound
import schedule
import io
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError
import subprocess

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Constants
MALIGNANCIES = [
    "ALL", "AML", "LYMPHOMA", "EWING", "OSTEO", "NEUROBLASTOMA",
    "BRAIN T", "RHABDO", "RETINO", "HEPATO", "GERM CELL"
]

DROPDOWN_OPTIONS = {
    "GENDER": ["MALE", "FEMALE"],
    "B_SYMPTOMS": ["FEVER", "NIGHT SWEAT", "WT LOSS", "NO B-SYMPTOMS"],
    "EXAMINATION": [
        "FEBRILE", "HEPATOMEGALLY", "SPLENOMEGALLY", "LYMPHADENOPATHY",
        "TESTICULAR ENLARGMENT", "GUM HYPATROPHY", "CRANIAL PULSY",
        "RACCON EYES", "BONY SWELLING", "JOINT SWELLING", "SUBCONJUNCTIVAL HE",
        "HIGH BP", "HYPOTENSION", "OTHERS"
    ],
    "SYMPTOMS": [
        "ASYMPTOMATIC", "FEVER", "BONE PAIN", "BRUSIS", "WT LOSS",
        "NIGHT SWEATTING", "NECK SWELLING", "JOINT SWELLING",
        "ABDOMINAL DISTENSION", "HEMATURIA", "FRACTURE", "PALLOR",
        "VOMITTING", "DIARRHEA", "CONSTIPATION", "EPISTAXIS",
        "HEMATEMESIS", "GUM BLEEDING", "MELENA", "HEMATOCHEZIA", "OTHERS"
    ],
    "SURGERY": ["INDICATED", "NOT INDICATED", "DONE"],
    "SR_TYPE": [
        "COMPLETE RESECTION", "DEBULKING", "PALLIATIVE",
        "INCOMPLETE RESECTION", "DIAGNOSTIC &/OR STAGING"
    ],
    "STATE": ["ALIVE", "DECEASED", "MISSED FOLLOW UP"],
    "STEROID_R": ["GOOD", "POOR"],
    "TREATMENT_STAGE": [
        "PRE PHASE", "INDUCTION", "CONSOLIDATION",
        "MAINTENANCE", "OFF THERAPY"
    ],
    "RISK_GROUP": ["LRG", "SRG", "IRG", "HRG"],
    "BMT": ["INDICATED", "NO INDICATION", "DONE", "UNKNOWN"],
    "RADIOTHERAPY": ["INDICATED", "NOT INDICATED", "DONE"],
    "NECROSIS_GRADE": [
        "GRADE 1 <50%", "GRADE 2 >50%", "GRADE 3 >90%", "GRADE 4 >100%"
    ],
    "CSF": [
        "INITIAL 0", "INITIAL 1", "INITIAL 2", "INITIAL 3",
        "RELPASE 1", "RELAPSE 2", "RELAPSE 3"
    ],
    "STAGE": [
        "1", "2", "3", "4", "4S", "5", "REFRACTORY 1",
        "REFRACTORY 2", "REFRACTORY 3", "LOCAL", "METS"
    ],
    "THERAPY_SIDE_EFFECTS": [
        "MTX TOXICITY", "LEUKOENCEPHALOPATHY", "HGE CYSTITIS",
        "PERIPHERAL NEUROPATHY", "PULMONARY FIBROSIS",
        "L-ASP. SENSITIVITY", "OTHERS"
    ]
}

MALIGNANCY_FIELDS = {
    "ALL": [
        "GROUP", "CYTOGENETIC", "CSF", "MRD", "STEROID_R", 
        "TREATMENT_STAGE", "CYCLE", "RADIOTHERAPY", "BMT", 
        "STATE", "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "AML": [
        "CYTOGENETIC", "CSF", "MRD", "TREATMENT_STAGE", 
        "CYCLE", "RADIOTHERAPY", "BMT", "STATE", 
        "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "LYMPHOMA": [
        "GROUP", "STAGE", "CYTOGENETIC", "CSF", "MRD", 
        "STEROID_R", "TREATMENT_STAGE", "CYCLE", 
        "RADIOTHERAPY", "BMT", "STATE", "THERAPY_SIDE_EFFECTS", 
        "NOTES"
    ],
    "EWING": [
        "STAGE", "HISTOPATHOLOGY", "CYTOGENETIC", "SURGERY", 
        "SR_TYPE", "SR_DATE", "TREATMENT_STAGE", "CYCLE", 
        "RADIOTHERAPY", "BMT", "STATE", "THERAPY_SIDE_EFFECTS", 
        "NOTES"
    ],
    "OSTEO": [
        "STAGE", "HISTOPATHOLOGY", "CYTOGENETIC", "SURGERY", 
        "SR_TYPE", "SR_DATE", "TREATMENT_STAGE", "CYCLE", 
        "NECROSIS_GRADE", "RADIOTHERAPY", "BMT", "STATE", 
        "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "NEUROBLASTOMA": [
        "STAGE", "HISTOPATHOLOGY", "CYTOGENETIC", "SURGERY", 
        "SR_TYPE", "SR_DATE", "TREATMENT_STAGE", "CYCLE", 
        "RADIOTHERAPY", "BMT", "STATE", "THERAPY_SIDE_EFFECTS", 
        "NOTES"
    ],
    "BRAIN T": [
        "GROUP", "STAGE", "GRADE", "HISTOPATHOLOGY", 
        "CYTOGENETIC", "SURGERY", "SR_TYPE", "SR_DATE", 
        "TREATMENT_STAGE", "CYCLE", "RADIOTHERAPY", "BMT", 
        "STATE", "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "RHABDO": [
        "GROUP", "STAGE", "HISTOPATHOLOGY", "CYTOGENETIC", 
        "SURGERY", "SR_TYPE", "SR_DATE", "TREATMENT_STAGE", 
        "CYCLE", "RADIOTHERAPY", "BMT", "STATE", 
        "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "RETINO": [
        "GROUP", "STAGE", "GRADE", "HISTOPATHOLOGY", 
        "CYTOGENETIC", "SURGERY", "SR_TYPE", "SR_DATE", 
        "TREATMENT_STAGE", "CYCLE", "RADIOTHERAPY", "BMT", 
        "STATE", "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "HEPATO": [
        "GROUP", "STAGE", "GRADE", "HISTOPATHOLOGY", 
        "CYTOGENETIC", "SURGERY", "SR_TYPE", "SR_DATE", 
        "TREATMENT_STAGE", "CYCLE", "RADIOTHERAPY", "BMT", 
        "STATE", "THERAPY_SIDE_EFFECTS", "NOTES"
    ],
    "GERM CELL": [
        "GROUP", "STAGE", "GRADE", "HISTOPATHOLOGY", 
        "CYTOGENETIC", "SURGERY", "SR_TYPE", "SR_DATE", 
        "TREATMENT_STAGE", "CYCLE", "RADIOTHERAPY", "BMT", 
        "STATE", "THERAPY_SIDE_EFFECTS", "NOTES"
    ]
}

COMMON_FIELDS = [
    "FILE NUMBER", "NAME", "GENDER", "DATE OF BIRTH", "NATIONALITY",
    "NATIONAL NUMBER", "ADDRESS", "PHONE NUMBER", "FAMILY HISTORY OF MALIGNANCY",
    "ASSOCIATED DISEASE", "HISTORY", "EXAMINATION", "SYMPTOMS", "DIAGNOSIS",
    "AGE ON DIAGNOSIS", "INITIAL WBC"
]

MALIGNANCY_COLORS = {
    "ALL": "#FFCCCC", "AML": "#CCE5FF", "LYMPHOMA": "#FFFFCC",
    "EWING": "#E5FFCC", "OSTEO": "#FFCCE5", "NEUROBLASTOMA": "#CCFFFF",
    "BRAIN T": "#E5CCFF", "RHABDO": "#FFE5CC", "RETINO": "#CCFFCC",
    "HEPATO": "#FFCCFF", "GERM CELL": "#CCE5E5"
}

# Google Drive API Scopes
SCOPES = ['https://www.googleapis.com/auth/drive']
# Google Drive folder ID for the app (will be created if not exists)
DRIVE_FOLDER_NAME = "OncoCare"
DRIVE_PATIENTS_FOLDER_NAME = "OncoCare_Patients"

# Chemo Stocks Google Sheet URL
CHEMO_STOCKS_URL = "https://docs.google.com/spreadsheets/d/1QxcuCM4JLxPyePbaCvB1fdQOmEMgHwAL-bqRCaMvMfM/edit?usp=sharing"

# Supported file extensions for sync
SUPPORTED_EXTENSIONS = {
    'documents': ['.doc', '.docx', '.txt', '.pdf', '.rtf'],
    'spreadsheets': ['.xls', '.xlsx', '.xlsm', '.csv'],
    'images': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff'],
    'data': ['.json', '.xml'],
    'presentations': ['.ppt', '.pptx'],
    'archives': ['.zip', '.rar']
}

class GoogleDriveManager:
    """Handles all Google Drive operations for OncoCare with improved file sync"""
    
    def __init__(self):
        self.service = None
        self.initialized = False
        self.app_folder_id = None
        self.patients_folder_id = None
        self.initialize_drive()
    
    def initialize_drive(self):
        """Initialize Google Drive connection with error handling"""
        try:
            creds = None
            # The file token.json stores the user's access and refresh tokens
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            
            # If there are no (valid) credentials available, let the user log in
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    if os.path.exists('credentials.json'):
                        flow = InstalledAppFlow.from_client_secrets_file(
                            'credentials.json', SCOPES)
                        creds = flow.run_local_server(port=0)
                    else:
                        print("Google Drive credentials not found. File operations will be local only.")
                        return
                
                # Save the credentials for the next run
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            
            self.service = build('drive', 'v3', credentials=creds)
            self.initialized = True
            
            # Set up the app folder structure
            self.setup_app_folders()
            
        except Exception as e:
            print(f"Google Drive initialization failed: {e}")
            self.initialized = False
    
    def setup_app_folders(self):
        """Set up the necessary folder structure in Google Drive"""
        if not self.initialized:
            return
            
        try:
            # Check if main app folder exists
            query = f"name='{DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            if items:
                self.app_folder_id = items[0]['id']
            else:
                # Create main app folder
                folder_metadata = {
                    'name': DRIVE_FOLDER_NAME,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'description': 'Main folder for OncoCare application data'
                }
                folder = self.service.files().create(body=folder_metadata, fields='id').execute()
                self.app_folder_id = folder.get('id')
            
            # Check if patients folder exists
            query = f"name='{DRIVE_PATIENTS_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and '{self.app_folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            if items:
                self.patients_folder_id = items[0]['id']
            else:
                # Create patients folder inside app folder
                folder_metadata = {
                    'name': DRIVE_PATIENTS_FOLDER_NAME,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [self.app_folder_id],
                    'description': 'Folder for storing patient data and documents'
                }
                folder = self.service.files().create(body=folder_metadata, fields='id').execute()
                self.patients_folder_id = folder.get('id')
                
        except Exception as e:
            print(f"Error setting up Google Drive folders: {e}")
            self.initialized = False
    
    def upload_file(self, file_path, file_name, folder_id=None):
        """Upload a file to Google Drive with improved error handling"""
        if not self.initialized:
            return False
        
        try:
            # Check if file already exists
            query = f"name='{file_name}' and trashed=false"
            if folder_id:
                query += f" and '{folder_id}' in parents"
            
            existing_files = self.service.files().list(q=query, fields="files(id, modifiedTime)").execute().get('files', [])
            
            file_metadata = {
                'name': file_name
            }
            
            if folder_id:
                file_metadata['parents'] = [folder_id]
            
            media = MediaFileUpload(file_path)
            
            if existing_files:
                # Get the existing file's modification time
                drive_mtime = datetime.strptime(existing_files[0].get('modifiedTime'), '%Y-%m-%dT%H:%M:%S.%fZ').timestamp()
                local_mtime = os.path.getmtime(file_path)
                
                # Only update if local file is newer
                if local_mtime > drive_mtime:
                    file_id = existing_files[0]['id']
                    file = self.service.files().update(
                        fileId=file_id,
                        body=file_metadata,
                        media_body=media,
                        fields='id'
                    ).execute()
                    print(f"Updated file: {file_name}")
                else:
                    print(f"File {file_name} is up to date in Drive")
                    return True
            else:
                # Create new file
                file = self.service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id'
                ).execute()
                print(f"Uploaded new file: {file_name}")
            
            return True
        except HttpError as error:
            print(f"An error occurred while uploading file: {error}")
            return False
        except Exception as e:
            print(f"Error uploading file {file_name}: {e}")
            return False
    
    def download_file(self, file_id, save_path):
        """Download a file from Google Drive by file ID"""
        if not self.initialized:
            return False
        
        try:
            request = self.service.files().get_media(fileId=file_id)
            
            # Ensure directory exists
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            
            fh = io.FileIO(save_path, 'wb')
            downloader = MediaIoBaseDownload(fh, request)
            
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print(f"Download {int(status.progress() * 100)}%")
            
            return True
        except HttpError as error:
            print(f"An error occurred while downloading file: {error}")
            return False
        except Exception as e:
            print(f"Error downloading file: {e}")
            return False
    
    def create_patient_folder(self, file_no):
        """Create a folder for a patient in Google Drive"""
        if not self.initialized or not self.patients_folder_id:
            return None
        
        folder_name = f"Patient_{file_no}"
        try:
            # Check if folder already exists
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and '{self.patients_folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields="files(id)").execute()
            items = results.get('files', [])
            
            if items:
                return items[0]['id']
            
            # Create new folder
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [self.patients_folder_id],
                'description': f'Patient folder for file number {file_no}'
            }
            folder = self.service.files().create(body=folder_metadata, fields='id').execute()
            return folder.get('id')
        except Exception as e:
            print(f"Error creating patient folder: {e}")
            return None
    
    def upload_patient_data(self, patient_data):
        """Upload patient data to Google Drive"""
        if not self.initialized:
            return False
        
        try:
            # Save patient data to a temporary file
            temp_file = "temp_patient_data.json"
            with open(temp_file, 'w') as f:
                json.dump(patient_data, f)
            
            # Upload to the app folder
            file_name = f"patient_{patient_data['FILE NUMBER']}.json"
            success = self.upload_file(temp_file, file_name, self.app_folder_id)
            
            # Remove temporary file
            os.remove(temp_file)
            
            return success
        except Exception as e:
            print(f"Error uploading patient data: {e}")
            return False
    
    def download_all_patient_data(self):
        """Download all patient data from Google Drive"""
        if not self.initialized or not self.app_folder_id:
            return None
        
        try:
            # List all patient data files
            query = f"'{self.app_folder_id}' in parents and name contains 'patient_' and name contains '.json' and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            all_patients = []
            
            for item in items:
                # Download each file
                request = self.service.files().get_media(fileId=item['id'])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                
                # Parse the JSON data
                fh.seek(0)
                patient_data = json.load(fh)
                all_patients.append(patient_data)
            
            return all_patients
        except Exception as e:
            print(f"Error downloading patient data: {e}")
            return None
    
    def sync_patient_files(self, file_no):
        """Sync all files for a specific patient between local and Google Drive with two-way sync"""
        if not self.initialized:
            return False, "Google Drive not initialized"
        
        local_folder = f"Patient_{file_no}"
        if not os.path.exists(local_folder):
            os.makedirs(local_folder, exist_ok=True)
        
        # Get or create the patient folder in Drive
        drive_folder_id = self.create_patient_folder(file_no)
        if not drive_folder_id:
            return False, "Failed to create/get patient folder in Drive"
        
        try:
            # List files in local folder
            local_files = {}
            for root, _, files in os.walk(local_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, local_folder)
                    local_files[rel_path] = {
                        'path': file_path,
                        'mtime': os.path.getmtime(file_path),
                        'size': os.path.getsize(file_path)
                    }
            
            # List files in Drive folder
            query = f"'{drive_folder_id}' in parents and trashed=false"
            results = self.service.files().list(
                q=query, 
                fields="files(id, name, modifiedTime, size, mimeType)"
            ).execute()
            drive_files = {
                item['name']: {
                    'id': item['id'],
                    'mtime': datetime.strptime(item['modifiedTime'], '%Y-%m-%dT%H:%M:%S.%fZ').timestamp(),
                    'size': int(item.get('size', 0)),
                    'mimeType': item['mimeType']
                } 
                for item in results.get('files', [])
            }
            
            # Process files for two-way sync
            upload_count = 0
            download_count = 0
            
            # Upload new or modified files to Drive
            for rel_path, local_info in local_files.items():
                file_name = os.path.basename(rel_path)
                
                # Skip unsupported file types
                if not self.is_supported_file(file_name):
                    continue
                
                if file_name not in drive_files:
                    # New file - upload to Drive
                    if self.upload_file(local_info['path'], file_name, drive_folder_id):
                        upload_count += 1
                else:
                    # File exists in both - compare modification times
                    drive_info = drive_files[file_name]
                    if local_info['mtime'] > drive_info['mtime']:
                        # Local file is newer - upload to Drive
                        if self.upload_file(local_info['path'], file_name, drive_folder_id):
                            upload_count += 1
            
            # Download files from Drive that don't exist locally or are newer
            for file_name, drive_info in drive_files.items():
                local_path = os.path.join(local_folder, file_name)
                
                if not os.path.exists(local_path):
                    # File doesn't exist locally - download
                    if self.download_file(drive_info['id'], local_path):
                        download_count += 1
                else:
                    # File exists - compare modification times
                    local_mtime = os.path.getmtime(local_path)
                    if drive_info['mtime'] > local_mtime:
                        # Drive file is newer - download
                        if self.download_file(drive_info['id'], local_path):
                            download_count += 1
            
            return True, f"Sync complete. Uploaded: {upload_count}, Downloaded: {download_count}"
        except Exception as e:
            print(f"Error syncing patient files: {e}")
            return False, f"Sync failed: {str(e)}"
    
    def is_supported_file(self, filename):
        """Check if file extension is supported for sync"""
        ext = os.path.splitext(filename)[1].lower()
        for category, extensions in SUPPORTED_EXTENSIONS.items():
            if ext in extensions:
                return True
        return False
    
    def get_drive_file_mtime(self, file_id):
        """Get the modification time of a file in Google Drive by file ID"""
        try:
            file = self.service.files().get(fileId=file_id, fields='modifiedTime').execute()
            modified_time = file.get('modifiedTime')
            if modified_time:
                return datetime.strptime(modified_time, '%Y-%m-%dT%H:%M:%S.%fZ').timestamp()
        except Exception as e:
            print(f"Error getting file modification time: {e}")
        
        return 0

class FirebaseManager:
    """Handles all Firebase operations including initialization and data synchronization"""
    
    def __init__(self):
        self.db = None
        self.initialized = False
        self.initialize_firebase()
    
    def initialize_firebase(self):
        """Initialize Firebase connection with error handling"""
        try:
            if not firebase_admin._apps:
                # Use a service account (you'll need to provide your own credentials file)
                cred_path = 'serviceAccountKey.json'
                if os.path.exists(cred_path):
                    cred = credentials.Certificate(cred_path)
                    firebase_admin.initialize_app(cred)
                    self.db = firestore.client()
                    self.initialized = True
                else:
                    print("Firebase credentials not found. Offline mode only.")
        except Exception as e:
            print(f"Firebase initialization failed: {e}")
            self.initialized = False
    
    def sync_patients(self, local_data):
        """Synchronize patient data with Firebase with improved conflict resolution"""
        if not self.initialized:
            return False, "Firebase not initialized"
        
        try:
            # Get the latest data from Firebase
            firebase_data = []
            docs = self.db.collection('patients').stream()
            for doc in docs:
                patient = doc.to_dict()
                patient['_firestore_id'] = doc.id  # Store document ID for updates
                firebase_data.append(patient)
            
            # Merge strategies with improved conflict resolution
            merged_data = self.merge_data(local_data, firebase_data)
            
            # Update Firebase with merged data
            batch = self.db.batch()
            patients_ref = self.db.collection('patients')
            
            for patient in merged_data:
                # Use existing document ID if available (for updates)
                doc_id = patient.pop('_firestore_id', None) or patient['FILE NUMBER']
                doc_ref = patients_ref.document(doc_id)
                batch.set(doc_ref, patient)
            
            batch.commit()
            
            return True, "Sync completed successfully"
        
        except FirebaseError as e:
            return False, f"Firebase error: {e}"
        except Exception as e:
            return False, f"Unexpected error: {e}"
    
    def merge_data(self, local_data, firebase_data):
        """Merge local and Firebase data with improved conflict resolution"""
        merged = []
        local_by_id = {p['FILE NUMBER']: p for p in local_data}
        firebase_by_id = {p['FILE NUMBER']: p for p in firebase_data}
        
        all_ids = set(local_by_id.keys()).union(set(firebase_by_id.keys()))
        
        for file_no in all_ids:
            local_patient = local_by_id.get(file_no)
            firebase_patient = firebase_by_id.get(file_no)
            
            if local_patient and not firebase_patient:
                # Only exists locally - add to merged
                merged.append(local_patient)
            elif firebase_patient and not local_patient:
                # Only exists in Firebase - add to merged
                merged.append(firebase_patient)
            else:
                # Exists in both - use the most recently modified version
                try:
                    local_date = datetime.strptime(local_patient['LAST_MODIFIED_DATE'], '%d/%m/%Y %H:%M:%S')
                    firebase_date = datetime.strptime(firebase_patient['LAST_MODIFIED_DATE'], '%d/%m/%Y %H:%M:%S')
                    
                    if local_date > firebase_date:
                        # Local is newer - keep local and preserve Firebase ID
                        local_patient['_firestore_id'] = firebase_patient.get('_firestore_id', file_no)
                        merged.append(local_patient)
                    else:
                        # Firebase is newer or same - keep Firebase version
                        merged.append(firebase_patient)
                except:
                    # If date parsing fails, keep both with local version marked
                    local_patient['_conflict'] = True
                    merged.append(local_patient)
                    merged.append(firebase_patient)
        
        return merged
    
    def get_all_patients(self):
        """Retrieve all patients from Firebase with error handling"""
        if not self.initialized:
            return None
        
        try:
            patients = []
            docs = self.db.collection('patients').stream()
            for doc in docs:
                patient = doc.to_dict()
                patient['_firestore_id'] = doc.id  # Store document ID for reference
                patients.append(patient)
            return patients
        except Exception as e:
            print(f"Error fetching patients from Firebase: {e}")
            return None

class OncologyApp:
    """Main application class for Pediatric Oncology Patient Management System"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("OncoCare - Pediatric Oncology Patient Management System")
        self.root.geometry("1200x800")
        self.root.state('zoomed')
        
        # Set window icon
        try:
            self.root.iconbitmap('icon.ico')  # Replace with your icon file
        except:
            pass
        
        # Initialize Firebase manager
        self.firebase = FirebaseManager()
        
        # Initialize Google Drive manager
        self.drive = GoogleDriveManager()
        
        # Initialize variables
        self.current_user = None
        self.current_results = []
        self.current_result_index = 0
        self.last_sync_time = None
        self.internet_connected = False
        self.sync_in_progress = False
        
        # Configure styles
        self.setup_styles()
        
        # Setup status bar
        self.setup_status_bar()
        
        # Check internet connection
        self.check_internet_connection()
        
        # Load users
        self.load_users()
        
        # Load patient data
        self.load_patient_data()
        
        # Start with login screen
        self.setup_login_screen()
        
        # Initialize thread pool for background tasks
        self.executor = ThreadPoolExecutor(max_workers=4)  # Increased workers for better sync
        
        # Handle window close properly
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    # ... [Previous methods remain the same until the sync_data method] ...
    #     
    def setup_styles(self):
        """Configure the visual styles for the application"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors
        self.primary_color = '#2c3e50'  # Dark blue
        self.secondary_color = '#3498db'  # Blue
        self.accent_color = '#e74c3c'  # Red
        self.light_color = '#ecf0f1'  # Light gray
        self.dark_color = '#2c3e50'  # Dark blue
        self.success_color = '#27ae60'  # Green
        self.warning_color = '#f39c12'  # Orange
        self.danger_color = '#e74c3c'  # Red
        
        # Configure styles
        self.style.configure('TFrame', background=self.light_color)
        self.style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=8)
        self.style.map('TButton',
            background=[('active', self.secondary_color)],
            foreground=[('active', 'white')]
        )
        
        # Custom button styles
        self.style.configure('Blue.TButton', background=self.secondary_color, foreground='white')
        self.style.map('Blue.TButton',
            background=[('active', '#2980b9')],
            foreground=[('active', 'white')]
        )
        
        self.style.configure('Green.TButton', background=self.success_color, foreground='white')
        self.style.map('Green.TButton',
            background=[('active', '#219653')],
            foreground=[('active', 'white')]
        )
        
        self.style.configure('Yellow.TButton', background=self.warning_color, foreground='white')
        self.style.map('Yellow.TButton',
            background=[('active', '#e67e22')],
            foreground=[('active', 'white')]
        )
        
        self.style.configure('Brown.TButton', background='#8B4513', foreground='white')
        self.style.map('Brown.TButton',
            background=[('active', '#A0522D')],
            foreground=[('active', 'white')]
        )
        
        self.style.configure('Red.TButton', background=self.danger_color, foreground='white')
        self.style.map('Red.TButton',
            background=[('active', '#c0392b')],
            foreground=[('active', 'white')]
        )
        
        self.style.configure('TLabel', font=('Helvetica', 10), background=self.light_color)
        self.style.configure('Header.TLabel', font=('Helvetica', 24, 'bold'), foreground=self.primary_color)
        self.style.configure('Subheader.TLabel', font=('Helvetica', 18), foreground=self.secondary_color)
        self.style.configure('TEntry', font=('Helvetica', 10), padding=5)
        self.style.configure('TCombobox', font=('Helvetica', 10))
        self.style.configure('Status.TFrame', background=self.primary_color)
        self.style.configure('Status.TLabel', background=self.primary_color, foreground='white')
        self.style.configure('Malignancy.TCombobox', font=('Helvetica', 12, 'bold'))
    
    def setup_status_bar(self):
        """Setup the status bar at the bottom of the window"""
        self.status_frame = ttk.Frame(self.root, style='Status.TFrame', height=30)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # User info
        self.user_label = ttk.Label(self.status_frame, text="Not logged in", style='Status.TLabel')
        self.user_label.pack(side=tk.LEFT, padx=10)
        
        # Internet status indicator
        self.internet_indicator = tk.Canvas(self.status_frame, width=20, height=20, highlightthickness=0)
        self.internet_indicator.pack(side=tk.LEFT, padx=5)
        self.update_internet_indicator()
        
        # Sync status
        self.sync_status = ttk.Label(self.status_frame, text="", style='Status.TLabel')
        self.sync_status.pack(side=tk.LEFT, padx=10)
        
        # Date and time
        self.datetime_label = ttk.Label(self.status_frame, text="", style='Status.TLabel')
        self.datetime_label.pack(side=tk.RIGHT, padx=10)
        
        # Sync button
        self.sync_btn = ttk.Button(self.status_frame, text="Synchronize Data", 
                                  command=self.sync_data, style='Blue.TButton')
        self.sync_btn.pack(side=tk.RIGHT, padx=10)
        
        # Update datetime
        self.update_datetime()
    
    def load_patient_data(self):
        """Load patient data from file or create empty list"""
        if not os.path.exists('patients_data.json'):
            self.patient_data = []
            return
            
        try:
            with open('patients_data.json', 'r') as f:
                self.patient_data = json.load(f)
        except json.JSONDecodeError:
            self.patient_data = []
    
    def save_patient_data(self):
        """Save patient data to file"""
        with open('patients_data.json', 'w') as f:
            json.dump(self.patient_data, f, indent=4)
    
    def update_internet_indicator(self):
        """Update the internet connection indicator"""
        color = "green" if self.internet_connected else "red"
        self.internet_indicator.delete("all")
        self.internet_indicator.create_oval(5, 5, 15, 15, fill=color, outline="")
    
    def update_datetime(self):
        """Update the date and time display"""
        if hasattr(self, 'datetime_label') and self.datetime_label.winfo_exists():
            now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            self.datetime_label.config(text=now)
            # Use after_idle to prevent "invalid command name" errors
            self.root.after(1000, self.update_datetime)
    
    def check_internet_connection(self):
        """Check internet connection status periodically"""
        def internet_check():
            try:
                socket.create_connection(("www.google.com", 80), timeout=5)
                self.internet_connected = True
            except:
                self.internet_connected = False
            
            self.update_internet_indicator()
            # Use after_idle to prevent "invalid command name" errors
            self.root.after(30000, internet_check)  # Check every 30 seconds
        
        internet_check()
    
    def load_users(self):
        """Load user data from file or create default users"""
        if os.path.exists("users_data.json"):
            try:
                with open("users_data.json", "r") as f:
                    self.users = json.load(f)
            except json.JSONDecodeError:
                # Handle empty or corrupted file
                self.users = self.create_default_users()
                self.save_users_to_file()
        else:
            self.users = self.create_default_users()
            self.save_users_to_file()
    
    def create_default_users(self):
        """Create default user data with hashed passwords"""
        return {
            "mej.esam": {
                "password": bcrypt.hashpw("wjap19527".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                "role": "admin"
            },
            "doctor1": {
                "password": bcrypt.hashpw("doc123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                "role": "editor"
            },
            "nurse1": {
                "password": bcrypt.hashpw("nur123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                "role": "viewer"
            },
            "pharmacist1": {
                "password": bcrypt.hashpw("pharm123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                "role": "pharmacist"
            },
            "seraj": {
                "password": bcrypt.hashpw("steve8288".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
                "role": "admin"
            }
        }
    
    def save_users_to_file(self):
        """Save user data to file"""
        with open("users_data.json", "w") as f:
            json.dump(self.users, f, indent=4)
    
    def on_close(self):
        """Handle window close event"""
        if messagebox.askyesno("Exit Confirmation", "Are you sure you want to exit?"):
            # Cancel all pending after events
            for after_id in self.root.tk.eval('after info').split():
                self.root.after_cancel(after_id)
            
            self.executor.shutdown(wait=False)
            self.root.destroy()
    
    def clear_frame(self):
        """Clear all widgets from the main frame except status bar"""
        for widget in self.root.winfo_children():
            if widget not in [self.status_frame]:
                widget.destroy()
    
    def setup_login_screen(self):
        """Setup the login screen with modern design"""
        self.clear_frame()
        
        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 36, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(40, 10))
        
        tk.Label(logo_frame, text="Pediatric Oncology Patient Management System", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with login form
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Login form container
        form_container = tk.Frame(right_frame, bg='white')
        form_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # Login header
        tk.Label(form_container, text="Login to OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='white').pack(pady=(0, 30))
        
        # Username field
        tk.Label(form_container, text="Username", font=('Helvetica', 12), 
                bg='white', anchor='w').pack(fill=tk.X, pady=(10, 0))
        self.entry_username = ttk.Entry(form_container, font=('Helvetica', 12))
        self.entry_username.pack(fill=tk.X, ipady=5, pady=(0, 20))
        
        # Password field
        tk.Label(form_container, text="Password", font=('Helvetica', 12), 
                bg='white', anchor='w').pack(fill=tk.X)
        self.entry_password = ttk.Entry(form_container, show="*", font=('Helvetica', 12))
        self.entry_password.pack(fill=tk.X, ipady=5, pady=(0, 30))
        
        # Login button
        login_btn = ttk.Button(form_container, text="Login", command=self.login, 
                             style='Blue.TButton')
        login_btn.pack(fill=tk.X, ipady=10, pady=(0, 20))
        
        # Exit button
        exit_btn = ttk.Button(form_container, text="Exit", command=self.root.quit)
        exit_btn.pack(fill=tk.X, ipady=10)
        
        # Focus on username field
        self.entry_username.focus()
        
        # Bind Enter key to login
        self.entry_password.bind('<Return>', lambda event: self.login())
    
    def login(self):
        """Handle user login"""
        username = self.entry_username.get()
        password = self.entry_password.get()

        if username in self.users:
            stored_hash = self.users[username]["password"].encode('utf-8')
            if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
                self.current_user = username
                self.user_label.config(text=f"User: {username} ({self.users[username]['role']})")
                messagebox.showinfo("Login Successful", f"Welcome, {username}!")
                self.main_menu()
                return

        messagebox.showerror("Login Failed", "Invalid username or password.")
        self.entry_password.delete(0, tk.END)
    
    def main_menu(self):
        """Display the main menu with organized buttons"""
        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text=f"Welcome, {self.current_user}", font=('Helvetica', 16, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="Main Menu", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with menu buttons
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Menu form container
        form_container = tk.Frame(right_frame, bg='white')
        form_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # Button grid
        btn_frame = ttk.Frame(form_container, style='TFrame')
        btn_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        # First row - Patient Management (Blue buttons)
        row1_frame = ttk.Frame(btn_frame, style='TFrame')
        row1_frame.pack(fill=tk.X, pady=5)
        
        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(row1_frame, text="Add New Patient", command=self.add_patient,
                      style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        ttk.Button(row1_frame, text="Search Patient", command=self.search_patient,
                  style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.users[self.current_user]["role"] == "admin":
            ttk.Button(row1_frame, text="View All Patients", command=self.view_all_patients,
                      style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Second row - Data Operations (Green buttons)
        row2_frame = ttk.Frame(btn_frame, style='TFrame')
        row2_frame.pack(fill=tk.X, pady=5)
        
        if self.users[self.current_user]["role"] == "admin":
            ttk.Button(row2_frame, text="Export All Data", command=self.export_all_data,
                      style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            
            ttk.Button(row2_frame, text="Backup Data", command=self.backup_data,
                      style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            
            if self.current_user == "mej.esam":
                ttk.Button(row2_frame, text="Restore Data", command=self.restore_data,
                          style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            else:
                # Empty space to maintain layout
                ttk.Frame(row2_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Third row - Clinical Tools (Yellow buttons)
        row3_frame = ttk.Frame(btn_frame, style='TFrame')
        row3_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(row3_frame, text="CHEMO PROTOCOLS", command=self.show_chemo_protocols,
                  style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        ttk.Button(row3_frame, text="CHEMO SHEETS", command=self.show_chemo_sheets,
                  style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.users[self.current_user]["role"] in ["admin", "pharmacist"]:
            ttk.Button(row3_frame, text="CHEMO STOCKS", command=self.show_chemo_stocks,
                      style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row3_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(row3_frame, text="Statistics", command=self.show_statistics,
                      style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row3_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Fourth row - User Management (Brown buttons)
        row4_frame = ttk.Frame(btn_frame, style='TFrame')
        row4_frame.pack(fill=tk.X, pady=5)
        
        if self.current_user == "mej.esam" or self.users[self.current_user]["role"] == "admin":
            ttk.Button(row4_frame, text="Manage Users", command=self.manage_users,
                      style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            
            ttk.Button(row4_frame, text="Change Password", command=self.change_password,
                      style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row4_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Fifth row - Logout (Red button)
        row5_frame = ttk.Frame(btn_frame, style='TFrame')
        row5_frame.pack(fill=tk.X, pady=(20, 5))
        
        ttk.Button(row5_frame, text="Logout", command=self.setup_login_screen,
                  style='Red.TButton').pack(fill=tk.X, ipady=10)

        # Signature
        signature_frame = ttk.Frame(form_container, style='TFrame')
        signature_frame.pack(pady=(30, 0))

        ttk.Label(signature_frame, text="Made by: DR. ESAM MEJRAB",
                 font=('Times New Roman', 14, 'italic'), 
                 foreground=self.primary_color).pack()
    
    def show_chemo_protocols(self):
        """Open the Protocols folder"""
        try:
            protocols_path = get_resource_path("Protocols")
            if os.path.exists(protocols_path):
                os.startfile(protocols_path)
            else:
                messagebox.showerror("Error", f"Protocols folder not found at: {protocols_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Protocols folder: {e}")

    def show_chemo_sheets(self):
        """Open the FULL CHEMO SHEET.xlsm file"""
        try:
            sheet_path = get_resource_path("FULL CHEMO SHEET.xlsm")
            if os.path.exists(sheet_path):
                os.startfile(sheet_path)
            else:
                messagebox.showerror("Error", f"FULL CHEMO SHEET.xlsm not found at: {sheet_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open chemo sheet: {e}")

    def show_chemo_stocks(self):
        """Open the Chemo Stocks Google Sheet"""
        try:
            webbrowser.open(CHEMO_STOCKS_URL)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Chemo Stocks: {e}")
    
    def show_statistics(self):
        """Show statistics window"""
        self.open_statistics_window()
    
    def create_scrollable_frame(self, parent):
        """Create a scrollable frame within the given parent"""
        container = ttk.Frame(parent)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        container.pack(fill="both", expand=True)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)
        
        return scrollable_frame
    
    def add_patient(self):
        """Display the add patient form with scrollable fields"""
        if self.users[self.current_user]["role"] not in ["admin", "editor"]:
            messagebox.showerror("Access Denied", "Only admins and editors can add patients")
            return

        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="Add New Patient", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with form
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create a container frame for the canvas and scrollbar
        container = ttk.Frame(right_frame)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Create a canvas
        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add a scrollbar
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure the canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create a frame inside the canvas
        form_container = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=form_container, anchor="nw")
        
        # Bind the canvas to the scroll region
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        form_container.bind("<Configure>", on_frame_configure)
        
        # Enable mouse wheel scrolling
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        # Malignancy selection
        malignancy_frame = ttk.LabelFrame(form_container, text="Select Malignancy Type", style='TFrame')
        malignancy_frame.pack(fill=tk.X, padx=10, pady=10)

        self.malignancy_var = tk.StringVar()
        
        # Create a bold combobox for malignancy selection
        ttk.Label(malignancy_frame, text="Malignancy Type:", 
                 font=('Helvetica', 12)).pack(pady=5, anchor="w")
        
        malignancy_combo = ttk.Combobox(malignancy_frame, textvariable=self.malignancy_var,
                                      values=MALIGNANCIES, state="readonly",
                                      font=('Helvetica', 12, 'bold'), style='Malignancy.TCombobox')
        malignancy_combo.pack(fill=tk.X, padx=5, pady=5)
        
        # Bind the selection to show malignancy fields
        malignancy_combo.bind("<<ComboboxSelected>>", lambda e: self.show_malignancy_fields(form_container))

        # Container for malignancy-specific fields
        self.malignancy_fields_frame = ttk.Frame(form_container, style='TFrame')
        self.malignancy_fields_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Button frame at the bottom (outside the scrollable area)
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Back", command=self.main_menu).pack(side=tk.LEFT, padx=10)
    
    def show_malignancy_fields(self, form_container):
        """Show fields specific to the selected malignancy type"""
        # Clear previous fields
        for widget in self.malignancy_fields_frame.winfo_children():
            widget.destroy()

        malignancy = self.malignancy_var.get()
        if not malignancy:
            return

        # Common fields
        common_frame = ttk.LabelFrame(self.malignancy_fields_frame, text="Common Information", style='TFrame')
        common_frame.pack(fill=tk.X, padx=10, pady=10)

        self.entries = {}
        row = 0

        # Create common fields
        for field in COMMON_FIELDS:
            ttk.Label(common_frame, text=f"{field}:", anchor="w").grid(row=row, column=0, padx=5, pady=5, sticky="w")
            
            if field == "GENDER":
                var = tk.StringVar()
                combobox = ttk.Combobox(common_frame, textvariable=var, 
                                       values=DROPDOWN_OPTIONS["GENDER"], state="readonly")
                combobox.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.entries[field] = var
            elif field in ["EXAMINATION", "SYMPTOMS"]:
                listbox_frame = ttk.Frame(common_frame)
                listbox_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4, exportselection=0)
                scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                listbox.configure(yscrollcommand=scrollbar.set)
                
                options = DROPDOWN_OPTIONS[field]
                for option in options:
                    listbox.insert(tk.END, option)
                
                listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                self.entries[field] = listbox
                
                # Add "Others" entry if needed
                if "OTHERS" in options:
                    ttk.Label(common_frame, text=f"{field} (Others):").grid(row=row+1, column=0, padx=5, pady=5, sticky="w")
                    others_entry = ttk.Entry(common_frame)
                    others_entry.grid(row=row+1, column=1, padx=5, pady=5, sticky="ew")
                    self.entries[f"{field}_OTHERS"] = others_entry
                    row += 1
            elif field == "DATE OF BIRTH":
                entry_frame = ttk.Frame(common_frame)
                entry_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                entry = ttk.Entry(entry_frame)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                cal_btn = ttk.Button(entry_frame, text="📅", width=3,
                                    command=lambda e=entry: self.show_calendar(e))
                cal_btn.pack(side=tk.LEFT, padx=5)
                self.entries[field] = entry
            else:
                entry = ttk.Entry(common_frame)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.entries[field] = entry
            
            row += 1

        # Malignancy-specific fields
        malignancy_frame = tk.LabelFrame(self.malignancy_fields_frame, 
                                       text=f"{malignancy} Specific Information", 
                                       bg=MALIGNANCY_COLORS.get(malignancy, "#FFFFFF"))
        malignancy_frame.pack(fill=tk.X, padx=10, pady=10)

        row = 0
        for field in MALIGNANCY_FIELDS.get(malignancy, []):
            ttk.Label(malignancy_frame, text=f"{field}:", anchor="w").grid(row=row, column=0, padx=5, pady=5, sticky="w")
            
            if field in DROPDOWN_OPTIONS:
                var = tk.StringVar()
                combobox = ttk.Combobox(malignancy_frame, textvariable=var, 
                                       values=DROPDOWN_OPTIONS[field], state="readonly")
                combobox.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.entries[field] = var
            elif field == "SR_DATE":
                entry_frame = ttk.Frame(malignancy_frame)
                entry_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                entry = ttk.Entry(entry_frame)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                cal_btn = ttk.Button(entry_frame, text="📅", width=3,
                                    command=lambda e=entry: self.show_calendar(e))
                cal_btn.pack(side=tk.LEFT, padx=5)
                self.entries[field] = entry
            elif field == "THERAPY_SIDE_EFFECTS":
                listbox_frame = ttk.Frame(malignancy_frame)
                listbox_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4, exportselection=0)
                scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                listbox.configure(yscrollcommand=scrollbar.set)
                
                options = DROPDOWN_OPTIONS[field]
                for option in options:
                    listbox.insert(tk.END, option)
                
                listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                self.entries[field] = listbox
                
                # Add "Others" entry if needed
                if "OTHERS" in options:
                    ttk.Label(malignancy_frame, text=f"{field} (Others):").grid(row=row+1, column=0, padx=5, pady=5, sticky="w")
                    others_entry = ttk.Entry(malignancy_frame)
                    others_entry.grid(row=row+1, column=1, padx=5, pady=5, sticky="ew")
                    self.entries[f"{field}_OTHERS"] = others_entry
                    row += 1
            else:
                entry = ttk.Entry(malignancy_frame)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.entries[field] = entry
            
            row += 1

        # Button frame at the bottom (inside the scrollable frame)
        btn_frame = ttk.Frame(self.malignancy_fields_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(btn_frame, text="Save Patient", command=self.save_patient,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Patient Folder", command=self.open_patient_folder,
                  style='Green.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Patient Report", command=self.create_patient_report,
                  style='Yellow.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu).pack(side=tk.LEFT, padx=10)
    
    def show_calendar(self, entry_widget):
        """Show a calendar popup for date selection"""
        top = tk.Toplevel(self.root)
        top.title("Select Date")
        top.geometry("300x300")
        
        cal = Calendar(top, selectmode='day', date_pattern='dd/mm/yyyy')
        cal.pack(pady=20)
        
        def set_date():
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, cal.get_date())
            top.destroy()
        
        ttk.Button(top, text="Select", command=set_date, style='Blue.TButton').pack(pady=10)
    
    def save_patient(self):
        """Save patient data to file and sync with Google Drive"""
        # Check if malignancy is selected
        malignancy = self.malignancy_var.get()
        if not malignancy:
            messagebox.showerror("Error", "Please select a malignancy type")
            return

        # Check mandatory fields
        mandatory_fields = [
            "FILE NUMBER", "NAME", "GENDER", "DATE OF BIRTH", "NATIONALITY",
            "DIAGNOSIS", "AGE ON DIAGNOSIS"
        ]
        
        missing_fields = []
        for field in mandatory_fields:
            value = self.entries[field].get() if isinstance(self.entries[field], (tk.Entry, tk.StringVar)) else None
            if not value:
                missing_fields.append(field)
        
        # Check at least one from examination, history, symptoms
        exam_selected = False
        if "EXAMINATION" in self.entries:
            exam_selected = len(self.entries["EXAMINATION"].curselection()) > 0 or \
                          ("EXAMINATION_OTHERS" in self.entries and self.entries["EXAMINATION_OTHERS"].get())
        
        history_selected = "HISTORY" in self.entries and self.entries["HISTORY"].get()
        symptoms_selected = False
        if "SYMPTOMS" in self.entries:
            symptoms_selected = len(self.entries["SYMPTOMS"].curselection()) > 0 or \
                              ("SYMPTOMS_OTHERS" in self.entries and self.entries["SYMPTOMS_OTHERS"].get())
        
        if not (exam_selected or history_selected or symptoms_selected):
            missing_fields.append("At least one from EXAMINATION, HISTORY, or SYMPTOMS")
        
        if missing_fields:
            messagebox.showerror("Error", f"The following fields are mandatory:\n{', '.join(missing_fields)}")
            return

        # Check file number uniqueness and numeric value
        file_no = self.entries["FILE NUMBER"].get()
        try:
            file_no_int = int(file_no)
        except ValueError:
            messagebox.showerror("Error", "File Number must be a numeric value")
            return
            
        for patient in self.patient_data:
            if patient.get("FILE NUMBER") == file_no:
                messagebox.showerror("Error", "Patient with this File Number already exists")
                return

        # Validate date format
        dob = self.entries["DATE OF BIRTH"].get()
        try:
            datetime.strptime(dob, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format for 'DATE OF BIRTH'. Please use dd/mm/yyyy")
            return

        # Collect patient data
        patient_data = {
            "MALIGNANCY": malignancy,
            "CREATED_BY": self.current_user,
            "CREATED_DATE": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "LAST_MODIFIED_BY": self.current_user,
            "LAST_MODIFIED_DATE": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }

        for field, widget in self.entries.items():
            if isinstance(widget, tk.Listbox):  # Multi-select fields
                selected = [widget.get(i) for i in widget.curselection()]
                if "OTHERS" in selected and f"{field}_OTHERS" in self.entries:
                    others_text = self.entries[f"{field}_OTHERS"].get()
                    if others_text:
                        selected[selected.index("OTHERS")] = f"OTHERS: {others_text}"
                patient_data[field] = ", ".join(selected) if selected else ""
            elif isinstance(widget, tk.StringVar):  # Combobox
                patient_data[field] = widget.get()
            else:  # Entry fields
                patient_data[field] = widget.get()

        # Save the patient data
        self.patient_data.append(patient_data)
        # Sort patient data by file number (numeric)
        self.patient_data.sort(key=lambda x: int(x.get("FILE NUMBER", 0)))
        self.save_patient_data()

        # Create patient folder and report
        self.create_patient_folder(file_no)
        self.create_patient_report(file_no, create_only=True)

        # Upload to Google Drive in background
        self.executor.submit(self.upload_patient_to_drive, patient_data)

        messagebox.showinfo("Success", "Patient data saved successfully!")
        self.main_menu()
    
    def upload_patient_to_drive(self, patient_data):
        """Upload patient data and files to Google Drive"""
        if not self.drive.initialized:
            return False
        
        try:
            # Upload patient data
            file_no = patient_data["FILE NUMBER"]
            self.drive.upload_patient_data(patient_data)
            
            # Sync patient folder
            self.drive.sync_patient_files(file_no)
            
            return True
        except Exception as e:
            print(f"Error uploading patient to Google Drive: {e}")
            return False
    
    def create_patient_folder(self, file_no):
        """Create a folder for the patient's documents"""
        folder_name = f"Patient_{file_no}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
    
    def open_patient_folder(self):
        """Open the patient's folder in file explorer"""
        file_no = self.entries["FILE NUMBER"].get()
        if not file_no:
            messagebox.showerror("Error", "Please enter a file number first")
            return
        
        folder_name = f"Patient_{file_no}"
        self.create_patient_folder(file_no)
        
        # Try to sync with Google Drive
        if self.drive.initialized:
            self.executor.submit(self.drive.sync_patient_files, file_no)
        
        try:
            os.startfile(folder_name)
        except:
            messagebox.showerror("Error", f"Could not open folder: {folder_name}")
    
    def create_patient_report(self, file_no=None, create_only=False):
        """Create a Word document report for the patient"""
        if not file_no:
            file_no = self.entries["FILE NUMBER"].get()
            if not file_no:
                messagebox.showerror("Error", "Please enter a file number first")
                return
        
        doc_name = f"Patient_{file_no}_Report.docx"
        doc_path = os.path.join(f"Patient_{file_no}", doc_name)
        
        if not os.path.exists(doc_path) or not create_only:
            doc = Document()
            
            # Add title
            doc.add_heading(f'Patient Report - File Number: {file_no}', 0)
            
            # Add basic info
            doc.add_heading('Basic Information', level=1)
            
            if not create_only:
                # If creating from current form data
                for field in COMMON_FIELDS:
                    if field in self.entries:
                        value = ""
                        if isinstance(self.entries[field], tk.Listbox):
                            value = ", ".join([self.entries[field].get(i) for i in self.entries[field].curselection()])
                        elif isinstance(self.entries[field], tk.StringVar):
                            value = self.entries[field].get()
                        else:
                            value = self.entries[field].get()
                        
                        doc.add_paragraph(f"{field}: {value}", style='List Bullet')
            else:
                # If creating from existing data
                patient = next((p for p in self.patient_data if p.get("FILE NUMBER") == file_no), None)
                if patient:
                    for field in COMMON_FIELDS:
                        if field in patient:
                            doc.add_paragraph(f"{field}: {patient[field]}", style='List Bullet')
            
            # Save document
            doc.save(doc_path)
        
        if not create_only:
            # Upload to Google Drive in background
            if self.drive.initialized:
                self.executor.submit(self.drive.upload_file, doc_path, doc_name, self.drive.create_patient_folder(file_no))
            
            try:
                os.startfile(doc_path)
            except:
                messagebox.showerror("Error", f"Could not open document: {doc_path}")

    def search_patient(self):
        """Display the patient search form"""
        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="Search Patient", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with search form
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Search form container
        form_container = ttk.Frame(right_frame, style='TFrame')
        form_container.place(relx=0.5, rely=0.5, anchor='center')

        ttk.Label(form_container, text="File Number:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        self.search_file_no_entry = ttk.Entry(form_container, width=40, 
                                            font=('Helvetica', 12))
        self.search_file_no_entry.pack(pady=5)

        ttk.Label(form_container, text="Name:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        self.search_name_entry = ttk.Entry(form_container, width=40, 
                                         font=('Helvetica', 12))
        self.search_name_entry.pack(pady=5)

        btn_frame = ttk.Frame(form_container, style='TFrame')
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="Search", command=self.perform_search,
                   style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Back", command=self.main_menu).pack(side=tk.LEFT, padx=10)

    def perform_search(self):
        """Perform patient search based on criteria"""
        file_no = self.search_file_no_entry.get().strip()
        name = self.search_name_entry.get().strip()

        if not self.patient_data:
            messagebox.showerror("Error", "No patient data available")
            return

        # Filter results based on search criteria
        matching_data = []
        for patient in self.patient_data:
            match = True
            
            if file_no and patient.get("FILE NUMBER") != file_no:
                match = False
            if name and name.lower() not in patient.get("NAME", "").lower():
                match = False
            
            if match:
                matching_data.append(patient)

        if not matching_data:
            messagebox.showerror("Not Found", "No patient found with the given criteria")
            return

        # Handle multiple results
        self.current_results = matching_data
        self.current_result_index = 0
        self.view_patient(self.current_results[self.current_result_index], 
                         multiple_results=(len(matching_data) > 1))

    def view_patient(self, patient_data, multiple_results=False):
        """View patient details"""
        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        malignancy = patient_data.get("MALIGNANCY", "")
        title = f"Patient Details - {malignancy}" if malignancy else "Patient Details"
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text=title, 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with patient details
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Main content
        form_frame = self.create_scrollable_frame(right_frame)

        # Display patient data in a clean layout with alternating colors
        bg_color = MALIGNANCY_COLORS.get(malignancy, "#e6f2ff")
        
        # User info
        user_frame = ttk.Frame(form_frame, style='TFrame')
        user_frame.pack(fill=tk.X, pady=5, padx=10)
        
        created_by = patient_data.get("CREATED_BY", "Unknown")
        created_date = patient_data.get("CREATED_DATE", "Unknown")
        modified_by = patient_data.get("LAST_MODIFIED_BY", "Unknown")
        modified_date = patient_data.get("LAST_MODIFIED_DATE", "Unknown")
        
        ttk.Label(user_frame, text=f"Added by: {created_by} on {created_date}", 
                 font=('Helvetica', 10, 'italic')).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(user_frame, text=f"Last modified by: {modified_by} on {modified_date}", 
                 font=('Helvetica', 10, 'italic')).pack(side=tk.RIGHT, padx=5)

        # Common fields
        common_frame = ttk.LabelFrame(form_frame, text="Common Information", style='TFrame')
        common_frame.pack(fill=tk.X, padx=10, pady=10)

        for i, field in enumerate(COMMON_FIELDS):
            row_frame = tk.Frame(common_frame, bg=bg_color if i % 2 == 0 else '#f0f7ff')
            row_frame.pack(fill=tk.X, ipady=5)

            tk.Label(row_frame, text=f"{field}:", width=25, anchor="w",
                     bg=row_frame['bg'], font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT, padx=5)
            
            tk.Label(row_frame, text=patient_data.get(field, ""), width=40, anchor="w",
                     bg=row_frame['bg'], font=('Helvetica', 10)).pack(side=tk.LEFT)

        # Malignancy-specific fields
        if malignancy in MALIGNANCY_FIELDS:
            specific_frame = tk.LabelFrame(form_frame, text=f"{malignancy} Specific Information", 
                                          bg=bg_color)
            specific_frame.pack(fill=tk.X, padx=10, pady=10)

            for i, field in enumerate(MALIGNANCY_FIELDS[malignancy]):
                row_frame = tk.Frame(specific_frame, bg=bg_color if i % 2 == 0 else '#f0f7ff')
                row_frame.pack(fill=tk.X, ipady=5)

                tk.Label(row_frame, text=f"{field}:", width=25, anchor="w",
                         bg=row_frame['bg'], font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT, padx=5)
                
                tk.Label(row_frame, text=patient_data.get(field, ""), width=40, anchor="w",
                         bg=row_frame['bg'], font=('Helvetica', 10)).pack(side=tk.LEFT)

        # Button frame
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)

        file_no = patient_data.get("FILE NUMBER", "")
        
        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(btn_frame, text="Edit", command=lambda: self.edit_patient(patient_data),
                      style='Blue.TButton').pack(side=tk.LEFT, padx=10)

        if self.current_user == "mej.esam":
            ttk.Button(btn_frame, text="Delete", command=lambda: self.delete_patient(patient_data),
                      style='Red.TButton').pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="Patient Folder", 
                  command=lambda: self.open_existing_patient_folder(file_no),
                  style='Green.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Patient Report", 
                  command=lambda: self.open_existing_patient_report(file_no),
                  style='Yellow.TButton').pack(side=tk.LEFT, padx=10)

        # Show "Previous" and "Next" buttons only if multiple results are available
        if multiple_results:
            ttk.Button(btn_frame, text="Previous", 
                      command=self.show_previous_result).pack(side=tk.LEFT, padx=10)
            
            ttk.Button(btn_frame, text="Next", 
                      command=self.show_next_result).pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="Back to Search", 
                  command=self.search_patient).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Main Menu", command=self.main_menu, 
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=10)

    def open_existing_patient_folder(self, file_no):
        """Open an existing patient's folder"""
        folder_name = f"Patient_{file_no}"
        if not os.path.exists(folder_name):
            self.create_patient_folder(file_no)
        
        # Try to sync with Google Drive
        if self.drive.initialized:
            self.executor.submit(self.drive.sync_patient_files, file_no)
        
        try:
            os.startfile(folder_name)
        except:
            messagebox.showerror("Error", f"Could not open folder: {folder_name}")

    def open_existing_patient_report(self, file_no):
        """Open an existing patient's report"""
        doc_name = f"Patient_{file_no}_Report.docx"
        doc_path = os.path.join(f"Patient_{file_no}", doc_name)
        
        if not os.path.exists(doc_path) and self.drive.initialized:
            self.drive.download_file(doc_name, None, doc_path)
        
        try:
            os.startfile(doc_path)
        except:
            messagebox.showerror("Error", f"Could not open document: {doc_path}")

    def show_previous_result(self):
        """Show the previous search result"""
        if self.current_result_index > 0:
            self.current_result_index -= 1
            self.view_patient(self.current_results[self.current_result_index], multiple_results=True)
        else:
            messagebox.showinfo("Info", "This is the first result.")

    def show_next_result(self):
        """Show the next search result"""
        if self.current_result_index < len(self.current_results) - 1:
            self.current_result_index += 1
            self.view_patient(self.current_results[self.current_result_index], multiple_results=True)
        else:
            messagebox.showinfo("Info", "This is the last result.")

    def edit_patient(self, patient_data):
        """Edit an existing patient's record"""
        if self.users[self.current_user]["role"] not in ["admin", "editor"]:
            messagebox.showerror("Access Denied", "Only admins and editors can edit patients")
            return
            
        self.clear_frame()
        
        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        malignancy = patient_data.get("MALIGNANCY", "")
        title = f"Edit Patient - {malignancy}" if malignancy else "Edit Patient"
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text=title, 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with form
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Form container
        form_frame = self.create_scrollable_frame(right_frame)
        
        # Malignancy display (can't change malignancy type after creation)
        malignancy_frame = ttk.LabelFrame(form_frame, text="Malignancy Type", style='TFrame')
        malignancy_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(malignancy_frame, text=malignancy, 
                 font=('Helvetica', 12)).pack(pady=5)
        
        # Common fields
        common_frame = ttk.LabelFrame(form_frame, text="Common Information", style='TFrame')
        common_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.edit_entries = {}
        row = 0
        
        # Create common fields
        for field in COMMON_FIELDS:
            ttk.Label(common_frame, text=f"{field}:", anchor="w").grid(row=row, column=0, padx=5, pady=5, sticky="w")
            
            current_value = patient_data.get(field, "")
            
            if field == "GENDER":
                var = tk.StringVar(value=current_value)
                combobox = ttk.Combobox(common_frame, textvariable=var, 
                                      values=DROPDOWN_OPTIONS["GENDER"], state="readonly")
                combobox.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.edit_entries[field] = var
            elif field in ["EXAMINATION", "SYMPTOMS"]:
                listbox_frame = ttk.Frame(common_frame)
                listbox_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4, exportselection=0)
                scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                listbox.configure(yscrollcommand=scrollbar.set)
                
                options = DROPDOWN_OPTIONS[field]
                for option in options:
                    listbox.insert(tk.END, option)
                
                # Select previously selected items
                selected_values = patient_data.get(field, "").split(", ") if patient_data.get(field) else []
                for i, item in enumerate(options):
                    if item in selected_values:
                        listbox.select_set(i)
                
                listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                self.edit_entries[field] = listbox
                
                # Add "Others" entry if needed
                if "OTHERS" in options:
                    others_value = ""
                    for val in selected_values:
                        if val.startswith("OTHERS:"):
                            others_value = val[8:]  # Remove "OTHERS: " prefix
                            break
                    
                    ttk.Label(common_frame, text=f"{field} (Others):").grid(row=row+1, column=0, padx=5, pady=5, sticky="w")
                    others_entry = ttk.Entry(common_frame)
                    others_entry.insert(0, others_value)
                    others_entry.grid(row=row+1, column=1, padx=5, pady=5, sticky="ew")
                    self.edit_entries[f"{field}_OTHERS"] = others_entry
                    row += 1
            elif field == "DATE OF BIRTH":
                entry_frame = ttk.Frame(common_frame)
                entry_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                
                entry = ttk.Entry(entry_frame)
                entry.insert(0, current_value)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                cal_btn = ttk.Button(entry_frame, text="📅", width=3,
                                    command=lambda e=entry: self.show_calendar(e))
                cal_btn.pack(side=tk.LEFT, padx=5)
                self.edit_entries[field] = entry
            else:
                entry = ttk.Entry(common_frame)
                entry.insert(0, current_value)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.edit_entries[field] = entry
            
            row += 1
        
        # Malignancy-specific fields
        if malignancy in MALIGNANCY_FIELDS:
            specific_frame = tk.LabelFrame(form_frame, text=f"{malignancy} Specific Information", 
                                          bg=MALIGNANCY_COLORS.get(malignancy, "#FFFFFF"))
            specific_frame.pack(fill=tk.X, padx=10, pady=10)
            
            row = 0
            for field in MALIGNANCY_FIELDS[malignancy]:
                ttk.Label(specific_frame, text=f"{field}:", anchor="w").grid(row=row, column=0, padx=5, pady=5, sticky="w")
                
                current_value = patient_data.get(field, "")
                
                if field in DROPDOWN_OPTIONS:
                    var = tk.StringVar(value=current_value)
                    combobox = ttk.Combobox(specific_frame, textvariable=var, 
                                          values=DROPDOWN_OPTIONS[field], state="readonly")
                    combobox.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                    self.edit_entries[field] = var
                elif field == "SR_DATE":
                    entry_frame = ttk.Frame(specific_frame)
                    entry_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                    
                    entry = ttk.Entry(entry_frame)
                    entry.insert(0, current_value)
                    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                    
                    cal_btn = ttk.Button(entry_frame, text="📅", width=3,
                                        command=lambda e=entry: self.show_calendar(e))
                    cal_btn.pack(side=tk.LEFT, padx=5)
                    self.edit_entries[field] = entry
                elif field == "THERAPY_SIDE_EFFECTS":
                    listbox_frame = ttk.Frame(specific_frame)
                    listbox_frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                    
                    listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4, exportselection=0)
                    scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                    listbox.configure(yscrollcommand=scrollbar.set)
                    
                    options = DROPDOWN_OPTIONS[field]
                    for option in options:
                        listbox.insert(tk.END, option)
                    
                    # Select previously selected items
                    selected_values = patient_data.get(field, "").split(", ") if patient_data.get(field) else []
                    for i, item in enumerate(options):
                        if item in selected_values:
                            listbox.select_set(i)
                    
                    listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    self.edit_entries[field] = listbox
                    
                    # Add "Others" entry if needed
                    if "OTHERS" in options:
                        others_value = ""
                        for val in selected_values:
                            if val.startswith("OTHERS:"):
                                others_value = val[8:]  # Remove "OTHERS: " prefix
                                break
                        
                        ttk.Label(specific_frame, text=f"{field} (Others):").grid(row=row+1, column=0, padx=5, pady=5, sticky="w")
                        others_entry = ttk.Entry(specific_frame)
                        others_entry.insert(0, others_value)
                        others_entry.grid(row=row+1, column=1, padx=5, pady=5, sticky="ew")
                        self.edit_entries[f"{field}_OTHERS"] = others_entry
                        row += 1
                else:
                    entry = ttk.Entry(specific_frame)
                    entry.insert(0, current_value)
                    entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                    self.edit_entries[field] = entry
                
                row += 1
        
        # Button frame
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="Save Changes", 
                  command=lambda: self.save_edited_patient(patient_data),
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Cancel", 
                  command=lambda: self.view_patient(patient_data)).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Main Menu", command=self.main_menu,
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=10)
    
    def save_edited_patient(self, original_data):
        """Save edited patient data"""
        # Validate date format
        dob = self.edit_entries["DATE OF BIRTH"].get()
        try:
            datetime.strptime(dob, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Please use dd/mm/yyyy")
            return
            
        updated_data = {
            "MALIGNANCY": original_data.get("MALIGNANCY", ""),
            "CREATED_BY": original_data.get("CREATED_BY", ""),
            "CREATED_DATE": original_data.get("CREATED_DATE", ""),
            "LAST_MODIFIED_BY": self.current_user,
            "LAST_MODIFIED_DATE": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }
        
        for field, widget in self.edit_entries.items():
            if isinstance(widget, tk.Listbox):  # Multi-select fields
                selected = [widget.get(i) for i in widget.curselection()]
                if "OTHERS" in selected and f"{field}_OTHERS" in self.edit_entries:
                    others_text = self.edit_entries[f"{field}_OTHERS"].get()
                    if others_text:
                        selected[selected.index("OTHERS")] = f"OTHERS: {others_text}"
                updated_data[field] = ", ".join(selected) if selected else ""
            elif isinstance(widget, tk.StringVar):  # Combobox
                updated_data[field] = widget.get()
            else:  # Entry fields
                updated_data[field] = widget.get()
        
        # Update the record
        file_no = original_data.get("FILE NUMBER")
        for i, patient in enumerate(self.patient_data):
            if patient.get("FILE NUMBER") == file_no:
                self.patient_data[i] = updated_data
                break
        
        # Sort patient data by file number (numeric)
        self.patient_data.sort(key=lambda x: int(x.get("FILE NUMBER", 0)))
        self.save_patient_data()
        
        # Update patient report
        self.create_patient_report(file_no, create_only=True)
        
        # Upload to Google Drive in background
        self.executor.submit(self.upload_patient_to_drive, updated_data)
        
        # Show confirmation message
        messagebox.showinfo("Success", "Patient data updated successfully!")
        
        # Force the message box to display and process events
        self.root.update()
        
        # Now show the updated patient view
        self.view_patient(updated_data)
                
    def delete_patient(self, patient_data):
        """Delete a patient record"""
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only 'mej.esam' can delete patients.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this patient?")
        if not confirm:
            return
        
        file_no = patient_data.get("FILE NUMBER")
        
        # Remove from local data
        self.patient_data = [patient for patient in self.patient_data if patient.get("FILE NUMBER") != file_no]
        self.save_patient_data()
        
        # Remove from Google Drive in background
        if self.drive.initialized:
            self.executor.submit(self.delete_patient_from_drive, file_no)
        
        # Delete local folder
        try:
            folder_name = f"Patient_{file_no}"
            if os.path.exists(folder_name):
                shutil.rmtree(folder_name)
        except Exception as e:
            print(f"Error deleting patient folder: {e}")
        
        messagebox.showinfo("Success", "Patient deleted successfully!")
        self.search_patient()
    
    def delete_patient_from_drive(self, file_no):
        """Delete patient data and folder from Google Drive"""
        if not self.drive.initialized:
            return False
        
        try:
            # Delete patient data file
            file_name = f"patient_{file_no}.json"
            query = f"name='{file_name}' and '{self.drive.app_folder_id}' in parents and trashed=false"
            results = self.drive.service.files().list(q=query, fields="files(id)").execute()
            items = results.get('files', [])
            
            if items:
                self.drive.service.files().delete(fileId=items[0]['id']).execute()
            
            # Delete patient folder
            folder_name = f"Patient_{file_no}"
            query = f"name='{folder_name}' and '{self.drive.patients_folder_id}' in parents and trashed=false"
            results = self.drive.service.files().list(q=query, fields="files(id)").execute()
            items = results.get('files', [])
            
            if items:
                self.drive.service.files().delete(fileId=items[0]['id']).execute()
            
            return True
        except Exception as e:
            print(f"Error deleting patient from Google Drive: {e}")
            return False
    
    def backup_data(self):
        """Create a backup of patient data"""
        if not self.patient_data:
            messagebox.showerror("Error", "No data available to back up.")
            return

        # Local backup
        backup_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Backup As"
        )
        if not backup_path:
            return

        try:
            with open(backup_path, 'w') as f:
                json.dump(self.patient_data, f, indent=4)
            
            # Upload backup to Google Drive
            if self.drive.initialized:
                backup_name = os.path.basename(backup_path)
                self.drive.upload_file(backup_path, backup_name, self.drive.app_folder_id)
            
            messagebox.showinfo("Success", f"Data backed up to {backup_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to back up data: {e}")
    
    def restore_data(self):
        """Restore patient data from backup"""
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only 'mej.esam' can restore data.")
            return

        restore_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Select Backup File"
        )
        if not restore_path:
            return

        try:
            with open(restore_path, 'r') as f:
                data = json.load(f)
            
            self.patient_data = data
            # Sort patient data by file number (numeric)
            self.patient_data.sort(key=lambda x: int(x.get("FILE NUMBER", 0)))
            self.save_patient_data()
            
            # Upload restored data to Google Drive
            if self.drive.initialized:
                for patient in self.patient_data:
                    self.drive.upload_patient_data(patient)
            
            messagebox.showinfo("Success", "Data restored successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to restore data: {e}")
    
    def export_all_data(self):
        """Export all patient data to Excel"""
        if self.users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can export all patient data.")
            return

        if not self.patient_data:
            messagebox.showerror("Error", "No data to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save All Patient Data As"
        )
        if not file_path:
            return

        try:
            # Create a Pandas Excel writer using XlsxWriter as the engine
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                # Create a sheet for each malignancy type
                for malignancy in MALIGNANCIES:
                    malignancy_data = [patient for patient in self.patient_data if patient.get("MALIGNANCY") == malignancy]
                    if malignancy_data:
                        df = pd.DataFrame(malignancy_data)
                        df.to_excel(writer, sheet_name=malignancy, index=False)
                
                # Create a sheet with all patients
                df = pd.DataFrame(self.patient_data)
                df.to_excel(writer, sheet_name="ALL PATIENTS", index=False)
            
            # Upload to Google Drive
            if self.drive.initialized:
                self.drive.upload_file(file_path, os.path.basename(file_path), self.drive.app_folder_id)
            
            messagebox.showinfo("Success", f"All patient data exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")
    
    def view_all_patients(self):
        """View all patient records in a table"""
        if self.users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can view all patient records.")
            return

        if not self.patient_data:
            messagebox.showerror("Error", "No data available to display.")
            return

        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="All Patients", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Filter frame
        filter_frame = ttk.Frame(right_frame, style='TFrame')
        filter_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(filter_frame, text="Filter by Malignancy:", 
                 font=('Helvetica', 12)).pack(side=tk.LEFT, padx=5)
        
        self.malignancy_filter = ttk.Combobox(filter_frame, values=MALIGNANCIES, state="readonly")
        self.malignancy_filter.pack(side=tk.LEFT, padx=5)

        ttk.Button(filter_frame, text="Apply Filter", command=self.apply_malignancy_filter,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        # View all button (outside the filter frame)
        ttk.Button(filter_frame, text="View All Patients", command=self.display_all_patients,
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=10)

        # Main content
        content_frame = ttk.Frame(right_frame, style='TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Create a frame for the canvas and scrollbars
        container = ttk.Frame(content_frame, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)

        # Create a canvas for scrolling
        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a vertical scrollbar
        y_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add a horizontal scrollbar with increased height
        x_scrollbar = ttk.Scrollbar(content_frame, orient="horizontal", command=canvas.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, pady=5, ipady=10)

        canvas.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)

        # Create a frame inside the canvas to hold the table
        self.table_frame = ttk.Frame(canvas, style='TFrame')
        canvas.create_window((0, 0), window=self.table_frame, anchor="nw")

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        self.table_frame.bind("<Configure>", on_configure)

        # Load and display all data by default
        self.display_all_patients()

        # Button frame
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                  style='Blue.TButton').pack(fill=tk.X, pady=10)
    
    def display_all_patients(self, malignancy_filter=None):
        """Display all patients in a table, optionally filtered by malignancy"""
        # Clear previous data
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        if not self.patient_data:
            ttk.Label(self.table_frame, text="No patient data available", style='TLabel').pack(pady=20)
            return

        # Filter by malignancy if needed and sort by file number (numeric)
        if malignancy_filter:
            data = sorted([patient for patient in self.patient_data 
                          if patient.get("MALIGNANCY") == malignancy_filter],
                         key=lambda x: int(x.get("FILE NUMBER", 0)))
        else:
            data = sorted(self.patient_data, key=lambda x: int(x.get("FILE NUMBER", 0)))

        # Create a table-like structure
        columns = ["FILE NUMBER", "NAME", "MALIGNANCY", "GENDER", "AGE ON DIAGNOSIS", "DIAGNOSIS"]

        # Set a fixed width for all cells
        cell_width = 20

        # Create header row
        header_row = ttk.Frame(self.table_frame, style='TFrame')
        header_row.pack(fill=tk.X)

        for col in columns:
            header_cell = ttk.Label(header_row, text=col, 
                                  font=('Helvetica', 10, 'bold'), 
                                  anchor="center", borderwidth=1, relief="solid", 
                                  width=cell_width)
            header_cell.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipadx=5, ipady=5)

        # Create data rows with alternating colors
        for i, patient in enumerate(data):
            row_frame = ttk.Frame(self.table_frame, style='TFrame')
            row_frame.pack(fill=tk.X)

            bg_color = '#f0f7ff' if i % 2 == 0 else '#e6f2ff'

            for col in columns:
                cell = tk.Label(row_frame, text=patient.get(col, ""), 
                              font=('Helvetica', 10), 
                              anchor="center", borderwidth=1, relief="solid", 
                              width=cell_width, bg=bg_color)
                cell.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipadx=5, ipady=5)
                # Bind double-click to view patient
                cell.bind("<Double-1>", lambda e, p=patient: self.view_patient(p))

    def apply_malignancy_filter(self):
        """Apply malignancy filter to the patient table"""
        malignancy = self.malignancy_filter.get()
        if malignancy:
            self.display_all_patients(malignancy)
    
    def open_statistics_window(self):
        """Open the statistics window"""
        if self.users[self.current_user]["role"] not in ["admin", "editor"]:
            messagebox.showerror("Access Denied", "Only admins and editors can access statistics.")
            return

        self.clear_frame()

        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="Statistics", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Back button
        btn_frame = ttk.Frame(right_frame, padding="20 20 20 20", style='TFrame')
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                   style='Blue.TButton').pack(fill=tk.X, pady=10, ipady=10)
        
        # Header
        header_frame = ttk.Frame(right_frame, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Statistics", 
                 font=('Helvetica', 18, 'bold'),
                 foreground=self.secondary_color).pack(side=tk.LEFT)

        # Main content
        stats_frame = ttk.Frame(right_frame, style='TFrame')
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Malignancy distribution
        ttk.Label(stats_frame, text="Malignancy Distribution", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)

        if not self.patient_data:
            ttk.Label(stats_frame, text="No patient data available", style='TLabel').pack(pady=20)
            return

        malignancy_counts = defaultdict(int)
        for patient in self.patient_data:
            malignancy = patient.get("MALIGNANCY", "Unknown")
            malignancy_counts[malignancy] += 1

        if not malignancy_counts:
            ttk.Label(stats_frame, text="No data to display", style='TLabel').pack(pady=20)
            return

        # Create statistics text
        total_patients = sum(malignancy_counts.values())
        stats_text = "Malignancy Statistics:\n\n"
        
        for malignancy, count in sorted(malignancy_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_patients) * 100
            stats_text += f"{malignancy}: {count} patients ({percentage:.1f}%)\n"
        
        stats_label = ttk.Label(stats_frame, text=stats_text, 
                              font=('Helvetica', 12), justify='left')
        stats_label.pack(pady=10)

        # Create pie chart
        fig, ax = plt.subplots(figsize=(8, 6))
        labels = list(malignancy_counts.keys())
        sizes = list(malignancy_counts.values())
        colors = [MALIGNANCY_COLORS.get(m, '#999999') for m in labels]

        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        ax.set_title('Malignancy Distribution')

        # Display the plot in a Tkinter window
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=stats_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def manage_users(self):
        """Manage user accounts"""
        if self.current_user != "mej.esam" and self.users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can manage users")
            return
            
        self.clear_frame()
        
        # Main container with gradient background
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo and app name
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # App logo and name
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # App name with modern font
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        
        tk.Label(logo_frame, text="User Management", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Main content
        main_frame = ttk.Frame(right_frame, padding="20 20 20 20", style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # User list
        list_frame = ttk.Frame(main_frame, style='TFrame')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.user_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                                  selectmode=tk.SINGLE, font=('Helvetica', 12),
                                  height=10, bg='white', bd=0, highlightthickness=0)
        self.user_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.user_list.yview)
        
        for username in self.users:
            self.user_list.insert(tk.END, f"{username} ({self.users[username]['role']})")
        
        # Add user form
        form_frame = ttk.Frame(main_frame, style='TFrame')
        form_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(form_frame, text="New Username:", 
                 font=('Helvetica', 11)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        self.new_username = ttk.Entry(form_frame, font=('Helvetica', 11))
        self.new_username.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Password:", 
                 font=('Helvetica', 11)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        self.new_password = ttk.Entry(form_frame, show="*", font=('Helvetica', 11))
        self.new_password.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Role:", 
                 font=('Helvetica', 11)).grid(row=2, column=0, padx=5, pady=5, sticky="e")
        
        self.new_role = ttk.Combobox(form_frame, values=["admin", "editor", "viewer", "pharmacist"], 
                                    font=('Helvetica', 11))
        self.new_role.grid(row=2, column=1, padx=5, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Add User", command=self.add_user, 
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        # Only mej.esam can delete users
        if self.current_user == "mej.esam":
            ttk.Button(btn_frame, text="Delete User", command=self.delete_user,
                      style='Red.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Change Password", command=self.change_user_password,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=5)

    def add_user(self):
        """Add a new user"""
        username = self.new_username.get()
        password = self.new_password.get()
        role = self.new_role.get()
        
        if not all([username, password, role]):
            messagebox.showerror("Error", "All fields are required")
            return
        
        if username in self.users:
            messagebox.showerror("Error", "Username already exists")
            return
        
        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        self.users[username] = {"password": hashed_password, "role": role}
        self.save_users_to_file()
        self.user_list.insert(tk.END, f"{username} ({role})")
        self.new_username.delete(0, tk.END)
        self.new_password.delete(0, tk.END)
        self.new_role.set("")
        
        messagebox.showinfo("Success", "User added successfully")
    
    def delete_user(self):
        """Delete a user"""
        selection = self.user_list.curselection()
        if not selection:
            messagebox.showerror("Error", "No user selected")
            return

        selected = self.user_list.get(selection[0])
        username = selected.split()[0]

        if username == self.current_user:
            messagebox.showerror("Error", "You cannot delete yourself.")
            return

        if username == "mej.esam":
            messagebox.showerror("Error", "Cannot delete 'mej.esam' user.")
            return

        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the user '{username}'?")
        if not confirm:
            return

        del self.users[username]
        self.save_users_to_file()
        self.user_list.delete(selection[0])
        messagebox.showinfo("Success", f"User '{username}' deleted successfully.")

    def change_user_password(self):
        """Change a user's password"""
        selection = self.user_list.curselection()
        if not selection:
            messagebox.showerror("Error", "No user selected")
            return

        selected = self.user_list.get(selection[0])
        username = selected.split()[0]

        if username == "mej.esam" and self.current_user != "mej.esam":
            messagebox.showerror("Error", "Only 'mej.esam' can change their own password.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text=f"Change Password for {username}", 
                 font=('Helvetica', 18, 'bold'),
                 foreground=self.secondary_color).pack(side=tk.LEFT)

        # Form
        form_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="New Password:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        new_password_entry = ttk.Entry(form_frame, show="*", font=('Helvetica', 12))
        new_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="Confirm New Password:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        confirm_password_entry = ttk.Entry(form_frame, show="*", font=('Helvetica', 12))
        confirm_password_entry.pack(pady=5)

        # Buttons
        btn_frame = ttk.Frame(form_frame, style='TFrame')
        btn_frame.pack(pady=20)

        def save_new_password():
            new_password = new_password_entry.get()
            confirm_password = confirm_password_entry.get()

            if new_password != confirm_password:
                messagebox.showerror("Error", "Passwords do not match.")
                return

            hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            self.users[username]["password"] = hashed_password
            self.save_users_to_file()
            messagebox.showinfo("Success", f"Password for {username} changed successfully!")
            self.manage_users()

        ttk.Button(btn_frame, text="Save", command=save_new_password, 
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Cancel", command=self.manage_users).pack(side=tk.LEFT, padx=10)

    def change_password(self):
        """Change the current user's password"""
        if self.users[self.current_user]["role"] == "viewer":
            messagebox.showerror("Access Denied", "Viewers cannot change passwords.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Change Password", 
                 font=('Helvetica', 18, 'bold'),
                 foreground=self.secondary_color).pack(side=tk.LEFT)

        # Form
        form_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="Current Password:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        current_password_entry = ttk.Entry(form_frame, show="*", font=('Helvetica', 12))
        current_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="New Password:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        new_password_entry = ttk.Entry(form_frame, show="*", font=('Helvetica', 12))
        new_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="Confirm New Password:", 
                 font=('Helvetica', 12)).pack(pady=5)
        
        confirm_password_entry = ttk.Entry(form_frame, show="*", font=('Helvetica', 12))
        confirm_password_entry.pack(pady=5)

        # Buttons
        btn_frame = ttk.Frame(form_frame, style='TFrame')
        btn_frame.pack(pady=20)

        def save_new_password():
            current_password = current_password_entry.get()
            new_password = new_password_entry.get()
            confirm_password = confirm_password_entry.get()

            stored_hash = self.users[self.current_user]["password"].encode('utf-8')
            if not bcrypt.checkpw(current_password.encode('utf-8'), stored_hash):
                messagebox.showerror("Error", "Current password is incorrect.")
                return

            if new_password != confirm_password:
                messagebox.showerror("Error", "New passwords do not match.")
                return

            hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            self.users[self.current_user]["password"] = hashed_password
            self.save_users_to_file()
            messagebox.showinfo("Success", "Password changed successfully!")
            self.main_menu()

        ttk.Button(btn_frame, text="Save", command=save_new_password, 
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Cancel", command=self.main_menu).pack(side=tk.LEFT, padx=10)
    
    def sync_data(self):
        """Synchronize data with Firebase and Google Drive with improved feedback"""
        if not self.internet_connected:
            messagebox.showerror("Error", "No internet connection available.")
            return
        
        if self.sync_in_progress:
            messagebox.showinfo("Info", "Sync already in progress.")
            return
        
        confirm = messagebox.askyesno("Confirm Sync", 
            "This will synchronize all data with Firebase and Google Drive.\n"
            "This may take some time depending on your internet speed and amount of data.\n\n"
            "Do you want to continue?")
        if not confirm:
            return
        
        self.sync_in_progress = True
        self.sync_status.config(text="Syncing... Please wait")
        self.sync_btn.config(state=tk.DISABLED)
        
        try:
            # Show progress window
            self.sync_progress_window = tk.Toplevel(self.root)
            self.sync_progress_window.title("Synchronization Progress")
            self.sync_progress_window.geometry("400x200")
            
            tk.Label(self.sync_progress_window, text="Synchronizing data...", 
                    font=('Helvetica', 12)).pack(pady=20)
            
            self.sync_progress_label = tk.Label(self.sync_progress_window, text="Starting...")
            self.sync_progress_label.pack(pady=10)
            
            self.sync_progress_bar = ttk.Progressbar(self.sync_progress_window, length=300, mode='indeterminate')
            self.sync_progress_bar.pack(pady=10)
            self.sync_progress_bar.start()
            
            # Close button (disabled during sync)
            self.sync_close_btn = ttk.Button(self.sync_progress_window, text="Close", 
                                           state=tk.DISABLED, command=self.sync_progress_window.destroy)
            self.sync_close_btn.pack(pady=10)
            
            # Perform sync in background
            def sync_task():
                try:
                    # Update progress
                    self.root.after(0, lambda: self.update_sync_progress("Syncing with Firebase..."))
                    
                    # Sync with Firebase
                    firebase_success = False
                    firebase_message = "Firebase not initialized"
                    if self.firebase.initialized:
                        firebase_success, firebase_message = self.firebase.sync_patients(self.patient_data)
                        
                        # Get updated data from Firebase
                        if firebase_success:
                            self.root.after(0, lambda: self.update_sync_progress("Getting updated data from Firebase..."))
                            firebase_data = self.firebase.get_all_patients()
                            if firebase_data:
                                self.patient_data = firebase_data
                                # Sort patient data by file number (numeric)
                                self.patient_data.sort(key=lambda x: int(x.get("FILE NUMBER", 0)))
                                self.save_patient_data()
                    
                    # Update progress
                    self.root.after(0, lambda: self.update_sync_progress("Syncing with Google Drive..."))
                    
                    # Sync with Google Drive
                    drive_success = False
                    drive_message = "Google Drive not initialized"
                    if self.drive.initialized:
                        # Upload all patient data
                        self.root.after(0, lambda: self.update_sync_progress("Uploading patient data to Google Drive..."))
                        for patient in self.patient_data:
                            self.drive.upload_patient_data(patient)
                        
                        # Sync all patient folders
                        total_patients = len(self.patient_data)
                        for i, patient in enumerate(self.patient_data):
                            file_no = patient["FILE NUMBER"]
                            self.root.after(0, lambda: self.update_sync_progress(
                                f"Syncing files for patient {file_no} ({i+1}/{total_patients})..."))
                            
                            # Create local folder if it doesn't exist
                            local_folder = f"Patient_{file_no}"
                            if not os.path.exists(local_folder):
                                os.makedirs(local_folder, exist_ok=True)
                            
                            # Perform two-way sync
                            success, message = self.drive.sync_patient_files(file_no)
                            if not success:
                                print(f"Error syncing patient {file_no}: {message}")
                        
                        drive_success = True
                        drive_message = "Google Drive sync completed"
                    
                    # Final status
                    self.last_sync_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    final_message = (
                        f"Firebase: {'Success' if firebase_success else 'Failed'} - {firebase_message}\n"
                        f"Google Drive: {'Success' if drive_success else 'Failed'} - {drive_message}"
                    )
                    
                    self.root.after(0, lambda: self.sync_complete(
                        firebase_success and drive_success, 
                        final_message))
                    
                except Exception as e:
                    self.root.after(0, lambda: self.sync_complete(False, f"Sync failed: {str(e)}"))
            
            # Start sync in background
            self.executor.submit(sync_task)
            
        except Exception as e:
            self.sync_complete(False, f"Sync failed to start: {str(e)}")
    
    def update_sync_progress(self, message):
        """Update the sync progress window with current status"""
        if hasattr(self, 'sync_progress_label') and self.sync_progress_label.winfo_exists():
            self.sync_progress_label.config(text=message)
            self.sync_progress_window.update()
    
    def sync_complete(self, success, message):
        """Handle sync completion with detailed feedback"""
        self.sync_in_progress = False
        self.sync_status.config(text="")
        self.sync_btn.config(state=tk.NORMAL)
        
        if hasattr(self, 'sync_progress_window') and self.sync_progress_window.winfo_exists():
            self.sync_progress_bar.stop()
            self.sync_progress_label.config(text="Sync completed!")
            self.sync_close_btn.config(state=tk.NORMAL)
        
        if success:
            messagebox.showinfo("Sync Complete", message)
        else:
            messagebox.showerror("Sync Failed", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = OncologyApp(root)
    root.mainloop()
