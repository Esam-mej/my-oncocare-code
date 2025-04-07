from docx.shared import Inches
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
import configparser

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Constants
LAB_RANGES = {
    # Hematology
    "HB": {"normal_range": "11.5-15.5 g/dL", "critical_low": "<7", "critical_high": ">20"},
    "PLT": {"normal_range": "150-450 x10³/μL", "critical_low": "<20", "critical_high": ">1000"},
"WBC": {"normal_range": "4.5-13.5 x10³/μL", "critical_low": "<1", "critical_high": ">30"},
"NEUTROPHILS": {"normal_range": "1.5-8.5 x10³/μL", "critical_low": "<0.5", "critical_high": ">15"},
"LYMPHOCYTES": {"normal_range": "1.5-7.0 x10³/μL", "critical_low": "<0.5", "critical_high": ">10"},
"HCT": {"normal_range": "35-45%", "critical_low": "<20", "critical_high": ">60"},

# Chemistry
"UREA": {"normal_range": "10-40 mg/dL", "critical_low": "<10", "critical_high": ">100"},
"CREATININE": {"normal_range": "0.3-1.0 mg/dL", "critical_low": "<0.3", "critical_high": ">2.0"},
"NA": {"normal_range": "135-145 mEq/L", "critical_low": "<120", "critical_high": ">160"},
"K": {"normal_range": "3.5-5.0 mEq/L", "critical_low": "<2.5", "critical_high": ">6.5"},
"CL": {"normal_range": "98-107 mEq/L", "critical_low": "<80", "critical_high": ">115"},
"CA": {"normal_range": "8.8-10.8 mg/dL", "critical_low": "<7", "critical_high": ">13"},
"MG": {"normal_range": "1.7-2.4 mg/dL", "critical_low": "<1", "critical_high": ">4"},
"PHOSPHORUS": {"normal_range": "3.0-6.0 mg/dL", "critical_low": "<1.5", "critical_high": ">8"},
"URIC ACID": {"normal_range": "2.5-7.0 mg/dL", "critical_low": "<1", "critical_high": ">10"},
"LDH": {"normal_range": "100-300 U/L", "critical_low": "<50", "critical_high": ">1000"},
"FERRITIN": {"normal_range": "10-300 ng/mL", "critical_low": "<10", "critical_high": ">1000"},

# Liver
"GPT (ALT)": {"normal_range": "5-45 U/L", "critical_low": "<5", "critical_high": ">200"},
"GOT (AST)": {"normal_range": "10-40 U/L", "critical_low": "<10", "critical_high": ">200"},
"ALK PHOS": {"normal_range": "50-400 U/L", "critical_low": "<50", "critical_high": ">1000"},
"T.BILIRUBIN": {"normal_range": "0.2-1.2 mg/dL", "critical_low": "<0.2", "critical_high": ">5"},
"D.BILIRUBIN": {"normal_range": "0-0.4 mg/dL", "critical_low": "<0", "critical_high": ">2"},

# Coagulation
"PT": {"normal_range": "11-14 sec", "critical_low": "<10", "critical_high": ">30"},
"APTT": {"normal_range": "25-35 sec", "critical_low": "<20", "critical_high": ">60"},
"INR": {"normal_range": "0.9-1.2", "critical_low": "<0.8", "critical_high": ">5"},
"FIBRINOGEN": {"normal_range": "200-400 mg/dL", "critical_low": "<100", "critical_high": ">700"},
"D-DIMER": {"normal_range": "<0.5 μg/mL", "critical_low": "<0.1", "critical_high": ">5"},

# ABG
"PH": {"normal_range": "7.35-7.45", "critical_low": "<7.2", "critical_high": ">7.6"},
"PCO2": {"normal_range": "35-45 mmHg", "critical_low": "<25", "critical_high": ">60"},
"PO2": {"normal_range": "80-100 mmHg", "critical_low": "<60", "critical_high": ">120"},
"HCO3": {"normal_range": "22-26 mEq/L", "critical_low": "<15", "critical_high": ">35"},
"BASE EXCESS": {"normal_range": "-2 to +2", "critical_low": "<-5", "critical_high": ">5"},

# Viral Screen
"HBSAG": {"normal_range": "Negative", "critical_low": "", "critical_high": "Positive"},
"HCV": {"normal_range": "Negative", "critical_low": "", "critical_high": "Positive"},
"HIV": {"normal_range": "Negative", "critical_low": "", "critical_high": "Positive"},
"CMV": {"normal_range": "Negative", "critical_low": "", "critical_high": "Positive"},
"EBV": {"normal_range": "Negative", "critical_low": "", "critical_high": "Positive"}
}

EF_NOTES_TEMPLATES = [
    "EF < 50%",
    ">10% drop from baseline",
    "Clinical symptoms",
    "Consider cardiology consult"
]

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
        self.setup_keyboard_shortcuts()
        self.setup_login_screen()

        # Set window icon
        try:
            self.root.iconbitmap('icon.ico')  # Replace with your icon file
        except:
            pass
        
        # Initialize F&N Documentation configuration
        self.setup_fn_documentation_config()

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
    
    def setup_keyboard_shortcuts(self):
        """Configure keyboard shortcuts for quick navigation."""
        # General shortcuts
        self.root.bind('<Control-q>', lambda e: self.on_close())  # Ctrl+Q to quit
        self.root.bind('<Control-l>', lambda e: self.setup_login_screen())  # Ctrl+L to logout
        self.root.bind('<F1>', lambda e: self.show_help())  # F1 for help
        
        # Patient management
        self.root.bind('<Control-n>', lambda e: self.add_patient())  # Ctrl+N for new patient
        self.root.bind('<Control-f>', lambda e: self.search_patient())  # Ctrl+F to search
        
        # Navigation
        self.root.bind('<Escape>', lambda e: self.main_menu())  # ESC to return to main menu

    def show_help(self):
        """Display keyboard shortcuts help dialog."""
        shortcuts = {
            "Ctrl+Q": "Quit the application",
            "Ctrl+L": "Log out",
            "F1": "Show this help dialog",
            "Ctrl+N": "Add new patient",
            "Ctrl+F": "Search patients",
            "Esc": "Return to main menu"
        }
        
        help_text = "Keyboard Shortcuts:\n\n" + "\n".join(
            f"{key}: {desc}" for key, desc in shortcuts.items()
        )
        
        messagebox.showinfo("Keyboard Shortcuts", help_text)

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
        
        self.style.configure('Purple.TButton', background='#9b59b6', foreground='white')
        self.style.map('Purple.TButton',
            background=[('active', '#8e44ad')],
            foreground=[('active', 'white')]
        )

        self.style.configure('Orange.TButton', background='#e67e22', foreground='white')
        self.style.map('Orange.TButton',
            background=[('active', '#d35400')],
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
        """Display the main menu with rearranged buttons"""
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

        # Button grid container
        btn_frame = ttk.Frame(form_container, style='TFrame')
        btn_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        # Row 1 - Patient Management (Blue buttons)
        new_row1_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row1_frame.pack(fill=tk.X, pady=5)

        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(new_row1_frame, text="Add New Patient", command=self.add_patient,
                       style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            ttk.Frame(new_row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Placeholder

        ttk.Button(new_row1_frame, text="Search Patient", command=self.search_patient,
                   style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        if self.users[self.current_user]["role"] == "admin":
            ttk.Button(new_row1_frame, text="View All Patients", command=self.view_all_patients,
                       style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            ttk.Frame(new_row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Placeholder

        if self.users[self.current_user]["role"] in ["admin", "editor", "pharmacist"]:
            ttk.Button(new_row1_frame, text="Lab & EF Documentation", 
                      command=self.show_lab_ef_window,
                      style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            ttk.Frame(new_row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Placeholder

        # Row 2 - Data Operations (Green buttons)
        new_row2_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row2_frame.pack(fill=tk.X, pady=5)

        if self.users[self.current_user]["role"] == "admin":
            ttk.Button(new_row2_frame, text="Export All Data", command=self.export_all_data,
                       style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

            ttk.Button(new_row2_frame, text="Backup Data", command=self.backup_data,
                       style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

            # Specific user check for Restore Data remains
            if self.current_user == "mej.esam":
                 ttk.Button(new_row2_frame, text="Restore Data", command=self.restore_data,
                           style='Green.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            else:
                 # Placeholder if admin but not mej.esam
                 ttk.Frame(new_row2_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
             # Placeholders if not admin (need 3 placeholders if admin has 3 buttons)
             ttk.Frame(new_row2_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
             ttk.Frame(new_row2_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
             ttk.Frame(new_row2_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)


        # Row 3 - Chemo Tools (Yellow buttons)
        new_row3_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row3_frame.pack(fill=tk.X, pady=5)

        ttk.Button(new_row3_frame, text="CHEMO PROTOCOLS", command=self.show_chemo_protocols,
                   style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        ttk.Button(new_row3_frame, text="CHEMO SHEETS", command=self.show_chemo_sheets,
                   style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        if self.users[self.current_user]["role"] in ["admin", "pharmacist"]:
            ttk.Button(new_row3_frame, text="CHEMO STOCKS", command=self.show_chemo_stocks,
                       style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            ttk.Frame(new_row3_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Placeholder


        # Row 4 - Statistics (Yellow button)
        new_row4_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row4_frame.pack(fill=tk.X, pady=5)

        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            # Pack the single button to expand and fill
            ttk.Button(new_row4_frame, text="Statistics", command=self.show_statistics,
                       style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # If the only button isn't shown, add a placeholder to maintain row height/spacing
             ttk.Frame(new_row4_frame, height=1).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5) # Minimal placeholder


        # Row 5 - Calculations (Purple button)
        new_row5_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row5_frame.pack(fill=tk.X, pady=5)

        ttk.Button(new_row5_frame, text="Medical Calculators", command=self.show_calculators,
                   style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)


        # Row 6 - Extravasation (Orange button)
        new_row6_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row6_frame.pack(fill=tk.X, pady=5)

        ttk.Button(new_row6_frame, text="Extravasation Management", command=self.show_extravasation_management,
                   style='Orange.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)


        # Row 7 - F&N Documentation (Purple button)
        new_row7_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row7_frame.pack(fill=tk.X, pady=5)

        ttk.Button(new_row7_frame, text="F&N Documentation", command=self.handle_fn_documentation,
                   style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)


        # Row 8 - User Management (Brown buttons)
        new_row8_frame = ttk.Frame(btn_frame, style='TFrame')
        new_row8_frame.pack(fill=tk.X, pady=5)

        # Check if user is admin OR the specific user 'mej.esam' for Manage Users/Change Password
        can_manage_users = self.users[self.current_user]["role"] == "admin" or self.current_user == "mej.esam"

        if can_manage_users:
            ttk.Button(new_row8_frame, text="Manage Users", command=self.manage_users,
                       style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            ttk.Button(new_row8_frame, text="Change Password", command=self.change_password,
                       style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Add two placeholders if user cannot manage users to maintain layout consistency
            ttk.Frame(new_row8_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            ttk.Frame(new_row8_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)


        # Row 9 - Logout (Red button)
        new_row9_frame = ttk.Frame(btn_frame, style='TFrame')
        # Add extra padding before the logout button as in the original idea
        new_row9_frame.pack(fill=tk.X, pady=(20, 5))

        ttk.Button(new_row9_frame, text="Logout", command=self.setup_login_screen,
                   style='Red.TButton').pack(fill=tk.X, ipady=10) # Use original packing for logout

        # Signature
        signature_frame = ttk.Frame(form_container, style='TFrame')
        signature_frame.pack(pady=(30, 0)) # Keep padding relative to the last element (Logout row)

        # Make sure signature label background matches the frame it's in
        ttk.Label(signature_frame, text="Made by: DR. ESAM MEJRAB",
                  font=('Times New Roman', 14, 'italic'),
                  foreground=self.primary_color, background='white').pack() # Specify background
            
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
    
    def show_lab_ef_window(self, patient_data=None):
        """Show the lab and EF documentation window with historical data viewing"""
        self.lab_ef_window = tk.Toplevel(self.root)
        self.lab_ef_window.title("Lab Investigations & EF Documentation")
        self.lab_ef_window.geometry("1100x850")
        
        # Store current patient data if provided
        self.current_lab_patient = patient_data
        self.current_lab_data_index = -1  # -1 means new entry, otherwise index of viewed data
        
        # Patient info frame with file number display
        patient_frame = ttk.Frame(self.lab_ef_window, style='TFrame')
        patient_frame.pack(fill=tk.X, padx=10, pady=10)
        
        if patient_data:
            patient_info = f"Patient: {patient_data.get('NAME', '')} (File#: {patient_data.get('FILE NUMBER', '')})"
            self.current_file_number = patient_data.get('FILE NUMBER', '')
        else:
            patient_info = "New Lab/EF Entry (Select Patient)"
            self.current_file_number = ""
            
        self.patient_info_label = ttk.Label(patient_frame, text=patient_info, font=('Helvetica', 12, 'bold'))
        self.patient_info_label.pack(side=tk.LEFT)
        
        # Add patient selection button if no patient provided
        if not patient_data:
                button = ttk.Button(patient_frame, text="Select Patient", 
                                    command=self.select_patient_for_labs)
                
                # Set the button size and color
                button.config(width=30, padding=15, style='Red.TButton')  # Adjust width and padding as needed
                
                # Create a style for the button
                style = ttk.Style()
                style.configure('Red.TButton', background='red', foreground='white')  # Set button color
                
                # Pack the button in the center, slightly to the left
                button.pack(side=tk.TOP, padx=5, pady=10, expand=True)
        
        # Historical data controls
        hist_frame = ttk.Frame(self.lab_ef_window)
        hist_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.hist_data_label = ttk.Label(hist_frame, text="Viewing: New Entry", font=('Helvetica', 10))
        self.hist_data_label.pack(side=tk.LEFT)
        
        self.hist_prev_btn = ttk.Button(hist_frame, text="◀ Previous", state='disabled',
                                       command=lambda: self.navigate_lab_ef_history(-1))
        self.hist_prev_btn.pack(side=tk.LEFT, padx=5)
        
        self.hist_next_btn = ttk.Button(hist_frame, text="Next ▶", state='disabled',
                                       command=lambda: self.navigate_lab_ef_history(1))
        self.hist_next_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(hist_frame, text="New Entry", command=self.create_new_lab_ef_entry).pack(side=tk.RIGHT)
        
        # Notebook for tabs
        notebook = ttk.Notebook(self.lab_ef_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Lab Investigations Tab
        self.lab_frame = ttk.Frame(notebook)
        notebook.add(self.lab_frame, text="Lab Investigations")
        self.setup_lab_tab(self.lab_frame)
        
        # EF Documentation Tab
        self.ef_frame = ttk.Frame(notebook)
        notebook.add(self.ef_frame, text="Ejection Fraction")
        self.setup_ef_tab(self.ef_frame)
        
        # Button frame
        btn_frame = ttk.Frame(self.lab_ef_window, style='TFrame')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Save", command=self.save_lab_ef_data,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Print", command=self.print_lab_ef_report,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Close", command=self.lab_ef_window.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Load historical data if patient has previous entries
        if patient_data and ("lab_results" in patient_data or "ef_data" in patient_data):
            self.load_lab_ef_history_controls()

    def load_lab_ef_history_controls(self):
        """Enable and setup historical data navigation controls"""
        if not hasattr(self, 'current_lab_patient') or not self.current_lab_patient:
            return
        
        lab_count = len(self.current_lab_patient.get("lab_results", []))
        ef_count = len(self.current_lab_patient.get("ef_data", []))
        total_entries = max(lab_count, ef_count)
        
        if total_entries > 0:
            self.hist_prev_btn.config(state='normal')
            self.hist_next_btn.config(state='normal')
            self.current_lab_data_index = total_entries - 1  # Start with most recent
            self.display_lab_ef_data(self.current_lab_data_index)
        else:
            self.hist_prev_btn.config(state='disabled')
            self.hist_next_btn.config(state='disabled')
            self.current_lab_data_index = -1
            self.hist_data_label.config(text="Viewing: New Entry")

    def navigate_lab_ef_history(self, direction):
        """Navigate through historical lab/EF data"""
        if not hasattr(self, 'current_lab_patient') or not self.current_lab_patient:
            return
        
        lab_count = len(self.current_lab_patient.get("lab_results", []))
        ef_count = len(self.current_lab_patient.get("ef_data", []))
        total_entries = max(lab_count, ef_count)
        
        new_index = self.current_lab_data_index + direction
        
        if 0 <= new_index < total_entries:
            self.current_lab_data_index = new_index
            self.display_lab_ef_data(self.current_lab_data_index)
        elif new_index == -1:  # Show new entry
            self.create_new_lab_ef_entry()
        else:
            return  # Don't go beyond available entries
        
        # Update button states
        self.hist_prev_btn.config(state='normal' if new_index > 0 else 'disabled')
        self.hist_next_btn.config(state='normal' if new_index < total_entries - 1 else 'disabled')

    def display_lab_ef_data(self, index):
        """Display historical lab and EF data at the given index"""
        if not hasattr(self, 'current_lab_patient') or not self.current_lab_patient:
            return
        
        lab_data = self.current_lab_patient.get("lab_results", [])
        ef_data = self.current_lab_patient.get("ef_data", [])
        
        # Clear current entries
        self.clear_lab_ef_entries()
        
        # Load lab data if available
        if index < len(lab_data):
            data = lab_data[index]
            self.lab_date_entry.insert(0, data.get("date", ""))
            
            values = data.get("values", {})
            for test, entry in self.lab_entries.items():
                if test in values:
                    entry.insert(0, values[test])
                    # Trigger critical value check
                    entry.event_generate("<FocusOut>")
        
        # Load EF data if available
        if index < len(ef_data):
            data = ef_data[index]
            baseline = data.get("baseline", {})
            self.ef_baseline_date.insert(0, baseline.get("date", ""))
            self.ef_baseline_value.insert(0, baseline.get("value", ""))
            
            serial = data.get("serial", [])
            for i, entry in enumerate(self.ef_serial_entries):
                if i < len(serial):
                    entry["date"].insert(0, serial[i].get("date", ""))
                    entry["value"].insert(0, serial[i].get("value", ""))
                    entry["change_label"].config(text=serial[i].get("change", ""))
            
            self.ef_notes.delete("1.0", tk.END)
            self.ef_notes.insert("1.0", data.get("notes", ""))
        
        # Update status label
        date_str = ""
        if index < len(lab_data) and "date" in lab_data[index]:
            date_str = lab_data[index]["date"]
        elif index < len(ef_data) and "baseline" in ef_data[index] and "date" in ef_data[index]["baseline"]:
            date_str = ef_data[index]["baseline"]["date"]
        
        self.hist_data_label.config(text=f"Viewing: Entry from {date_str if date_str else 'Unknown date'}")
        self.current_lab_data_index = index

    def clear_lab_ef_entries(self):
        """Clear all lab and EF entries"""
        self.lab_date_entry.delete(0, tk.END)
        for entry in self.lab_entries.values():
            entry.delete(0, tk.END)
        
        self.ef_baseline_date.delete(0, tk.END)
        self.ef_baseline_value.delete(0, tk.END)
        for entry in self.ef_serial_entries:
            entry["date"].delete(0, tk.END)
            entry["value"].delete(0, tk.END)
            entry["change_label"].config(text="")
        self.ef_notes.delete("1.0", tk.END)

    def create_new_lab_ef_entry(self):
        """Prepare the window for a new lab/EF entry"""
        self.clear_lab_ef_entries()
        self.current_lab_data_index = -1
        self.hist_data_label.config(text="Viewing: New Entry")
        
        # Set today's date as default
        today = datetime.now().strftime("%d/%m/%Y")
        self.lab_date_entry.insert(0, today)
        self.ef_baseline_date.insert(0, today)
        
        # Update button states
        lab_count = len(self.current_lab_patient.get("lab_results", [])) if hasattr(self, 'current_lab_patient') else 0
        ef_count = len(self.current_lab_patient.get("ef_data", [])) if hasattr(self, 'current_lab_patient') else 0
        total_entries = max(lab_count, ef_count)
        
        self.hist_prev_btn.config(state='normal' if total_entries > 0 else 'disabled')
        self.hist_next_btn.config(state='disabled')

    def select_patient_for_labs(self):
        """Open patient selection dialog for lab/EF entries"""
        self.patient_selection_window = tk.Toplevel(self.lab_ef_window)
        self.patient_selection_window.title("Select Patient")
        self.patient_selection_window.geometry("600x400")
        
        # Search frame
        search_frame = ttk.Frame(self.patient_selection_window)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        search_entry = ttk.Entry(search_frame, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.focus()
        
        ttk.Button(search_frame, text="Search", 
                  command=lambda: self.search_patients_for_labs(search_entry.get())).pack(side=tk.LEFT)
        
        # Results treeview
        columns = ("file_number", "name", "age", "gender")
        self.lab_patient_tree = ttk.Treeview(self.patient_selection_window, columns=columns, show="headings")
        
        self.lab_patient_tree.heading("file_number", text="File #")
        self.lab_patient_tree.heading("name", text="Name")
        self.lab_patient_tree.heading("age", text="Age")
        self.lab_patient_tree.heading("gender", text="Gender")
        
        self.lab_patient_tree.column("file_number", width=80)
        self.lab_patient_tree.column("name", width=200)
        self.lab_patient_tree.column("age", width=60)
        self.lab_patient_tree.column("gender", width=80)
        
        scrollbar = ttk.Scrollbar(self.patient_selection_window, orient="vertical", command=self.lab_patient_tree.yview)
        self.lab_patient_tree.configure(yscrollcommand=scrollbar.set)
        
        self.lab_patient_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate with all patients initially
        self.search_patients_for_labs("")
        
        # Select button
        btn_frame = ttk.Frame(self.patient_selection_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Select", 
                  command=self.assign_selected_patient_for_labs).pack(side=tk.RIGHT)

    def search_patients_for_labs(self, search_term):
        """Search patients for lab assignment"""
        self.lab_patient_tree.delete(*self.lab_patient_tree.get_children())
        
        for patient in self.patient_data:
            if (search_term.lower() in patient.get("NAME", "").lower() or 
                search_term.lower() in patient.get("FILE NUMBER", "").lower()):
                self.lab_patient_tree.insert("", "end", values=(
                    patient.get("FILE NUMBER", ""),
                    patient.get("NAME", ""),
                    patient.get("AGE", ""),
                    patient.get("GENDER", "")
                ))

    def assign_selected_patient_for_labs(self):
        """Assign selected patient to the lab/EF window"""
        selected_item = self.lab_patient_tree.focus()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a patient first")
            return
            
        patient_values = self.lab_patient_tree.item(selected_item, "values")
        file_number = patient_values[0]
        
        # Find the complete patient data
        for patient in self.patient_data:
            if patient.get("FILE NUMBER") == file_number:
                self.current_lab_patient = patient
                self.current_file_number = file_number
                
                # Update the patient info display
                self.patient_info_label.config(
                    text=f"Patient: {patient.get('NAME', '')} (File#: {file_number})"
                )
                
                # Load historical data controls
                self.load_lab_ef_history_controls()
                
                self.patient_selection_window.destroy()
                return
        
        messagebox.showerror("Error", "Patient data not found")

    def setup_lab_tab(self, parent):
        """Setup the lab investigations tab"""
        # Create scrollable frame
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", on_frame_configure)
        
        # Lab entries dictionary
        self.lab_entries = {}
        self.lab_crit_labels = {}  # To store critical value labels
        
        # Date entry
        date_frame = ttk.Frame(scrollable_frame)
        date_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(date_frame, text="Date:").pack(side=tk.LEFT, padx=5)
        self.lab_date_entry = ttk.Entry(date_frame, width=15)
        self.lab_date_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(date_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(self.lab_date_entry)).pack(side=tk.LEFT, padx=5)
        
        # Add lab tests
        for test, ranges in LAB_RANGES.items():
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Test name and normal range
            info_label = f"{test} ({ranges['normal_range']})"
            ttk.Label(frame, text=info_label, width=30, anchor="w").pack(side=tk.LEFT, padx=5)
            
            # Value entry
            entry = ttk.Entry(frame, width=10)
            entry.pack(side=tk.LEFT, padx=5)
            self.lab_entries[test] = entry
            
            # Critical value indicator
            crit_label = ttk.Label(frame, text="", width=20)
            crit_label.pack(side=tk.LEFT, padx=5)
            self.lab_crit_labels[test] = crit_label
            
            # Bind validation
            entry.bind("<FocusOut>", lambda e, t=test: self.check_lab_critical_value(t))

    def setup_ef_tab(self, parent):
        """Setup the ejection fraction documentation tab"""
        # Create scrollable frame
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", on_frame_configure)
        
        # EF documentation fields
        self.ef_entries = {}
        
        # Baseline EF
        baseline_frame = ttk.LabelFrame(scrollable_frame, text="Baseline EF (Before Anthracyclines)")
        baseline_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(baseline_frame, text="Date:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.ef_baseline_date = ttk.Entry(baseline_frame, width=15)
        self.ef_baseline_date.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(baseline_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(self.ef_baseline_date)).grid(row=0, column=2, padx=5)
        
        ttk.Label(baseline_frame, text="EF (%):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.ef_baseline_value = ttk.Entry(baseline_frame, width=10)
        self.ef_baseline_value.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Serial EF measurements
        serial_frame = ttk.LabelFrame(scrollable_frame, text="Serial EF Measurements")
        serial_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Table headers
        ttk.Label(serial_frame, text="Date", font=('Helvetica', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(serial_frame, text="EF (%)", font=('Helvetica', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(serial_frame, text="Change from Baseline", font=('Helvetica', 10, 'bold')).grid(row=0, column=2, padx=5, pady=5)
        
        # Add 5 measurement rows
        self.ef_serial_entries = []
        for i in range(1, 6):
            date_entry = ttk.Entry(serial_frame, width=15)
            date_entry.grid(row=i, column=0, padx=5, pady=5)
            
            ef_entry = ttk.Entry(serial_frame, width=10)
            ef_entry.grid(row=i, column=1, padx=5, pady=5)
            
            change_label = ttk.Label(serial_frame, text="", width=20)
            change_label.grid(row=i, column=2, padx=5, pady=5)
            
            self.ef_serial_entries.append({
                "date": date_entry,
                "value": ef_entry,
                "change_label": change_label
            })
            
            # Add calendar buttons
            ttk.Button(serial_frame, text="📅", width=3,
                      command=lambda e=date_entry: self.show_calendar(e)).grid(row=i, column=3, padx=5)
        
        # Add button to calculate changes
        ttk.Button(serial_frame, text="Calculate Changes", 
                  command=self.calculate_ef_changes).grid(row=6, column=0, columnspan=4, pady=10)
        
        # EF monitoring notes
        notes_frame = ttk.LabelFrame(scrollable_frame, text="Monitoring Notes")
        notes_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.ef_notes = tk.Text(notes_frame, height=5, wrap=tk.WORD)
        self.ef_notes.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add predefined alerts
        alert_frame = ttk.Frame(notes_frame)
        alert_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(alert_frame, text="Quick Alerts:").pack(side=tk.LEFT, padx=5)
        
        for alert in EF_NOTES_TEMPLATES:
            ttk.Button(alert_frame, text=alert, width=20,
                      command=lambda a=alert: self.add_ef_alert(a)).pack(side=tk.LEFT, padx=2)

    def check_lab_critical_value(self, test):
        """Check if lab value is critical and update label"""
        try:
            value = float(self.lab_entries[test].get())
            ranges = LAB_RANGES[test]
            label = self.lab_crit_labels[test]
            
            # Parse normal range (simplified parsing)
            normal_range = ranges["normal_range"]
            if "-" in normal_range:
                low, high = map(float, normal_range.split("-")[:2])
            elif "<" in normal_range:
                low, high = 0, float(normal_range[1:])
            else:
                low, high = 0, float('inf')
            
            # Check critical values
            if ranges["critical_low"] and value < float(ranges["critical_low"][1:]):
                label.config(text="CRITICALLY LOW!", foreground="red")
            elif ranges["critical_high"] and value > float(ranges["critical_high"][1:]):
                label.config(text="CRITICALLY HIGH!", foreground="red")
            elif value < low or value > high:
                label.config(text="Abnormal", foreground="orange")
            else:
                label.config(text="Normal", foreground="green")
                
        except ValueError:
            self.lab_crit_labels[test].config(text="", foreground="black")

    def calculate_ef_changes(self):
        """Calculate EF changes from baseline"""
        try:
            baseline = float(self.ef_baseline_value.get())
            
            for entry in self.ef_serial_entries:
                if entry["value"].get():
                    current = float(entry["value"].get())
                    change = current - baseline
                    percent_change = (change / baseline) * 100
                    
                    alert = ""
                    if current < 50:
                        alert = "EF < 50%!"
                    elif percent_change < -10:
                        alert = ">10% drop!"
                    
                    entry["change_label"].config(
                        text=f"{change:.1f} ({percent_change:.1f}%) {alert}",
                        foreground="red" if alert else "black"
                    )
                    
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")

    def add_ef_alert(self, alert_text):
        """Add a predefined alert to the notes"""
        self.ef_notes.insert(tk.END, alert_text + "\n")

    def save_lab_ef_data(self):
        """Save lab and EF data to patient record"""
        if not hasattr(self, 'current_file_number') or not self.current_file_number:
            messagebox.showwarning("Warning", "Please select a patient first")
            return
        
        # Find patient in data
        patient_found = False
        for patient in self.patient_data:
            if patient.get("FILE NUMBER") == self.current_file_number:
                patient_found = True
                
                # Prepare lab data
                lab_data = {
                    "date": self.lab_date_entry.get(),
                    "values": {test: entry.get() for test, entry in self.lab_entries.items()}
                }
                
                # Prepare EF data
                ef_data = {
                    "baseline": {
                        "date": self.ef_baseline_date.get(),
                        "value": self.ef_baseline_value.get()
                    },
                    "serial": [
                        {
                            "date": entry["date"].get(),
                            "value": entry["value"].get(),
                            "change": entry["change_label"].cget("text")
                        }
                        for entry in self.ef_serial_entries
                    ],
                    "notes": self.ef_notes.get("1.0", tk.END)
                }
                
                # Add to patient record
                if "lab_results" not in patient:
                    patient["lab_results"] = []
                if "ef_data" not in patient:
                    patient["ef_data"] = []
                
                if self.current_lab_data_index == -1:  # New entry
                    patient["lab_results"].append(lab_data)
                    patient["ef_data"].append(ef_data)
                else:  # Update existing entry
                    if self.current_lab_data_index < len(patient["lab_results"]):
                        patient["lab_results"][self.current_lab_data_index] = lab_data
                    else:
                        patient["lab_results"].append(lab_data)
                    
                    if self.current_lab_data_index < len(patient["ef_data"]):
                        patient["ef_data"][self.current_lab_data_index] = ef_data
                    else:
                        patient["ef_data"].append(ef_data)
                
                # Update last modified info
                patient["LAST_MODIFIED_BY"] = self.current_user
                patient["LAST_MODIFIED_DATE"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
                # Update the history controls
                self.load_lab_ef_history_controls()
                break
        
        if not patient_found:
            messagebox.showerror("Error", f"Patient with file number {self.current_file_number} not found")
            return
        
        self.save_patient_data()
        messagebox.showinfo("Success", f"Lab and EF data saved for patient {self.current_file_number}")

    def print_lab_ef_report(self):
        """Print the lab and EF report"""
        # Create a Word document
        doc = Document()
        
        # Add title
        doc.add_heading('Lab Investigations & EF Report', level=1)
        
        # Add patient info if available
        if hasattr(self, 'current_file_number') and self.current_file_number:
            doc.add_paragraph(f"Patient File Number: {self.current_file_number}")
            if hasattr(self, 'current_lab_patient') and self.current_lab_patient:
                doc.add_paragraph(f"Patient Name: {self.current_lab_patient.get('NAME', 'N/A')}")
        
        # Add date info
        if self.current_lab_data_index == -1:
            doc.add_paragraph("Report Date: " + datetime.now().strftime("%d/%m/%Y"))
        else:
            doc.add_paragraph(f"Report for entry dated: {self.lab_date_entry.get()}")
        
        # Add lab data section
        doc.add_heading('Lab Investigations', level=2)
        doc.add_paragraph(f"Date: {self.lab_date_entry.get()}")
        
        # Create table for lab results
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        
        # Header row
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Test'
        hdr_cells[1].text = 'Value'
        hdr_cells[2].text = 'Status'
        
        # Add lab results
        for test, entry in self.lab_entries.items():
            if entry.get():
                row_cells = table.add_row().cells
                row_cells[0].text = test
                row_cells[1].text = entry.get()
                
                # Get status from the label
                status = self.lab_crit_labels[test].cget("text")
                row_cells[2].text = status if status else "N/A"
        
        # Add EF data section
        doc.add_heading('Ejection Fraction', level=2)
        
        # Baseline EF
        doc.add_heading('Baseline EF', level=3)
        doc.add_paragraph(f"Date: {self.ef_baseline_date.get()}")
        doc.add_paragraph(f"Value: {self.ef_baseline_value.get()}%")
        
        # Serial EF measurements
        doc.add_heading('Serial Measurements', level=3)
        serial_table = doc.add_table(rows=1, cols=3)
        serial_table.style = 'Table Grid'
        
        # Header row
        hdr_cells = serial_table.rows[0].cells
        hdr_cells[0].text = 'Date'
        hdr_cells[1].text = 'EF (%)'
        hdr_cells[2].text = 'Change from Baseline'
        
        # Add measurements
        for entry in self.ef_serial_entries:
            if entry["date"].get() or entry["value"].get():
                row_cells = serial_table.add_row().cells
                row_cells[0].text = entry["date"].get()
                row_cells[1].text = entry["value"].get()
                row_cells[2].text = entry["change_label"].cget("text")
        
        # Add notes
        doc.add_heading('Notes', level=3)
        doc.add_paragraph(self.ef_notes.get("1.0", tk.END))
        
        # Save to temporary file and print
        temp_file = "temp_lab_ef_report.docx"
        doc.save(temp_file)
        
        try:
            os.startfile(temp_file, "print")
            messagebox.showinfo("Print", "Report sent to printer")
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print document: {str(e)}")
        finally:
            # Clean up after printing
            def delete_temp_file():
                try:
                    os.remove(temp_file)
                except:
                    pass
            
            # Schedule file deletion after a delay
            self.root.after(5000, delete_temp_file)
                        
    def show_statistics(self):
        """Show statistics window"""
        self.open_statistics_window()

    def setup_fn_documentation_config(self):
        """Setup configuration for F&N Documentation software"""
        self.config = configparser.ConfigParser()
        self.config_file = 'onco_config.ini'
        
        # Create config file if it doesn't exist
        if not os.path.exists(self.config_file):
            self.config['FN_DOCUMENTATION'] = {'path': 'C:\\Path\\To\\Your\\Software.exe'}
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)
        else:
            self.config.read(self.config_file)
        
        # Set default path if not configured
        if 'FN_DOCUMENTATION' not in self.config:
            self.config['FN_DOCUMENTATION'] = {'path': 'C:\\Path\\To\\Your\\Software.exe'}
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)

    def get_fn_documentation_path(self):
        """Get the path to F&N Documentation software"""
        self.config.read(self.config_file)
        return self.config['FN_DOCUMENTATION']['path']

    def set_fn_documentation_path(self, new_path):
        """Set new path for F&N Documentation software"""
        self.config['FN_DOCUMENTATION']['path'] = new_path
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def handle_fn_documentation(self):
        """Handle F&N Documentation button click"""
        # Get the current path
        software_path = self.get_fn_documentation_path()
        
        # Create a custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("F&N Documentation")
        dialog.geometry("400x200")
        
        # Main message
        msg = f"Do you want to start F&N Documentation?\n\nCurrent path: {software_path}"
        ttk.Label(dialog, text=msg, wraplength=380).pack(pady=20)
        
        # Button frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        # Launch button
        ttk.Button(btn_frame, text="Launch", command=lambda: self.launch_fn_documentation(dialog),
                  style='Green.TButton').pack(side=tk.LEFT, padx=10)
        
        # Change path button (only for mej.esam)
        if self.current_user == "mej.esam":
            ttk.Button(btn_frame, text="Change Path", 
                      command=lambda: self.change_fn_documentation_path(dialog),
                      style='Blue.TButton').pack(side=tk.LEFT, padx=10)
        
        # Cancel button
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=10)

    def launch_fn_documentation(self, dialog=None):
        """Launch the F&N Documentation software"""
        try:
            software_path = self.get_fn_documentation_path()
            if not os.path.exists(software_path):
                messagebox.showerror("Error", f"Software not found at:\n{software_path}")
                return
            
            # Try to run the software
            subprocess.Popen(software_path)
            if dialog:
                dialog.destroy()
            messagebox.showinfo("Success", "F&N Documentation started successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not start software: {str(e)}")

    def change_fn_documentation_path(self, dialog):
        """Change the path to F&N Documentation software"""
        new_path = filedialog.askopenfilename(
            title="Select F&N Documentation executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        
        if new_path:
            self.set_fn_documentation_path(new_path)
            messagebox.showinfo("Success", f"Path updated to:\n{new_path}")
            dialog.destroy()

    def show_extravasation_management(self):
        """Display chemotherapy extravasation management with improved layout and print function"""
        self.clear_frame()

        # Main container
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        logo_frame = tk.Frame(left_frame, bg='#3498db')
        logo_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        tk.Label(logo_frame, text="OncoCare", font=('Helvetica', 24, 'bold'), 
                bg='#3498db', fg='white').pack(pady=(0, 10))
        tk.Label(logo_frame, text="Extravasation Management", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create scrollable frame
        canvas = tk.Canvas(right_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", on_frame_configure)
        
        # Mouse wheel scrolling
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)
        
        # Add content with new layout
        self.add_extravasation_content(scrollable_frame)
        
        # Print button (outside scrollable area)
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)
        
        print_btn = ttk.Button(btn_frame, text="Print Protocol", 
                             command=self.print_extravasation_protocol,
                             style='Blue.TButton')
        print_btn.pack(side=tk.RIGHT, padx=10)
        
        back_btn = ttk.Button(btn_frame, text="Back to Menu", 
                            command=self.main_menu,
                            style='Blue.TButton')
        back_btn.pack(side=tk.LEFT, padx=10)

    def add_extravasation_content(self, parent):
        """Add extravasation content with improved layout"""
        # Overview frame (top left)
        overview_frame = ttk.LabelFrame(parent, text="Chemotherapy Extravasation Overview", 
                                      style='TFrame')
        overview_frame.pack(fill=tk.X, padx=10, pady=5, anchor='nw')
        
        overview_text = (
            "Extravasation is the leakage of chemotherapy drugs into surrounding tissues. "
            "In pediatric patients, risk is higher due to smaller veins and movement.\n\n"
            "SEVERITY CLASSIFICATION:\n"
            "• Grade 1: Mild (erythema, swelling <1cm)\n"
            "• Grade 2: Moderate (pain, swelling 1-5cm)\n"
            "• Grade 3: Severe (ulceration, necrosis)\n"
            "• Grade 4: Life-threatening (compartment syndrome)"
        )
        ttk.Label(overview_frame, text=overview_text, justify='left', 
                 wraplength=700).pack(padx=10, pady=5)
        
        # General protocol frame (under overview)
        general_frame = ttk.LabelFrame(parent, text="General Management Protocol", 
                                     style='TFrame')
        general_frame.pack(fill=tk.X, padx=10, pady=5, anchor='nw')
        
        steps = [
        ("1. STOP INFUSION IMMEDIATELY", 
         "• Clamp the tubing closest to the IV site\n"
         "• Discontinue the infusion pump\n"
         "• Note exact time of recognition"),
        
        ("2. LEAVE CATHETER IN PLACE INITIALLY",
         "• Attempt to aspirate residual drug (use 1-3mL syringe for better suction)\n"
         "• Gently aspirate for at least 1 minute before removing\n"
         "• If unsuccessful, remove catheter after aspiration attempt"),
        
        ("3. ASSESS EXTENT OF EXTRAVASATION",
         "• Measure area of induration/erythema in cm\n"
         "• Photograph the site if possible\n"
         "• Grade severity (1-4) based on symptoms"),
        
        ("4. ADMINISTER SPECIFIC ANTIDOTE",
         "• See drug-specific protocols below\n"
         "• Prepare antidote within 15 minutes of recognition\n"
         "• Use 25-27G needle for subcutaneous administration"),
        
        ("5. APPLY TOPICAL MANAGEMENT",
         "• Warm compress (40-42°C) for vinca alkaloids\n"
         "• Cold compress (ice pack wrapped in cloth) for anthracyclines\n"
         "• Elevate extremity above heart level"),
        
        ("6. DOCUMENT THOROUGHLY",
         "• Complete incident report form\n"
         "• Record drug, concentration, volume extravasated\n"
         "• Document patient response and interventions"),
        
        ("7. NOTIFY PHYSICIAN/ONCOLOGIST",
         "• Immediate notification for grade 3-4 extravasation\n"
         "• Consider surgical consult for severe cases\n"
         "• Report to pharmacy for quality improvement")
        ]
        
        for step in steps:
            ttk.Label(general_frame, text=step, anchor='w').pack(fill=tk.X, padx=10, pady=2)
        
        # Drug selection frame
        drug_frame = ttk.LabelFrame(parent, text="Drug-Specific Protocol Selection",
                                  style='TFrame')
        drug_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Drug selection dropdown
        ttk.Label(drug_frame, text="Select Chemotherapy Drug:").pack(side=tk.LEFT, padx=5)
        
        self.drug_var = tk.StringVar()
        drug_dropdown = ttk.Combobox(drug_frame, textvariable=self.drug_var, 
                                   state="readonly", width=30)
        drug_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Initialize drug data with detailed protocols
        self.drug_protocols = {
            "Anthracyclines (Doxorubicin, Daunorubicin, Epirubicin)": {
                "Risk": "Vesicant (High Risk)\n• Severe tissue necrosis\n• Delayed onset (weeks to months)",
                "Antidote": (
                    "Dexrazoxane (Totect)\n"
                    "• Dose: 10:1 ratio to anthracycline dose (max 1000mg/m²/day)\n"
                    "• Schedule:\n"
                    "  - First dose: Within 6 hours of extravasation\n"
                    "  - Second dose: 24 hours after first dose\n"
                    "  - Third dose: 24 hours after second dose\n"
                    "• Infuse over 1-2 hours in large vein\n"
                    "• Preparation: Reconstitute with 25mL sterile water (50mg/mL)"
                ),
                "Local Care": (
                    "• Apply cold compress immediately (15-20 minutes QID x 3 days)\n"
                    "• Avoid pressure on affected area\n"
                    "• Consider topical DMSO 99% (apply thin layer Q8H x 7 days)"
                ),
                "Monitoring": (
                    "• Daily wound assessment for 7 days\n"
                    "• Monitor for delayed necrosis (may appear 1-4 weeks post)\n"
                    "• Consider MRI for deep tissue assessment"
                ),
                "Special Notes": (
                    "• Delayed tissue damage common (weeks to months later)\n"
                    "• Surgical debridement often required for necrosis\n"
                    "• Higher risk in infants and young children"
                )
            },
            "Vinca Alkaloids (Vincristine, Vinblastine, Vinorelbine)": {
                "Risk": "Vesicant (Moderate Risk)\n• Neurotoxicity possible\n• More severe with vincristine",
                "Antidote": (
                    "Hyaluronidase (Vitrase, Amphadase)\n"
                    "• Dose: 150-1500 units (1mL of 150 units/mL solution)\n"
                    "• Administration:\n"
                    "  - Inject subcutaneously around extravasation site\n"
                    "  - Use 25G needle, change puncture sites for multiple injections\n"
                    "  - May repeat after 1 hour if symptoms persist"
                ),
                "Local Care": (
                    "• Apply WARM compress (40-42°C) for 15-20 minutes QID x 3 days\n"
                    "• Elevate extremity\n"
                    "• Gentle massage to promote dispersion"
                ),
                "Monitoring": (
                    "• Assess neurovascular status Q2H x 24 hours\n"
                    "• Monitor for compartment syndrome\n"
                    "• Document sensory/motor function"
                ),
                "Special Notes": (
                    "• Neurotoxicity may occur without visible skin changes\n"
                    "• Consider nerve conduction studies if neurological symptoms\n"
                    "• More severe in patients with pre-existing neuropathy"
                )
            },
            "Platinum Compounds (Cisplatin, Carboplatin, Oxaliplatin)": {
                "Risk": "Irritant (Low-Moderate Risk)\n• Cisplatin more toxic than carboplatin\n• Oxaliplatin may cause cold-induced neuropathy",
                "Antidote": (
                    "No specific antidote\n"
                    "• Sodium thiosulfate 1/6M may be used (off-label)\n"
                    "• For cisplatin: Consider systemic sodium thiosulfate"
                ),
                "Local Care": (
                    "• Apply cold compress for 15-20 minutes QID x 24 hours\n"
                    "• Elevate extremity\n"
                    "• For oxaliplatin: Avoid cold exposure to affected area"
                ),
                "Monitoring": (
                    "• Daily assessment for 3 days\n"
                    "• Monitor for hypersensitivity reactions\n"
                    "• Check renal function if systemic absorption suspected"
                ),
                "Special Notes": (
                    "• Tissue damage usually resolves within 2 weeks\n"
                    "• Oxaliplatin extravasation may cause prolonged cold sensitivity\n"
                    "• Higher risk with concentrated solutions"
                )
            },
            "Alkylating Agents (Cyclophosphamide, Ifosfamide)": {
                "Risk": "Irritant (Low Risk)\n• Usually mild tissue damage\n• Higher risk with concentrated solutions",
                "Antidote": "No specific antidote",
                "Local Care": (
                    "• Cold compress for 15-20 minutes QID x 24 hours\n"
                    "• Topical corticosteroids may reduce inflammation\n"
                    "• Saline flush of affected area (off-label use)"
                ),
                "Monitoring": (
                    "• Assess site Q8H x 48 hours\n"
                    "• Monitor for infection\n"
                    "• Check urine for hemorrhagic cystitis (ifosfamide)"
                ),
                "Special Notes": (
                    "• Mesna not effective for local tissue damage\n"
                    "• Healing typically occurs within 1-2 weeks\n"
                    "• Rarely requires surgical intervention"
                )
            },
            "Taxanes (Paclitaxel, Docetaxel)": {
                "Risk": "Irritant (Moderate Risk)\n• Cremophor-containing formulations more irritating\n• Delayed reactions may occur 3-10 days post",
                "Antidote": "No specific antidote\n• Hyaluronidase may be considered (off-label)",
                "Local Care": (
                    "• Cold compress for 15-20 minutes QID x 24 hours\n"
                    "• Topical hydrocortisone 1% cream BID\n"
                    "• For severe reactions: Consider oral corticosteroids"
                ),
                "Monitoring": (
                    "• Assess for hypersensitivity reactions\n"
                    "• Monitor for neuropathic pain\n"
                    "• Document resolution of erythema"
                ),
                "Special Notes": (
                    "• More severe in patients with prior radiation to site\n"
                    "• Albumin-bound paclitaxel less likely to cause severe reactions\n"
                    "• May cause recall reactions at previous extravasation sites"
                )
            },
            "Etoposide": {
                "Risk": "Irritant (Low Risk)\n• Usually mild reactions\n• Rarely causes tissue necrosis",
                "Antidote": "No specific antidote",
                "Local Care": (
                    "• Cold compress for 15-20 minutes QID x 24 hours\n"
                    "• Topical corticosteroids for persistent inflammation"
                ),
                "Monitoring": "• Monitor site for 48 hours\n• Watch for hypersensitivity reactions",
                "Special Notes": "• Reactions typically resolve within 3-5 days\n• Higher risk with concentrated solutions"
            },
            "Methotrexate": {
                "Risk": "Irritant (Low Risk)\n• Mild local reaction\n• Rarely causes tissue damage",
                "Antidote": "Consider systemic leucovorin\n• Dose based on methotrexate exposure",
                "Local Care": (
                    "• Cold compress for 15-20 minutes QID x 24 hours\n"
                    "• Elevate extremity\n"
                    "• Alkalinization of urine if systemic absorption suspected"
                ),
                "Monitoring": (
                    "• Assess site Q8H x 48h\n"
                    "• Monitor renal function if significant absorption\n"
                    "• Check methotrexate levels if concern for systemic exposure"
                ),
                "Special Notes": (
                    "• Usually mild tissue damage\n"
                    "• Rarely requires surgical intervention\n"
                    "• Higher risk in patients with third spacing"
                )
            },
            "Bleomycin": {
                "Risk": "Non-vesicant\n• Rarely causes tissue damage\n• Minimal local reaction",
                "Antidote": "Not applicable",
                "Local Care": "• Observation only\n• Routine wound care if skin breakdown occurs",
                "Monitoring": "• Minimal monitoring required\n• Assess for infection if skin breaks",
                "Special Notes": "• Does not typically require specific treatment\n• Rarely causes significant tissue damage"
            },
            "Asparaginase": {
                "Risk": "Non-vesicant\n• Primary concern is hypersensitivity\n• Minimal local tissue effects",
                "Antidote": "Not applicable\n• Treat hypersensitivity reactions if they occur",
                "Local Care": "• Routine wound care\n• Cold compress if inflammation present",
                "Monitoring": "• Watch for allergic reactions\n• Monitor for infection",
                "Special Notes": "• Tissue damage extremely rare\n• More concern for systemic reactions than local effects"
            }
        }
        
        drug_dropdown['values'] = list(self.drug_protocols.keys())
        drug_dropdown.bind("<<ComboboxSelected>>", self.display_drug_protocol)
        
        # Protocol display frame
        self.protocol_display = ttk.LabelFrame(parent, text="Selected Drug Protocol",
                                             style='TFrame')
        self.protocol_display.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        
        # Initialize with empty display
        self.display_drug_protocol()

    def display_drug_protocol(self, event=None):
        """Display protocol for selected drug"""
        selected_drug = self.drug_var.get()
        
        # Clear previous content
        for widget in self.protocol_display.winfo_children():
            widget.destroy()
        
        if not selected_drug:
            ttk.Label(self.protocol_display, 
                     text="Please select a chemotherapy drug from the dropdown above",
                     style='TLabel').pack(padx=10, pady=20)
            return
        
        protocol = self.drug_protocols.get(selected_drug, {})
        
        # Create notebook for different sections
        notebook = ttk.Notebook(self.protocol_display)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Risk Assessment tab
        risk_frame = ttk.Frame(notebook)
        notebook.add(risk_frame, text="Risk Assessment")
        ttk.Label(risk_frame, text=protocol.get("Risk", "N/A"), 
                 justify='left', wraplength=650).pack(padx=10, pady=10)
        
        # Antidote tab
        antidote_frame = ttk.Frame(notebook)
        notebook.add(antidote_frame, text="Antidote")
        ttk.Label(antidote_frame, text=protocol.get("Antidote", "N/A"), 
                 justify='left', wraplength=650).pack(padx=10, pady=10)
        
        # Local Care tab
        local_frame = ttk.Frame(notebook)
        notebook.add(local_frame, text="Local Care")
        ttk.Label(local_frame, text=protocol.get("Local Care", "N/A"), 
                 justify='left', wraplength=650).pack(padx=10, pady=10)
        
        # Monitoring tab
        monitor_frame = ttk.Frame(notebook)
        notebook.add(monitor_frame, text="Monitoring")
        ttk.Label(monitor_frame, text=protocol.get("Monitoring", "N/A"), 
                 justify='left', wraplength=650).pack(padx=10, pady=10)
        
        # Special Notes tab
        notes_frame = ttk.Frame(notebook)
        notebook.add(notes_frame, text="Special Notes")
        ttk.Label(notes_frame, text=protocol.get("Special Notes", "N/A"), 
                 justify='left', wraplength=650).pack(padx=10, pady=10)

    def print_extravasation_protocol(self):
        """Print the current extravasation protocol"""
        selected_drug = self.drug_var.get()
        if not selected_drug:
            messagebox.showwarning("Print", "Please select a chemotherapy drug first")
            return
        
        protocol = self.drug_protocols.get(selected_drug, {})
        
        # Create a printable document
        doc = Document()
        doc.add_heading(f'Chemotherapy Extravasation Protocol: {selected_drug}', level=1)
        
        # Add overview
        doc.add_heading('Overview', level=2)
        doc.add_paragraph(
            "Extravasation is the leakage of chemotherapy drugs into surrounding tissues. "
            "Immediate recognition and proper management are crucial to minimize complications."
        )
        
        # Add general protocol
        doc.add_heading('General Management Protocol', level=2)
        general_steps = doc.add_paragraph()
        general_steps.add_run("1. STOP infusion immediately\n")
        general_steps.add_run("2. Aspirate residual drug\n")
        general_steps.add_run("3. Assess extent (measure, photograph, grade severity)\n")
        general_steps.add_run("4. Administer specific antidote\n")
        general_steps.add_run("5. Apply appropriate compress (warm/cold)\n")
        general_steps.add_run("6. Elevate extremity\n")
        general_steps.add_run("7. Document thoroughly\n")
        general_steps.add_run("8. Notify physician/oncologist")
        
        # Add drug-specific protocol
        doc.add_heading(f'{selected_drug} Specific Protocol', level=2)
        
        doc.add_heading('Risk Assessment', level=3)
        doc.add_paragraph(protocol.get("Risk", "N/A"))
        
        doc.add_heading('Antidote', level=3)
        doc.add_paragraph(protocol.get("Antidote", "N/A"))
        
        doc.add_heading('Local Care', level=3)
        doc.add_paragraph(protocol.get("Local Care", "N/A"))
        
        doc.add_heading('Monitoring', level=3)
        doc.add_paragraph(protocol.get("Monitoring", "N/A"))
        
        doc.add_heading('Special Notes', level=3)
        doc.add_paragraph(protocol.get("Special Notes", "N/A"))
        
        # Add footer
        doc.add_paragraph("\nGenerated by OncoCare Pediatric Oncology System")
        doc.add_paragraph(datetime.now().strftime("%Y-%m-%d %H:%M"))
        
        # Save and print
        temp_file = "temp_extravasation_protocol.docx"
        doc.save(temp_file)
        
        try:
            os.startfile(temp_file, "print")
            messagebox.showinfo("Print", "Protocol sent to printer")
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print document: {str(e)}")
        finally:
            # Clean up after printing
            def delete_temp_file():
                try:
                    os.remove(temp_file)
                except:
                    pass
            
            # Schedule file deletion after a delay
            self.root.after(5000, delete_temp_file)

    def show_calculators(self):
        """Display the medical calculators window"""
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
        
        tk.Label(logo_frame, text="Medical Calculators", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Notebook for different calculators
        notebook = ttk.Notebook(right_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # BSA Calculator Tab
        bsa_frame = ttk.Frame(notebook)
        notebook.add(bsa_frame, text="BSA Calculator")
        self.setup_bsa_calculator(bsa_frame)
        
        # IV Fluid Calculator Tab
        iv_frame = ttk.Frame(notebook)
        notebook.add(iv_frame, text="IV Fluid Calculator")
        self.setup_iv_calculator(iv_frame)
        
        # Dosage Calculator Tab
        dosage_frame = ttk.Frame(notebook)
        notebook.add(dosage_frame, text="Chemo Dosage Calculator")
        self.setup_dosage_calculator(dosage_frame)
        
        # Antibiotics Calculator Tab
        abx_frame = ttk.Frame(notebook)
        notebook.add(abx_frame, text="Antibiotics Calculator")
        self.setup_antibiotics_calculator(abx_frame)
        
        # Button frame
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                  style='Blue.TButton').pack(fill=tk.X, pady=10)

    def setup_antibiotics_calculator(self, parent):
        """Setup the antibiotics calculator interface"""
        frame = self.create_scrollable_frame(parent)
        
        # Antibiotics data dictionary (same as before)
        self.antibiotics_data = {
            "MEROPENEM": {
                "dose_range": "10-40 mg/kg/dose",
                "min_dose": 10,
                "max_dose": 40,
                "max_total_dose": 2000,
                "unit": "mg",
                "frequency": "Every 8 hours",
                "per_day": False,
                "incompatible_solutions": ["Dextrose solutions", "Other beta-lactams"],
                "interactions": "Probenecid may decrease renal clearance. Valproic acid levels may be reduced.",
                "notes": "Adjust dose in renal impairment. CNS side effects at high doses."
            },
            "AMIKACIN": {
                "dose_range": "15-22.5 mg/kg/day",
                "min_dose": 15,
                "max_dose": 22.5,
                "max_total_dose": 1500,
                "unit": "mg",
                "frequency": "Every 24 hours",
                "per_day": True,
                "incompatible_solutions": ["Penicillins", "Cephalosporins"],
                "interactions": "Nephrotoxic when combined with other aminoglycosides, vancomycin, or diuretics.",
                "notes": "Monitor levels. Adjust dose in renal impairment."
            },
            "CEFTRIAXONE (ROCEPHINE)": {
                "dose_range": "50-100 mg/kg/day",
                "min_dose": 50,
                "max_dose": 100,
                "max_total_dose": 4000,
                "unit": "mg",
                "frequency": "Every 12-24 hours",
                "per_day": True,
                "incompatible_solutions": ["Calcium-containing solutions", "Aminoglycosides"],
                "interactions": "May increase warfarin effect. Calcium precipitates in neonates.",
                "notes": "Do not mix with calcium-containing solutions in neonates."
            },
            "CIPROFLOXACIN": {
                "dose_range": "10-20 mg/kg/dose",
                "min_dose": 10,
                "max_dose": 20,
                "max_total_dose": 800,
                "unit": "mg",
                "frequency": "Every 12 hours",
                "per_day": False,
                "incompatible_solutions": ["Divalent cation solutions"],
                "interactions": "Antacids reduce absorption. May increase theophylline levels.",
                "notes": "Avoid in growing children due to cartilage toxicity risk."
            },
            "TAZOCIN": {
                "dose_range": "80-100 mg/kg/dose (piperacillin component)",
                "min_dose": 80,
                "max_dose": 100,
                "max_total_dose": 4000,
                "unit": "mg",
                "frequency": "Every 6-8 hours",
                "per_day": False,
                "incompatible_solutions": ["Aminoglycosides"],
                "interactions": "Probenecid may increase levels. May increase bleeding risk with anticoagulants.",
                "notes": "Contains piperacillin/tazobactam 8:1 ratio."
            },
            "VANCOMYCIN": {
                "dose_range": "10-15 mg/kg/dose",
                "min_dose": 10,
                "max_dose": 15,
                "max_total_dose": 1000,
                "unit": "mg",
                "frequency": "Every 6-8 hours",
                "per_day": False,
                "incompatible_solutions": ["Alkaline solutions", "Beta-lactams"],
                "interactions": "Nephrotoxic with aminoglycosides. Ototoxic with loop diuretics.",
                "notes": "Monitor levels. Red man syndrome risk with rapid infusion."
            },
            "METRONIDAZOLE (FLAGYL)": {
                "dose_range": "7.5-10 mg/kg/dose",
                "min_dose": 7.5,
                "max_dose": 10,
                "max_total_dose": 500,
                "unit": "mg",
                "frequency": "Every 8 hours",
                "per_day": False,
                "incompatible_solutions": ["Aluminum-containing solutions"],
                "interactions": "Disulfiram-like reaction with alcohol. Increases warfarin effect.",
                "notes": "For anaerobic infections. CNS toxicity at high doses."
            },
            "FLUCONAZOLE": {
                "dose_range": "6-12 mg/kg/day",
                "min_dose": 6,
                "max_dose": 12,
                "max_total_dose": 800,
                "unit": "mg",
                "frequency": "Every 24 hours",
                "per_day": True,
                "incompatible_solutions": ["None known"],
                "interactions": "Increases phenytoin, warfarin levels. Rifampin decreases levels.",
                "notes": "Adjust dose in renal impairment."
            },
            "AMPHOTERICIN B": {
                "dose_range": "1-1.5 mg/kg/day",
                "min_dose": 1,
                "max_dose": 1.5,
                "max_total_dose": 50,
                "unit": "mg",
                "frequency": "Every 24 hours",
                "per_day": True,
                "incompatible_solutions": ["Saline solutions"],
                "interactions": "Nephrotoxic with cyclosporine, aminoglycosides. Hypokalemia with diuretics.",
                "notes": "Pre-medicate for infusion reactions. Monitor renal function."
            },
            "VORICONAZOLE": {
                "dose_range": "7-8 mg/kg/dose",
                "min_dose": 7,
                "max_dose": 8,
                "max_total_dose": 400,
                "unit": "mg",
                "frequency": "Every 12 hours",
                "per_day": False,
                "incompatible_solutions": ["None known"],
                "interactions": "Many CYP450 interactions. Avoid with rifampin, carbamazepine.",
                "notes": "Therapeutic drug monitoring recommended."
            },
            "AUGMENTIN": {
                "dose_range": "25-45 mg/kg/dose (amoxicillin component)",
                "min_dose": 25,
                "max_dose": 45,
                "max_total_dose": 1750,
                "unit": "mg",
                "frequency": "Every 8 hours",
                "per_day": False,
                "incompatible_solutions": ["None known"],
                "interactions": "Probenecid increases levels. Allopurinol increases rash risk.",
                "notes": "Contains amoxicillin/clavulanate 7:1 ratio."
            },
            "ACYCLOVIR": {
                "dose_range": "10-20 mg/kg/dose",
                "min_dose": 10,
                "max_dose": 20,
                "max_total_dose": 800,
                "unit": "mg",
                "frequency": "Every 8 hours",
                "per_day": False,
                "incompatible_solutions": ["Alkaline solutions"],
                "interactions": "Probenecid increases levels. Nephrotoxic with other nephrotoxic drugs.",
                "notes": "Hydrate well. Adjust dose in renal impairment."
            },
            "AMPICILLIN": {
                "dose_range": "25-50 mg/kg/dose",
                "min_dose": 25,
                "max_dose": 50,
                "max_total_dose": 2000,
                "unit": "mg",
                "frequency": "Every 6 hours",
                "per_day": False,
                "incompatible_solutions": ["Aminoglycosides"],
                "interactions": "Probenecid increases levels. Allopurinol increases rash risk.",
                "notes": "Monitor for rash. Adjust dose in renal impairment."
            },
            "CLARITHROMYCIN": {
                "dose_range": "7.5-15 mg/kg/dose",
                "min_dose": 7.5,
                "max_dose": 15,
                "max_total_dose": 1000,
                "unit": "mg",
                "frequency": "Every 12 hours",
                "per_day": False,
                "incompatible_solutions": ["None known"],
                "interactions": "Many CYP450 interactions. Increases digoxin, theophylline levels.",
                "notes": "QT prolongation risk at high doses."
            },
            "CEFTAZIDIME": {
                "dose_range": "30-50 mg/kg/dose",
                "min_dose": 30,
                "max_dose": 50,
                "max_total_dose": 2000,
                "unit": "mg",
                "frequency": "Every 8 hours",
                "per_day": False,
                "incompatible_solutions": ["Aminoglycosides"],
                "interactions": "Probenecid increases levels. Nephrotoxic with loop diuretics.",
                "notes": "Pseudomonas coverage. Adjust dose in renal impairment."
            },
            "CLOXACILLIN": {
                "dose_range": "25-50 mg/kg/dose",
                "min_dose": 25,
                "max_dose": 50,
                "max_total_dose": 2000,
                "unit": "mg",
                "frequency": "Every 6 hours",
                "per_day": False,
                "incompatible_solutions": ["Aminoglycosides"],
                "interactions": "Probenecid increases levels. Allopurinol increases rash risk.",
                "notes": "For MSSA infections. Adjust dose in renal impairment."
            },
            "CLINDAMYCIN": {
                "dose_range": "5-10 mg/kg/dose",
                "min_dose": 5,
                "max_dose": 10,
                "max_total_dose": 600,
                "unit": "mg",
                "frequency": "Every 6-8 hours",
                "per_day": False,
                "incompatible_solutions": ["Aminoglycosides"],
                "interactions": "Neuromuscular blockade with paralytics. Erythromycin antagonizes effect.",
                "notes": "C. diff risk. Good anaerobic coverage."
            },
            "COLISTIN": {
                "dose_range": "2.5-5 mg/kg/day (CBA)",
                "min_dose": 2.5,
                "max_dose": 5,
                "max_total_dose": 300,
                "unit": "mg",
                "frequency": "Every 8-12 hours",
                "per_day": True,
                "incompatible_solutions": ["None known"],
                "interactions": "Nephrotoxic with other nephrotoxic drugs. Neuromuscular blockade with paralytics.",
                "notes": "For MDR gram-negative infections. Monitor renal function."
            },
            "GENTAMICIN": {
                "dose_range": "5-7.5 mg/kg/day",
                "min_dose": 5,
                "max_dose": 7.5,
                "max_total_dose": 400,
                "unit": "mg",
                "frequency": "Every 24 hours",
                "per_day": True,
                "incompatible_solutions": ["Penicillins", "Cephalosporins"],
                "interactions": "Nephrotoxic when combined with other aminoglycosides, vancomycin, or diuretics.",
                "notes": "Monitor levels. Adjust dose in renal impairment."
            }
        }
        
        # Title
        ttk.Label(frame, text="Antibiotics Calculator", font=('Helvetica', 16, 'bold')).grid(row=0, column=0, columnspan=6, pady=10)
        
        # Single weight input for all antibiotics
        weight_frame = ttk.Frame(frame)
        weight_frame.grid(row=1, column=0, columnspan=6, pady=(0,20), sticky="ew")
        
        self.antibiotics_weight_var = tk.StringVar()
        ttk.Label(weight_frame, text="*Patient Weight (kg):", font=('Helvetica', 10, 'bold')).grid(row=0, column=0, padx=5, sticky="e")
        ttk.Entry(weight_frame, textvariable=self.antibiotics_weight_var, width=10, font=('Helvetica', 10)).grid(row=0, column=1, padx=5, sticky="w")
        
        # Instructions
        ttk.Label(frame, text="Select up to 5 antibiotics. Dose fields are editable.", 
                 font=('Helvetica', 10)).grid(row=2, column=0, columnspan=6, pady=(0,20))
        
        # Create frames for each antibiotic calculation
        self.antibiotic_frames = []
        self.antibiotic_vars = []
        
        for i in range(5):
            # Frame for each antibiotic
            abx_frame = ttk.LabelFrame(frame, text=f"Antibiotic {i+1}", padding=(10,5))
            abx_frame.grid(row=3+i*5, column=0, columnspan=6, sticky="ew", padx=5, pady=5)
            self.antibiotic_frames.append(abx_frame)
            
            # Variables for this antibiotic
            var_dict = {
                "drug_var": tk.StringVar(),
                "dose_var": tk.StringVar(),
                "result_var": tk.StringVar(),
                "result_label": None,
                "max_reached": tk.BooleanVar(value=False),
                "default_dose": tk.StringVar()  # To store the default dose
            }
            self.antibiotic_vars.append(var_dict)
            
            # Drug selection
            ttk.Label(abx_frame, text="*Drug:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            drug_combo = ttk.Combobox(abx_frame, textvariable=var_dict["drug_var"], 
                                     values=sorted(self.antibiotics_data.keys()), 
                                     state="readonly", width=20)
            drug_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
            drug_combo.bind("<<ComboboxSelected>>", lambda e, idx=i: self.update_antibiotic_info(idx))
            
            # Editable dose field (shows default but can be changed)
            ttk.Label(abx_frame, text="Dose (mg/kg):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
            dose_entry = ttk.Entry(abx_frame, textvariable=var_dict["dose_var"], width=10)
            dose_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
            
            # Information display area (same as before)
            info_frame = ttk.Frame(abx_frame)
            info_frame.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(5,0))
            
            ttk.Label(info_frame, text="Dose Range:").grid(row=0, column=0, padx=5, sticky="e")
            ttk.Label(info_frame, text="", width=20, anchor="w").grid(row=0, column=1, padx=5, sticky="w")
            
            ttk.Label(info_frame, text="Frequency:").grid(row=0, column=2, padx=5, sticky="e")
            ttk.Label(info_frame, text="", width=15, anchor="w").grid(row=0, column=3, padx=5, sticky="w")
            
            ttk.Label(info_frame, text="Per:").grid(row=0, column=4, padx=5, sticky="e")
            ttk.Label(info_frame, text="", width=10, anchor="w").grid(row=0, column=5, padx=5, sticky="w")
            
            ttk.Label(info_frame, text="Incompatible Solutions:").grid(row=1, column=0, padx=5, sticky="e")
            ttk.Label(info_frame, text="", wraplength=300, anchor="w").grid(row=1, column=1, columnspan=3, padx=5, sticky="w")
            
            ttk.Label(info_frame, text="Interactions:").grid(row=2, column=0, padx=5, sticky="ne")
            ttk.Label(info_frame, text="", wraplength=400, anchor="w").grid(row=2, column=1, columnspan=5, padx=5, sticky="w")
            
            ttk.Label(info_frame, text="Notes:").grid(row=3, column=0, padx=5, sticky="ne")
            ttk.Label(info_frame, text="", wraplength=400, anchor="w").grid(row=3, column=1, columnspan=5, padx=5, sticky="w")
            
            # Result
            result_frame = ttk.Frame(abx_frame)
            result_frame.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(5,0))
            
            ttk.Label(result_frame, text="Calculated Dose:").grid(row=0, column=0, padx=5, sticky="e")
            result_label = ttk.Label(result_frame, textvariable=var_dict["result_var"], 
                                   font=('Helvetica', 10, 'bold'))
            result_label.grid(row=0, column=1, padx=5, sticky="w")
            var_dict["result_label"] = result_label
            
            # Separator
            ttk.Separator(abx_frame, orient='horizontal').grid(row=3, column=0, columnspan=6, sticky="ew", pady=5)
        
        # Button frame on the side
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3+5*5, column=4, columnspan=2, sticky="ne", padx=10, pady=20)
        
        # Calculate button
        calc_btn = ttk.Button(button_frame, text="Calculate All", 
                            command=self.calculate_all_antibiotics,
                            style='Large.TButton', width=15)
        calc_btn.pack(pady=10, fill=tk.X)
        
        # Clear button
        clear_btn = ttk.Button(button_frame, text="Clear All", 
                             command=self.clear_all_antibiotics,
                             style='TButton', width=15)
        clear_btn.pack(pady=10, fill=tk.X)
        
        # Back to menu button
        back_btn = ttk.Button(button_frame, text="Back to Menu", 
                             command=self.main_menu,
                             style='TButton', width=15)
        back_btn.pack(pady=10, fill=tk.X)
        
        # Interaction check frame
        interaction_frame = ttk.LabelFrame(frame, text="Antibiotic Interactions Check", padding=10)
        interaction_frame.grid(row=3+5*5, column=0, columnspan=4, sticky="nsew", padx=5, pady=10)
        
        self.interaction_text = tk.Text(interaction_frame, height=8, width=60, wrap=tk.WORD,
                                      state=tk.DISABLED, font=('Helvetica', 9))
        self.interaction_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)
        frame.columnconfigure(4, weight=1)
        frame.columnconfigure(5, weight=1)
        
        # Create a custom style for the large button
        style = ttk.Style()
        style.configure('Large.TButton', font=('Helvetica', 12, 'bold'), padding=10)

    def update_antibiotic_info(self, index):
        """Update the information display when an antibiotic is selected"""
        drug = self.antibiotic_vars[index]["drug_var"].get()
        
        if drug in self.antibiotics_data:
            data = self.antibiotics_data[drug]
            frame = self.antibiotic_frames[index]
            
            # Update the information labels
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=1)[0].config(text=data["dose_range"])
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=3)[0].config(text=data["frequency"])
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=5)[0].config(
                text="DAY" if data["per_day"] else "DOSE")
            
            # Set default dose to max recommended dose (editable)
            self.antibiotic_vars[index]["dose_var"].set(str(data["max_dose"]))
            self.antibiotic_vars[index]["default_dose"].set(str(data["max_dose"]))
            
            # Update incompatible solutions
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=1, column=1)[0].config(
                text=", ".join(data["incompatible_solutions"]))
            
            # Update interactions
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=2, column=1)[0].config(text=data["interactions"])
            
            # Update notes
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=3, column=1)[0].config(text=data["notes"])
            
            # Clear any previous result
            self.antibiotic_vars[index]["result_var"].set("")
            self.antibiotic_vars[index]["max_reached"].set(False)
            self.antibiotic_vars[index]["result_label"].config(foreground='black')
            
            # Check for interactions
            self.check_antibiotic_interactions()

    def calculate_all_antibiotics(self):
        """Calculate doses for all antibiotics using the single weight value"""
        weight_str = self.antibiotics_weight_var.get()
        
        if not weight_str:
            messagebox.showerror("Error", "Please enter patient weight first")
            return
            
        try:
            weight = float(weight_str)
            if weight <= 0:
                raise ValueError("Weight must be positive")
                
            for i in range(5):
                self.calculate_single_antibiotic(i, weight)
            
            # Check interactions after calculation
            self.check_antibiotic_interactions()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid weight value: {str(e)}")

    def calculate_single_antibiotic(self, index, weight=None):
        """Calculate dose for a single antibiotic"""
        if weight is None:
            weight_str = self.antibiotics_weight_var.get()
            if not weight_str:
                return
            try:
                weight = float(weight_str)
            except ValueError:
                return
                
        drug = self.antibiotic_vars[index]["drug_var"].get()
        dose_str = self.antibiotic_vars[index]["dose_var"].get()
        
        if not drug:
            return  # Skip if no drug selected
            
        try:
            dose = float(dose_str)
            data = self.antibiotics_data[drug]
            
            # Warn if dose is outside recommended range
            if dose < data["min_dose"] or dose > data["max_dose"]:
                messagebox.showwarning("Warning", 
                                     f"Dose for {drug} is outside recommended range ({data['dose_range']})")
                
            # Calculate total dose
            calculated_dose = dose * weight
            
            # Check against max total dose
            max_reached = calculated_dose > data["max_total_dose"]
            if max_reached:
                calculated_dose = data["max_total_dose"]
                self.antibiotic_vars[index]["max_reached"].set(True)
                self.antibiotic_vars[index]["result_label"].config(foreground='red')
            else:
                self.antibiotic_vars[index]["max_reached"].set(False)
                self.antibiotic_vars[index]["result_label"].config(foreground='black')
            
            # Format the result
            per_text = "per day" if data["per_day"] else "per dose"
            max_warning = " (MAX DOSE REACHED!)" if max_reached else ""
            result_text = f"{calculated_dose:.2f} {data['unit']} {per_text}{max_warning}"
            
            self.antibiotic_vars[index]["result_var"].set(result_text)
            
        except ValueError:
            messagebox.showerror("Error", f"Invalid dose value for {drug}")

    def check_antibiotic_interactions(self):
        """Check for interactions between selected antibiotics"""
        selected_drugs = []
        interaction_text = ""
        
        # Get all selected drugs
        for i in range(5):
            drug = self.antibiotic_vars[i]["drug_var"].get()
            if drug and drug in self.antibiotics_data:
                selected_drugs.append(drug)
        
        # Check for known interactions
        if len(selected_drugs) >= 2:
            interaction_text = "Potential Interactions:\n"
            
            # Check each pair of drugs
            for i in range(len(selected_drugs)):
                for j in range(i+1, len(selected_drugs)):
                    drug1 = selected_drugs[i]
                    drug2 = selected_drugs[j]
                    
                    # Get interactions for drug1
                    interactions1 = self.antibiotics_data[drug1]["interactions"]
                    
                    # Check if drug2 is mentioned in drug1's interactions
                    if drug2.lower() in interactions1.lower():
                        interaction_text += f"- {drug1} + {drug2}: {interactions1}\n"
                    
                    # Also check the reverse
                    interactions2 = self.antibiotics_data[drug2]["interactions"]
                    if drug1.lower() in interactions2.lower():
                        interaction_text += f"- {drug2} + {drug1}: {interactions2}\n"
            
            if interaction_text == "Potential Interactions:\n":
                interaction_text += "No significant interactions detected between selected antibiotics."
        
        # Update the interaction text widget
        self.interaction_text.config(state=tk.NORMAL)
        self.interaction_text.delete(1.0, tk.END)
        self.interaction_text.insert(tk.END, interaction_text)
        self.interaction_text.config(state=tk.DISABLED)

    def clear_all_antibiotics(self):
        """Clear all antibiotic calculation fields"""
        for i in range(5):
            self.antibiotic_vars[i]["drug_var"].set("")
            self.antibiotic_vars[i]["dose_var"].set("")
            self.antibiotic_vars[i]["result_var"].set("")
            self.antibiotic_vars[i]["max_reached"].set(False)
            self.antibiotic_vars[i]["result_label"].config(foreground='black')
            
            # Clear information labels
            frame = self.antibiotic_frames[i]
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=1)[0].config(text="")
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=3)[0].config(text="")
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=0, column=5)[0].config(text="")
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=1, column=1)[0].config(text="")
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=2, column=1)[0].config(text="")
            frame.grid_slaves(row=1, column=0)[0].grid_slaves(row=3, column=1)[0].config(text="")
        
        # Clear interaction text
        self.interaction_text.config(state=tk.NORMAL)
        self.interaction_text.delete(1.0, tk.END)
        self.interaction_text.config(state=tk.DISABLED)

    def setup_bsa_calculator(self, parent):
        """Setup the BSA calculator interface"""
        frame = ttk.Frame(parent, padding="10 10 10 10", style='TFrame')
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Variables
        self.height_var = tk.StringVar()
        self.weight_var = tk.StringVar()
        self.bsa_result_var = tk.StringVar()
        self.bsa_formula_var = tk.StringVar(value="Mosteller")
        
        # Input fields
        ttk.Label(frame, text="Height (cm):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.height_var).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Weight (kg):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.weight_var).grid(row=1, column=1, padx=5, pady=5)
        
        # Formula selection
        ttk.Label(frame, text="Formula:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ttk.Combobox(frame, textvariable=self.bsa_formula_var, 
                    values=["Mosteller", "DuBois"], state="readonly").grid(row=2, column=1, padx=5, pady=5)
        
        # Calculate button
        ttk.Button(frame, text="Calculate BSA", command=self.calculate_bsa,
                  style='Blue.TButton').grid(row=3, column=0, columnspan=2, pady=10)
        
        # Result
        ttk.Label(frame, text="BSA (m²):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(frame, textvariable=self.bsa_result_var, font=('Helvetica', 12, 'bold')).grid(row=4, column=1, padx=5, pady=5, sticky="w")

    def calculate_bsa(self):
        """Calculate Body Surface Area using selected formula"""
        try:
            height = float(self.height_var.get())
            weight = float(self.weight_var.get())
            formula = self.bsa_formula_var.get()
            
            if formula == "Mosteller":
                bsa = ((height * weight) / 3600) ** 0.5
            else:  # DuBois
                bsa = 0.007184 * (height ** 0.725) * (weight ** 0.425)
            
            self.bsa_result_var.set(f"{bsa:.4f}")
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values for height and weight")

    def setup_iv_calculator(self, parent):
        """Setup the IV fluid calculator interface"""
        frame = ttk.Frame(parent, padding="10 10 10 10", style='TFrame')
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Variables
        self.iv_bsa_var = tk.StringVar()
        self.iv_weight_var = tk.StringVar()  # New weight variable
        self.iv_fluid_type_var = tk.StringVar(value="4000ml per BSA")
        self.iv_drug_dilution_var = tk.StringVar(value="0")
        self.iv_result_var = tk.StringVar()
        
        # Input fields
        ttk.Label(frame, text="BSA (m²):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.iv_bsa_var).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Weight (kg):").grid(row=1, column=0, padx=5, pady=5, sticky="e")  # New weight field
        ttk.Entry(frame, textvariable=self.iv_weight_var).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Fluid Type:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ttk.Combobox(frame, textvariable=self.iv_fluid_type_var, 
                    values=["4000ml per BSA", "3000ml per BSA", "2000ml per BSA", 
                           "400ml per BSA", "Full Maintenance", "Half Maintenance"],
                    state="readonly").grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Drug Dilution (ml):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.iv_drug_dilution_var).grid(row=3, column=1, padx=5, pady=5)
        
        # Calculate button
        ttk.Button(frame, text="Calculate IV Rate", command=self.calculate_iv_rate,
                  style='Blue.TButton').grid(row=4, column=0, columnspan=2, pady=10)
        
        # Result
        ttk.Label(frame, text="IV Rate (ml):").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(frame, textvariable=self.iv_result_var, font=('Helvetica', 12, 'bold')).grid(row=5, column=1, padx=5, pady=5, sticky="w")

    def calculate_iv_rate(self):
        """Calculate IV fluid rate based on BSA/weight and selected protocol"""
        try:
            bsa = float(self.iv_bsa_var.get())
            weight = float(self.iv_weight_var.get())  # Get weight from the new field
            fluid_type = self.iv_fluid_type_var.get()
            drug_dilution = float(self.iv_drug_dilution_var.get() or 0)
            
            if fluid_type == "4000ml per BSA":
                rate = 4000 * bsa - drug_dilution
            elif fluid_type == "3000ml per BSA":
                rate = 3000 * bsa - drug_dilution
            elif fluid_type == "2000ml per BSA":
                rate = 2000 * bsa - drug_dilution
            elif fluid_type == "400ml per BSA":
                rate = 400 * bsa - drug_dilution
            elif fluid_type == "Full Maintenance":
                # Holliday-Segar formula using actual weight
                if weight <= 10:
                    rate = weight * 100 - drug_dilution
                elif weight <= 20:
                    rate = 1000 + (weight - 10) * 50 - drug_dilution
                else:
                    rate = 1500 + (weight - 20) * 20 - drug_dilution
            else:  # Half Maintenance
                if weight <= 10:
                    rate = (weight * 100) / 2 - drug_dilution
                elif weight <= 20:
                    rate = (1000 + (weight - 10) * 50) / 2 - drug_dilution
                else:
                    rate = (1500 + (weight - 20) * 20) / 2 - drug_dilution
            
            self.iv_result_var.set(f"{rate:.2f}")
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")

    def setup_dosage_calculator(self, parent):
        """Setup the dosage calculator interface"""
        frame = ttk.Frame(parent, padding="10 10 10 10", style='TFrame')
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Chemotherapy drugs data (drug: [dose_per_kg, max_dose, unit])
        self.chemo_drugs = {
            "Vincristine": [1.5, 2, "mg"],
            "Dactinomycin": [0.015, 2.5, "mg"],
            "Doxorubicin": [30, 60, "mg/m²"],
            "Cyclophosphamide": [1000, 2000, "mg/m²"],
            "Methotrexate": [12, 15, "g/m²"],
            "Cisplatin": [100, 120, "mg/m²"],
            "Carboplatin": [600, 800, "mg/m²"],
            "Etoposide": [100, 500, "mg/m²"],
            "Ifosfamide": [1800, 3000, "mg/m²"],
            "Cytarabine": [100, 300, "mg/m²"]
        }
        
        # Variables
        self.drug_var = tk.StringVar()
        self.dose_var = tk.StringVar()
        self.max_dose_var = tk.StringVar()
        self.patient_weight_var = tk.StringVar()
        self.bsa_var = tk.StringVar()
        self.dosage_result_var = tk.StringVar()
        
        # Drug selection
        ttk.Label(frame, text="Drug:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        drug_combo = ttk.Combobox(frame, textvariable=self.drug_var, 
                                values=list(self.chemo_drugs.keys()), state="readonly")
        drug_combo.grid(row=0, column=1, padx=5, pady=5)
        drug_combo.bind("<<ComboboxSelected>>", self.update_drug_dosage)
        
        # Dose and max dose (editable)
        ttk.Label(frame, text="Dose per kg/m²:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.dose_var).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Max Dose:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.max_dose_var).grid(row=2, column=1, padx=5, pady=5)
        
        # Patient parameters
        ttk.Label(frame, text="Patient Weight (kg):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.patient_weight_var).grid(row=3, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="BSA (m²):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(frame, textvariable=self.bsa_var).grid(row=4, column=1, padx=5, pady=5)
        
        # Calculate button
        ttk.Button(frame, text="Calculate Dosage", command=self.calculate_dosage,
                  style='Blue.TButton').grid(row=5, column=0, columnspan=2, pady=10)
        
        # Result
        ttk.Label(frame, text="Dosage:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(frame, textvariable=self.dosage_result_var, font=('Helvetica', 12, 'bold')).grid(row=6, column=1, padx=5, pady=5, sticky="w")

    def update_drug_dosage(self, event=None):
        """Update dose and max dose fields when drug is selected"""
        drug = self.drug_var.get()
        if drug in self.chemo_drugs:
            dose, max_dose, unit = self.chemo_drugs[drug]
            self.dose_var.set(str(dose))
            self.max_dose_var.set(str(max_dose))

    def calculate_dosage(self):
        """Calculate chemotherapy dosage with weight and BSA considerations"""
        try:
            drug = self.drug_var.get()
            dose_per_kg = float(self.dose_var.get())
            max_dose = float(self.max_dose_var.get())
            weight = float(self.patient_weight_var.get())
            bsa = float(self.bsa_var.get())
            
            if drug in self.chemo_drugs:
                unit = self.chemo_drugs[drug][2]
                
                # Calculate based on unit
                if unit.endswith("/m²"):
                    calculated_dose = dose_per_kg * bsa
                else:
                    calculated_dose = dose_per_kg * weight
                
                # Apply max dose
                final_dose = min(calculated_dose, max_dose)
                
                # Special consideration for patients under 10kg
                if weight < 10:
                    final_dose = final_dose * 1.1  # 10% increase for very small patients
                    warning = "\n(Warning: Patient under 10kg - dose increased by 10%)"
                else:
                    warning = ""
                
                self.dosage_result_var.set(f"{final_dose:.2f} {unit}{warning}")
            else:
                messagebox.showerror("Error", "Please select a valid drug")
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")

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

        if self.users[self.current_user]["role"] in ["admin", "editor", "pharmacist"]:
            ttk.Button(btn_frame, text="Lab & EF", 
                      command=lambda: self.show_lab_ef_window(patient_data),
                      style='Purple.TButton').pack(side=tk.LEFT, padx=10)

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
        """Open the enhanced statistics window with multiple analysis options"""
        if self.users[self.current_user]["role"] not in ["admin", "editor"]:
            messagebox.showerror("Access Denied", "Only admins and editors can access statistics.")
            return

        self.clear_frame()
        self.root.minsize(800, 600)

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
        
        tk.Label(logo_frame, text="Advanced Statistics", 
                font=('Helvetica', 14), bg='#3498db', fg='white').pack(pady=(0, 40))
        
        # Right side with content
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Analysis options frame
        options_frame = ttk.Frame(right_frame, style='TFrame')
        options_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(options_frame, text="Select Analysis:", 
                 font=('Helvetica', 12)).pack(side=tk.LEFT, padx=5)
        
        self.analysis_var = tk.StringVar(value="Malignancy Distribution")
        analysis_options = [
            "Malignancy Distribution",
            "Age at Diagnosis",
            "Gender Distribution",
            "Treatment Outcomes",
            "B-Symptoms Analysis",
            "Risk Group Analysis",
            "Diagnosis Timeline",
            "Survival Analysis"
        ]
        
        analysis_menu = ttk.Combobox(options_frame, textvariable=self.analysis_var,
                                   values=analysis_options, state="readonly")
        analysis_menu.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(options_frame, text="Generate", command=self.generate_analysis,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=10)

        # Filter frame
        filter_frame = ttk.Frame(right_frame, style='TFrame')
        filter_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(filter_frame, text="Filter by Malignancy:", 
                 font=('Helvetica', 12)).pack(side=tk.LEFT, padx=5)
        
        self.stats_malignancy_filter = ttk.Combobox(filter_frame, values=["All"] + MALIGNANCIES, 
                                                   state="readonly")
        self.stats_malignancy_filter.set("All")
        self.stats_malignancy_filter.pack(side=tk.LEFT, padx=5)

        # Date range filter
        ttk.Label(filter_frame, text="Date Range:", 
                 font=('Helvetica', 12)).pack(side=tk.LEFT, padx=5)
        
        self.start_date_var = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.start_date_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(self.start_date_var)).pack(side=tk.LEFT, padx=0)
        
        ttk.Label(filter_frame, text="to").pack(side=tk.LEFT, padx=5)
        
        self.end_date_var = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.end_date_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(self.end_date_var)).pack(side=tk.LEFT, padx=0)

        # Main content area for statistics
        self.stats_content_frame = ttk.Frame(right_frame, style='TFrame')
        self.stats_content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Button frame at bottom
        btn_frame = ttk.Frame(right_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Button(btn_frame, text="Export to Excel", command=self.export_statistics,
                  style='Green.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Print Report", command=self.print_statistics_report,
                  style='Yellow.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=10)

        # Generate initial analysis
        self.generate_analysis()

    def generate_analysis(self):
        """Generate the selected analysis with current filters"""
        # Clear previous content
        for widget in self.stats_content_frame.winfo_children():
            widget.destroy()

        if not self.patient_data:
            ttk.Label(self.stats_content_frame, text="No patient data available", 
                     style='TLabel').pack(pady=20)
            return

        # Apply filters
        filtered_data = self.apply_statistics_filters()

        if not filtered_data:
            ttk.Label(self.stats_content_frame, text="No data matches the selected filters", 
                     style='TLabel').pack(pady=20)
            return

        analysis_type = self.analysis_var.get()
        
        # Create main container
        container = ttk.Frame(self.stats_content_frame, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)

        # Create canvas with improved scrolling
        canvas = tk.Canvas(container, highlightthickness=0, bg='white')
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        x_scrollbar = ttk.Scrollbar(self.stats_content_frame, orient="horizontal", 
                                  command=canvas.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, pady=5, ipady=10)

        canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        # Create content frame
        content_frame = ttk.Frame(canvas, style='TFrame')
        canvas.create_window((0, 0), window=content_frame, anchor="nw", tags="content_frame")

        def on_frame_configure(event):
            """Update scroll region"""
            bbox = canvas.bbox("all")
            if bbox:
                canvas.configure(scrollregion=bbox)

        content_frame.bind("<Configure>", on_frame_configure)

        # UNIVERSAL MOUSE SCROLLING (NEW IMPROVEMENT)
        def on_mousewheel(event):
            """Handle mouse wheel scrolling from anywhere in the window"""
            # Get current scroll position
            x1, y1, x2, y2 = canvas.bbox("all")
            
            # Calculate scroll amount (more sensitive)
            scroll_speed = 2 if event.state & 1 else 1  # Faster with Shift key
            
            if event.delta > 0 or event.num == 4:  # Scroll up/left
                canvas.yview_scroll(-1 * scroll_speed, "units")
            elif event.delta < 0 or event.num == 5:  # Scroll down/right
                canvas.yview_scroll(1 * scroll_speed, "units")
            
            # Horizontal scrolling with Shift
            if event.state & 1:  # Shift key pressed
                if event.delta > 0:
                    canvas.xview_scroll(-1 * scroll_speed, "units")
                else:
                    canvas.xview_scroll(1 * scroll_speed, "units")

        # Bind mouse wheel events to the ENTIRE WINDOW
        self.root.bind_all("<MouseWheel>", on_mousewheel)  # Windows/Mac
        self.root.bind_all("<Button-4>", on_mousewheel)    # Linux up
        self.root.bind_all("<Button-5>", on_mousewheel)    # Linux down
        
        # Middle-click drag scrolling
        def start_drag(event):
            if event.num == 2:  # Middle mouse button
                canvas.scan_mark(event.x, event.y)
                canvas.config(cursor="fleur")

        def drag_scroll(event):
            if canvas.cget("cursor") == "fleur":
                canvas.scan_dragto(event.x, event.y, gain=1)

        def end_drag(event):
            canvas.config(cursor="")

        canvas.bind("<ButtonPress-2>", start_drag)
        canvas.bind("<B2-Motion>", drag_scroll)
        canvas.bind("<ButtonRelease-2>", end_drag)

        # Generate the selected analysis
        if analysis_type == "Malignancy Distribution":
            self.generate_malignancy_distribution(content_frame, filtered_data)
        elif analysis_type == "Age at Diagnosis":
            self.generate_age_at_diagnosis(content_frame, filtered_data)
        elif analysis_type == "Gender Distribution":
            self.generate_gender_distribution(content_frame, filtered_data)
        elif analysis_type == "Treatment Outcomes":
            self.generate_treatment_outcomes(content_frame, filtered_data)
        elif analysis_type == "B-Symptoms Analysis":
            self.generate_b_symptoms_analysis(content_frame, filtered_data)
        elif analysis_type == "Risk Group Analysis":
            self.generate_risk_group_analysis(content_frame, filtered_data)
        elif analysis_type == "Diagnosis Timeline":
            self.generate_diagnosis_timeline(content_frame, filtered_data)
        elif analysis_type == "Survival Analysis":
            self.generate_survival_analysis(content_frame, filtered_data)

        # Initial configuration
        content_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure("content_frame", width=max(
            content_frame.winfo_reqwidth(), 
            canvas.winfo_width()
        ))
                
    def apply_statistics_filters(self):
        """Apply the current filters to the patient data"""
        filtered_data = self.patient_data.copy()
        
        # Malignancy filter
        malignancy_filter = self.stats_malignancy_filter.get()
        if malignancy_filter != "All":
            filtered_data = [p for p in filtered_data if p.get("MALIGNANCY") == malignancy_filter]
        
        # Date range filter
        start_date = self.start_date_var.get()
        end_date = self.end_date_var.get()
        
        if start_date or end_date:
            try:
                if start_date:
                    start_date = datetime.strptime(start_date, "%d/%m/%Y")
                if end_date:
                    end_date = datetime.strptime(end_date, "%d/%m/%Y")
                
                def date_in_range(patient):
                    try:
                        # Try to get diagnosis date first, fall back to created date
                        date_str = patient.get("CREATED_DATE", "")
                        if not date_str:  # Handle empty or None
                            return False
                        
                        # Extract date part (assuming format is "dd/mm/YYYY HH:MM:SS")
                        date_part = date_str.split()[0]
                        patient_date = datetime.strptime(date_part, "%d/%m/%Y")
                        
                        if start_date and end_date:
                            return start_date <= patient_date <= end_date
                        elif start_date:
                            return patient_date >= start_date
                        elif end_date:
                            return patient_date <= end_date
                        return True
                    except:
                        return False
                
                filtered_data = [p for p in filtered_data if date_in_range(p)]
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Please use dd/mm/yyyy")
                return self.patient_data
        
        return filtered_data

    def generate_malignancy_distribution(self, parent, data):
        """Generate malignancy distribution analysis with enhanced visualizations"""
        # Calculate malignancy counts
        malignancy_counts = defaultdict(int)
        for patient in data:
            malignancy = patient.get("MALIGNANCY", "Unknown")
            malignancy_counts[malignancy] += 1

        if not malignancy_counts:
            ttk.Label(parent, text="No data available for malignancy distribution", 
                     style='TLabel').pack(pady=20)
            return

        total_patients = sum(malignancy_counts.values())
        
        # Create statistics text
        stats_text = "Malignancy Distribution Analysis:\n\n"
        stats_text += f"Total Patients: {total_patients}\n\n"
        
        for malignancy, count in sorted(malignancy_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_patients) * 100
            stats_text += f"{malignancy}: {count} patients ({percentage:.1f}%)\n"
        
        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create pie chart
        fig1, ax1 = plt.subplots(figsize=(8, 6))
        labels = list(malignancy_counts.keys())
        sizes = list(malignancy_counts.values())
        colors = [MALIGNANCY_COLORS.get(m, '#999999') for m in labels]

        ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax1.axis('equal')
        ax1.set_title('Malignancy Distribution')
        fig1.tight_layout()

        # Create bar chart
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        y_pos = range(len(labels))
        
        ax2.barh(y_pos, sizes, color=colors)
        ax2.set_yticks(y_pos)
        ax2.set_yticklabels(labels)
        ax2.invert_yaxis()
        ax2.set_xlabel('Number of Patients')
        ax2.set_title('Malignancy Distribution (Count)')
        fig2.tight_layout()

        # Display the plots in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Pie chart
        pie_canvas = FigureCanvasTkAgg(fig1, master=viz_frame)
        pie_canvas.draw()
        pie_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bar chart
        bar_canvas = FigureCanvasTkAgg(fig2, master=viz_frame)
        bar_canvas.draw()
        bar_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Store figures for export/print
        self.current_figures = [fig1, fig2]

    def generate_age_at_diagnosis(self, parent, data):
        """Generate age at diagnosis analysis with distribution and statistics"""
        ages = []
        for patient in data:
            try:
                age = float(patient.get("AGE ON DIAGNOSIS", 0))
                if age > 0:
                    ages.append(age)
            except (ValueError, TypeError):
                continue

        if not ages:
            ttk.Label(parent, text="No valid age data available", 
                     style='TLabel').pack(pady=20)
            return

        # Calculate statistics
        min_age = min(ages)
        max_age = max(ages)
        mean_age = sum(ages) / len(ages)
        median_age = sorted(ages)[len(ages) // 2]
        
        # Create age groups
        age_groups = {
            "0-1 years": 0,
            "1-5 years": 0,
            "5-10 years": 0,
            "10-15 years": 0,
            "15+ years": 0
        }
        
        for age in ages:
            if age <= 1:
                age_groups["0-1 years"] += 1
            elif age <= 5:
                age_groups["1-5 years"] += 1
            elif age <= 10:
                age_groups["5-10 years"] += 1
            elif age <= 15:
                age_groups["10-15 years"] += 1
            else:
                age_groups["15+ years"] += 1

        # Create statistics text
        stats_text = "Age at Diagnosis Analysis:\n\n"
        stats_text += f"Total Patients with Age Data: {len(ages)}\n"
        stats_text += f"Minimum Age: {min_age:.1f} years\n"
        stats_text += f"Maximum Age: {max_age:.1f} years\n"
        stats_text += f"Mean Age: {mean_age:.1f} years\n"
        stats_text += f"Median Age: {median_age:.1f} years\n\n"
        stats_text += "Age Group Distribution:\n"
        
        for group, count in age_groups.items():
            percentage = (count / len(ages)) * 100
            stats_text += f"{group}: {count} patients ({percentage:.1f}%)\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create histogram
        fig1, ax1 = plt.subplots(figsize=(8, 6))
        ax1.hist(ages, bins=20, color='#3498db', edgecolor='black')
        ax1.set_xlabel('Age at Diagnosis (years)')
        ax1.set_ylabel('Number of Patients')
        ax1.set_title('Age Distribution at Diagnosis')
        fig1.tight_layout()

        # Create age group pie chart
        fig2, ax2 = plt.subplots(figsize=(8, 6))
        ax2.pie(age_groups.values(), labels=age_groups.keys(), autopct='%1.1f%%', startangle=90)
        ax2.axis('equal')
        ax2.set_title('Age Group Distribution')
        fig2.tight_layout()

        # Display the plots in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Histogram
        hist_canvas = FigureCanvasTkAgg(fig1, master=viz_frame)
        hist_canvas.draw()
        hist_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Pie chart
        pie_canvas = FigureCanvasTkAgg(fig2, master=viz_frame)
        pie_canvas.draw()
        pie_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Store figures for export/print
        self.current_figures = [fig1, fig2]

    def generate_gender_distribution(self, parent, data):
        """Generate gender distribution analysis"""
        gender_counts = defaultdict(int)
        for patient in data:
            gender = patient.get("GENDER", "Unknown")
            gender_counts[gender] += 1

        if not gender_counts:
            ttk.Label(parent, text="No gender data available", 
                     style='TLabel').pack(pady=20)
            return

        total_patients = sum(gender_counts.values())
        
        # Create statistics text
        stats_text = "Gender Distribution Analysis:\n\n"
        stats_text += f"Total Patients: {total_patients}\n\n"
        
        for gender, count in gender_counts.items():
            percentage = (count / total_patients) * 100
            stats_text += f"{gender}: {count} patients ({percentage:.1f}%)\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create pie chart
        fig1, ax1 = plt.subplots(figsize=(8, 6))
        ax1.pie(gender_counts.values(), labels=gender_counts.keys(), autopct='%1.1f%%', startangle=90)
        ax1.axis('equal')
        ax1.set_title('Gender Distribution')
        fig1.tight_layout()

        # Create bar chart
        fig2, ax2 = plt.subplots(figsize=(8, 6))
        ax2.bar(gender_counts.keys(), gender_counts.values(), color=['#3498db', '#e74c3c'])
        ax2.set_xlabel('Gender')
        ax2.set_ylabel('Number of Patients')
        ax2.set_title('Gender Distribution (Count)')
        fig2.tight_layout()

        # Display the plots in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Pie chart
        pie_canvas = FigureCanvasTkAgg(fig1, master=viz_frame)
        pie_canvas.draw()
        pie_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bar chart
        bar_canvas = FigureCanvasTkAgg(fig2, master=viz_frame)
        bar_canvas.draw()
        bar_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Store figures for export/print
        self.current_figures = [fig1, fig2]

    def generate_treatment_outcomes(self, parent, data):
        """Generate treatment outcomes analysis"""
        outcome_counts = defaultdict(int)
        for patient in data:
            outcome = patient.get("STATE", "Unknown")
            outcome_counts[outcome] += 1

        if not outcome_counts:
            ttk.Label(parent, text="No treatment outcome data available", 
                     style='TLabel').pack(pady=20)
            return

        total_patients = sum(outcome_counts.values())
        
        # Create statistics text
        stats_text = "Treatment Outcomes Analysis:\n\n"
        stats_text += f"Total Patients: {total_patients}\n\n"
        
        for outcome, count in outcome_counts.items():
            percentage = (count / total_patients) * 100
            stats_text += f"{outcome}: {count} patients ({percentage:.1f}%)\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create pie chart
        fig1, ax1 = plt.subplots(figsize=(8, 6))
        ax1.pie(outcome_counts.values(), labels=outcome_counts.keys(), autopct='%1.1f%%', startangle=90)
        ax1.axis('equal')
        ax1.set_title('Treatment Outcomes')
        fig1.tight_layout()

        # Create bar chart
        fig2, ax2 = plt.subplots(figsize=(8, 6))
        ax2.bar(outcome_counts.keys(), outcome_counts.values(), color=['#2ecc71', '#e74c3c', '#f39c12'])
        ax2.set_xlabel('Outcome')
        ax2.set_ylabel('Number of Patients')
        ax2.set_title('Treatment Outcomes (Count)')
        fig2.tight_layout()

        # Display the plots in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Pie chart
        pie_canvas = FigureCanvasTkAgg(fig1, master=viz_frame)
        pie_canvas.draw()
        pie_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bar chart
        bar_canvas = FigureCanvasTkAgg(fig2, master=viz_frame)
        bar_canvas.draw()
        bar_canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Store figures for export/print
        self.current_figures = [fig1, fig2]

    def generate_b_symptoms_analysis(self, parent, data):
        """Generate B-symptoms analysis"""
        symptom_counts = defaultdict(int)
        total_with_data = 0
        
        for patient in data:
            symptoms = patient.get("B_SYMPTOMS", "")
            if symptoms:
                total_with_data += 1
                for symptom in symptoms.split(", "):
                    symptom_counts[symptom] += 1

        if not symptom_counts:
            ttk.Label(parent, text="No B-symptoms data available", 
                     style='TLabel').pack(pady=20)
            return

        # Create statistics text
        stats_text = "B-Symptoms Analysis:\n\n"
        stats_text += f"Total Patients with B-Symptoms Data: {total_with_data}\n\n"
        
        for symptom, count in sorted(symptom_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_with_data) * 100
            stats_text += f"{symptom}: {count} patients ({percentage:.1f}%)\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create horizontal bar chart
        fig, ax = plt.subplots(figsize=(10, 6))
        
        symptoms = list(symptom_counts.keys())
        counts = list(symptom_counts.values())
        
        y_pos = range(len(symptoms))
        ax.barh(y_pos, counts, color='#3498db')
        ax.set_yticks(y_pos)
        ax.set_yticklabels(symptoms)
        ax.invert_yaxis()
        ax.set_xlabel('Number of Patients')
        ax.set_title('B-Symptoms Distribution')
        fig.tight_layout()

        # Display the plot in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        canvas = FigureCanvasTkAgg(fig, master=viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Store figure for export/print
        self.current_figures = [fig]

    def generate_risk_group_analysis(self, parent, data):
        """Generate risk group analysis"""
        risk_counts = defaultdict(int)
        total_with_data = 0
        
        for patient in data:
            risk_group = patient.get("RISK_GROUP", "")
            if risk_group:
                total_with_data += 1
                risk_counts[risk_group] += 1

        if not risk_counts:
            ttk.Label(parent, text="No risk group data available", 
                     style='TLabel').pack(pady=20)
            return

        # Create statistics text
        stats_text = "Risk Group Analysis:\n\n"
        stats_text += f"Total Patients with Risk Group Data: {total_with_data}\n\n"
        
        for group, count in sorted(risk_counts.items(), key=lambda x: x[0]):
            percentage = (count / total_with_data) * 100
            stats_text += f"{group}: {count} patients ({percentage:.1f}%)\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create pie chart
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.pie(risk_counts.values(), labels=risk_counts.keys(), autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        ax.set_title('Risk Group Distribution')
        fig.tight_layout()

        # Display the plot in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        canvas = FigureCanvasTkAgg(fig, master=viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Store figure for export/print
        self.current_figures = [fig]

    def generate_diagnosis_timeline(self, parent, data):
        """Generate diagnosis timeline analysis"""
        diagnosis_dates = []
        
        for patient in data:
            try:
                date_str = patient.get("CREATED_DATE", "")
                if not date_str:
                    continue
                
                # Extract date part (assuming format is "dd/mm/YYYY HH:MM:SS")
                date_part = date_str.split()[0]
                date = datetime.strptime(date_part, "%d/%m/%Y")
                diagnosis_dates.append(date)
            except (ValueError, IndexError):
                continue

        if not diagnosis_dates:
            ttk.Label(parent, text="No valid diagnosis date data available", 
                     style='TLabel').pack(pady=20)
            return

        # Group by month
        monthly_counts = defaultdict(int)
        for date in diagnosis_dates:
            month_year = f"{date.year}-{date.month:02d}"
            monthly_counts[month_year] += 1

        # Sort by date
        sorted_months = sorted(monthly_counts.keys())
        sorted_counts = [monthly_counts[m] for m in sorted_months]

        # Create statistics text
        stats_text = "Diagnosis Timeline Analysis:\n\n"
        stats_text += f"Total Patients with Diagnosis Dates: {len(diagnosis_dates)}\n"
        stats_text += f"First Diagnosis Date: {min(diagnosis_dates).strftime('%d/%m/%Y')}\n"
        stats_text += f"Last Diagnosis Date: {max(diagnosis_dates).strftime('%d/%m/%Y')}\n\n"
        stats_text += "Monthly Diagnosis Counts:\n"
        
        for month, count in zip(sorted_months, sorted_counts):
            stats_text += f"{month}: {count} patients\n"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create line chart
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(sorted_months, sorted_counts, marker='o', color='#3498db')
        ax.set_xlabel('Month-Year')
        ax.set_ylabel('Number of Diagnoses')
        ax.set_title('Monthly Diagnosis Counts')
        plt.xticks(rotation=45)
        fig.tight_layout()

        # Display the plot in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        canvas = FigureCanvasTkAgg(fig, master=viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Store figure for export/print
        self.current_figures = [fig]

    def generate_survival_analysis(self, parent, data):
        """Generate basic survival analysis"""
        # This is a simplified survival analysis - in a real application you would need
        # more detailed data including diagnosis date, last follow-up date, and status
        
        outcomes = defaultdict(int)
        for patient in data:
            outcome = patient.get("STATE", "Unknown")
            outcomes[outcome] += 1

        if not outcomes:
            ttk.Label(parent, text="No outcome data available for survival analysis", 
                     style='TLabel').pack(pady=20)
            return

        total_patients = sum(outcomes.values())
        alive = outcomes.get("ALIVE", 0)
        deceased = outcomes.get("DECEASED", 0)
        lost = outcomes.get("MISSED FOLLOW UP", 0)
        
        # Calculate survival rate (simplified)
        if (alive + deceased) > 0:
            survival_rate = (alive / (alive + deceased)) * 100
        else:
            survival_rate = 0

        # Create statistics text
        stats_text = "Survival Analysis (Simplified):\n\n"
        stats_text += f"Total Patients: {total_patients}\n"
        stats_text += f"Alive: {alive} patients\n"
        stats_text += f"Deceased: {deceased} patients\n"
        stats_text += f"Lost to Follow-up: {lost} patients\n"
        stats_text += f"\nEstimated Survival Rate: {survival_rate:.1f}%\n"
        stats_text += "(Based on patients with known outcomes, excluding lost to follow-up)"

        # Text summary
        text_frame = ttk.Frame(parent, style='TFrame')
        text_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(text_frame, text=stats_text, 
                 font=('Helvetica', 12), justify='left').pack(anchor='w')

        # Create visualizations frame
        viz_frame = ttk.Frame(parent, style='TFrame')
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create bar chart
        fig, ax = plt.subplots(figsize=(8, 6))
        
        categories = ["Alive", "Deceased", "Lost to Follow-up"]
        counts = [alive, deceased, lost]
        colors = ['#2ecc71', '#e74c3c', '#f39c12']
        
        ax.bar(categories, counts, color=colors)
        ax.set_xlabel('Outcome')
        ax.set_ylabel('Number of Patients')
        ax.set_title('Patient Outcomes')
        fig.tight_layout()

        # Display the plot in Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        canvas = FigureCanvasTkAgg(fig, master=viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Store figure for export/print
        self.current_figures = [fig]

    def export_statistics(self):
        """Export the current statistics to Excel with graphs"""
        if not hasattr(self, 'analysis_var'):
            messagebox.showerror("Error", "No analysis to export")
            return

        analysis_type = self.analysis_var.get()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title=f"Save {analysis_type} As"
        )
        if not file_path:
            return

        try:
            # Get filtered data
            filtered_data = self.apply_statistics_filters()
            
            # Create a Pandas Excel writer
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                # Write raw data sheet
                df = pd.DataFrame(filtered_data)
                df.to_excel(writer, sheet_name="Patient Data", index=False)
                
                # Create analysis sheet
                workbook = writer.book
                worksheet = workbook.add_worksheet(analysis_type[:31])
                
                # Add some analysis text
                bold = workbook.add_format({'bold': True})
                
                worksheet.write(0, 0, f"{analysis_type} Analysis", bold)
                worksheet.write(1, 0, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Add basic statistics
                if analysis_type == "Malignancy Distribution":
                    malignancy_counts = defaultdict(int)
                    for patient in filtered_data:
                        malignancy = patient.get("MALIGNANCY", "Unknown")
                        malignancy_counts[malignancy] += 1
                    
                    worksheet.write(3, 0, "Malignancy", bold)
                    worksheet.write(3, 1, "Count", bold)
                    worksheet.write(3, 2, "Percentage", bold)
                    
                    row = 4
                    total = sum(malignancy_counts.values())
                    for malignancy, count in sorted(malignancy_counts.items(), key=lambda x: x[1], reverse=True):
                        worksheet.write(row, 0, malignancy)
                        worksheet.write(row, 1, count)
                        worksheet.write(row, 2, f"{(count/total)*100:.1f}%")
                        row += 1
                
                # Add more analysis types as needed...
                
                # Add matplotlib figures to Excel
                if hasattr(self, 'current_figures'):
                    for i, fig in enumerate(self.current_figures, start=1):
                        # Save figure to a temporary image file
                        temp_img = f"temp_fig_{i}.png"
                        fig.savefig(temp_img, dpi=300, bbox_inches='tight')
                        
                        # Insert image into Excel
                        worksheet.insert_image(f'D{row + (i-1)*20}', temp_img)
                        
                        # Remove temporary image
                        os.remove(temp_img)
            
            messagebox.showinfo("Success", f"Statistics exported to {file_path}")
            
            # Upload to Google Drive if available
            if self.drive.initialized:
                self.executor.submit(
                    self.drive.upload_file, 
                    file_path, 
                    os.path.basename(file_path), 
                    self.drive.app_folder_id
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export statistics: {str(e)}")

    def print_statistics_report(self):
        """Print the current statistics report with graphs"""
        if not hasattr(self, 'analysis_var'):
            messagebox.showerror("Error", "No analysis to print")
            return

        try:
            # Create a Word document
            doc = Document()
            
            # Add title
            doc.add_heading(f'OncoCare Statistics Report', 0)
            doc.add_paragraph(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            doc.add_paragraph(f"Analysis Type: {self.analysis_var.get()}")
            
            # Add filters info
            filters = []
            malignancy_filter = self.stats_malignancy_filter.get()
            if malignancy_filter != "All":
                filters.append(f"Malignancy: {malignancy_filter}")
            
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            if start_date or end_date:
                date_range = f"Date Range: {start_date if start_date else 'Start'} to {end_date if end_date else 'End'}"
                filters.append(date_range)
            
            if filters:
                doc.add_paragraph("Filters Applied:")
                for f in filters:
                    doc.add_paragraph(f, style='List Bullet')
            
            # Add some basic statistics
            doc.add_heading('Summary Statistics', level=1)
            
            # Get filtered data
            filtered_data = self.apply_statistics_filters()
            doc.add_paragraph(f"Total Patients in Analysis: {len(filtered_data)}")
            
            # Add matplotlib figures to Word document
            if hasattr(self, 'current_figures'):
                for fig in self.current_figures:
                    # Save figure to a temporary image file
                    temp_img = "temp_fig.png"
                    fig.savefig(temp_img, dpi=300, bbox_inches='tight')
                    
                    # Add image to document
                    doc.add_picture(temp_img, width=Inches(6))
                    doc.add_paragraph()  # Add space after image
                    
                    # Remove temporary image
                    os.remove(temp_img)
            
            # Save to temporary file and print
            temp_file = "temp_statistics_report.docx"
            doc.save(temp_file)
            
            try:
                os.startfile(temp_file, "print")
                messagebox.showinfo("Print", "Report sent to printer")
            except Exception as e:
                messagebox.showerror("Print Error", f"Could not print document: {str(e)}")
            finally:
                # Clean up after printing
                def delete_temp_file():
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                
                # Schedule file deletion after a delay
                self.root.after(5000, delete_temp_file)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

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
