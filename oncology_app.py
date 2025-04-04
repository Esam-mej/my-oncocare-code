import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import sys
import os
import pandas as pd
from datetime import datetime, timedelta
from tkcalendar import Calendar
import bcrypt
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from collections import defaultdict
import webbrowser
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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
import hashlib
from cryptography.fernet import Fernet
import csv
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

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

# Default dropdown options that can be edited by admin
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

# Lab tests and reference ranges (age-adjusted)
LAB_TESTS = {
    "WBC": {"unit": "x10³/μL", "pediatric_range": "5.0-17.0"},
    "Hemoglobin": {"unit": "g/dL", "pediatric_range": "10.0-14.0"},
    "Platelets": {"unit": "x10³/μL", "pediatric_range": "150-450"},
    "Neutrophils": {"unit": "x10³/μL", "pediatric_range": "1.5-8.5"},
    "Lymphocytes": {"unit": "x10³/μL", "pediatric_range": "1.0-4.8"},
    "Urea": {"unit": "mg/dL", "pediatric_range": "5-18"},
    "Creatinine": {"unit": "mg/dL", "pediatric_range": "0.2-0.7"},
    "Sodium": {"unit": "mEq/L", "pediatric_range": "135-145"},
    "Potassium": {"unit": "mEq/L", "pediatric_range": "3.5-5.0"},
    "Chloride": {"unit": "mEq/L", "pediatric_range": "98-107"},
    "Calcium": {"unit": "mg/dL", "pediatric_range": "8.8-10.8"},
    "Magnesium": {"unit": "mg/dL", "pediatric_range": "1.7-2.3"},
    "pH": {"unit": "", "pediatric_range": "7.35-7.45"},
    "Uric Acid": {"unit": "mg/dL", "pediatric_range": "2.5-5.5"},
    "LDH": {"unit": "U/L", "pediatric_range": "140-280"},
    "Ferritin": {"unit": "ng/mL", "pediatric_range": "10-60"},
    "PT": {"unit": "sec", "pediatric_range": "11-13.5"},
    "APTT": {"unit": "sec", "pediatric_range": "25-35"},
    "INR": {"unit": "", "pediatric_range": "0.9-1.1"},
    "Fibrinogen": {"unit": "mg/dL", "pediatric_range": "200-400"},
    "D-Dimer": {"unit": "μg/mL", "pediatric_range": "<0.5"},
    "AST": {"unit": "U/L", "pediatric_range": "10-40"},
    "ALT": {"unit": "U/L", "pediatric_range": "7-35"},
    "ALP": {"unit": "U/L", "pediatric_range": "100-320"},
    "Total Bilirubin": {"unit": "mg/dL", "pediatric_range": "0.2-1.2"},
    "Direct Bilirubin": {"unit": "mg/dL", "pediatric_range": "0.0-0.3"},
    "EF": {"unit": "%", "pediatric_range": ">55%"}
}

# Common medications with standard doses (mg/m²)
MEDICATIONS = {
    "Vincristine": {"standard_dose": 1.5, "max_dose": 2.0},
    "Methotrexate": {"standard_dose": 5000, "max_dose": 12000},
    "Doxorubicin": {"standard_dose": 30, "max_dose": 300},
    "Cyclophosphamide": {"standard_dose": 1000, "max_dose": 4000},
    "Prednisone": {"standard_dose": 40, "max_dose": 60},
    "L-Asparaginase": {"standard_dose": 6000, "max_dose": 10000},
    "Cytarabine": {"standard_dose": 100, "max_dose": 3000},
    "Daunorubicin": {"standard_dose": 45, "max_dose": 300},
    "Etoposide": {"standard_dose": 100, "max_dose": 300},
    "Ifosfamide": {"standard_dose": 1800, "max_dose": 3600},
    "Carboplatin": {"standard_dose": 400, "max_dose": 800},
    "Cisplatin": {"standard_dose": 60, "max_dose": 100}
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
DRIVE_FOLDER_NAME = "OncoCare"
DRIVE_PATIENTS_FOLDER_NAME = "OncoCare_Patients"
CHEMO_STOCKS_URL = "https://docs.google.com/spreadsheets/d/1QxcuCM4JLxPyePbaCvB1fdQOmEMgHwAL-bqRCaMvMfM/edit?usp=sharing"

# Encryption key for sensitive data
try:
    with open('encryption_key.key', 'rb') as key_file:
        ENCRYPTION_KEY = key_file.read()
except:
    ENCRYPTION_KEY = Fernet.generate_key()
    with open('encryption_key.key', 'wb') as key_file:
        key_file.write(ENCRYPTION_KEY)

cipher_suite = Fernet(ENCRYPTION_KEY)

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
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            
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
                
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            
            self.service = build('drive', 'v3', credentials=creds)
            self.initialized = True
            self.setup_app_folders()
            
        except Exception as e:
            print(f"Google Drive initialization failed: {e}")
            self.initialized = False
    
    def setup_app_folders(self):
        """Set up the necessary folder structure in Google Drive"""
        if not self.initialized:
            return
            
        try:
            query = f"name='{DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            if items:
                self.app_folder_id = items[0]['id']
            else:
                folder_metadata = {
                    'name': DRIVE_FOLDER_NAME,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'description': 'Main folder for OncoCare application data'
                }
                folder = self.service.files().create(body=folder_metadata, fields='id').execute()
                self.app_folder_id = folder.get('id')
            
            query = f"name='{DRIVE_PATIENTS_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and '{self.app_folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            if items:
                self.patients_folder_id = items[0]['id']
            else:
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
                drive_mtime = datetime.strptime(existing_files[0].get('modifiedTime'), '%Y-%m-%dT%H:%M:%S.%fZ').timestamp()
                local_mtime = os.path.getmtime(file_path)
                
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
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and '{self.patients_folder_id}' in parents and trashed=false"
            results = self.service.files().list(q=query, fields="files(id)").execute()
            items = results.get('files', [])
            
            if items:
                return items[0]['id']
            
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
            temp_file = "temp_patient_data.json"
            with open(temp_file, 'w') as f:
                json.dump(patient_data, f)
            
            file_name = f"patient_{patient_data['FILE NUMBER']}.json"
            success = self.upload_file(temp_file, file_name, self.app_folder_id)
            
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
            query = f"'{self.app_folder_id}' in parents and name contains 'patient_' and name contains '.json' and trashed=false"
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            
            all_patients = []
            
            for item in items:
                request = self.service.files().get_media(fileId=item['id'])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                
                fh.seek(0)
                patient_data = json.load(fh)
                all_patients.append(patient_data)
            
            return all_patients
        except Exception as e:
            print(f"Error downloading patient data: {e}")
            return None
    
    def sync_patient_files(self, file_no):
        """Sync all files for a specific patient between local and Google Drive"""
        if not self.initialized:
            return False, "Google Drive not initialized"
        
        local_folder = f"Patient_{file_no}"
        if not os.path.exists(local_folder):
            os.makedirs(local_folder, exist_ok=True)
        
        drive_folder_id = self.create_patient_folder(file_no)
        if not drive_folder_id:
            return False, "Failed to create/get patient folder in Drive"
        
        try:
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
            
            upload_count = 0
            download_count = 0
            
            for rel_path, local_info in local_files.items():
                file_name = os.path.basename(rel_path)
                
                if not self.is_supported_file(file_name):
                    continue
                
                if file_name not in drive_files:
                    if self.upload_file(local_info['path'], file_name, drive_folder_id):
                        upload_count += 1
                else:
                    drive_info = drive_files[file_name]
                    if local_info['mtime'] > drive_info['mtime']:
                        if self.upload_file(local_info['path'], file_name, drive_folder_id):
                            upload_count += 1
            
            for file_name, drive_info in drive_files.items():
                local_path = os.path.join(local_folder, file_name)
                
                if not os.path.exists(local_path):
                    if self.download_file(drive_info['id'], local_path):
                        download_count += 1
                else:
                    local_mtime = os.path.getmtime(local_path)
                    if drive_info['mtime'] > local_mtime:
                        if self.download_file(drive_info['id'], local_path):
                            download_count += 1
            
            return True, f"Sync complete. Uploaded: {upload_count}, Downloaded: {download_count}"
        except Exception as e:
            print(f"Error syncing patient files: {e}")
            return False, f"Sync failed: {str(e)}"
    
    def is_supported_file(self, filename):
        """Check if file extension is supported for sync"""
        ext = os.path.splitext(filename)[1].lower()
        supported_extensions = [
            '.doc', '.docx', '.txt', '.pdf', '.rtf', '.xls', '.xlsx', '.xlsm', '.csv',
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.json', '.xml',
            '.ppt', '.pptx', '.zip', '.rar'
        ]
        return ext in supported_extensions
    
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
            firebase_data = []
            docs = self.db.collection('patients').stream()
            for doc in docs:
                patient = doc.to_dict()
                patient['_firestore_id'] = doc.id
                firebase_data.append(patient)
            
            merged_data = self.merge_data(local_data, firebase_data)
            
            batch = self.db.batch()
            patients_ref = self.db.collection('patients')
            
            for patient in merged_data:
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
                merged.append(local_patient)
            elif firebase_patient and not local_patient:
                merged.append(firebase_patient)
            else:
                try:
                    local_date = datetime.strptime(local_patient['LAST_MODIFIED_DATE'], '%d/%m/%Y %H:%M:%S')
                    firebase_date = datetime.strptime(firebase_patient['LAST_MODIFIED_DATE'], '%d/%m/%Y %H:%M:%S')
                    
                    if local_date > firebase_date:
                        local_patient['_firestore_id'] = firebase_patient.get('_firestore_id', file_no)
                        merged.append(local_patient)
                    else:
                        merged.append(firebase_patient)
                except:
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
                patient['_firestore_id'] = doc.id
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
        
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # Initialize managers
        self.firebase = FirebaseManager()
        self.drive = GoogleDriveManager()
        
        # Initialize variables
        self.current_user = None
        self.current_results = []
        self.current_result_index = 0
        self.last_sync_time = None
        self.internet_connected = False
        self.sync_in_progress = False
        self.fn_doc_path = None  # Path for F&N documentation executable
        self.load_settings()
        
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
        self.executor = ThreadPoolExecutor(max_workers=4)
        
        # Handle window close properly
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def setup_styles(self):
        """Configure the visual styles for the application"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors
        self.primary_color = '#2c3e50'
        self.secondary_color = '#3498db'
        self.accent_color = '#e74c3c'
        self.light_color = '#ecf0f1'
        self.dark_color = '#2c3e50'
        self.success_color = '#27ae60'
        self.warning_color = '#f39c12'
        self.danger_color = '#e74c3c'
        
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
    
    def load_settings(self):
        """Load application settings from encrypted file"""
        try:
            if os.path.exists('settings.enc'):
                with open('settings.enc', 'rb') as f:
                    encrypted_data = f.read()
                    decrypted_data = cipher_suite.decrypt(encrypted_data)
                    settings = json.loads(decrypted_data.decode())
                    self.fn_doc_path = settings.get('fn_doc_path')
        except Exception as e:
            print(f"Error loading settings: {e}")
    
    def save_settings(self):
        """Save application settings to encrypted file"""
        try:
            settings = {
                'fn_doc_path': self.fn_doc_path,
                'dropdown_options': DROPDOWN_OPTIONS
            }
            encrypted_data = cipher_suite.encrypt(json.dumps(settings).encode())
            with open('settings.enc', 'wb') as f:
                f.write(encrypted_data)
        except Exception as e:
            print(f"Error saving settings: {e}")
    
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
            self.root.after(30000, internet_check)
        
        internet_check()
    
    def load_users(self):
        """Load user data from file or create default users"""
        if os.path.exists("users_data.json"):
            try:
                with open("users_data.json", "r") as f:
                    self.users = json.load(f)
            except json.JSONDecodeError:
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
            ttk.Frame(row1_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        ttk.Button(row1_frame, text="Search Patient", command=self.search_patient,
                  style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.users[self.current_user]["role"] == "admin":
            ttk.Button(row1_frame, text="View All Patients", command=self.view_all_patients,
                      style='Blue.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
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
            ttk.Frame(row3_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(row3_frame, text="Statistics", command=self.show_statistics,
                      style='Yellow.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
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
            ttk.Frame(row4_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Fifth row - Special Features (Purple buttons)
        row5_frame = ttk.Frame(btn_frame, style='TFrame')
        row5_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(row5_frame, text="F&N Documentation", command=self.run_fn_documentation,
                  style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.current_user == "mej.esam":
            ttk.Button(row5_frame, text="Settings", command=self.open_settings,
                      style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            ttk.Frame(row5_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Sixth row - Logout (Red button)
        row6_frame = ttk.Frame(btn_frame, style='TFrame')
        row6_frame.pack(fill=tk.X, pady=(20, 5))
        
        ttk.Button(row6_frame, text="Logout", command=self.setup_login_screen,
                  style='Red.TButton').pack(fill=tk.X, ipady=10)

        # Signature
        signature_frame = ttk.Frame(form_container, style='TFrame')
        signature_frame.pack(pady=(30, 0))

        ttk.Label(signature_frame, text="Made by: DR. ESAM MEJRAB",
                 font=('Times New Roman', 14, 'italic'), 
                 foreground=self.primary_color).pack()
    
    def run_fn_documentation(self):
        """Run the F&N documentation application if path is set"""
        if not self.fn_doc_path:
            if self.current_user == "mej.esam":
                self.set_fn_documentation_path()
            else:
                messagebox.showerror("Error", "F&N Documentation path not set. Please contact admin.")
            return
        
        try:
            subprocess.Popen(self.fn_doc_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not run F&N Documentation: {e}")
    
    def set_fn_documentation_path(self):
        """Set the path for F&N documentation executable"""
        path = filedialog.askopenfilename(title="Select F&N Documentation Executable", 
                                        filetypes=[("Executable files", "*.exe")])
        if path:
            self.fn_doc_path = path
            self.save_settings()
            messagebox.showinfo("Success", "F&N Documentation path set successfully!")
    
    def open_settings(self):
        """Open settings window (only for mej.esam)"""
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only mej.esam can access settings.")
            return
        
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Application Settings")
        settings_window.geometry("600x400")
        
        # Main container
        main_frame = ttk.Frame(settings_window, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab control
        tab_control = ttk.Notebook(main_frame)
        tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Dropdown Options Tab
        dropdown_tab = ttk.Frame(tab_control)
        tab_control.add(dropdown_tab, text="Dropdown Options")
        
        # List of dropdown categories
        dropdown_categories = list(DROPDOWN_OPTIONS.keys())
        self.dropdown_category_var = tk.StringVar(value=dropdown_categories[0])
        
        # Category selection
        category_frame = ttk.Frame(dropdown_tab, style='TFrame')
        category_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(category_frame, text="Category:").pack(side=tk.LEFT, padx=5)
        category_combo = ttk.Combobox(category_frame, textvariable=self.dropdown_category_var,
                                     values=dropdown_categories, state="readonly")
        category_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        category_combo.bind("<<ComboboxSelected>>", self.update_dropdown_options_view)
        
        # Options listbox with scrollbar
        listbox_frame = ttk.Frame(dropdown_tab, style='TFrame')
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.options_listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE, height=10)
        self.options_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.options_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.options_listbox.config(yscrollcommand=scrollbar.set)
        
        # Add/Remove options
        option_edit_frame = ttk.Frame(dropdown_tab, style='TFrame')
        option_edit_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.new_option_entry = ttk.Entry(option_edit_frame)
        self.new_option_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(option_edit_frame, text="Add", command=self.add_dropdown_option,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(option_edit_frame, text="Remove", command=self.remove_dropdown_option,
                  style='Red.TButton').pack(side=tk.LEFT, padx=5)
        
        # System Settings Tab
        system_tab = ttk.Frame(tab_control)
        tab_control.add(system_tab, text="System Settings")
        
        # F&N Documentation Path
        fn_frame = ttk.Frame(system_tab, style='TFrame')
        fn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(fn_frame, text="F&N Documentation Path:").pack(side=tk.LEFT)
        self.fn_path_var = tk.StringVar(value=self.fn_doc_path or "Not set")
        ttk.Entry(fn_frame, textvariable=self.fn_path_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(fn_frame, text="Change", command=self.set_fn_documentation_path,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        # Data Management Tab
        data_tab = ttk.Frame(tab_control)
        tab_control.add(data_tab, text="Data Management")
        
        # Delete all local patient folders
        ttk.Button(data_tab, text="Delete All Local Patient Folders", 
                  command=self.delete_all_patient_folders,
                  style='Red.TButton').pack(fill=tk.X, padx=5, pady=5)
        
        # Button frame
        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings_from_ui,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Close", command=settings_window.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Initialize the options view
        self.update_dropdown_options_view()
    
    def update_dropdown_options_view(self, event=None):
        """Update the listbox with options for the selected category"""
        category = self.dropdown_category_var.get()
        self.options_listbox.delete(0, tk.END)
        
        for option in DROPDOWN_OPTIONS.get(category, []):
            self.options_listbox.insert(tk.END, option)
    
    def add_dropdown_option(self):
        """Add a new option to the selected dropdown category"""
        category = self.dropdown_category_var.get()
        new_option = self.new_option_entry.get().strip()
        
        if not new_option:
            messagebox.showerror("Error", "Please enter an option to add")
            return
        
        if new_option in DROPDOWN_OPTIONS[category]:
            messagebox.showerror("Error", "This option already exists")
            return
        
        DROPDOWN_OPTIONS[category].append(new_option)
        self.update_dropdown_options_view()
        self.new_option_entry.delete(0, tk.END)
        messagebox.showinfo("Success", "Option added successfully")
    
    def remove_dropdown_option(self):
        """Remove the selected option from the dropdown category"""
        category = self.dropdown_category_var.get()
        selection = self.options_listbox.curselection()
        
        if not selection:
            messagebox.showerror("Error", "Please select an option to remove")
            return
        
        option_to_remove = self.options_listbox.get(selection[0])
        DROPDOWN_OPTIONS[category].remove(option_to_remove)
        self.update_dropdown_options_view()
        messagebox.showinfo("Success", "Option removed successfully")
    
    def save_settings_from_ui(self):
        """Save settings from the UI"""
        self.save_settings()
        messagebox.showinfo("Success", "Settings saved successfully!")
    
    def delete_all_patient_folders(self):
        """Delete all local patient folders (admin only)"""
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only mej.esam can perform this action.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", 
                                    "This will delete ALL local patient folders and their contents.\n"
                                    "This action cannot be undone!\n\n"
                                    "Are you sure you want to continue?")
        if not confirm:
            return
        
        try:
            deleted_count = 0
            for folder in os.listdir('.'):
                if folder.startswith('Patient_'):
                    shutil.rmtree(folder)
                    deleted_count += 1
            
            messagebox.showinfo("Success", f"Deleted {deleted_count} patient folders")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete folders: {e}")
    
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
            file_no = patient_data["FILE NUMBER"]
            self.drive.upload_patient_data(patient_data)
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
                patient = next((p for p in self.patient_data if p.get("FILE NUMBER") == file_no), None)
                if patient:
                    for field in COMMON_FIELDS:
                        if field in patient:
                            doc.add_paragraph(f"{field}: {patient[field]}", style='List Bullet')
            
            # Save document
            doc.save(doc_path)
        
        if not create_only:
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
        
        # Add Medication Management button
        ttk.Button(btn_frame, text="Medication Management", 
                  command=lambda: self.open_medication_management(patient_data["FILE NUMBER"]),
                  style='Green.TButton').pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Cancel", 
                  command=lambda: self.view_patient(patient_data)).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="Main Menu", command=self.main_menu,
                  style='Blue.TButton').pack(side=tk.RIGHT, padx=10)

    def open_medication_management(self, file_no):
        """Open medication management window for a patient"""
        self.medication_window = tk.Toplevel(self.root)
        self.medication_window.title(f"Medication Management - Patient {file_no}")
        self.medication_window.geometry("1000x700")
        
        # Get patient data
        patient = next((p for p in self.patient_data if p.get("FILE NUMBER") == file_no), None)
        if not patient:
            messagebox.showerror("Error", "Patient not found")
            self.medication_window.destroy()
            return
        
        # Main container
        main_frame = ttk.Frame(self.medication_window, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Current Medications tab
        current_meds_frame = ttk.Frame(notebook, style='TFrame')
        notebook.add(current_meds_frame, text="Current Medications")
        self.setup_current_medications_tab(current_meds_frame, patient)
        
        # Medication History tab
        history_frame = ttk.Frame(notebook, style='TFrame')
        notebook.add(history_frame, text="Medication History")
        self.setup_medication_history_tab(history_frame, patient)
        
        # BSA Calculator tab
        bsa_frame = ttk.Frame(notebook, style='TFrame')
        notebook.add(bsa_frame, text="BSA Calculator")
        self.setup_bsa_calculator_tab(bsa_frame, patient)
        
        # Lab Results tab
        lab_frame = ttk.Frame(notebook, style='TFrame')
        notebook.add(lab_frame, text="Lab Results")
        self.setup_lab_results_tab(lab_frame, patient)
        
        # Close button
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Close", command=self.medication_window.destroy,
                  style='Blue.TButton').pack()

    def setup_current_medications_tab(self, parent, patient):
        """Setup the current medications tab"""
        # Load medication data or initialize if not exists
        if "MEDICATIONS" not in patient:
            patient["MEDICATIONS"] = {"current": [], "history": []}
            self.save_patient_data()
        
        # Create scrollable frame
        container = ttk.Frame(parent, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Header
        ttk.Label(scrollable_frame, text="Current Medications", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Add new medication button
        ttk.Button(scrollable_frame, text="Add New Medication", 
                  command=lambda: self.add_new_medication(patient),
                  style='Green.TButton').pack(pady=10)
        
        # Display current medications
        if not patient["MEDICATIONS"]["current"]:
            ttk.Label(scrollable_frame, text="No current medications", 
                     style='TLabel').pack(pady=20)
        else:
            for med in patient["MEDICATIONS"]["current"]:
                self.create_medication_card(scrollable_frame, med, patient)
    
    def create_medication_card(self, parent, medication, patient):
        """Create a medication card display"""
        card_frame = ttk.Frame(parent, style='Card.TFrame', padding=10)
        card_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Medication name and details
        ttk.Label(card_frame, text=medication["name"], 
                 font=('Helvetica', 12, 'bold')).pack(anchor="w")
        
        details_frame = ttk.Frame(card_frame, style='TFrame')
        details_frame.pack(fill=tk.X, pady=5)
        
        # Left side - details
        left_frame = ttk.Frame(details_frame, style='TFrame')
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(left_frame, text=f"Dose: {medication['dose']} {medication['unit']}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Frequency: {medication['frequency']}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Route: {medication['route']}").pack(anchor="w")
        
        # Right side - dates and actions
        right_frame = ttk.Frame(details_frame, style='TFrame')
        right_frame.pack(side=tk.RIGHT, fill=tk.X)
        
        ttk.Label(right_frame, text=f"Start: {medication['start_date']}").pack(anchor="e")
        if 'end_date' in medication:
            ttk.Label(right_frame, text=f"End: {medication['end_date']}").pack(anchor="e")
        
        # Action buttons
        btn_frame = ttk.Frame(card_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Discontinue", 
                  command=lambda m=medication: self.discontinue_medication(m, patient),
                  style='Red.TButton').pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame, text="Edit", 
                  command=lambda m=medication: self.edit_medication(m, patient),
                  style='Blue.TButton').pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame, text="Administer", 
                  command=lambda m=medication: self.record_administration(m, patient),
                  style='Green.TButton').pack(side=tk.LEFT, padx=2)
    
    def add_new_medication(self, patient):
        """Open dialog to add new medication"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Medication")
        dialog.geometry("600x500")
        
        # Form frame
        form_frame = ttk.Frame(dialog, style='TFrame', padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Medication name
        ttk.Label(form_frame, text="Medication Name:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = ttk.Entry(form_frame)
        name_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        # Dose
        ttk.Label(form_frame, text="Dose:").grid(row=1, column=0, sticky="w", pady=5)
        dose_frame = ttk.Frame(form_frame)
        dose_frame.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        dose_entry = ttk.Entry(dose_frame)
        dose_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        unit_var = tk.StringVar(value="mg")
        unit_combo = ttk.Combobox(dose_frame, textvariable=unit_var, 
                                values=["mg", "mcg", "g", "mg/m²", "mcg/m²", "IU", "mL"])
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        # Frequency
        ttk.Label(form_frame, text="Frequency:").grid(row=2, column=0, sticky="w", pady=5)
        freq_var = tk.StringVar(value="Daily")
        freq_combo = ttk.Combobox(form_frame, textvariable=freq_var,
                                values=["Daily", "BID", "TID", "QID", "QHS", "QOD", 
                                       "Weekly", "Monthly", "Other"])
        freq_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        
        # Route
        ttk.Label(form_frame, text="Route:").grid(row=3, column=0, sticky="w", pady=5)
        route_var = tk.StringVar(value="PO")
        route_combo = ttk.Combobox(form_frame, textvariable=route_var,
                                 values=["PO", "IV", "IM", "SC", "Topical", "PR", "Other"])
        route_combo.grid(row=3, column=1, sticky="ew", pady=5, padx=5)
        
        # Start date
        ttk.Label(form_frame, text="Start Date:").grid(row=4, column=0, sticky="w", pady=5)
        start_frame = ttk.Frame(form_frame)
        start_frame.grid(row=4, column=1, sticky="ew", pady=5, padx=5)
        
        start_entry = ttk.Entry(start_frame)
        start_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        start_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        ttk.Button(start_frame, text="📅", width=3,
                  command=lambda e=start_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # Indication
        ttk.Label(form_frame, text="Indication:").grid(row=5, column=0, sticky="w", pady=5)
        indication_entry = ttk.Entry(form_frame)
        indication_entry.grid(row=5, column=1, sticky="ew", pady=5, padx=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").grid(row=6, column=0, sticky="w", pady=5)
        notes_entry = tk.Text(form_frame, height=4, wrap=tk.WORD)
        notes_entry.grid(row=6, column=1, sticky="ew", pady=5, padx=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=10)
        
        def save_medication():
            new_med = {
                "name": name_entry.get(),
                "dose": dose_entry.get(),
                "unit": unit_var.get(),
                "frequency": freq_var.get(),
                "route": route_var.get(),
                "start_date": start_entry.get(),
                "indication": indication_entry.get(),
                "notes": notes_entry.get("1.0", tk.END).strip(),
                "administrations": []
            }
            
            if not new_med["name"]:
                messagebox.showerror("Error", "Medication name is required")
                return
                
            patient["MEDICATIONS"]["current"].append(new_med)
            self.save_patient_data()
            
            # Update the medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_medication,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def discontinue_medication(self, medication, patient):
        """Discontinue a medication"""
        confirm = messagebox.askyesno("Confirm", 
                                     f"Discontinue {medication['name']}?")
        if not confirm:
            return
            
        # Set end date
        medication["end_date"] = datetime.now().strftime("%d/%m/%Y")
        
        # Move to history
        patient["MEDICATIONS"]["history"].append(medication)
        patient["MEDICATIONS"]["current"].remove(medication)
        
        self.save_patient_data()
        
        # Refresh the medication window
        if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
            self.medication_window.destroy()
            self.open_medication_management(patient["FILE NUMBER"])
    
    def edit_medication(self, medication, patient):
        """Edit an existing medication"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit {medication['name']}")
        dialog.geometry("600x500")
        
        # Form frame
        form_frame = ttk.Frame(dialog, style='TFrame', padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Medication name
        ttk.Label(form_frame, text="Medication Name:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = ttk.Entry(form_frame)
        name_entry.insert(0, medication["name"])
        name_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        # Dose
        ttk.Label(form_frame, text="Dose:").grid(row=1, column=0, sticky="w", pady=5)
        dose_frame = ttk.Frame(form_frame)
        dose_frame.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        dose_entry = ttk.Entry(dose_frame)
        dose_entry.insert(0, medication["dose"])
        dose_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        unit_var = tk.StringVar(value=medication.get("unit", "mg"))
        unit_combo = ttk.Combobox(dose_frame, textvariable=unit_var, 
                                values=["mg", "mcg", "g", "mg/m²", "mcg/m²", "IU", "mL"])
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        # Frequency
        ttk.Label(form_frame, text="Frequency:").grid(row=2, column=0, sticky="w", pady=5)
        freq_var = tk.StringVar(value=medication.get("frequency", "Daily"))
        freq_combo = ttk.Combobox(form_frame, textvariable=freq_var,
                                values=["Daily", "BID", "TID", "QID", "QHS", "QOD", 
                                       "Weekly", "Monthly", "Other"])
        freq_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        
        # Route
        ttk.Label(form_frame, text="Route:").grid(row=3, column=0, sticky="w", pady=5)
        route_var = tk.StringVar(value=medication.get("route", "PO"))
        route_combo = ttk.Combobox(form_frame, textvariable=route_var,
                                 values=["PO", "IV", "IM", "SC", "Topical", "PR", "Other"])
        route_combo.grid(row=3, column=1, sticky="ew", pady=5, padx=5)
        
        # Start date
        ttk.Label(form_frame, text="Start Date:").grid(row=4, column=0, sticky="w", pady=5)
        start_frame = ttk.Frame(form_frame)
        start_frame.grid(row=4, column=1, sticky="ew", pady=5, padx=5)
        
        start_entry = ttk.Entry(start_frame)
        start_entry.insert(0, medication.get("start_date", datetime.now().strftime("%d/%m/%Y")))
        start_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(start_frame, text="📅", width=3,
                  command=lambda e=start_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # End date (if exists)
        if "end_date" in medication:
            ttk.Label(form_frame, text="End Date:").grid(row=5, column=0, sticky="w", pady=5)
            end_frame = ttk.Frame(form_frame)
            end_frame.grid(row=5, column=1, sticky="ew", pady=5, padx=5)
            
            end_entry = ttk.Entry(end_frame)
            end_entry.insert(0, medication["end_date"])
            end_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            ttk.Button(end_frame, text="📅", width=3,
                      command=lambda e=end_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # Indication
        ttk.Label(form_frame, text="Indication:").grid(row=6, column=0, sticky="w", pady=5)
        indication_entry = ttk.Entry(form_frame)
        indication_entry.insert(0, medication.get("indication", ""))
        indication_entry.grid(row=6, column=1, sticky="ew", pady=5, padx=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").grid(row=7, column=0, sticky="w", pady=5)
        notes_entry = tk.Text(form_frame, height=4, wrap=tk.WORD)
        notes_entry.insert("1.0", medication.get("notes", ""))
        notes_entry.grid(row=7, column=1, sticky="ew", pady=5, padx=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=8, column=0, columnspan=2, pady=10)
        
        def save_changes():
            # Update medication details
            medication["name"] = name_entry.get()
            medication["dose"] = dose_entry.get()
            medication["unit"] = unit_var.get()
            medication["frequency"] = freq_var.get()
            medication["route"] = route_var.get()
            medication["start_date"] = start_entry.get()
            if "end_date" in medication:
                medication["end_date"] = end_entry.get()
            medication["indication"] = indication_entry.get()
            medication["notes"] = notes_entry.get("1.0", tk.END).strip()
            
            self.save_patient_data()
            
            # Update the medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_changes,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def record_administration(self, medication, patient):
        """Record administration of a medication"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Record Administration - {medication['name']}")
        dialog.geometry("500x400")
        
        # Form frame
        form_frame = ttk.Frame(dialog, style='TFrame', padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Date
        ttk.Label(form_frame, text="Administration Date:").pack(pady=5)
        date_frame = ttk.Frame(form_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        date_entry = ttk.Entry(date_frame)
        date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(date_frame, text="📅", width=3,
                  command=lambda e=date_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # Time
        ttk.Label(form_frame, text="Time:").pack(pady=5)
        time_entry = ttk.Entry(form_frame)
        time_entry.insert(0, datetime.now().strftime("%H:%M"))
        time_entry.pack(fill=tk.X, pady=5)
        
        # Administered by
        ttk.Label(form_frame, text="Administered by:").pack(pady=5)
        admin_by_entry = ttk.Entry(form_frame)
        admin_by_entry.insert(0, self.current_user)
        admin_by_entry.pack(fill=tk.X, pady=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_entry = tk.Text(form_frame, height=5, wrap=tk.WORD)
        notes_entry.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def record_admin():
            admin_record = {
                "date": date_entry.get(),
                "time": time_entry.get(),
                "administered_by": admin_by_entry.get(),
                "notes": notes_entry.get("1.0", tk.END).strip()
            }
            
            if "administrations" not in medication:
                medication["administrations"] = []
                
            medication["administrations"].append(admin_record)
            self.save_patient_data()
            
            messagebox.showinfo("Success", "Administration recorded")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Record", command=record_admin,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def setup_medication_history_tab(self, parent, patient):
        """Setup the medication history tab"""
        # Create scrollable frame
        container = ttk.Frame(parent, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Header
        ttk.Label(scrollable_frame, text="Medication History", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Display history
        if not patient["MEDICATIONS"]["history"]:
            ttk.Label(scrollable_frame, text="No medication history", 
                     style='TLabel').pack(pady=20)
        else:
            for med in patient["MEDICATIONS"]["history"]:
                self.create_history_medication_card(scrollable_frame, med, patient)
    
    def create_history_medication_card(self, parent, medication, patient):
        """Create a medication history card display"""
        card_frame = ttk.Frame(parent, style='Card.TFrame', padding=10)
        card_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Medication name and details
        ttk.Label(card_frame, text=medication["name"], 
                 font=('Helvetica', 12, 'bold')).pack(anchor="w")
        
        details_frame = ttk.Frame(card_frame, style='TFrame')
        details_frame.pack(fill=tk.X, pady=5)
        
        # Left side - details
        left_frame = ttk.Frame(details_frame, style='TFrame')
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(left_frame, text=f"Dose: {medication['dose']} {medication['unit']}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Frequency: {medication['frequency']}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Route: {medication['route']}").pack(anchor="w")
        
        # Right side - dates
        right_frame = ttk.Frame(details_frame, style='TFrame')
        right_frame.pack(side=tk.RIGHT, fill=tk.X)
        
        ttk.Label(right_frame, text=f"Start: {medication['start_date']}").pack(anchor="e")
        ttk.Label(right_frame, text=f"End: {medication['end_date']}").pack(anchor="e")
        
        # Administration history
        if "administrations" in medication and medication["administrations"]:
            admin_frame = ttk.Frame(card_frame, style='TFrame')
            admin_frame.pack(fill=tk.X, pady=(5, 0))
            
            ttk.Label(admin_frame, text="Administrations:", 
                     font=('Helvetica', 10, 'bold')).pack(anchor="w")
            
            for admin in medication["administrations"]:
                admin_label = ttk.Label(admin_frame, 
                                      text=f"{admin['date']} {admin['time']} - {admin['administered_by']}")
                admin_label.pack(anchor="w", padx=10)
                
                if admin["notes"]:
                    notes_label = ttk.Label(admin_frame, text=f"Notes: {admin['notes']}",
                                          font=('Helvetica', 9))
                    notes_label.pack(anchor="w", padx=20)
        
        # Reinstate button
        btn_frame = ttk.Frame(card_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Reinstate", 
                  command=lambda m=medication: self.reinstate_medication(m, patient),
                  style='Green.TButton').pack(side=tk.LEFT, padx=2)
    
    def reinstate_medication(self, medication, patient):
        """Reinstate a discontinued medication"""
        confirm = messagebox.askyesno("Confirm", 
                                    f"Reinstate {medication['name']}?")
        if not confirm:
            return
            
        # Remove end date
        if "end_date" in medication:
            del medication["end_date"]
            
        # Move back to current medications
        patient["MEDICATIONS"]["current"].append(medication)
        patient["MEDICATIONS"]["history"].remove(medication)
        
        self.save_patient_data()
        
        # Refresh the medication window
        if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
            self.medication_window.destroy()
            self.open_medication_management(patient["FILE NUMBER"])
    
    def setup_bsa_calculator_tab(self, parent, patient):
        """Setup the BSA calculator tab"""
        # Main frame
        main_frame = ttk.Frame(parent, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        ttk.Label(main_frame, text="Body Surface Area Calculator", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Input frame
        input_frame = ttk.Frame(main_frame, style='TFrame')
        input_frame.pack(fill=tk.X, pady=10)
        
        # Height
        ttk.Label(input_frame, text="Height (cm):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.height_entry = ttk.Entry(input_frame)
        self.height_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # Weight
        ttk.Label(input_frame, text="Weight (kg):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.weight_entry = ttk.Entry(input_frame)
        self.weight_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Formula selection
        ttk.Label(input_frame, text="Formula:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.formula_var = tk.StringVar(value="Mosteller")
        formula_combo = ttk.Combobox(input_frame, textvariable=self.formula_var,
                                   values=["Mosteller", "DuBois", "Haycock"])
        formula_combo.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Calculate button
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Calculate BSA", 
                  command=self.calculate_bsa,
                  style='Blue.TButton').pack()
        
        # Results frame
        self.results_frame = ttk.Frame(main_frame, style='TFrame')
        self.results_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Historical BSA records
        if "BSA_HISTORY" in patient:
            history_frame = ttk.LabelFrame(main_frame, text="BSA History", style='TFrame')
            history_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            for record in patient["BSA_HISTORY"]:
                record_frame = ttk.Frame(history_frame, style='TFrame')
                record_frame.pack(fill=tk.X, pady=2)
                
                ttk.Label(record_frame, 
                         text=f"{record['date']}: {record['bsa']:.2f} m² (H: {record['height']} cm, W: {record['weight']} kg)").pack(anchor="w")
    
    def calculate_bsa(self):
        """Calculate body surface area"""
        try:
            height = float(self.height_entry.get())
            weight = float(self.weight_entry.get())
            
            if height <= 0 or weight <= 0:
                raise ValueError("Height and weight must be positive numbers")
                
            formula = self.formula_var.get()
            
            if formula == "Mosteller":
                bsa = math.sqrt(height * weight / 3600)
            elif formula == "DuBois":
                bsa = 0.007184 * math.pow(height, 0.725) * math.pow(weight, 0.425)
            elif formula == "Haycock":
                bsa = 0.024265 * math.pow(height, 0.3964) * math.pow(weight, 0.5378)
            else:
                bsa = math.sqrt(height * weight / 3600)  # Default to Mosteller
            
            # Clear previous results
            for widget in self.results_frame.winfo_children():
                widget.destroy()
            
            # Display results
            ttk.Label(self.results_frame, 
                     text=f"BSA: {bsa:.2f} m² ({formula} formula)",
                     font=('Helvetica', 12, 'bold')).pack(pady=10)
            
            # Save button
            ttk.Button(self.results_frame, text="Save to Patient Record",
                      command=lambda: self.save_bsa_result(height, weight, bsa, formula),
                      style='Green.TButton').pack(pady=10)
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")
    
    def save_bsa_result(self, height, weight, bsa, formula):
        """Save BSA result to patient record"""
        # Get current patient from medication window title
        title = self.medication_window.title()
        file_no = title.split("Patient ")[1].split(")")[0]
        
        patient = next((p for p in self.patient_data if p.get("FILE NUMBER") == file_no), None)
        if not patient:
            messagebox.showerror("Error", "Patient not found")
            return
            
        if "BSA_HISTORY" not in patient:
            patient["BSA_HISTORY"] = []
            
        patient["BSA_HISTORY"].append({
            "date": datetime.now().strftime("%d/%m/%Y"),
            "height": height,
            "weight": weight,
            "bsa": bsa,
            "formula": formula
        })
        
        self.save_patient_data()
        messagebox.showinfo("Success", "BSA result saved to patient record")
        
        # Refresh the BSA tab
        if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
            self.medication_window.destroy()
            self.open_medication_management(file_no)
    
    def setup_lab_results_tab(self, parent, patient):
        """Setup the lab results tracking tab"""
        # Main frame
        main_frame = ttk.Frame(parent, style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        ttk.Label(main_frame, text="Lab Results Tracking", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Add new lab button
        ttk.Button(main_frame, text="Add New Lab Result", 
                  command=lambda: self.add_lab_result(patient),
                  style='Blue.TButton').pack(pady=10)
        
        # Lab results display
        if "LAB_RESULTS" not in patient or not patient["LAB_RESULTS"]:
            ttk.Label(main_frame, text="No lab results recorded", 
                     style='TLabel').pack(pady=20)
            return
            
        # Create scrollable frame
        container = ttk.Frame(main_frame, style='TFrame')
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Group labs by date
        labs_by_date = defaultdict(list)
        for lab in patient["LAB_RESULTS"]:
            labs_by_date[lab["date"]].append(lab)
            
        # Display labs sorted by date (newest first)
        for date in sorted(labs_by_date.keys(), reverse=True):
            date_frame = ttk.LabelFrame(scrollable_frame, text=date, style='TFrame')
            date_frame.pack(fill=tk.X, padx=5, pady=5)
            
            for lab in labs_by_date[date]:
                self.create_lab_result_display(date_frame, lab, patient)
    
    def create_lab_result_display(self, parent, lab, patient):
        """Create a display for a lab result"""
        lab_frame = ttk.Frame(parent, style='TFrame')
        lab_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Test name and value
        ttk.Label(lab_frame, text=f"{lab['test']}: {lab['value']} {lab.get('unit', '')}",
                 font=('Helvetica', 10)).pack(side=tk.LEFT)
        
        # Abnormal flag
        if lab.get("abnormal", False):
            ttk.Label(lab_frame, text="(Abnormal)", 
                     foreground="red").pack(side=tk.LEFT, padx=5)
        
        # Edit button (for admins/editors)
        if self.users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(lab_frame, text="Edit", 
                      command=lambda l=lab: self.edit_lab_result(l, patient),
                      style='Blue.TButton').pack(side=tk.RIGHT, padx=2)
    
    def add_lab_result(self, patient):
        """Add a new lab result"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Lab Result")
        dialog.geometry("500x400")
        
        # Form frame
        form_frame = ttk.Frame(dialog, style='TFrame', padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Test name
        ttk.Label(form_frame, text="Test Name:").pack(pady=5)
        test_combo = ttk.Combobox(form_frame, 
                                values=["WBC", "Hemoglobin", "Platelets", "Neutrophils", 
                                       "Lymphocytes", "Urea", "Creatinine", "Sodium", 
                                       "Potassium", "Chloride", "Calcium", "Magnesium",
                                       "AST", "ALT", "Total Bilirubin", "Direct Bilirubin",
                                       "ALP", "LDH", "Ferritin", "PT", "APTT", "INR",
                                       "Fibrinogen", "D-Dimer", "EF%"])
        test_combo.pack(fill=tk.X, pady=5)
        
        # Value
        ttk.Label(form_frame, text="Value:").pack(pady=5)
        value_entry = ttk.Entry(form_frame)
        value_entry.pack(fill=tk.X, pady=5)
        
        # Unit
        ttk.Label(form_frame, text="Unit:").pack(pady=5)
        unit_entry = ttk.Entry(form_frame)
        unit_entry.pack(fill=tk.X, pady=5)
        
        # Date
        ttk.Label(form_frame, text="Date:").pack(pady=5)
        date_frame = ttk.Frame(form_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        date_entry = ttk.Entry(date_frame)
        date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(date_frame, text="📅", width=3,
                  command=lambda e=date_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # Abnormal checkbox
        abnormal_var = tk.BooleanVar()
        abnormal_check = ttk.Checkbutton(form_frame, text="Abnormal Result",
                                       variable=abnormal_var)
        abnormal_check.pack(pady=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_entry = tk.Text(form_frame, height=5, wrap=tk.WORD)
        notes_entry.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def save_lab():
            new_lab = {
                "test": test_combo.get(),
                "value": value_entry.get(),
                "unit": unit_entry.get(),
                "date": date_entry.get(),
                "abnormal": abnormal_var.get(),
                "notes": notes_entry.get("1.0", tk.END).strip(),
                "entered_by": self.current_user,
                "entry_date": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            }
            
            if not new_lab["test"] or not new_lab["value"]:
                messagebox.showerror("Error", "Test name and value are required")
                return
                
            if "LAB_RESULTS" not in patient:
                patient["LAB_RESULTS"] = []
                
            patient["LAB_RESULTS"].append(new_lab)
            self.save_patient_data()
            
            # Update the lab results tab
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_lab,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def edit_lab_result(self, lab, patient):
        """Edit an existing lab result"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Lab Result - {lab['test']}")
        dialog.geometry("500x400")
        
        # Form frame
        form_frame = ttk.Frame(dialog, style='TFrame', padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Test name
        ttk.Label(form_frame, text="Test Name:").pack(pady=5)
        test_entry = ttk.Entry(form_frame)
        test_entry.insert(0, lab["test"])
        test_entry.pack(fill=tk.X, pady=5)
        
        # Value
        ttk.Label(form_frame, text="Value:").pack(pady=5)
        value_entry = ttk.Entry(form_frame)
        value_entry.insert(0, lab["value"])
        value_entry.pack(fill=tk.X, pady=5)
        
        # Unit
        ttk.Label(form_frame, text="Unit:").pack(pady=5)
        unit_entry = ttk.Entry(form_frame)
        unit_entry.insert(0, lab.get("unit", ""))
        unit_entry.pack(fill=tk.X, pady=5)
        
        # Date
        ttk.Label(form_frame, text="Date:").pack(pady=5)
        date_frame = ttk.Frame(form_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        date_entry = ttk.Entry(date_frame)
        date_entry.insert(0, lab["date"])
        date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(date_frame, text="📅", width=3,
                  command=lambda e=date_entry: self.show_calendar(e)).pack(side=tk.LEFT, padx=5)
        
        # Abnormal checkbox
        abnormal_var = tk.BooleanVar(value=lab.get("abnormal", False))
        abnormal_check = ttk.Checkbutton(form_frame, text="Abnormal Result",
                                       variable=abnormal_var)
        abnormal_check.pack(pady=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_entry = tk.Text(form_frame, height=5, wrap=tk.WORD)
        notes_entry.insert("1.0", lab.get("notes", ""))
        notes_entry.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def save_changes():
            lab["test"] = test_entry.get()
            lab["value"] = value_entry.get()
            lab["unit"] = unit_entry.get()
            lab["date"] = date_entry.get()
            lab["abnormal"] = abnormal_var.get()
            lab["notes"] = notes_entry.get("1.0", tk.END).strip()
            
            self.save_patient_data()
            
            # Update the lab results tab
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_changes,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Delete", 
                  command=lambda: self.delete_lab_result(lab, patient, dialog),
                  style='Red.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def delete_lab_result(self, lab, patient, dialog):
        """Delete a lab result"""
        confirm = messagebox.askyesno("Confirm", 
                                     f"Delete {lab['test']} result from {lab['date']}?")
        if not confirm:
            return
            
        patient["LAB_RESULTS"].remove(lab)
        self.save_patient_data()
        
        # Update the lab results tab
        if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
            self.medication_window.destroy()
            self.open_medication_management(patient["FILE NUMBER"])
        
        dialog.destroy()

    def delete_lab_result(self, lab, patient, dialog):
        """Delete a lab result"""
        confirm = messagebox.askyesno("Confirm", 
                                     f"Delete {lab['test']} result from {lab['date']}?")
        if not confirm:
            return
            
        patient["LAB_RESULTS"].remove(lab)
        self.save_patient_data()
        
        # Update the lab results tab
        if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
            self.medication_window.destroy()
            self.open_medication_management(patient["FILE NUMBER"])
        
        dialog.destroy()

    # ==============================================
    # F&N Documentation Feature (Requested Addition)
    # ==============================================
    def setup_fn_documentation(self):
        """Setup F&N documentation feature"""
        if not hasattr(self, 'fn_path'):
            # Initialize with default path if not set
            self.fn_path = ""
            
        # Create dialog for path selection
        dialog = tk.Toplevel(self.root)
        dialog.title("F&N Documentation Setup")
        dialog.geometry("600x200")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current path display
        ttk.Label(main_frame, text="Current F&N Documentation Path:").pack(pady=5)
        path_label = ttk.Label(main_frame, text=self.fn_path if self.fn_path else "Not set")
        path_label.pack(pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        def browse_path():
            path = filedialog.askopenfilename(title="Select F&N Documentation Executable",
                                             filetypes=[("Executable files", "*.exe")])
            if path:
                self.fn_path = path
                path_label.config(text=path)
                # Save to config file
                self.save_config()
                
        def run_fn():
            if not self.fn_path:
                messagebox.showerror("Error", "No F&N documentation path set")
                return
            try:
                subprocess.Popen(self.fn_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not launch F&N documentation: {e}")
        
        ttk.Button(btn_frame, text="Browse", command=browse_path).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Run F&N Documentation", command=run_fn,
                  style='Blue.TButton').pack(side=tk.LEFT, padx=5)
        
        # Only show edit button for mej.esam
        if self.current_user == "mej.esam":
            def edit_path():
                new_path = simpledialog.askstring("Edit Path", 
                                                 "Enter new F&N Documentation path:",
                                                 initialvalue=self.fn_path)
                if new_path is not None:
                    self.fn_path = new_path
                    path_label.config(text=new_path)
                    self.save_config()
                    
            ttk.Button(btn_frame, text="Edit Path", command=edit_path,
                      style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def save_config(self):
        """Save configuration settings"""
        config = {
            'fn_path': self.fn_path,
            # Add other config settings here
        }
        with open('config.json', 'w') as f:
            json.dump(config, f)

    def load_config(self):
        """Load configuration settings"""
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
                self.fn_path = config.get('fn_path', "")
        except FileNotFoundError:
            self.fn_path = ""

    # ==============================================
    # Settings Management (Requested Addition)
    # ==============================================
    def open_settings(self):
        """Open settings window (only for mej.esam)"""
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only mej.esam can access settings")
            return
            
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Application Settings")
        settings_window.geometry("800x600")
        
        # Notebook for different setting categories
        notebook = ttk.Notebook(settings_window)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Dropdown Lists Settings Tab
        dropdown_frame = ttk.Frame(notebook)
        notebook.add(dropdown_frame, text="Dropdown Lists")
        self.setup_dropdown_settings(dropdown_frame)
        
        # Path Settings Tab
        path_frame = ttk.Frame(notebook)
        notebook.add(path_frame, text="Paths")
        self.setup_path_settings(path_frame)
        
        # Data Management Tab
        data_frame = ttk.Frame(notebook)
        notebook.add(data_frame, text="Data Management")
        self.setup_data_management_settings(data_frame)
        
        # Close button
        btn_frame = ttk.Frame(settings_window)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Close", command=settings_window.destroy,
                  style='Blue.TButton').pack()

    def setup_dropdown_settings(self, parent):
        """Setup dropdown list editing interface"""
        # Create scrollable frame
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Header
        ttk.Label(scrollable_frame, text="Edit Dropdown Options", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Create editable fields for each dropdown
        for category, options in DROPDOWN_OPTIONS.items():
            frame = ttk.LabelFrame(scrollable_frame, text=category)
            frame.pack(fill=tk.X, padx=5, pady=5)
            
            # Current options display
            ttk.Label(frame, text="Current Options:").pack(anchor="w")
            current_text = tk.Text(frame, height=4, width=50)
            current_text.insert("1.0", "\n".join(options))
            current_text.config(state=tk.DISABLED)
            current_text.pack(fill=tk.X, padx=5, pady=5)
            
            # Edit field
            ttk.Label(frame, text="New Options (one per line):").pack(anchor="w")
            edit_text = tk.Text(frame, height=4, width=50)
            edit_text.pack(fill=tk.X, padx=5, pady=5)
            
            # Update button
            def update_options(cat=category, text_widget=edit_text):
                new_options = text_widget.get("1.0", tk.END).strip().split("\n")
                new_options = [opt.strip() for opt in new_options if opt.strip()]
                DROPDOWN_OPTIONS[cat] = new_options
                self.save_dropdown_options()
                messagebox.showinfo("Success", f"{cat} options updated")
                
            ttk.Button(frame, text="Update", command=update_options,
                      style='Blue.TButton').pack(pady=5)

    def save_dropdown_options(self):
        """Save dropdown options to file"""
        with open('dropdown_options.json', 'w') as f:
            json.dump(DROPDOWN_OPTIONS, f, indent=4)

    def setup_path_settings(self, parent):
        """Setup path management interface"""
        ttk.Label(parent, text="Application Path Settings", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # F&N Documentation Path
        fn_frame = ttk.Frame(parent)
        fn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(fn_frame, text="F&N Documentation Path:").pack(side=tk.LEFT)
        self.fn_path_entry = ttk.Entry(fn_frame)
        self.fn_path_entry.insert(0, self.fn_path)
        self.fn_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(fn_frame, text="Browse", 
                  command=self.browse_fn_path).pack(side=tk.LEFT)
        
        # Save button
        ttk.Button(parent, text="Save Path Settings", 
                  command=self.save_path_settings,
                  style='Green.TButton').pack(pady=10)

    def browse_fn_path(self):
        """Browse for F&N documentation executable"""
        path = filedialog.askopenfilename(title="Select F&N Documentation Executable",
                                        filetypes=[("Executable files", "*.exe")])
        if path:
            self.fn_path_entry.delete(0, tk.END)
            self.fn_path_entry.insert(0, path)

    def save_path_settings(self):
        """Save path settings"""
        self.fn_path = self.fn_path_entry.get()
        self.save_config()
        messagebox.showinfo("Success", "Path settings saved")

    def setup_data_management_settings(self, parent):
        """Setup data management interface"""
        ttk.Label(parent, text="Data Management", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Delete local patient folders
        ttk.Button(parent, text="Delete All Local Patient Folders", 
                  command=self.confirm_delete_folders,
                  style='Red.TButton').pack(pady=10)
        
        # Export/Import settings
        ttk.Label(parent, text="Data Transfer:").pack(pady=5)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="Export All Data", 
                  command=self.export_all_data).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Import Data", 
                  command=self.import_data).pack(side=tk.LEFT, padx=5)

    def confirm_delete_folders(self):
        """Confirm deletion of all local patient folders"""
        confirm = messagebox.askyesno("Confirm", 
                                     "This will delete ALL local patient folders.\n"
                                     "Make sure all data is backed up!\n\n"
                                     "Are you absolutely sure?")
        if confirm:
            self.delete_all_patient_folders()

    def delete_all_patient_folders(self):
        """Delete all local patient folders"""
        try:
            for patient in self.patient_data:
                folder_name = f"Patient_{patient['FILE NUMBER']}"
                if os.path.exists(folder_name):
                    shutil.rmtree(folder_name)
            messagebox.showinfo("Success", "All local patient folders deleted")
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete folders: {e}")

    # ==============================================
    # Main Menu Updates (Requested Additions)
    # ==============================================
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
        form_container = ttk.Frame(right_frame, style='TFrame')
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

        # Fourth row - New Features (Purple buttons)
        row4_frame = ttk.Frame(btn_frame, style='TFrame')
        row4_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(row4_frame, text="F&N Documentation", command=self.setup_fn_documentation,
                  style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        if self.current_user == "mej.esam":
            ttk.Button(row4_frame, text="Settings", command=self.open_settings,
                      style='Purple.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row4_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Fifth row - User Management (Brown buttons)
        row5_frame = ttk.Frame(btn_frame, style='TFrame')
        row5_frame.pack(fill=tk.X, pady=5)
        
        if self.current_user == "mej.esam" or self.users[self.current_user]["role"] == "admin":
            ttk.Button(row5_frame, text="Manage Users", command=self.manage_users,
                      style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            
            ttk.Button(row5_frame, text="Change Password", command=self.change_password,
                      style='Brown.TButton').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        else:
            # Empty space to maintain layout
            ttk.Frame(row5_frame, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # Sixth row - Logout (Red button)
        row6_frame = ttk.Frame(btn_frame, style='TFrame')
        row6_frame.pack(fill=tk.X, pady=(20, 5))
        
        ttk.Button(row6_frame, text="Logout", command=self.setup_login_screen,
                  style='Red.TButton').pack(fill=tk.X, ipady=10)

    # ... [Previous code in main_menu()] ...

        # Signature
        signature_frame = ttk.Frame(form_container, style='TFrame')
        signature_frame.pack(pady=(30, 0))

        ttk.Label(signature_frame, text="Made by: DR. ESAM MEJRAB",
                 font=('Times New Roman', 14, 'italic'), 
                 foreground=self.primary_color).pack()

    # ==============================================
    # Existing Methods (Keep these unchanged)
    # ==============================================

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
        form_container = ttk.Frame(right_frame, bg='white')
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

    def clear_frame(self):
        """Clear all widgets from the main frame except status bar"""
        for widget in self.root.winfo_children():
            if widget not in [self.status_frame]:
                widget.destroy()

    def on_close(self):
        """Handle window close event"""
        if messagebox.askyesno("Exit Confirmation", "Are you sure you want to exit?"):
            # Cancel all pending after events
            for after_id in self.root.tk.eval('after info').split():
                self.root.after_cancel(after_id)
            
            self.executor.shutdown(wait=False)
            self.root.destroy()

    def update_datetime(self):
        """Update the date and time display"""
        if hasattr(self, 'datetime_label') and self.datetime_label.winfo_exists():
            now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            self.datetime_label.config(text=now)
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
            self.root.after(30000, internet_check)  # Check every 30 seconds
        
        internet_check()

    def update_internet_indicator(self):
        """Update the internet connection indicator"""
        color = "green" if self.internet_connected else "red"
        self.internet_indicator.delete("all")
        self.internet_indicator.create_oval(5, 5, 15, 15, fill=color, outline="")

    def load_users(self):
        """Load user data from file or create default users"""
        if os.path.exists("users_data.json"):
            try:
                with open("users_data.json", "r") as f:
                    self.users = json.load(f)
            except json.JSONDecodeError:
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
        self.style.configure('Green.TButton', background=self.success_color, foreground='white')
        self.style.configure('Yellow.TButton', background=self.warning_color, foreground='white')
        self.style.configure('Brown.TButton', background='#8B4513', foreground='white')
        self.style.configure('Red.TButton', background=self.danger_color, foreground='white')
        self.style.configure('Purple.TButton', background='#8A2BE2', foreground='white')
        
        # Card style for medication displays
        self.style.configure('Card.TFrame', background='white', borderwidth=1, relief='solid')

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

if __name__ == "__main__":
    root = tk.Tk()
    app = OncologyApp(root)
    root.mainloop()
