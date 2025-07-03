import matplotlib
matplotlib.use('TkAgg')  # Explicitly set the backend to TkAgg

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import pandas as pd
from tkinter.scrolledtext import ScrolledText
import re
from datetime import datetime
from tkcalendar import Calendar
from collections import Counter
import matplotlib.pyplot as plt
import bcrypt

def save_users_to_file():
    with open("users_data.json", "w") as f:
        json.dump(users, f, indent=4)

if os.path.exists("users_data.json"):
    with open("users_data.json", "r") as f:
        users = json.load(f)
else:
    # Default user data with hashed passwords
    users = {
        "mej.esam": {"password": bcrypt.hashpw("wjap19527".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'), "role": "admin"},
        "doctor1": {"password": bcrypt.hashpw("doc123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'), "role": "editor"},
        "nurse1": {"password": bcrypt.hashpw("nur123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'), "role": "viewer"},
        "seraj": {"password": bcrypt.hashpw("steve8288".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'), "role": "admin"}
    }
    save_users_to_file()  # Save the default users to the file

# Define all dropdown options with corrected spellings and full words
DROPDOWN_OPTIONS = {
    "Neutropenia": ["Febrile Non Neutropenic", "Afebrile Neutropenic", "Febrile Neutropenic"],
    "Toxicity Grade": ["1", "2", "3", "4"],  # Updated values
    "Symptoms": ["Asymptomatic", "Vomiting", "Diarrhea", "Low Blood Pressure", 
                 "High Blood Pressure", "Septicemia", "Septic Shock", "Skin Rash", 
                 "Mouth Ulcer", "Dysphagia", "Bleeding", "cough", "abdominal pain", "dysurea", "headache", "fatigue"],
    "AB Sensitivity": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Resistance": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Sensitivity 2": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Resistance 2": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Sensitivity 3": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Resistance 3": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Sensitivity 4": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Resistance 4": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Sensitivity 5": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "AB Resistance 5": [
        "AMIKA", "AMPICILLIN", "AMOXIL", "AUGMENTIN", "AZETREONAM", "CARBINCILIN", 
        "CEFALEXIN", "CEFALOTIN", "CEFAMANADOLE", "CEFOTAXIM", "CEFOXITIN", 
        "CRFTRIAXONE", "CEFORUIME", "CEFTAZIDIME", "CHLORAMPHENICOL", "CIPRO", 
        "CLINDA", "CLOXA", "COLISTIN", "ERYTHRO", "FORAZOLIDONE", "FUCIDIC ACID", 
        "DAPTOMYCIN", "SULFAMETHOXAZOLE", "SEPTRIN (COTRIMOXAZOLE)", "STREPTOMYCIN", 
        "TIGECYCLIN", "KANAMYCIN", "LEVOFLOXACIN", "METHICILLINNALIDIXIC ACID", 
        "NEOMYCIN", "NITROFURANTOIN", "NORABIOCIN", "OXACILLIN", "OLEADOMYCIN", 
        "PENICILLIN", "PIPRACILLIN", "PIPRACILLIN/TAZOBACTAM", "RIFAMPICIN", 
        "RIFAMYCIN", "SEPTRIN", "TETRACYCLIN", "TECARCILLIN/CLAVULANATE", 
        "TOBRAMYCIN", "VANCOM", "DOXYCYCLIN", "AMPICILLIN SULBACTAM", "ERTAPENEM", 
        "MEROPENEM", "CEFEPIME", "TRIMETHOPRIM/SULFAMETHOXAZOLE", "GENTA", 
        "IMIPENEM", "KANAMYCIN", "OXACILLIN", "GENTAMICIN"
    ],
    "Used Antibiotics": [
        "MERO", "AMIKA", "GENTA", "ROCOPHINE", "CIPRO", "TAZO", "VANCO", 
        "FLAGYL", "FLUCONAZOLE", "AMPHOTERICIN", "VORICONAZOLE", "AUGMENTIN", "ACYCLOVIR",
        "AMPICILLIN", "CLARITHROMYCIN", "CEFTA", "CLOXA", "CLINDAMYCIN", "COLISTIN",
        "GENTAMYCIN", "SUPRAX ORALLY", "ZOMAX ORALLY", "AUGMENTIN ORALLY", "CIPRO ORALLY"
    ],
    "Outcome": ["Recovered", "Deceased", "Unknown"],
    "Gender": ["Male", "Female"],
    "Stage": ["1", "2", "3", "4", "4s", "5"],
    "Risk Group": ["LRG", "SRG", "IRG", "HRG"]
}

def save_users_to_file():
    with open("users_data.json", "w") as f:
        json.dump(users, f, indent=4)

class FNApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pediatric Oncology F&N Documentation")
        self.root.geometry("1200x800")
        self.root.configure(bg='#e6f2ff')

        # Bind the close button to the confirmation dialog
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Start in full-screen mode
        self.root.state('zoomed')

        # Custom styling
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Configure styles
        self.style.configure('TFrame', background='#e6f2ff')
        self.style.configure('TButton', font=('Arial Rounded MT Bold', 10), padding=5, 
                           background='#4d94ff', foreground='white')
        self.style.map('TButton', background=[('active', '#0066cc')])
        self.style.configure('TLabel', font=('Arial Rounded MT Bold', 10), 
                           background='#e6f2ff', foreground='#003366')
        self.style.configure('Header.TLabel', font=('Arial Rounded MT Bold', 20), 
                           foreground='#003366')
        self.style.configure('Accent.TButton', font=('Arial Rounded MT Bold', 10), 
                          background='#009933', foreground='white')
        self.style.map('Accent.TButton', background=[('active', '#006622')])
        self.style.configure('TEntry', font=('Arial', 10), padding=5)
        self.style.configure('TCombobox', font=('Arial', 10))

        self.current_user = None
        self.current_results = []
        self.current_result_index = 0
        self.setup_login_screen()

    def on_close(self):
        """Handle the close button click event."""
        if messagebox.askyesno("Exit Confirmation", "Are you sure you want to exit?"):
            self.root.destroy()

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def setup_login_screen(self):
        self.clear_frame()
        
        # Main container with gradient background
        main_frame = ttk.Frame(self.root, padding="40 40 40 40", style='TFrame')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Logo/Header
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(pady=(0, 30))
        
        ttk.Label(header_frame, text="Pediatric Oncology", font=('Arial Rounded MT Bold', 24), 
                 foreground='#003366', style='Header.TLabel').pack()
        ttk.Label(header_frame, text="Fever & Neutropenia Documentation", 
                 font=('Arial Rounded MT Bold', 18), foreground='#0066cc', 
                 style='Header.TLabel').pack()
        
        # Login form
        login_frame = ttk.Frame(main_frame, style='TFrame')
        login_frame.pack(pady=20)
        
        ttk.Label(login_frame, text="Username:", font=('Arial Rounded MT Bold', 12)).grid(row=0, column=0, padx=5, pady=10, sticky="e")
        self.entry_username = ttk.Entry(login_frame, width=30, font=('Arial', 12))
        self.entry_username.grid(row=0, column=1, padx=5, pady=10)
        
        ttk.Label(login_frame, text="Password:", font=('Arial Rounded MT Bold', 12)).grid(row=1, column=0, padx=5, pady=10, sticky="e")
        self.entry_password = ttk.Entry(login_frame, show="*", width=30, font=('Arial', 12))
        self.entry_password.grid(row=1, column=1, padx=5, pady=10)
        
        # Button frame
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Login", command=self.login, style='Accent.TButton').pack(side=tk.LEFT, padx=10, ipadx=20)
        ttk.Button(btn_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT, padx=10, ipadx=20)
        
        # Focus on username field
        self.entry_username.focus()
    
    def login(self):
        username = self.entry_username.get()
        password = self.entry_password.get()

        if username in users:
            stored_hash = users[username]["password"].encode('utf-8')
            if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
                self.current_user = username
                messagebox.showinfo("Login Successful", f"Welcome, {username}!")
                self.main_menu()
                return

        messagebox.showerror("Login Failed", "Invalid username or password.")
        self.entry_password.delete(0, tk.END)
    
    def main_menu(self):
        self.clear_frame()

        # Main container
        main_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Header
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(pady=(0, 30))

        ttk.Label(header_frame, text=f"Welcome, {self.current_user}",
                  font=('Arial Rounded MT Bold', 16), foreground='#003366').pack()
        ttk.Label(header_frame, text="Main Menu",
                  font=('Arial Rounded MT Bold', 20), foreground='#0066cc').pack()

        # Button grid
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(pady=20)

        # Determine available actions based on user role
        menu_buttons = []

        if users[self.current_user]["role"] in ["admin", "editor", "viewer"]:
            menu_buttons.append(("Search Patient", self.search_patient, "#4da6ff"))

        if users[self.current_user]["role"] in ["admin", "editor"]:
            menu_buttons.append(("Add New Patient", self.add_patient, "#66b3ff"))
            menu_buttons.append(("Backup Data", self.backup_data, "#ffcc00"))

        if self.current_user == "mej.esam":
            menu_buttons.append(("Restore Data", self.restore_data, "#ff6600"))
            menu_buttons.append(("Manage Users", self.manage_users, "#ff9933"))  # Only for mej.esam

        if users[self.current_user]["role"] == "admin":
            menu_buttons.append(("Export All Data", self.export_all_data, "#ffcc00"))
            menu_buttons.append(("View All Patients", self.view_all_patients, "#33cc33"))
            menu_buttons.append(("Statistics", self.open_statistics_window, "#ff9966"))

        if users[self.current_user]["role"] in ["admin", "editor"]:
            menu_buttons.append(("Change Password", self.change_password, "#ff9966"))
        menu_buttons.append(("Logout", self.setup_login_screen, "#ff6666"))

        # Create buttons in a grid
        for i, (text, command, color) in enumerate(menu_buttons):
            btn = tk.Button(btn_frame, text=text, command=command,
                            font=('Arial Rounded MT Bold', 12), bg=color, fg='white',
                            activebackground=color, activeforeground='white',
                            relief='flat', bd=0, padx=20, pady=15, width=20)
            btn.grid(row=i // 2, column=i % 2, padx=10, pady=10, sticky="nsew")
            btn.bind("<Enter>", lambda e, b=btn: b.config(relief='groove'))
            btn.bind("<Leave>", lambda e, b=btn: b.config(relief='flat'))

        # Signature
        signature_frame = ttk.Frame(main_frame, style='TFrame')
        signature_frame.pack(pady=(30, 0))

        ttk.Label(signature_frame, text="Made by: DR.ESAM MEJRAB",
                  font=('Times New Roman', 14, 'italic'), foreground='#003366').pack()

    def backup_data(self):
        if not os.path.exists('patients_data.json'):
            messagebox.showerror("Error", "No data available to back up.")
            return

        backup_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Backup As"
        )
        if not backup_path:
            return

        try:
            with open('patients_data.json', 'r') as f:
                data = json.load(f)
            with open(backup_path, 'w') as f:
                json.dump(data, f, indent=4)
            messagebox.showinfo("Success", f"Data backed up to {backup_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to back up data: {e}")

    def restore_data(self):
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
            with open('patients_data.json', 'w') as f:
                json.dump(data, f, indent=4)
            messagebox.showinfo("Success", "Data restored successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to restore data: {e}")

    def create_scrollable_frame(self, parent):
        container = ttk.Frame(parent)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        container.pack(fill="both", expand=True)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling anywhere on the screen
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)
        
        return scrollable_frame
    
    def add_patient(self):
        if users[self.current_user]["role"] == "viewer":
            messagebox.showerror("Access Denied", "Viewers cannot add patients")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Add New Patient", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Main form area
        form_frame = self.create_scrollable_frame(self.root)

        # Grouped fields configuration
        sections = [
            {
                "title": "Basic Information",
                "fields": [
                    ["File No", "text"], ["Name (English/Arabic)", "text"],
                    ["Date of Birth", "calendar"], ["Date of Admission", "calendar"],
                    ["Address", "text"], ["Gender", "dropdown"]
                ]
            },
            {
                "title": "Medical Information",
                "fields": [
                    ["Malignancy", "text"], ["Type", "text"], ["Stage", "dropdown"],
                    ["Risk Group", "dropdown"], ["Protocol", "text"], ["Cycle", "text"],
                    ["Days Post Last Cycle", "text"], ["Duration of Admission (days)", "text"],
                    ["Outcome", "dropdown"],
                    ["Neutropenia", "dropdown"],
                    ["Toxicity Grade", "dropdown"],
                    ["Symptoms", "multidropdown"], ["WBC Count", "text"],
                    ["Neutrophils", "text"], ["Hemoglobin", "text"], ["Platelets", "text"]
                ]
            },
            {
                "title": "Blood Culture",
                "fields": [
                    ["Blood Culture & Sensitivity", "text"],
                    ["AB Sensitivity", "multidropdown"],
                    ["AB Resistance", "multidropdown"]
                ]
            },
            {
                "title": "Urine Culture",
                "fields": [
                    ["Urine Culture & Sensitivity", "text"],
                    ["AB Sensitivity 2", "multidropdown"],
                    ["AB Resistance 2", "multidropdown"]
                ]
            },
            {
                "title": "Stool Culture",
                "fields": [
                    ["Stool Culture & Sensitivity", "text"],
                    ["AB Sensitivity 3", "multidropdown"],
                    ["AB Resistance 3", "multidropdown"]
                ]
            },
            {
                "title": "Throat Swab",
                "fields": [
                    ["Throat Swab", "text"],
                    ["AB Sensitivity 4", "multidropdown"],
                    ["AB Resistance 4", "multidropdown"]
                ]
            },
            {
                "title": "Other Swab",
                "fields": [
                    ["Other Swab Site", "text"],
                    ["AB Sensitivity 5", "multidropdown"],
                    ["AB Resistance 5", "multidropdown"]
                ]
            },
            {
                "title": "Used Antibiotics",
                "fields": [["Used Antibiotics", "multidropdown"]]
            }
        ]

        self.entries = {}

        # Create sections for better organization
        for section in sections:
            section_frame = ttk.LabelFrame(form_frame, text=section["title"], style='TFrame')
            section_frame.pack(fill=tk.X, padx=10, pady=10)

            for field_group in section["fields"]:
                if isinstance(field_group[0], list):  # Handle grouped fields in the same row
                    row_frame = ttk.Frame(section_frame, style='TFrame')
                    row_frame.pack(fill=tk.X, pady=5, padx=10)

                    for field, field_type in field_group:
                        ttk.Label(row_frame, text=field, width=15, anchor="w").pack(side=tk.LEFT, padx=5)

                        if field_type == "dropdown":
                            var = tk.StringVar()
                            options = DROPDOWN_OPTIONS.get(field, [])  # Access the options directly using the field name
                            combobox = ttk.Combobox(row_frame, textvariable=var, values=options, width=27)
                            combobox.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                            self.entries[field] = var
                else:  # Handle single fields
                    row_frame = ttk.Frame(section_frame, style='TFrame')
                    row_frame.pack(fill=tk.X, pady=5, padx=10)

                    field, field_type = field_group
                    ttk.Label(row_frame, text=field, width=25, anchor="w").pack(side=tk.LEFT, padx=5)

                    if field_type == "text":
                        entry = ttk.Entry(row_frame, width=30)
                        entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                        self.entries[field] = entry
                    elif field_type == "calendar":
                        cal_frame = ttk.Frame(row_frame)
                        cal_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

                        entry = ttk.Entry(cal_frame, width=20)
                        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

                        cal_btn = ttk.Button(cal_frame, text="📅", width=3,
                                             command=lambda e=entry: self.show_calendar(e))
                        cal_btn.pack(side=tk.LEFT, padx=5)
                        self.entries[field] = entry
                    elif field_type == "dropdown":
                        var = tk.StringVar()
                        options = DROPDOWN_OPTIONS.get(field, [])  # Access the options directly using the field name
                        combobox = ttk.Combobox(row_frame, textvariable=var, values=options, width=27)
                        combobox.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                        self.entries[field] = var
                    elif field_type == "multidropdown":
                        listbox_frame = ttk.Frame(row_frame)
                        listbox_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

                        listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4,
                                             exportselection=0, width=30, font=('Arial', 9))
                        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                        listbox.configure(yscrollcommand=scrollbar.set)

                        options = DROPDOWN_OPTIONS.get(field, [])
                        for option in options:
                            listbox.insert(tk.END, option)

                        listbox.pack(side=tk.LEFT, expand=True, fill=tk.X)
                        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                        self.entries[field] = listbox

        # Button frame
        btn_frame = ttk.Frame(self.root, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Save Patient", command=self.save_patient,
                   style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Back", command=self.main_menu).pack(side=tk.LEFT, padx=10)

    def show_calendar(self, entry_widget):
        top = tk.Toplevel(self.root)
        top.title("Select Date")
        top.geometry("300x300")  # Adjusted size to ensure button visibility
        
        cal = Calendar(top, selectmode='day', date_pattern='dd/mm/yyyy')
        cal.pack(pady=20)
        
        def set_date():
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, cal.get_date())
            top.destroy()
        
        # Fixed-size "Select" button (approximately 10mm by 20mm)
        select_btn = ttk.Button(top, text="Select", command=set_date, style='Accent.TButton')
        select_btn.pack(pady=10, ipadx=20, ipady=10)  # Adjusted padding for fixed size

    def validate_date(self, event):
        entry = event.widget
        text = entry.get()
        
        # Allow only digits and slashes
        if not re.match(r'^[0-9/]*$', text):
            entry.delete(0, tk.END)
            entry.insert(0, text[:-1])
            return
        
        # Auto-insert slashes
        if len(text) == 2 or len(text) == 5:
            entry.insert(tk.END, '/')
        
        # Validate day, month, year as they're being entered
        parts = text.split('/')
        if len(parts) > 0 and parts[0]:  # Day
            day = parts[0]
            if len(day) > 2 or (day and not 1 <= int(day) <= 31):
                entry.config(foreground='red')
            else:
                entry.config(foreground='black')
        
        if len(parts) > 1 and parts[1]:  # Month
            month = parts[1]
            if len(month) > 2 or (month and not 1 <= int(month) <= 12):
                entry.config(foreground='red')
            else:
                entry.config(foreground='black')
        
        if len(parts) > 2 and parts[2]:  # Year
            year = parts[2]
            if len(year) > 4:
                entry.delete(8, tk.END)
                entry.config(foreground='red')
            else:
                entry.config(foreground='black')
    
    def save_patient(self):
        # Mandatory fields
        mandatory_fields = [
            "File No", "Name (English/Arabic)", "Date of Admission", "Gender",
            "Malignancy", "Protocol", "Neutropenia", "WBC Count"
        ]

        # Check for missing mandatory fields
        missing_fields = []
        for field in mandatory_fields:
            value = self.entries[field].get() if isinstance(self.entries[field], (tk.Entry, tk.StringVar)) else None
            if not value:  # Check if the field is empty
                missing_fields.append(field)

        if missing_fields:
            messagebox.showerror("Error", f"The following fields are mandatory and cannot be empty:\n{', '.join(missing_fields)}")
            return

        # Check for duplicate file number
        file_no = self.entries["File No"].get()
        if os.path.exists('patients_data.json'):
            with open('patients_data.json', 'r') as f:
                data = json.load(f)

            for patient in data:
                if patient.get("File No") == file_no:
                    messagebox.showerror("Error", "Patient with this File No already exists")
                    return
        else:
            data = []

        # Validate date format for "Date of Admission"
        date_of_admission = self.entries["Date of Admission"].get()
        try:
            datetime.strptime(date_of_admission, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format for 'Date of Admission'. Please use dd/mm/yyyy")
            return

        # Collect patient data
        patient_data = {}
        for field, widget in self.entries.items():
            if isinstance(widget, tk.Listbox):  # Multi-select fields
                patient_data[field] = ", ".join([widget.get(i) for i in widget.curselection()])
            elif isinstance(widget, tk.StringVar):  # Combobox
                patient_data[field] = widget.get()
            else:  # Entry fields
                patient_data[field] = widget.get()

        # Save the patient data
        data.append(patient_data)
        with open('patients_data.json', 'w') as f:
            json.dump(data, f, indent=4)

        messagebox.showinfo("Success", "Patient data saved successfully!")
        self.main_menu()
    
    def search_patient(self):
        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Search Patient", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Search form
        form_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="File No:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        self.search_file_no_entry = ttk.Entry(form_frame, width=40, font=('Arial', 12))
        self.search_file_no_entry.pack(pady=5)

        ttk.Label(form_frame, text="Name:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        self.search_name_entry = ttk.Entry(form_frame, width=40, font=('Arial', 12))
        self.search_name_entry.pack(pady=5)

        btn_frame = ttk.Frame(form_frame, style='TFrame')
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="Search", command=self.perform_search,
                   style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Back", command=self.main_menu).pack(side=tk.LEFT, padx=10)

    def perform_search(self):
        file_no = self.search_file_no_entry.get().strip()
        name = self.search_name_entry.get().strip()

        if not os.path.exists('patients_data.json'):
            messagebox.showerror("Error", "No patient data available")
            return

        with open('patients_data.json', 'r') as f:
            data = json.load(f)

        # Filter results based on file number or name
        matching_data = []
        for patient in data:
            if file_no and patient.get("File No") == file_no:
                matching_data = [patient]  # If searching by file number, return only one result
                break
            elif name and name.lower() in patient.get("Name (English/Arabic)", "").lower():
                matching_data.append(patient)

        if not matching_data:
            messagebox.showerror("Not Found", "No patient found with the given criteria")
            return

        # Handle multiple results
        self.current_results = matching_data
        self.current_result_index = 0
        self.view_patient(self.current_results[self.current_result_index], multiple_results=(not file_no))

    def view_patient(self, patient_data, multiple_results=False):
        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Patient Details", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Main content
        form_frame = self.create_scrollable_frame(self.root)

        # Display patient data in a clean layout with alternating colors
        for i, (field, value) in enumerate(patient_data.items()):
            row_frame = ttk.Frame(form_frame, style='TFrame')
            row_frame.pack(fill=tk.X, pady=2, padx=10)

            # Alternate row colors for better readability
            bg_color = '#f0f7ff' if i % 2 == 0 else '#e6f2ff'

            field_frame = tk.Frame(row_frame, bg=bg_color)
            field_frame.pack(fill=tk.X, ipady=5)

            tk.Label(field_frame, text=f"{field}:", width=25, anchor="w",
                     bg=bg_color, font=('Arial Rounded MT Bold', 10)).pack(side=tk.LEFT, padx=5)
            tk.Label(field_frame, text=value, width=40, anchor="w",
                     bg=bg_color, font=('Arial', 10)).pack(side=tk.LEFT)

        # Button frame
        btn_frame = ttk.Frame(self.root, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)

        if users[self.current_user]["role"] in ["admin", "editor"]:
            ttk.Button(btn_frame, text="Edit", command=lambda: self.edit_patient(patient_data),
                       style='Accent.TButton').pack(side=tk.LEFT, padx=10)

        if self.current_user == "mej.esam":
            ttk.Button(btn_frame, text="Delete", command=lambda: self.delete_patient(patient_data),
                       style='Accent.TButton').pack(side=tk.LEFT, padx=10)

        # Show "Previous" and "Next" buttons only if multiple results are available
        if multiple_results:
            ttk.Button(btn_frame, text="Previous", command=self.show_previous_result).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="Next", command=self.show_next_result).pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="Back", command=self.search_patient).pack(side=tk.LEFT, padx=10)

    def show_previous_result(self):
        if self.current_result_index > 0:
            self.current_result_index -= 1
            self.view_patient(self.current_results[self.current_result_index], multiple_results=True)
        else:
            messagebox.showinfo("Info", "This is the first result.")

    def show_next_result(self):
        if self.current_result_index < len(self.current_results) - 1:
            self.current_result_index += 1
            self.view_patient(self.current_results[self.current_result_index], multiple_results=True)
        else:
            messagebox.showinfo("Info", "This is the last result.")

    def delete_patient(self, patient_data):
        if self.current_user != "mej.esam":
            messagebox.showerror("Access Denied", "Only 'mej.esam' can delete patients.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this patient?")
        if not confirm:
            return
        
        file_no = patient_data.get("File No")
        if not os.path.exists('patients_data.json'):
            messagebox.showerror("Error", "No patient data available")
            return
        
        with open('patients_data.json', 'r') as f:
            data = json.load(f)
        
        data = [patient for patient in data if patient.get("File No") != file_no]
        
        with open('patients_data.json', 'w') as f:
            json.dump(data, f, indent=4)
        
        messagebox.showinfo("Success", "Patient deleted successfully!")
        self.search_patient()
    
    def edit_patient(self, patient_data):
        if users[self.current_user]["role"] == "viewer":
            messagebox.showerror("Access Denied", "Viewers cannot edit patients")
            return
            
        self.clear_frame()
        
        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(header_frame, text="Edit Patient", font=('Arial Rounded MT Bold', 18), 
                 foreground='#0066cc').pack(side=tk.LEFT)
        
        # Main form area
        form_frame = self.create_scrollable_frame(self.root)
        
        # Form fields configuration - reorganized to group related fields
        fields = [
            ("File No", "text"), ("Name (English/Arabic)", "text"), 
            ("Date of Birth", "calendar"), ("Date of Admission", "calendar"),  # New field
            ("Address", "text"), ("Gender", "dropdown"),
            ("Malignancy", "text"), ("Type", "text"), ("Stage", "dropdown"),  # Added Stage
            ("Risk Group", "dropdown"),  # Added Risk Group
            ("Protocol", "text"), ("Cycle", "text"), ("Days Post Last Cycle", "text"), 
            ("Duration of Admission (days)", "text"), ("Outcome", "dropdown"),
            
            # Blood Culture section
            ("Blood Culture & Sensitivity", "text"), 
            ("AB Sensitivity", "multidropdown"), 
            ("AB Resistance", "multidropdown"), 
            
            # Urine Culture section
            ("Urine Culture & Sensitivity", "text"), 
            ("AB Sensitivity 2", "multidropdown"), 
            ("AB Resistance 2", "multidropdown"), 
            
            # Stool Culture section
            ("Stool Culture & Sensitivity", "text"), 
            ("AB Sensitivity 3", "multidropdown"), 
            ("AB Resistance 3", "multidropdown"), 
            
            # Throat Swab section
            ("Throat Swab", "text"), 
            ("AB Sensitivity 4", "multidropdown"), 
            ("AB Resistance 4", "multidropdown"), 
            
            # Other Swab section
            ("Other Swab Site", "text"), 
            ("AB Sensitivity 5", "multidropdown"), 
            ("AB Resistance 5", "multidropdown"), 
            
            # Used Antibiotics
            ("Used Antibiotics", "multidropdown")
        ]
        
        self.edit_entries = {}
        for field, field_type in fields:
            row_frame = ttk.Frame(form_frame, style='TFrame')
            row_frame.pack(fill=tk.X, pady=5, padx=10)
            
            ttk.Label(row_frame, text=field, width=25, anchor="w").pack(side=tk.LEFT, padx=5)
            
            current_value = patient_data.get(field, "")
            
            if field_type == "text":
                entry = ttk.Entry(row_frame, width=40)
                entry.insert(0, current_value)
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                self.edit_entries[field] = entry
            elif field_type == "calendar":
                cal_frame = ttk.Frame(row_frame)
                cal_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                
                entry = ttk.Entry(cal_frame, width=30)
                entry.insert(0, current_value)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                cal_btn = ttk.Button(cal_frame, text="📅", width=3,
                                     command=lambda e=entry: self.show_calendar(e))
                cal_btn.pack(side=tk.LEFT, padx=5)
                self.edit_entries[field] = entry
            elif field_type == "dropdown":
                var = tk.StringVar(value=current_value)
                options = DROPDOWN_OPTIONS.get(field.replace(" ", ""), [])
                combobox = ttk.Combobox(row_frame, textvariable=var, values=options, width=37)
                combobox.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                self.edit_entries[field] = var
            elif field_type == "multidropdown":
                listbox_frame = ttk.Frame(row_frame)
                listbox_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                
                listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=4, 
                                    exportselection=0, width=40, font=('Arial', 9))
                scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
                listbox.configure(yscrollcommand=scrollbar.set)
                
                # Use the exact field name to look up options in DROPDOWN_OPTIONS
                options = DROPDOWN_OPTIONS.get(field, [])
                for option in options:
                    listbox.insert(tk.END, option)
                
                # Select previously selected items
                selected_values = patient_data.get(field, "").split(", ") if patient_data.get(field) else []
                for i, item in enumerate(options):
                    if item in selected_values:
                        listbox.select_set(i)
                
                listbox.pack(side=tk.LEFT, expand=True, fill=tk.X)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                self.edit_entries[field] = listbox
        
        # Button frame
        btn_frame = ttk.Frame(self.root, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="Save Changes", 
                  command=lambda: self.save_edited_patient(patient_data),
                  style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", 
                  command=lambda: self.view_patient(patient_data)).pack(side=tk.LEFT, padx=10)
    
    def save_edited_patient(self, original_data):
        # Validate date format
        dob = self.edit_entries["Date of Birth"].get()
        try:
            datetime.strptime(dob, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Please use dd/mm/yyyy")
            return
            
        updated_data = {}
        for field, widget in self.edit_entries.items():
            if isinstance(widget, tk.Listbox):  # Multi-select fields
                updated_data[field] = ", ".join([widget.get(i) for i in widget.curselection()])
            elif isinstance(widget, tk.StringVar):  # Combobox
                updated_data[field] = widget.get()
            else:  # Entry fields
                updated_data[field] = widget.get()
        
        # Load existing data
        with open('patients_data.json', 'r') as f:
            data = json.load(f)
        
        # Update the record
        for i, patient in enumerate(data):
            if patient.get("File No") == original_data.get("File No"):
                data[i] = updated_data
                break
        
        # Save back
        with open('patients_data.json', 'w') as f:
            json.dump(data, f, indent=4)
        
        messagebox.showinfo("Success", "Patient data updated successfully!")
        self.view_patient(updated_data)
    
    def manage_users(self):
        if users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can manage users")
            return
            
        self.clear_frame()
        
        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(header_frame, text="User  Management", font=('Arial Rounded MT Bold', 18), 
                 foreground='#0066cc').pack(side=tk.LEFT)
        
        # Main content
        main_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # User list
        list_frame = ttk.Frame(main_frame, style='TFrame')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.user_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                                  selectmode=tk.SINGLE, font=('Arial', 12),
                                  height=10, bg='white', bd=0, highlightthickness=0)
        self.user_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.user_list.yview)
        
        for username in users:
            self.user_list.insert(tk.END, f"{username} ({users[username]['role']})")
        
        # Add user form
        form_frame = ttk.Frame(main_frame, style='TFrame')
        form_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(form_frame, text="New Username:", font=('Arial Rounded MT Bold', 11)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.new_username = ttk.Entry(form_frame, font=('Arial', 11))
        self.new_username.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Password:", font=('Arial Rounded MT Bold', 11)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.new_password = ttk.Entry(form_frame, show="*", font=('Arial', 11))
        self.new_password.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(form_frame, text="Role:", font=('Arial Rounded MT Bold', 11)).grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.new_role = ttk.Combobox(form_frame, values=["admin", "editor", "viewer"], 
                                    font=('Arial', 11))
        self.new_role.grid(row=2, column=1, padx=5, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Add User", command=self.add_user, 
                  style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete User", command=self.delete_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Change Password", command=self.change_user_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Back", command=self.main_menu).pack(side=tk.LEFT, padx=5)

    def change_user_password(self):
        selection = self.user_list.curselection()
        if not selection:
            messagebox.showerror("Error", "No user selected")
            return

        selected = self.user_list.get(selection[0])
        username = selected.split()[0]

        if username == "mej.esam":
            messagebox.showerror("Error", "Cannot change password for 'mej.esam'.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text=f"Change Password for {username}", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Form
        form_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="New Password:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        new_password_entry = ttk.Entry(form_frame, show="*", font=('Arial', 12))
        new_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="Confirm New Password:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        confirm_password_entry = ttk.Entry(form_frame, show="*", font=('Arial', 12))
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
            users[username]["password"] = hashed_password
            save_users_to_file()
            messagebox.showinfo("Success", f"Password for {username} changed successfully!")
            self.manage_users()

        ttk.Button(btn_frame, text="Save", command=save_new_password, style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=self.manage_users).pack(side=tk.LEFT, padx=10)

    def change_password(self):
        if users[self.current_user]["role"] == "viewer":
            messagebox.showerror("Access Denied", "Viewers cannot change passwords.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Change Password", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Form
        form_frame = ttk.Frame(self.root, padding="20 20 20 20", style='TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(form_frame, text="Current Password:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        current_password_entry = ttk.Entry(form_frame, show="*", font=('Arial', 12))
        current_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="New Password:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        new_password_entry = ttk.Entry(form_frame, show="*", font=('Arial', 12))
        new_password_entry.pack(pady=5)

        ttk.Label(form_frame, text="Confirm New Password:", font=('Arial Rounded MT Bold', 12)).pack(pady=5)
        confirm_password_entry = ttk.Entry(form_frame, show="*", font=('Arial', 12))
        confirm_password_entry.pack(pady=5)

        # Buttons
        btn_frame = ttk.Frame(form_frame, style='TFrame')
        btn_frame.pack(pady=20)

        def save_new_password():
            current_password = current_password_entry.get()
            new_password = new_password_entry.get()
            confirm_password = confirm_password_entry.get()

            stored_hash = users[self.current_user]["password"].encode('utf-8')
            if not bcrypt.checkpw(current_password.encode('utf-8'), stored_hash):
                messagebox.showerror("Error", "Current password is incorrect.")
                return

            if new_password != confirm_password:
                messagebox.showerror("Error", "New passwords do not match.")
                return

            hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            users[self.current_user]["password"] = hashed_password
            save_users_to_file()
            messagebox.showinfo("Success", "Password changed successfully!")
            self.main_menu()

        ttk.Button(btn_frame, text="Save", command=save_new_password, style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=self.main_menu).pack(side=tk.LEFT, padx=10)

    def add_user(self):
        username = self.new_username.get()
        password = self.new_password.get()
        role = self.new_role.get()
        
        if not all([username, password, role]):
            messagebox.showerror("Error", "All fields are required")
            return
        
        if username in users:
            messagebox.showerror("Error", "Username already exists")
            return
        
        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        users[username] = {"password": hashed_password, "role": role}
        save_users_to_file()  # Save the updated users to the file
        self.user_list.insert(tk.END, f"{username} ({role})")
        self.new_username.delete(0, tk.END)
        self.new_password.delete(0, tk.END)
        self.new_role.set("")
        
        messagebox.showinfo("Success", "User added successfully")
    
    def delete_user(self):
        selection = self.user_list.curselection()
        if not selection:
            messagebox.showerror("Error", "No user selected")
            return

        selected = self.user_list.get(selection[0])
        username = selected.split()[0]

        if username == self.current_user:
            messagebox.showerror("Error", "You cannot delete yourself.")
            return

        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the user '{username}'?")
        if not confirm:
            return

        del users[username]
        save_users_to_file()  # Save the updated users to the file
        self.user_list.delete(selection[0])
        messagebox.showinfo("Success", f"User '{username}' deleted successfully.")

    def export_all_data(self):
        if users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can export all patient data.")
            return

        if not os.path.exists('patients_data.json'):
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
            with open('patients_data.json', 'r') as f:
                data = json.load(f)
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"All patient data exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")

    def view_all_patients(self):
        if users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can view all patient records.")
            return

        if not os.path.exists('patients_data.json'):
            messagebox.showerror("Error", "No data available to display.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="All Saved Data", font=('Arial Rounded MT Bold', 18), 
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Main content
        content_frame = ttk.Frame(self.root, style='TFrame')
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
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, pady=5, ipady=10)  # Increased height

        canvas.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)

        # Create a frame inside the canvas to hold the table
        table_frame = ttk.Frame(canvas, style='TFrame')
        canvas.create_window((0, 0), window=table_frame, anchor="nw")

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        table_frame.bind("<Configure>", on_configure)

        # Load data
        with open('patients_data.json', 'r') as f:
            data = json.load(f)

        if not data:
            ttk.Label(table_frame, text="No patient data available", style='TLabel').pack(pady=20)
            return

        # Create a table-like structure
        columns = list(data[0].keys()) if data else []

        # Set a fixed width for all cells
        cell_width = 20

        # Create header row with filter entries
        header_row = ttk.Frame(table_frame, style='TFrame')
        header_row.pack(fill=tk.X)
        
        # Store filter variables
        self.filter_vars = {}
        
        for col in columns:
            header_cell_frame = ttk.Frame(header_row, style='TFrame')
            header_cell_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Column label
            ttk.Label(header_cell_frame, text=col, font=('Arial Rounded MT Bold', 10), 
                     anchor="center", borderwidth=1, relief="solid", width=cell_width).pack(fill=tk.X)
            
            # Filter entry - use combobox for fields with dropdown options, entry for others
            if col in DROPDOWN_OPTIONS:
                filter_var = tk.StringVar()
                self.filter_vars[col] = filter_var
                filter_entry = ttk.Combobox(header_cell_frame, textvariable=filter_var, 
                                          values=DROPDOWN_OPTIONS[col], state="readonly",
                                          font=('Arial', 8))
                filter_entry.set("")  # Empty default
                filter_entry.pack(fill=tk.X, padx=2, pady=2)
            else:
                filter_var = tk.StringVar()
                self.filter_vars[col] = filter_var
                filter_entry = ttk.Entry(header_cell_frame, textvariable=filter_var, 
                                       font=('Arial', 8))
                filter_entry.pack(fill=tk.X, padx=2, pady=2)

        # Create data rows with alternating colors
        self.data_rows = []
        for i, record in enumerate(data):
            row_frame = ttk.Frame(table_frame, style='TFrame')
            row_frame.pack(fill=tk.X)
            self.data_rows.append((row_frame, record))
            
            # Alternate row colors
            bg_color = '#f0f7ff' if i % 2 == 0 else '#e6f2ff'
            
            for col in columns:
                cell = tk.Label(row_frame, text=record.get(col, ""), font=('Arial', 10), 
                              anchor="center", borderwidth=1, relief="solid", width=cell_width,
                              bg=bg_color)
                cell.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipadx=5, ipady=5)
                # Bind double-click to view patient
                cell.bind("<Double-1>", lambda e, r=record: self.view_patient(r))

        # Filter and export buttons
        btn_frame = ttk.Frame(self.root, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="Apply Filters", command=self.apply_filters,
                   style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        
        # Export button will be shown only when filters are applied
        self.export_btn = ttk.Button(btn_frame, text="Export Filtered Data", 
                                    command=self.export_filtered_data,
                                    style='Accent.TButton')
        self.export_btn.pack_forget()  # Initially hidden
        
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu).pack(side=tk.LEFT, padx=10)

    def apply_filters(self):
        """Apply the current filters to the data rows."""
        if not hasattr(self, 'data_rows'):
            return
            
        # Get all filter values
        filter_values = {col: var.get().strip() for col, var in self.filter_vars.items()}
        
        # Check if any filter is applied
        filters_applied = any(value for value in filter_values.values())
        
        for row_frame, record in self.data_rows:
            match = True
            for col, value in filter_values.items():
                if value:  # Only apply filter if value is not empty
                    record_value = str(record.get(col, "")).lower()
                    if value.lower() not in record_value:
                        match = False
                        break
            
            # Show or hide row based on filter match
            if match:
                row_frame.pack(fill=tk.X)
            else:
                row_frame.pack_forget()
        
        # Show export button if filters are applied
        if filters_applied:
            self.export_btn.pack(side=tk.LEFT, padx=10)
        else:
            self.export_btn.pack_forget()

    def export_filtered_data(self):
        """Export the currently filtered data to Excel."""
        if users[self.current_user]["role"] != "admin":
            messagebox.showerror("Access Denied", "Only admins can export patient data.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Filtered Patient Data As"
        )
        if not file_path:
            return

        try:
            # Get all filter values
            filter_values = {col: var.get().strip() for col, var in self.filter_vars.items()}
            
            # Load all data
            with open('patients_data.json', 'r') as f:
                data = json.load(f)
            
            # Apply filters to data
            filtered_data = []
            for record in data:
                match = True
                for col, value in filter_values.items():
                    if value:  # Only apply filter if value is not empty
                        record_value = str(record.get(col, "")).lower()
                        if value.lower() not in record_value:
                            match = False
                            break
                if match:
                    filtered_data.append(record)
            
            if not filtered_data:
                messagebox.showerror("Error", "No data to export after filtering.")
                return
            
            df = pd.DataFrame(filtered_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Filtered patient data exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")

    def open_statistics_window(self):
        if users[self.current_user]["role"] not in ["admin", "editor"]:
            messagebox.showerror("Access Denied", "Only admins and editors can access statistics.")
            return

        self.clear_frame()

        # Header
        header_frame = ttk.Frame(self.root, style='TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(header_frame, text="Statistics", font=('Arial Rounded MT Bold', 18),
                  foreground='#0066cc').pack(side=tk.LEFT)

        # Main content
        stats_frame = ttk.Frame(self.root, style='TFrame')
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Create six selection fields
        selectable_fields = [field for field in DROPDOWN_OPTIONS.keys() if field not in ["Name", "File No"]]
        self.field_selections = []
        self.value_selections = []

        for i in range(6):
            input_frame = ttk.Frame(stats_frame, style='TFrame')
            input_frame.pack(fill=tk.X, pady=5)

            ttk.Label(input_frame, text=f"Select Field {i + 1}:", font=('Arial Rounded MT Bold', 12)).pack(side=tk.LEFT, padx=5)
            field_selection = ttk.Combobox(input_frame, values=selectable_fields, state="readonly", font=('Arial', 12), width=20)
            field_selection.pack(side=tk.LEFT, padx=5)
            self.field_selections.append(field_selection)

            ttk.Label(input_frame, text=f"Enter Value {i + 1}:", font=('Arial Rounded MT Bold', 12)).pack(side=tk.LEFT, padx=5)
            value_selection = ttk.Combobox(input_frame, state="normal", font=('Arial', 12), width=30)
            value_selection.pack(side=tk.LEFT, padx=5)
            self.value_selections.append(value_selection)

            def update_value_selection(event, index=i):
                selected_field = self.field_selections[index].get()
                if selected_field in DROPDOWN_OPTIONS:
                    self.value_selections[index].config(values=DROPDOWN_OPTIONS[selected_field], state="readonly")
                    self.value_selections[index].set("")
                else:
                    self.value_selections[index].config(values=[], state="normal")

            field_selection.bind("<<ComboboxSelected>>", update_value_selection)

        # Chart type selection
        chart_type_frame = ttk.Frame(stats_frame, style='TFrame')
        chart_type_frame.pack(fill=tk.X, pady=10)

        ttk.Label(chart_type_frame, text="Select Chart Type:", font=('Arial Rounded MT Bold', 12)).pack(side=tk.LEFT, padx=5)
        self.chart_type = tk.StringVar(value="Pie")
        ttk.Radiobutton(chart_type_frame, text="Pie Chart", variable=self.chart_type, value="Pie").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(chart_type_frame, text="Bar Chart", variable=self.chart_type, value="Bar").pack(side=tk.LEFT, padx=5)

        # Search button
        ttk.Button(stats_frame, text="Search", command=self.perform_detailed_statistics_search,
                   style='Accent.TButton').pack(fill=tk.X, pady=10)

        # Back button
        ttk.Button(stats_frame, text="Back to Menu", command=self.main_menu,
                   style='Accent.TButton').pack(fill=tk.X, pady=10)

    def perform_detailed_statistics_search(self):
        if not os.path.exists('patients_data.json'):
            messagebox.showerror("Error", "No patient data available")
            return

        with open('patients_data.json', 'r') as f:
            data = json.load(f)

        # Filter data based on selected fields and values
        filtered_data = data
        for i in range(6):
            field = self.field_selections[i].get()
            value = self.value_selections[i].get().strip()
            if field and value:
                filtered_data = [patient for patient in filtered_data if value.lower() in str(patient.get(field, "")).lower()]

        count = len(filtered_data)
        total = len(data)
        labels = [f'Matching Criteria', 'Others']
        values = [count, total - count]

        if count == 0:
            messagebox.showinfo("No Results", "No records found matching the given criteria.")
            return

        # Display result
        messagebox.showinfo("Statistics", f"Total Patients: {total}\nMatching Records: {count}")

        # Generate chart based on user selection
        chart_type = self.chart_type.get()
        if chart_type == "Pie":
            self.generate_pie_chart(labels, values, total, count)
        elif chart_type == "Bar":
            self.generate_bar_chart(labels, values, total, count)

    def generate_pie_chart(self, labels, values, total, count):
        fig, ax = plt.subplots(figsize=(6, 6))
        ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, colors=['#ff9999', '#66b3ff'])
        ax.set_title(f"Total Patients: {total}\nMatching Records: {count}")
        plt.show()

    def generate_bar_chart(self, labels, values, total, count):
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(labels, values, color=['#ff9999', '#66b3ff'])
        ax.set_title(f"Total Patients: {total}\nMatching Records: {count}")
        ax.set_ylabel("Count")
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = FNApp(root)
    root.mainloop()
