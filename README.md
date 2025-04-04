
First of all I want u to **Add and edit**: > I want drop lists to be coded to be editable by mej.esam can add and delete values >add a calculators button  >Add a button in home screen named F&N documentation, its function when clicked is to request from the user to enter the path on the users computer of an applications exe file to run , and that bath is saved and can not be edited by any user except mej.esam >Add a button in the home screen viewed only to mej.esam to configure system settings, no any other user can access it , its function is to make edits to a lot of software functions like drop lists to add and delete and edit its input values , change the bath of app that gets runed by clicking F&N documentation button , delete all local patients records folders and files


1.	Security & User Privileges System Current Issues:
Privilege checks are inconsistent
No proper audit logging
Password security can be improved
Improvements Needed:
A. Enhanced User Roles & Privileges
>Super Admin (mej.esam) Exclusive Privileges:
Full user management (add/edit/delete any user)
Delete single or all patient records permanently
Edit all dropdown lists (malignancies, symptoms, etc.)
Configure system settings (make edits to a lot of software functions like drop lists to add and delete and edit its input values , change the bath of app that gets runed by clicking F&N documentation button , delete all local patients records folders and files…etc)
Override any clinical warnings/protocols
Edit medication doses
Edit medication maximum doses
Restore from any backup
Export raw statistical data and figures
>Admin Users:
Add/edit patients (no deletion)
Manual data sync
Export data/backups can not restore
View/export statistics
Cannot modify protocols or doses
>Editors:
Add/edit patients
Generate reports
View statistics (no export)
Manual data sync
Access chemo protocols
can change their user password
Cannot modify system configurations
>Pharmacists:
Medication management , edit and change drugs dose and max doses
Dosage calculations 
Chemo stock management
Cannot modify patient demographics cant add, edit or delete
Cant view statistics
Can change their user passwords
>Viewers (Nurses):
Read-only access cant view statistics, chemo sheets or stocks
No export/print capabilities
Cannot change passwords
Cant calculate
Only can access search patients and chemo protocols 

Implementation Approach:
Create a privilege matrix JSON file defining exact permissions
Implement decorator pattern for privilege checks
Store last-modified-by info for all sensitive operations
Password policy: 12+ chars, special chars, 90-day expiration
B. Audit Logging
Log all privileged actions (user, timestamp, action)
Store logs both locally and in Firebase
Make logs viewable only by super admin
2.	Clinical Workflow Enhancements A. Medication Management System

BSA Calculator:
Multiple formula support (Mosteller/DuBois)
IV fluid rate calculator (used after calculating body surface area)
4000ml per BSA (-drug dilution if needed)
3000ml per BSA (-drug dilution if needed)
2000ml per BSA (-drug dilution if needed)
400ml per BSA (-drug dilution if needed)
Full maintenance (+deficit if needed) (-drug dilution if needed) and half maintenance (-drug dilution if needed)

Auto-save to patient record
Historical trend visualization
Dosage Management:
Protocol-based auto-calculation for all pediatric chemotherapies with special consideration for patients weight under 10 kg
Renal/hepatic dose adjustments
Cumulative dose tracking
Interactive dose rounding (nearest vial size)
Medication Timeline:
Visual treatment phases
Toxicity markers
Administration history
B. Lab Tracking System
Core Features:
Age-adjusted reference ranges
CTCAE toxicity grading
Nadir prediction
Transfusion triggers
Workflow Integration:
Pre-chemo clearance checklists
Critical value alerts
Protocol-specific requirements
Trend Visualization:
Sparkline graphs for key parameters
Exportable trend reports
Protocol phase correlation
C. Treatment Scheduling
Protocol Management:
Phase-based templates
Auto-generated calendars
Milestone tracking
Appointment System:
Integrated with treatment phases
Reminder system
Conflict checking

D. Chemotherapy extravasation managements
3.	Data Management Improvements 
A. Backup & Sync
Enhanced Sync Logic:
Conflict resolution protocols
Differential sync
Offline mode support
Backup Types:
Full system backups
Patient-specific snapshots
Automated nightly backups
B. Data Validation
Input Validation:
Field-level validation rules
Cross-field validation (e.g., age vs diagnosis date)
Protocol-specific requirements
Clinical Safety Checks:
Dose-range validation
Drug-drug interaction alerts
Cumulative dose warnings massage
4.	UI/UX Improvements A. Navigation & Workflow
Keyboard Shortcuts:
F2: New patient
Ctrl+F: Search
Alt+S: Save
Form Improvements:
Smart tab ordering
Auto-focus critical fields
Inline validation messages
B. Visual Design
Clinical Status Indicators:
Color-coded treatment phases
Toxicity severity badges
Protocol compliance markers
Information Hierarchy:
Priority-based layout
Clinical decision support prominence
Progressive disclosure
C. Performance Optimizations
Lazy Loading:
Paginated patient lists
On-demand image loading
Background data prefetching
Memory Management:
Clean up unused resources
Optimized image handling
Efficient data structures
5.	Technical Architecture A. Code Organization
Separation of Concerns:
Business logic layer
Data access layer
Presentation layer
Dependency Injection:
Configurable service providers
Mockable interfaces for testing
B. Testing Strategy
Unit Tests:
Core calculations
Validation logic
Privilege checks
Integration Tests:
Firebase sync
Google Drive operations
Clinical workflows
C. Error Handling
Recovery Protocols:
Transaction rollbacks
Auto-reconnect for cloud services
Graceful degradation
6.	Statistics & Reporting A. Enhanced Analytics
Clinical Metrics:
Protocol compliance rates
Toxicity profiles
Survival analysis
Operational Metrics:
Patient volume trends
Resource utilization
Follow-up rates
B. Visualization Tools
Interactive Charts:
Filterable by time/protocol
Export as PNG/PDF
Protocol benchmarking
Custom Reports:
Protocol-specific templates
Survivorship reports
Quality metrics
7.	Configuration System A. Super Admin Controls
Dynamic Dropdown Management:
Add/remove malignancy types
Edit symptom lists
Configure risk groups
System Settings:
EXE path configuration
Backup locations
Default protocols
B. Clinical Configuration
Protocol Management:
Dose limits
Required labs
Phase definitions
Alert Thresholds:
Critical lab values
Dose modification rules
Follow-up intervals
Implementation Roadmap Phase 1: Security & Core Architecture
Implement privilege matrix system
Build audit logging framework
Set up dependency injection
Phase 2: Clinical Workflows
Develop medication management
Implement lab tracking
Build treatment scheduler
Phase 3: UI/UX Overhaul
Redesign navigation
Add clinical status indicators
Optimize form workflows
Phase 4: Data & Analytics
Enhance statistics module
Build reporting engine
Implement advanced visualizations
Phase 5: Deployment & Monitoring
Performance benchmarking
User training materials
Usage analytics
Key Considerations Clinical Safety:
All clinical calculations require dual verification
Maintain complete audit trails
Never allow silent overrides
Data Integrity:
Cryptographic hashing for sensitive data
Write-once audit logs
Regular consistency checks
Regulatory Compliance:
HIPAA/GDPR considerations
21 CFR Part 11 for electronic records
Data retention policies

we already started editting some parts and we edited till this point in the code coming next

import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import sys
import os
import bcrypt
from datetime import datetime
import socket
import subprocess
from cryptography.fernet import Fernet

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
    # ... (other dropdown options remain the same)
}

# User roles and permissions matrix
USER_ROLES = {
    "super_admin": {
        "manage_users": True,
        "delete_patients": True,
        "edit_dropdowns": True,
        "configure_system": True,
        "override_warnings": True,
        "edit_meds": True,
        "restore_backups": True,
        "export_stats": True,
        "change_password": True
    },
    "admin": {
        "manage_users": False,
        "delete_patients": False,
        "edit_dropdowns": False,
        "configure_system": False,
        "override_warnings": False,
        "edit_meds": False,
        "restore_backups": False,
        "export_stats": True,
        "change_password": True
    },
    "editor": {
        "manage_users": False,
        "delete_patients": False,
        "edit_dropdowns": False,
        "configure_system": False,
        "override_warnings": False,
        "edit_meds": False,
        "restore_backups": False,
        "export_stats": False,
        "change_password": True
    },
    "pharmacist": {
        "manage_users": False,
        "delete_patients": False,
        "edit_dropdowns": False,
        "configure_system": False,
        "override_warnings": False,
        "edit_meds": True,
        "restore_backups": False,
        "export_stats": False,
        "change_password": True
    },
    "viewer": {
        "manage_users": False,
        "delete_patients": False,
        "edit_dropdowns": False,
        "configure_system": False,
        "override_warnings": False,
        "edit_meds": False,
        "restore_backups": False,
        "export_stats": False,
        "change_password": False
    }
}

class PrivilegeChecker:
    """Decorator class for privilege checking"""
    @staticmethod
    def check_privilege(privilege_name):
        def decorator(func):
            def wrapper(self, *args, **kwargs):
                if not self.current_user:
                    messagebox.showerror("Access Denied", "Not logged in")
                    return
                
                user_role = self.users[self.current_user]["role"]
                if USER_ROLES.get(user_role, {}).get(privilege_name, False):
                    return func(self, *args, **kwargs)
                else:
                    messagebox.showerror("Access Denied", 
                                       f"You don't have permission to {privilege_name.replace('_', ' ')}")
            return wrapper
        return decorator

class OncologyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OncoCare - Pediatric Oncology Patient Management System")
        self.root.geometry("1200x800")
        self.root.state('zoomed')
        
        # Initialize variables
        self.current_user = None
        self.fn_doc_path = None
        self.load_settings()
        self.load_users()
        self.load_patient_data()
        
        # Setup styles and UI
        self.setup_styles()
        self.setup_status_bar()
        self.setup_login_screen()
        
    def load_settings(self):
        """Load application settings from encrypted file"""
        try:
            if os.path.exists('settings.enc'):
                with open('settings.enc', 'rb') as f:
                    encrypted_data = f.read()
                    decrypted_data = cipher_suite.decrypt(encrypted_data)
                    settings = json.loads(decrypted_data.decode())
                    self.fn_doc_path = settings.get('fn_doc_path')
                    DROPDOWN_OPTIONS.update(settings.get('dropdown_options', DROPDOWN_OPTIONS))
        except Exception as e:
            print(f"Error loading settings: {e}")
            # Generate new encryption key if not exists
            try:
                with open('encryption_key.key', 'rb') as key_file:
                    ENCRYPTION_KEY = key_file.read()
            except:
                ENCRYPTION_KEY = Fernet.generate_key()
                with open('encryption_key.key', 'wb') as key_file:
                    key_file.write(ENCRYPTION_KEY)
            cipher_suite = Fernet(ENCRYPTION_KEY)
    
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
    
    def main_menu(self):
        """Display the main menu with organized buttons"""
        self.clear_frame()

        # Main container
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Left side with logo
        left_frame = tk.Frame(main_frame, bg='#3498db')
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        # Right side with menu buttons
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Menu form container
        form_container = ttk.Frame(right_frame, style='TFrame')
        form_container.place(relx=0.5, rely=0.5, anchor='center')

        # Button grid
        btn_frame = ttk.Frame(form_container, style='TFrame')
        btn_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        # Add F&N Documentation button (visible to all)
        ttk.Button(btn_frame, text="F&N Documentation", 
                 command=self.run_fn_documentation,
                 style='Purple.TButton').pack(fill=tk.X, pady=5)
        
        # Add Calculators button (visible to pharmacists and above)
        if self.has_privilege("edit_meds"):
            ttk.Button(btn_frame, text="Calculators", 
                     command=self.show_calculators,
                     style='Blue.TButton').pack(fill=tk.X, pady=5)
        
        # Add Settings button (only for mej.esam)
        if self.current_user == "mej.esam":
            ttk.Button(btn_frame, text="System Settings", 
                     command=self.open_settings,
                     style='Red.TButton').pack(fill=tk.X, pady=5)
    
    @PrivilegeChecker.check_privilege("configure_system")
    def open_settings(self):
        """Open system settings window"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("System Settings")
        settings_window.geometry("800x600")
        
        # Notebook for tabs
        notebook = ttk.Notebook(settings_window)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Dropdown Management Tab
        dropdown_tab = ttk.Frame(notebook)
        notebook.add(dropdown_tab, text="Dropdown Options")
        self.setup_dropdown_management(dropdown_tab)
        
        # System Configuration Tab
        system_tab = ttk.Frame(notebook)
        notebook.add(system_tab, text="System Config")
        self.setup_system_config(system_tab)
        
        # Close button
        btn_frame = ttk.Frame(settings_window)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="Close", command=settings_window.destroy).pack()
    
    def setup_dropdown_management(self, parent):
        """Setup dropdown options management interface"""
        # Category selection
        category_frame = ttk.Frame(parent)
        category_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(category_frame, text="Category:").pack(side=tk.LEFT)
        self.dropdown_category = ttk.Combobox(category_frame, 
                                            values=list(DROPDOWN_OPTIONS.keys()))
        self.dropdown_category.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.dropdown_category.bind("<<ComboboxSelected>>", self.update_dropdown_display)
        
        # Current options display
        options_frame = ttk.LabelFrame(parent, text="Current Options")
        options_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.options_listbox = tk.Listbox(options_frame, selectmode=tk.MULTIPLE)
        self.options_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Edit controls
        edit_frame = ttk.Frame(parent)
        edit_frame.pack(fill=tk.X, pady=10)
        
        self.new_option_entry = ttk.Entry(edit_frame)
        self.new_option_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(edit_frame, text="Add", 
                  command=self.add_dropdown_option,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(edit_frame, text="Remove Selected", 
                  command=self.remove_dropdown_options,
                  style='Red.TButton').pack(side=tk.LEFT, padx=5)
        
        # Initialize display
        if self.dropdown_category.get():
            self.update_dropdown_display()
    
    def update_dropdown_display(self, event=None):
        """Update the listbox with current dropdown options"""
        category = self.dropdown_category.get()
        self.options_listbox.delete(0, tk.END)
        
        for option in DROPDOWN_OPTIONS.get(category, []):
            self.options_listbox.insert(tk.END, option)
    
    def add_dropdown_option(self):
        """Add a new option to the selected dropdown"""
        category = self.dropdown_category.get()
        new_option = self.new_option_entry.get().strip()
        
        if not category:
            messagebox.showerror("Error", "Please select a category first")
            return
            
        if not new_option:
            messagebox.showerror("Error", "Please enter an option to add")
            return
            
        if new_option in DROPDOWN_OPTIONS[category]:
            messagebox.showerror("Error", "This option already exists")
            return
            
        DROPDOWN_OPTIONS[category].append(new_option)
        self.update_dropdown_display()
        self.new_option_entry.delete(0, tk.END)
        self.save_settings()
    
    def remove_dropdown_options(self):
        """Remove selected options from dropdown"""
        category = self.dropdown_category.get()
        selected = self.options_listbox.curselection()
        
        if not category:
            messagebox.showerror("Error", "Please select a category first")
            return
            
        if not selected:
            messagebox.showerror("Error", "Please select options to remove")
            return
            
        # Remove in reverse order to avoid index issues
        for i in reversed(selected):
            del DROPDOWN_OPTIONS[category][i]
            
        self.update_dropdown_display()
        self.save_settings()
    
    def setup_system_config(self, parent):
        """Setup system configuration interface"""
        # F&N Documentation Path
        fn_frame = ttk.LabelFrame(parent, text="F&N Documentation Path")
        fn_frame.pack(fill=tk.X, pady=10)
        
        self.fn_path_entry = ttk.Entry(fn_frame)
        self.fn_path_entry.insert(0, self.fn_doc_path or "")
        self.fn_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(fn_frame, text="Browse", 
                  command=self.browse_fn_path).pack(side=tk.LEFT, padx=5)
        
        # Save button
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="Save Settings", 
                  command=self.save_system_config,
                  style='Green.TButton').pack()
    
    def browse_fn_path(self):
        """Browse for F&N documentation executable"""
        path = filedialog.askopenfilename(title="Select F&N Documentation Executable",
                                        filetypes=[("Executable files", "*.exe")])
        if path:
            self.fn_path_entry.delete(0, tk.END)
            self.fn_path_entry.insert(0, path)
    
    def save_system_config(self):
        """Save system configuration"""
        self.fn_doc_path = self.fn_path_entry.get()
        self.save_settings()
        messagebox.showinfo("Success", "System configuration saved")
    
    def run_fn_documentation(self):
        """Run the F&N documentation application"""
        if not self.fn_doc_path:
            messagebox.showerror("Error", "F&N Documentation path not configured")
            return
            
        try:
            subprocess.Popen(self.fn_doc_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not run F&N Documentation: {e}")
    
    def show_calculators(self):
        """Show calculators menu"""
        calculators_window = tk.Toplevel(self.root)
        calculators_window.title("Clinical Calculators")
        calculators_window.geometry("600x400")
        
        # BSA Calculator
        ttk.Button(calculators_window, text="Body Surface Area (BSA) Calculator",
                  command=self.open_bsa_calculator,
                  style='Blue.TButton').pack(fill=tk.X, pady=5)
        
        # IV Fluid Calculator
        ttk.Button(calculators_window, text="IV Fluid Rate Calculator",
                  command=self.open_iv_fluid_calculator,
                  style='Blue.TButton').pack(fill=tk.X, pady=5)
        
        # Maintenance Fluid Calculator
        ttk.Button(calculators_window, text="Maintenance Fluid Calculator",
                  command=self.open_maintenance_fluid_calculator,
                  style='Blue.TButton').pack(fill=tk.X, pady=5)
    
    def open_bsa_calculator(self):
        """Open BSA calculator"""
        bsa_window = tk.Toplevel(self.root)
        bsa_window.title("BSA Calculator")
        bsa_window.geometry("400x300")
        
        # Input fields
        ttk.Label(bsa_window, text="Height (cm):").pack()
        height_entry = ttk.Entry(bsa_window)
        height_entry.pack()
        
        ttk.Label(bsa_window, text="Weight (kg):").pack()
        weight_entry = ttk.Entry(bsa_window)
        weight_entry.pack()
        
        ttk.Label(bsa_window, text="Formula:").pack()
        formula_var = tk.StringVar(value="Mosteller")
        formula_menu = ttk.OptionMenu(bsa_window, formula_var, 
                                    "Mosteller", "Mosteller", "DuBois", "Haycock")
        formula_menu.pack()
        
        # Result display
        result_label = ttk.Label(bsa_window, text="BSA: ")
        result_label.pack(pady=10)
        
        def calculate():
            try:
                height = float(height_entry.get())
                weight = float(weight_entry.get())
                formula = formula_var.get()
                
                if formula == "Mosteller":
                    bsa = math.sqrt(height * weight / 3600)
                elif formula == "DuBois":
                    bsa = 0.007184 * (height**0.725) * (weight**0.425)
                elif formula == "Haycock":
                    bsa = 0.024265 * (height**0.3964) * (weight**0.5378)
                
                result_label.config(text=f"BSA: {bsa:.4f} m² ({formula} formula)")
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers")
        
        ttk.Button(bsa_window, text="Calculate", 
                  command=calculate,
                  style='Green.TButton').pack()
    
    def has_privilege(self, privilege_name):
        """Check if current user has a specific privilege"""
        if not self.current_user:
            return False
            
        user_role = self.users[self.current_user]["role"]
        return USER_ROLES.get(user_role, {}).get(privilege_name, False)
    
    # ... (other existing methods remain the same)

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
        
        # Initialize medication data if not exists
        if "MEDICATIONS" not in patient:
            patient["MEDICATIONS"] = {
                "current": [],
                "history": [],
                "allergies": [],
                "interactions": []
            }
            self.save_patient_data()
        
        # Notebook for tabs
        notebook = ttk.Notebook(self.medication_window)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Current Medications Tab
        current_tab = ttk.Frame(notebook)
        notebook.add(current_tab, text="Current Medications")
        self.setup_current_medications_tab(current_tab, patient)
        
        # Medication History Tab
        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="Medication History")
        self.setup_medication_history_tab(history_tab, patient)
        
        # Allergies Tab
        allergies_tab = ttk.Frame(notebook)
        notebook.add(allergies_tab, text="Allergies")
        self.setup_allergies_tab(allergies_tab, patient)
        
        # Interactions Tab
        interactions_tab = ttk.Frame(notebook)
        notebook.add(interactions_tab, text="Interactions")
        self.setup_interactions_tab(interactions_tab, patient)
        
        # Calculators Tab
        if self.has_privilege("edit_meds"):
            calculators_tab = ttk.Frame(notebook)
            notebook.add(calculators_tab, text="Calculators")
            self.setup_medication_calculators_tab(calculators_tab, patient)
        
        # Close button
        btn_frame = ttk.Frame(self.medication_window)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="Close", command=self.medication_window.destroy).pack()

    def setup_current_medications_tab(self, parent, patient):
        """Setup current medications tab with enhanced features"""
        # Create scrollable frame
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container)
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
        ttk.Label(scrollable_frame, text="Current Medications", 
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Add new medication button
        ttk.Button(scrollable_frame, text="Add New Medication", 
                  command=lambda: self.add_new_medication(patient),
                  style='Green.TButton').pack(pady=10)
        
        # Display current medications
        if not patient["MEDICATIONS"]["current"]:
            ttk.Label(scrollable_frame, text="No current medications").pack(pady=20)
        else:
            for med in patient["MEDICATIONS"]["current"]:
                self.create_medication_card(scrollable_frame, med, patient)

    def create_medication_card(self, parent, medication, patient):
        """Create an enhanced medication card display"""
        card_frame = ttk.Frame(parent, style='Card.TFrame', padding=10)
        card_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Header with name and status
        header_frame = ttk.Frame(card_frame)
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text=medication["name"], 
                 font=('Helvetica', 12, 'bold')).pack(side=tk.LEFT)
        
        # Status indicators
        status_frame = ttk.Frame(header_frame)
        status_frame.pack(side=tk.RIGHT)
        
        if medication.get("prn", False):
            ttk.Label(status_frame, text="PRN", 
                     foreground="blue").pack(side=tk.LEFT, padx=2)
        if medication.get("abnormal", False):
            ttk.Label(status_frame, text="Abnormal", 
                     foreground="red").pack(side=tk.LEFT, padx=2)
        
        # Details frame
        details_frame = ttk.Frame(card_frame)
        details_frame.pack(fill=tk.X, pady=5)
        
        # Left column - prescription details
        left_frame = ttk.Frame(details_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(left_frame, text=f"Dose: {medication['dose']} {medication.get('unit', '')}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Frequency: {medication['frequency']}").pack(anchor="w")
        ttk.Label(left_frame, text=f"Route: {medication['route']}").pack(anchor="w")
        
        # Right column - dates and prescriber
        right_frame = ttk.Frame(details_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.X)
        
        ttk.Label(right_frame, text=f"Start: {medication['start_date']}").pack(anchor="e")
        if 'end_date' in medication:
            ttk.Label(right_frame, text=f"End: {medication['end_date']}").pack(anchor="e")
        ttk.Label(right_frame, text=f"Prescriber: {medication.get('prescriber', '')}").pack(anchor="e")
        
        # Administration history button
        admin_frame = ttk.Frame(card_frame)
        admin_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(admin_frame, text="View Administration History", 
                  command=lambda: self.show_administration_history(medication, patient),
                  style='Blue.TButton').pack(side=tk.LEFT)
        
        # Action buttons
        action_frame = ttk.Frame(card_frame)
        action_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(action_frame, text="Edit", 
                  command=lambda: self.edit_medication(medication, patient),
                  style='Blue.TButton').pack(side=tk.LEFT, padx=2)
        
        ttk.Button(action_frame, text="Discontinue", 
                  command=lambda: self.discontinue_medication(medication, patient),
                  style='Red.TButton').pack(side=tk.LEFT, padx=2)
        
        ttk.Button(action_frame, text="Administer", 
                  command=lambda: self.record_administration(medication, patient),
                  style='Green.TButton').pack(side=tk.LEFT, padx=2)
        
        # Only show dose adjustment for pharmacists and above
        if self.has_privilege("edit_meds"):
            ttk.Button(action_frame, text="Adjust Dose", 
                      command=lambda: self.adjust_medication_dose(medication, patient),
                      style='Yellow.TButton').pack(side=tk.LEFT, padx=2)

    def add_new_medication(self, patient):
        """Open dialog to add new medication with enhanced fields"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Medication")
        dialog.geometry("700x600")
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook for sections
        notebook = ttk.Notebook(form_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Prescription Info Tab
        presc_tab = ttk.Frame(notebook)
        notebook.add(presc_tab, text="Prescription")
        
        # Medication name
        ttk.Label(presc_tab, text="Medication Name:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = ttk.Entry(presc_tab)
        name_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Dose
        ttk.Label(presc_tab, text="Dose:").grid(row=1, column=0, sticky="w", pady=5)
        dose_frame = ttk.Frame(presc_tab)
        dose_frame.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        dose_entry = ttk.Entry(dose_frame)
        dose_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        unit_var = tk.StringVar(value="mg")
        unit_combo = ttk.Combobox(dose_frame, textvariable=unit_var, 
                                values=["mg", "mcg", "g", "mg/m²", "mcg/m²", "IU", "mL", "mg/kg"])
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        # Frequency
        ttk.Label(presc_tab, text="Frequency:").grid(row=2, column=0, sticky="w", pady=5)
        freq_var = tk.StringVar(value="Daily")
        freq_combo = ttk.Combobox(presc_tab, textvariable=freq_var,
                                values=["Daily", "BID", "TID", "QID", "QHS", "QOD", 
                                       "Weekly", "Monthly", "PRN", "Other"])
        freq_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Route
        ttk.Label(presc_tab, text="Route:").grid(row=3, column=0, sticky="w", pady=5)
        route_var = tk.StringVar(value="PO")
        route_combo = ttk.Combobox(presc_tab, textvariable=route_var,
                                 values=["PO", "IV", "IM", "SC", "Topical", "PR", "SL", "Other"])
        route_combo.grid(row=3, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Start date
        ttk.Label(presc_tab, text="Start Date:").grid(row=4, column=0, sticky="w", pady=5)
        start_frame = ttk.Frame(presc_tab)
        start_frame.grid(row=4, column=1, sticky="ew", pady=5, padx=5)
        
        start_entry = ttk.Entry(start_frame)
        start_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        start_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(start_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(start_entry)).pack(side=tk.LEFT, padx=5)
        
        # Prescriber
        ttk.Label(presc_tab, text="Prescriber:").grid(row=5, column=0, sticky="w", pady=5)
        prescriber_entry = ttk.Entry(presc_tab)
        prescriber_entry.insert(0, self.current_user)
        prescriber_entry.grid(row=5, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Indication
        ttk.Label(presc_tab, text="Indication:").grid(row=6, column=0, sticky="w", pady=5)
        indication_entry = ttk.Entry(presc_tab)
        indication_entry.grid(row=6, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # PRN checkbox
        prn_var = tk.BooleanVar()
        prn_check = ttk.Checkbutton(presc_tab, text="PRN Medication", variable=prn_var)
        prn_check.grid(row=7, column=1, sticky="w", pady=5, padx=5)
        
        # Additional Info Tab
        info_tab = ttk.Frame(notebook)
        notebook.add(info_tab, text="Additional Info")
        
        # Notes
        ttk.Label(info_tab, text="Notes:").pack(pady=5)
        notes_text = tk.Text(info_tab, height=8, wrap=tk.WORD)
        notes_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def save_medication():
            new_med = {
                "name": name_entry.get(),
                "dose": dose_entry.get(),
                "unit": unit_var.get(),
                "frequency": freq_var.get(),
                "route": route_var.get(),
                "start_date": start_entry.get(),
                "prescriber": prescriber_entry.get(),
                "indication": indication_entry.get(),
                "prn": prn_var.get(),
                "notes": notes_text.get("1.0", tk.END).strip(),
                "administrations": [],
                "created_by": self.current_user,
                "created_date": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            }
            
            if not new_med["name"] or not new_med["dose"]:
                messagebox.showerror("Error", "Medication name and dose are required")
                return
                
            patient["MEDICATIONS"]["current"].append(new_med)
            self.save_patient_data()
            
            # Refresh medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_medication,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def edit_medication(self, medication, patient):
        """Edit an existing medication with enhanced fields"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit {medication['name']}")
        dialog.geometry("700x600")
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook for sections
        notebook = ttk.Notebook(form_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Prescription Info Tab
        presc_tab = ttk.Frame(notebook)
        notebook.add(presc_tab, text="Prescription")
        
        # Medication name
        ttk.Label(presc_tab, text="Medication Name:").grid(row=0, column=0, sticky="w", pady=5)
        name_entry = ttk.Entry(presc_tab)
        name_entry.insert(0, medication["name"])
        name_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Dose
        ttk.Label(presc_tab, text="Dose:").grid(row=1, column=0, sticky="w", pady=5)
        dose_frame = ttk.Frame(presc_tab)
        dose_frame.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        dose_entry = ttk.Entry(dose_frame)
        dose_entry.insert(0, medication["dose"])
        dose_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        unit_var = tk.StringVar(value=medication.get("unit", "mg"))
        unit_combo = ttk.Combobox(dose_frame, textvariable=unit_var, 
                                values=["mg", "mcg", "g", "mg/m²", "mcg/m²", "IU", "mL", "mg/kg"])
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        # Frequency
        ttk.Label(presc_tab, text="Frequency:").grid(row=2, column=0, sticky="w", pady=5)
        freq_var = tk.StringVar(value=medication.get("frequency", "Daily"))
        freq_combo = ttk.Combobox(presc_tab, textvariable=freq_var,
                                values=["Daily", "BID", "TID", "QID", "QHS", "QOD", 
                                       "Weekly", "Monthly", "PRN", "Other"])
        freq_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Route
        ttk.Label(presc_tab, text="Route:").grid(row=3, column=0, sticky="w", pady=5)
        route_var = tk.StringVar(value=medication.get("route", "PO"))
        route_combo = ttk.Combobox(presc_tab, textvariable=route_var,
                                 values=["PO", "IV", "IM", "SC", "Topical", "PR", "SL", "Other"])
        route_combo.grid(row=3, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Start date
        ttk.Label(presc_tab, text="Start Date:").grid(row=4, column=0, sticky="w", pady=5)
        start_frame = ttk.Frame(presc_tab)
        start_frame.grid(row=4, column=1, sticky="ew", pady=5, padx=5)
        
        start_entry = ttk.Entry(start_frame)
        start_entry.insert(0, medication.get("start_date", datetime.now().strftime("%d/%m/%Y")))
        start_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(start_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(start_entry)).pack(side=tk.LEFT, padx=5)
        
        # End date if exists
        if 'end_date' in medication:
            ttk.Label(presc_tab, text="End Date:").grid(row=5, column=0, sticky="w", pady=5)
            end_frame = ttk.Frame(presc_tab)
            end_frame.grid(row=5, column=1, sticky="ew", pady=5, padx=5)
            
            end_entry = ttk.Entry(end_frame)
            end_entry.insert(0, medication["end_date"])
            end_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            ttk.Button(end_frame, text="📅", width=3,
                      command=lambda: self.show_calendar(end_entry)).pack(side=tk.LEFT, padx=5)
        
        # Prescriber
        ttk.Label(presc_tab, text="Prescriber:").grid(row=6, column=0, sticky="w", pady=5)
        prescriber_entry = ttk.Entry(presc_tab)
        prescriber_entry.insert(0, medication.get("prescriber", self.current_user))
        prescriber_entry.grid(row=6, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # Indication
        ttk.Label(presc_tab, text="Indication:").grid(row=7, column=0, sticky="w", pady=5)
        indication_entry = ttk.Entry(presc_tab)
        indication_entry.insert(0, medication.get("indication", ""))
        indication_entry.grid(row=7, column=1, sticky="ew", pady=5, padx=5, columnspan=2)
        
        # PRN checkbox
        prn_var = tk.BooleanVar(value=medication.get("prn", False))
        prn_check = ttk.Checkbutton(presc_tab, text="PRN Medication", variable=prn_var)
        prn_check.grid(row=8, column=1, sticky="w", pady=5, padx=5)
        
        # Additional Info Tab
        info_tab = ttk.Frame(notebook)
        notebook.add(info_tab, text="Additional Info")
        
        # Notes
        ttk.Label(info_tab, text="Notes:").pack(pady=5)
        notes_text = tk.Text(info_tab, height=8, wrap=tk.WORD)
        notes_text.insert("1.0", medication.get("notes", ""))
        notes_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def save_changes():
            # Update medication details
            medication["name"] = name_entry.get()
            medication["dose"] = dose_entry.get()
            medication["unit"] = unit_var.get()
            medication["frequency"] = freq_var.get()
            medication["route"] = route_var.get()
            medication["start_date"] = start_entry.get()
            if 'end_date' in medication:
                medication["end_date"] = end_entry.get()
            medication["prescriber"] = prescriber_entry.get()
            medication["indication"] = indication_entry.get()
            medication["prn"] = prn_var.get()
            medication["notes"] = notes_text.get("1.0", tk.END).strip()
            medication["last_modified_by"] = self.current_user
            medication["last_modified_date"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            
            self.save_patient_data()
            
            # Refresh medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_changes,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Only show delete button for admins
        if self.has_privilege("delete_patients"):
            ttk.Button(btn_frame, text="Delete", 
                      command=lambda: self.delete_medication(medication, patient, dialog),
                      style='Red.TButton').pack(side=tk.LEFT, padx=5)

    def adjust_medication_dose(self, medication, patient):
        """Adjust medication dose with renal/hepatic considerations"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Adjust Dose - {medication['name']}")
        dialog.geometry("500x400")
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current dose display
        ttk.Label(form_frame, 
                 text=f"Current Dose: {medication['dose']} {medication.get('unit', '')}",
                 font=('Helvetica', 11, 'bold')).pack(pady=5)
        
        # New dose
        ttk.Label(form_frame, text="New Dose:").pack(pady=5)
        new_dose_frame = ttk.Frame(form_frame)
        new_dose_frame.pack(fill=tk.X, pady=5)
        
        new_dose_entry = ttk.Entry(new_dose_frame)
        new_dose_entry.insert(0, medication["dose"])
        new_dose_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        unit_var = tk.StringVar(value=medication.get("unit", "mg"))
        unit_combo = ttk.Combobox(new_dose_frame, textvariable=unit_var, 
                                values=["mg", "mcg", "g", "mg/m²", "mcg/m²", "IU", "mL", "mg/kg"],
                                state="readonly")
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        # Adjustment reason
        ttk.Label(form_frame, text="Adjustment Reason:").pack(pady=5)
        reason_var = tk.StringVar()
        reason_combo = ttk.Combobox(form_frame, textvariable=reason_var,
                                  values=["Renal impairment", "Hepatic impairment", 
                                         "Toxicity", "Other clinical reason"])
        reason_combo.pack(fill=tk.X, pady=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_text = tk.Text(form_frame, height=5, wrap=tk.WORD)
        notes_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def save_adjustment():
            new_dose = new_dose_entry.get()
            if not new_dose:
                messagebox.showerror("Error", "Please enter a new dose")
                return
                
            # Create adjustment record
            adjustment = {
                "date": datetime.now().strftime("%d/%m/%Y"),
                "previous_dose": medication["dose"],
                "new_dose": new_dose,
                "unit": unit_var.get(),
                "reason": reason_var.get(),
                "notes": notes_text.get("1.0", tk.END).strip(),
                "adjusted_by": self.current_user
            }
            
            # Update medication
            medication["dose"] = new_dose
            medication["unit"] = unit_var.get()
            
            if "dose_adjustments" not in medication:
                medication["dose_adjustments"] = []
            medication["dose_adjustments"].append(adjustment)
            
            self.save_patient_data()
            messagebox.showinfo("Success", "Dose adjustment saved")
            dialog.destroy()
            
            # Refresh medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
        
        ttk.Button(btn_frame, text="Save Adjustment", 
                  command=save_adjustment,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", 
                  command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def discontinue_medication(self, medication, patient):
        """Discontinue a medication with reason"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Discontinue {medication['name']}")
        dialog.geometry("400x300")
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Confirmation
        ttk.Label(form_frame, 
                 text=f"Discontinue {medication['name']} {medication['dose']}{medication.get('unit', '')}?",
                 font=('Helvetica', 11, 'bold')).pack(pady=5)
        
        # Discontinuation reason
        ttk.Label(form_frame, text="Reason:").pack(pady=5)
        reason_var = tk.StringVar()
        reason_combo = ttk.Combobox(form_frame, textvariable=reason_var,
                                  values=["Completed course", "Adverse reaction", 
                                         "Ineffective", "Patient request", "Other"])
        reason_combo.pack(fill=tk.X, pady=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_entry = tk.Text(form_frame, height=5, wrap=tk.WORD)
        notes_entry.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def confirm_discontinue():
            if not reason_var.get():
                messagebox.showerror("Error", "Please select a reason")
                return
                
            # Set end date
            medication["end_date"] = datetime.now().strftime("%d/%m/%Y")
            
            # Add discontinuation record
            discontinuation = {
                "date": medication["end_date"],
                "reason": reason_var.get(),
                "notes": notes_entry.get("1.0", tk.END).strip(),
                "discontinued_by": self.current_user
            }
            
            if "discontinuations" not in medication:
                medication["discontinuations"] = []
            medication["discontinuations"].append(discontinuation)
            
            # Move to history
            patient["MEDICATIONS"]["history"].append(medication)
            patient["MEDICATIONS"]["current"].remove(medication)
            
            self.save_patient_data()
            messagebox.showinfo("Success", "Medication discontinued")
            dialog.destroy()
            
            # Refresh medication window
            if hasattr(self, 'medication_window') and self.medication_window.winfo_exists():
                self.medication_window.destroy()
                self.open_medication_management(patient["FILE NUMBER"])
        
        ttk.Button(btn_frame, text="Confirm Discontinue", 
                  command=confirm_discontinue,
                  style='Red.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", 
                  command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def record_administration(self, medication, patient):
        """Record administration of a medication with enhanced fields"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Administer {medication['name']}")
        dialog.geometry("500x500")
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Medication info
        ttk.Label(form_frame, 
                 text=f"{medication['name']} {medication['dose']}{medication.get('unit', '')}",
                 font=('Helvetica', 11, 'bold')).pack(pady=5)
        
        # Date and time
        datetime_frame = ttk.Frame(form_frame)
        datetime_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(datetime_frame, text="Date:").pack(side=tk.LEFT)
        date_entry = ttk.Entry(datetime_frame)
        date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(datetime_frame, text="📅", width=3,
                  command=lambda: self.show_calendar(date_entry)).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(datetime_frame, text="Time:").pack(side=tk.LEFT, padx=10)
        time_entry = ttk.Entry(datetime_frame, width=8)
        time_entry.insert(0, datetime.now().strftime("%H:%M"))
        time_entry.pack(side=tk.LEFT)
        
        # Administered by
        ttk.Label(form_frame, text="Administered by:").pack(pady=5)
        admin_by_entry = ttk.Entry(form_frame)
        admin_by_entry.insert(0, self.current_user)
        admin_by_entry.pack(fill=tk.X, pady=5)
        
        # Administration details
        details_frame = ttk.LabelFrame(form_frame, text="Administration Details")
        details_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Route (default to medication route but can be overridden)
        ttk.Label(details_frame, text="Route:").grid(row=0, column=0, sticky="w", pady=5)
        route_var = tk.StringVar(value=medication.get("route", ""))
        route_combo = ttk.Combobox(details_frame, textvariable=route_var,
                                 values=["PO", "IV", "IM", "SC", "Topical", "PR", "SL", "Other"])
        route_combo.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        # Site (for injections)
        ttk.Label(details_frame, text="Site:").grid(row=1, column=0, sticky="w", pady=5)
        site_entry = ttk.Entry(details_frame)
        site_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        # Notes
        ttk.Label(form_frame, text="Notes:").pack(pady=5)
        notes_text = tk.Text(form_frame, height=8, wrap=tk.WORD)
        notes_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Button frame
        btn_frame = ttk.Frame(form_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def record_admin():
            admin_record = {
                "date": date_entry.get(),
                "time": time_entry.get(),
                "administered_by": admin_by_entry.get(),
                "route": route_var.get(),
                "site": site_entry.get(),
                "notes": notes_text.get("1.0", tk.END).strip(),
                "recorded_by": self.current_user,
                "recorded_date": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            }
            
            if "administrations" not in medication:
                medication["administrations"] = []
                
            medication["administrations"].append(admin_record)
            self.save_patient_data()
            
            messagebox.showinfo("Success", "Administration recorded")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Record Administration", 
                  command=record_admin,
                  style='Green.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", 
                  command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def show_administration_history(self, medication, patient):
        """Show full administration history for a medication"""
        history_window = tk.Toplevel(self.root)
        history_window.title(f"Administration History - {medication['name']}")
        history_window.geometry("800x600")
        
        # Create scrollable frame
        container = ttk.Frame(history_window)
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container)
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
        ttk.Label(scrollable_frame, 
                 text=f"Administration History for {medication['name']}",
                 font=('Helvetica', 14, 'bold')).pack(pady=10)
        
        # Display administrations sorted by date (newest first)
        if "administrations" not in medication or not medication["administrations"]:
            ttk.Label(scrollable_frame, text="No administration history found").pack(pady=20)
        else:
            # Group by date
            admins_by_date = {}
            for admin in medication["administrations"]:
                date = admin["date"]
                if date not in admins_by_date:
                    admins_by_date[date] = []
                admins_by_date[date].append(admin)
            
            # Display by date (sorted newest to oldest)
            for date in sorted(admins_by_date.keys(), reverse=True):
                date_frame = ttk.LabelFrame(scrollable_frame, text=date)
                date_frame.pack(fill=tk.X, padx=5, pady=5)
                
                for admin in admins_by_date[date]:
                    admin_frame = ttk.Frame(date_frame)
                    admin_frame.pack(fill=tk.X, padx=5, pady=2)
                    
                    # Time and admin info
                    ttk.Label(admin_frame, 
                             text=f"{admin['time']} - {admin['administered_by']}",
                             font=('Helvetica', 10)).pack(side=tk.LEFT)
                    
                    # Route and site
                    details = []
                    if admin.get("route"):
                        details.append(f"Route: {admin['route']}")
                    if admin.get("site"):
                        details.append(f"Site: {admin['site']}")
                    
                    if details:
                        ttk.Label(admin_frame, 
                                 text=", ".join(details),
                                 font=('Helvetica', 9)).pack(side=tk.LEFT, padx=10)
                    
                    # Notes if available
                    if admin.get("notes"):
                        notes_frame = ttk.Frame(date_frame)
                        notes_frame.pack(fill=tk.X, padx=20, pady=2)
                        
                        ttk.Label(notes_frame, 
                                 text=f"Notes: {admin['notes']}",
                                 font=('Helvetica', 9)).pack(anchor="w")
        
        # Close button
        btn_frame = ttk.Frame(history_window)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="Close", command=history_window.destroy).pack()


