
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
        
        # Button frame at bottom
        btn_frame = ttk.Frame(right_frame, padding="10 10 10 10", style='TFrame')
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        ttk.Button(btn_frame, text="Back to Menu", command=self.main_menu,
                  style='Blue.TButton').pack(fill=tk.X, pady=10)
        
        # Notebook for different calculators (above button)
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
        
    def setup_antibiotics_calculator(self, parent):
        """Setup the antibiotics calculator interface for pediatric oncology patients"""
        # Create main container with scrollable canvas
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(main_frame, borderwidth=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Configure age groups
        self.age_groups = [
            "0-30 days (Neonate)",
            "1 month-1 year (Infant)",
            "1-6 years (Young child)",
            "6-12 years (Older child)",
            "12+ years (Adolescent)"
        ]
        
        # Expanded antibiotics data with age-specific dosing for oncology patients
        self.antibiotics_data = {
                "MEROPENEM": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 20, "max_dose": 40, "max_total": 2000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 20, "max_dose": 40, "max_total": 2000, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 20, "max_dose": 40, "max_total": 2000, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 20, "max_dose": 40, "max_total": 2000, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 20, "max_dose": 40, "max_total": 2000, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Dextrose >5%", "Other beta-lactams", "Aminoglycosides"],
                        "interactions": {
                                "Probenecid": "Decreases renal clearance",
                                "Valproic acid": "Reduces levels (avoid combination)",
                                "Warfarin": "Increased INR risk"
                        },
                        "notes": "First-line for febrile neutropenia. Higher doses for CNS infections. Adjust for renal impairment."
                },

                "AMIKACIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 15, "max_dose": 20, "max_total": 1500, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 15, "max_dose": 22.5, "max_total": 1500, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 15, "max_dose": 22.5, "max_total": 1500, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 15, "max_dose": 22.5, "max_total": 1500, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 15, "max_dose": 22.5, "max_total": 1500, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Penicillins", "Cephalosporins", "Heparin"],
                        "interactions": {
                                "Vancomycin": "Increased nephrotoxicity",
                                "Diuretics": "Ototoxicity risk",
                                "Cisplatin": "Increased toxicity"
                        },
                        "notes": "Monitor levels (trough <5 mg/L). Adjust dose based on renal function."
                },

                "CEFTRIAXONE (ROCEPHINE)": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 50, "max_dose": 80, "max_total": 4000, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 50, "max_dose": 100, "max_total": 4000, "frequency": "Every 12-24 hours"},
                                "1-6 years (Young child)": {"min_dose": 50, "max_dose": 100, "max_total": 4000, "frequency": "Every 12-24 hours"},
                                "6-12 years (Older child)": {"min_dose": 50, "max_dose": 100, "max_total": 4000, "frequency": "Every 12-24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 50, "max_dose": 100, "max_total": 4000, "frequency": "Every 12-24 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Calcium-containing solutions", "Aminoglycosides"],
                        "interactions": {
                                "Warfarin": "Increased effect",
                                "Calcium": "Precipitation risk (especially neonates)",
                                "Probenecid": "Increased levels"
                        },
                        "notes": "Avoid in hyperbilirubinemic neonates. Do not mix with calcium-containing solutions."
                },

                "CIPROFLOXACIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 10, "max_dose": 15, "max_total": 800, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 12 hours"},
                                "6-12 years (Older child)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 12 hours"},
                                "12+ years (Adolescent)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Divalent cation solutions", "TPN"],
                        "interactions": {
                                "Antacids": "Reduced absorption",
                                "Theophylline": "Increased levels",
                                "Warfarin": "Increased INR"
                        },
                        "notes": "Reserve for resistant gram-negative infections. Risk of tendonitis. Avoid in growing children when possible."
                },

                "TAZOCIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 80, "max_dose": 100, "max_total": 4000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 80, "max_dose": 100, "max_total": 4000, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 80, "max_dose": 100, "max_total": 4000, "frequency": "Every 6 hours"},
                                "6-12 years (Older child)": {"min_dose": 80, "max_dose": 100, "max_total": 4000, "frequency": "Every 6 hours"},
                                "12+ years (Adolescent)": {"min_dose": 80, "max_dose": 100, "max_total": 4000, "frequency": "Every 6 hours"}
                        },
                        "unit": "mg (piperacillin component)",
                        "incompatible": ["Aminoglycosides", "Vancomycin"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Anticoagulants": "Increased bleeding risk",
                                "Methotrexate": "Increased toxicity"
                        },
                        "notes": "Good Pseudomonas coverage. Contains piperacillin/tazobactam 8:1 ratio."
                },

                "VANCOMYCIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 10, "max_dose": 15, "max_total": 1000, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 10, "max_dose": 15, "max_total": 1000, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 10, "max_dose": 15, "max_total": 1000, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 10, "max_dose": 15, "max_total": 1000, "frequency": "Every 6 hours"},
                                "12+ years (Adolescent)": {"min_dose": 10, "max_dose": 15, "max_total": 1000, "frequency": "Every 6 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Alkaline solutions", "Beta-lactams", "Heparin"],
                        "interactions": {
                                "Aminoglycosides": "Nephrotoxicity risk",
                                "Loop diuretics": "Ototoxicity risk",
                                "NSAIDs": "Nephrotoxicity"
                        },
                        "notes": "Monitor trough levels (15-20 mg/L for serious infections). Pre-medicate for red man syndrome."
                },

                "METRONIDAZOLE (FLAGYL)": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 7.5, "max_dose": 10, "max_total": 500, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 7.5, "max_dose": 10, "max_total": 500, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 7.5, "max_dose": 10, "max_total": 500, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 7.5, "max_dose": 10, "max_total": 500, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 7.5, "max_dose": 10, "max_total": 500, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aluminum-containing solutions"],
                        "interactions": {
                                "Alcohol": "Disulfiram-like reaction",
                                "Warfarin": "Increased effect",
                                "Phenytoin": "Increased levels"
                        },
                        "notes": "For anaerobic infections. CNS toxicity at high doses. IV form contains sodium."
                },

                "FLUCONAZOLE": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 6, "max_dose": 12, "max_total": 800, "frequency": "Every 48 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 6, "max_dose": 12, "max_total": 800, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 6, "max_dose": 12, "max_total": 800, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 6, "max_dose": 12, "max_total": 800, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 6, "max_dose": 12, "max_total": 800, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Phenytoin": "Increased levels",
                                "Warfarin": "Increased INR",
                                "Rifampin": "Decreased fluconazole levels"
                        },
                        "notes": "Adjust dose in renal impairment. Loading dose recommended for serious infections."
                },

                "AMPHOTERICIN B": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 0.5, "max_dose": 1, "max_total": 50, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 0.5, "max_dose": 1, "max_total": 50, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 0.5, "max_dose": 1, "max_total": 50, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 0.5, "max_dose": 1, "max_total": 50, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 0.5, "max_dose": 1, "max_total": 50, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg/kg",
                        "incompatible": ["Saline solutions", "Other antifungals"],
                        "interactions": {
                                "Cyclosporine": "Increased nephrotoxicity",
                                "Diuretics": "Hypokalemia risk",
                                "Azoles": "Antagonism"
                        },
                        "notes": "Pre-medicate with antipyretics/antihistamines. Monitor renal function and electrolytes."
                },

                "VORICONAZOLE": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 7, "max_dose": 8, "max_total": 400, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 7, "max_dose": 8, "max_total": 400, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 7, "max_dose": 8, "max_total": 400, "frequency": "Every 12 hours"},
                                "6-12 years (Older child)": {"min_dose": 7, "max_dose": 8, "max_total": 400, "frequency": "Every 12 hours"},
                                "12+ years (Adolescent)": {"min_dose": 7, "max_dose": 8, "max_total": 400, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Rifampin": "Decreased levels",
                                "Carbamazepine": "Decreased levels",
                                "Warfarin": "Increased INR"
                        },
                        "notes": "Therapeutic drug monitoring recommended. Visual disturbances common."
                },

                "AUGMENTIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 25, "max_dose": 30, "max_total": 1750, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 25, "max_dose": 45, "max_total": 1750, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 25, "max_dose": 45, "max_total": 1750, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 25, "max_dose": 45, "max_total": 1750, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 25, "max_dose": 45, "max_total": 1750, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg (amoxicillin component)",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Allopurinol": "Increased rash risk",
                                "Warfarin": "Increased INR"
                        },
                        "notes": "Contains amoxicillin/clavulanate 7:1 ratio. Higher diarrhea risk than plain amoxicillin."
                },

                "ACYCLOVIR": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 8 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 10, "max_dose": 20, "max_total": 800, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Alkaline solutions"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Nephrotoxic drugs": "Increased toxicity",
                                "Zidovudine": "Neurotoxicity"
                        },
                        "notes": "Hydrate well. Adjust dose in renal impairment. Higher doses for HSV encephalitis."
                },

                "AMPICILLIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "1-6 years (Young child)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "6-12 years (Older child)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "12+ years (Adolescent)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aminoglycosides"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Allopurinol": "Increased rash risk",
                                "Warfarin": "Increased INR"
                        },
                        "notes": "Monitor for rash. Adjust dose in renal impairment. Listeria coverage."
                },

                "CLARITHROMYCIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 7.5, "max_dose": 10, "max_total": 1000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 7.5, "max_dose": 15, "max_total": 1000, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 7.5, "max_dose": 15, "max_total": 1000, "frequency": "Every 12 hours"},
                                "6-12 years (Older child)": {"min_dose": 7.5, "max_dose": 15, "max_total": 1000, "frequency": "Every 12 hours"},
                                "12+ years (Adolescent)": {"min_dose": 7.5, "max_dose": 15, "max_total": 1000, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Digoxin": "Increased levels",
                                "Theophylline": "Increased levels",
                                "Warfarin": "Increased INR"
                        },
                        "notes": "QT prolongation risk at high doses. Atypical pathogen coverage."
                },

                "CEFTAZIDIME": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aminoglycosides", "Vancomycin"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Loop diuretics": "Nephrotoxicity risk"
                        },
                        "notes": "Pseudomonas coverage. Adjust dose in renal impairment."
                },

                "CLOXACILLIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "1-6 years (Young child)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "6-12 years (Older child)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"},
                                "12+ years (Adolescent)": {"min_dose": 25, "max_dose": 50, "max_total": 2000, "frequency": "Every 6 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aminoglycosides"],
                        "interactions": {
                                "Probenecid": "Increased levels",
                                "Allopurinol": "Increased rash risk"
                        },
                        "notes": "For MSSA infections. Adjust dose in renal impairment."
                },

                "CLINDAMYCIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 5, "max_dose": 7.5, "max_total": 600, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 5, "max_dose": 10, "max_total": 600, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 5, "max_dose": 10, "max_total": 600, "frequency": "Every 6-8 hours"},
                                "6-12 years (Older child)": {"min_dose": 5, "max_dose": 10, "max_total": 600, "frequency": "Every 6-8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 5, "max_dose": 10, "max_total": 600, "frequency": "Every 6-8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aminoglycosides"],
                        "interactions": {
                                "Neuromuscular blockers": "Enhanced blockade",
                                "Erythromycin": "Antagonism"
                        },
                        "notes": "C. diff risk. Good anaerobic and soft tissue coverage."
                },

                "COLISTIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 2.5, "max_dose": 3, "max_total": 300, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 2.5, "max_dose": 5, "max_total": 300, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 2.5, "max_dose": 5, "max_total": 300, "frequency": "Every 12 hours"},
                                "6-12 years (Older child)": {"min_dose": 2.5, "max_dose": 5, "max_total": 300, "frequency": "Every 12 hours"},
                                "12+ years (Adolescent)": {"min_dose": 2.5, "max_dose": 5, "max_total": 300, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg (CBA)",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Aminoglycosides": "Increased nephrotoxicity",
                                "Neuromuscular blockers": "Enhanced blockade"
                        },
                        "notes": "For MDR gram-negative infections. Monitor renal function closely."
                },

                "GENTAMICIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 5, "max_dose": 7.5, "max_total": 400, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 5, "max_dose": 7.5, "max_total": 400, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 5, "max_dose": 7.5, "max_total": 400, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 5, "max_dose": 7.5, "max_total": 400, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 5, "max_dose": 7.5, "max_total": 400, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Penicillins", "Cephalosporins", "Vancomycin"],
                        "interactions": {
                                "Vancomycin": "Synergistic nephrotoxicity",
                                "Loop diuretics": "Ototoxicity",
                                "NSAIDs": "Nephrotoxicity"
                        },
                        "notes": "Monitor levels (trough <1 mg/L). Once daily dosing preferred."
                },

                # Additional critical agents for oncology patients
                "CEFEPIME": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 30, "max_dose": 50, "max_total": 2000, "frequency": "Every 8 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["Aminoglycosides", "Metronidazole"],
                        "interactions": {
                                "Aminoglycosides": "Nephrotoxicity",
                                "Probenecid": "Increased levels"
                        },
                        "notes": "4th gen cephalosporin with better gram-positive coverage than ceftazidime."
                },

                "LINEZOLID": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 10, "max_dose": 12, "max_total": 600, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 10, "max_dose": 12, "max_total": 600, "frequency": "Every 8 hours"},
                                "1-6 years (Young child)": {"min_dose": 10, "max_dose": 12, "max_total": 600, "frequency": "Every 8 hours"},
                                "6-12 years (Older child)": {"min_dose": 10, "max_dose": 12, "max_total": 600, "frequency": "Every 8 hours"},
                                "12+ years (Adolescent)": {"min_dose": 10, "max_dose": 12, "max_total": 600, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "SSRIs": "Serotonin syndrome risk",
                                "MAOIs": "Hypertensive crisis",
                                "Adrenergics": "Increased pressor response"
                        },
                        "notes": "For VRE and MRSA. Monitor CBC for myelosuppression. Limited course recommended."
                },

                "DAPTOMYCIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 4, "max_dose": 6, "max_total": 500, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 4, "max_dose": 6, "max_total": 500, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 4, "max_dose": 6, "max_total": 500, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 4, "max_dose": 6, "max_total": 500, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 4, "max_dose": 6, "max_total": 500, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg/kg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Statins": "Increased rhabdomyolysis risk",
                                "Tobramycin": "Increased CPK"
                        },
                        "notes": "For MRSA and VRE. Monitor CPK weekly. Not for pulmonary infections."
                },

                "CASPOFUNGIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 25, "max_dose": 50, "max_total": 70, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 25, "max_dose": 50, "max_total": 70, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 25, "max_dose": 50, "max_total": 70, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 25, "max_dose": 50, "max_total": 70, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 50, "max_dose": 70, "max_total": 70, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg/m²",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Cyclosporine": "Increased caspofungin levels",
                                "Tacrolimus": "Decreased tacrolimus levels"
                        },
                        "notes": "For invasive candidiasis and aspergillosis. Loading dose required."
                },

                "MICAFUNGIN": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 2, "max_dose": 4, "max_total": 150, "frequency": "Every 24 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 2, "max_dose": 4, "max_total": 150, "frequency": "Every 24 hours"},
                                "1-6 years (Young child)": {"min_dose": 2, "max_dose": 4, "max_total": 150, "frequency": "Every 24 hours"},
                                "6-12 years (Older child)": {"min_dose": 2, "max_dose": 4, "max_total": 150, "frequency": "Every 24 hours"},
                                "12+ years (Adolescent)": {"min_dose": 2, "max_dose": 4, "max_total": 150, "frequency": "Every 24 hours"}
                        },
                        "unit": "mg/kg",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Sirolimus": "Increased levels",
                                "Nifedipine": "Increased levels"
                        },
                        "notes": "For candidemia and prophylaxis. No loading dose needed."
                },

                "TRIMETHOPRIM-SULFAMETHOXAZOLE": {
                        "age_groups": {
                                "0-30 days (Neonate)": {"min_dose": 5, "max_dose": 10, "max_total": 320, "frequency": "Every 12 hours"},
                                "1 month-1 year (Infant)": {"min_dose": 5, "max_dose": 10, "max_total": 320, "frequency": "Every 12 hours"},
                                "1-6 years (Young child)": {"min_dose": 5, "max_dose": 10, "max_total": 320, "frequency": "Every 12 hours"},
                                "6-12 years (Older child)": {"min_dose": 5, "max_dose": 10, "max_total": 320, "frequency": "Every 12 hours"},
                                "12+ years (Adolescent)": {"min_dose": 5, "max_dose": 10, "max_total": 320, "frequency": "Every 12 hours"}
                        },
                        "unit": "mg/kg (TMP component)",
                        "incompatible": ["None known"],
                        "interactions": {
                                "Warfarin": "Increased INR",
                                "Phenytoin": "Increased levels",
                                "Methotrexate": "Increased toxicity"
                        },
                        "notes": "For PJP prophylaxis and treatment. Monitor CBC for myelosuppression."
                }
        }
        
        # UI Elements
        ttk.Label(scrollable_frame, text="Pediatric Oncology Antibiotics Calculator", 
                 font=('Helvetica', 16, 'bold')).grid(row=0, column=0, columnspan=6, pady=10)
        
        # Patient details frame
        input_frame = ttk.Frame(scrollable_frame)
        input_frame.grid(row=1, column=0, columnspan=6, pady=10, sticky="ew")
        
        # Weight input
        self.antibiotics_weight_var = tk.StringVar()
        ttk.Label(input_frame, text="Weight (kg):").grid(row=0, column=0, padx=5)
        ttk.Entry(input_frame, textvariable=self.antibiotics_weight_var, width=8).grid(row=0, column=1, padx=5)
        
        # Age group selection
        self.age_group_var = tk.StringVar()
        ttk.Label(input_frame, text="Age Group:").grid(row=0, column=2, padx=5)
        age_combo = ttk.Combobox(input_frame, textvariable=self.age_group_var, 
                                values=self.age_groups, state="readonly", width=25)
        age_combo.grid(row=0, column=3, padx=5)
        
        # Antibiotic selection frames
        self.antibiotic_frames = []
        for i in range(5):
            abx_frame = ttk.LabelFrame(scrollable_frame, text=f"Antibiotic {i+1}")
            abx_frame.grid(row=2+i, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
            
            # Drug selection
            drug_var = tk.StringVar()
            drug_combo = ttk.Combobox(abx_frame, textvariable=drug_var, 
                                     values=sorted(self.antibiotics_data.keys()), 
                                     state="readonly", width=25)
            drug_combo.grid(row=0, column=0, padx=5, pady=2)
            
            # Dose information
            dose_frame = ttk.Frame(abx_frame)
            dose_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
            
            ttk.Label(dose_frame, text="Dose Range:").grid(row=0, column=0, padx=2)
            dose_range_var = tk.StringVar()
            ttk.Label(dose_frame, textvariable=dose_range_var, width=20).grid(row=0, column=1, padx=2)
            
            ttk.Label(dose_frame, text="Frequency:").grid(row=0, column=2, padx=2)
            frequency_var = tk.StringVar()
            ttk.Label(dose_frame, textvariable=frequency_var, width=15).grid(row=0, column=3, padx=2)
            
            # Dose input
            ttk.Label(abx_frame, text="Enter Dose:").grid(row=2, column=0, padx=5)
            dose_var = tk.StringVar()
            dose_entry = ttk.Entry(abx_frame, textvariable=dose_var, width=8)
            dose_entry.grid(row=2, column=1, padx=5)
            
            # Calculation results
            result_var = tk.StringVar()
            ttk.Label(abx_frame, textvariable=result_var, width=25).grid(row=2, column=2, padx=5)
            
            # Store variables
            self.antibiotic_frames.append({
                "frame": abx_frame,
                "drug_var": drug_var,
                "dose_range_var": dose_range_var,
                "frequency_var": frequency_var,
                "dose_var": dose_var,
                "result_var": result_var
            })
            
            # Bind drug selection to update fields
            drug_var.trace_add("write", lambda *args, i=i: self.update_antibiotic_fields(i))
        
        # Interaction panel
        interaction_frame = ttk.LabelFrame(scrollable_frame, text="Drug Interactions & Warnings")
        interaction_frame.grid(row=2, column=4, rowspan=5, sticky="nsew", padx=10, pady=5)
        
        self.interaction_text = tk.Text(interaction_frame, wrap=tk.WORD, width=50, height=25)
        self.interaction_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=7, column=0, columnspan=6, pady=10)
        
        ttk.Button(btn_frame, text="Calculate", command=self.calculate_antibiotics).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Clear", command=self.clear_antibiotics).grid(row=0, column=1, padx=5)
        
        # Configure grid weights
        scrollable_frame.columnconfigure(4, weight=1)
        scrollable_frame.rowconfigure(6, weight=1)

    def update_antibiotic_fields(self, index):
        """Update fields when antibiotic is selected"""
        abx = self.antibiotic_frames[index]
        drug = abx["drug_var"].get()
        age_group = self.age_group_var.get()
        
        if drug and age_group:
            data = self.antibiotics_data[drug]["age_groups"][age_group]
            unit = self.antibiotics_data[drug]["unit"]
            
            # Update dose range and frequency
            abx["dose_range_var"].set(f"{data['min_dose']}-{data['max_dose']} {unit}/kg")
            abx["frequency_var"].set(data["frequency"])
            
            # Auto-fill with max dose
            abx["dose_var"].set(data["max_dose"])

    def calculate_antibiotics(self):
        """Calculate doses and check interactions"""
        try:
            weight = float(self.antibiotics_weight_var.get())
            age_group = self.age_group_var.get()
            
            for abx in self.antibiotic_frames:
                drug = abx["drug_var"].get()
                if not drug or not age_group:
                    continue
                
                dose_data = self.antibiotics_data[drug]["age_groups"][age_group]
                unit = self.antibiotics_data[drug]["unit"]
                
                try:
                    dose = float(abx["dose_var"].get())
                except ValueError:
                    dose = dose_data["max_dose"]  # Use max dose if invalid input
                
                # Calculate total dose
                calculated_dose = weight * dose
                if calculated_dose > dose_data["max_total"]:
                    calculated_dose = dose_data["max_total"]
                
                # Update results
                abx["result_var"].set(f"Total: {calculated_dose:.1f} {unit}\n(Max: {dose_data['max_total']} {unit})")
            
            self.check_interactions()
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid weight and select age group")
        except KeyError:
            messagebox.showerror("Error", "Please select valid age group for all medications")

    def get_incompatibility_reason(self, drug, substance):
        """Get detailed reason for incompatibility based on drug properties"""
        reasons = {
            # Solution-based incompatibilities
            "Dextrose >5%": "High dextrose concentrations alter pH balance, affecting drug stability",
            "Calcium-containing solutions": "Risk of precipitation (especially with ceftriaxone in neonates)",
            "Divalent cation solutions": "Cations (Ca²⁺, Mg²⁺, Zn²⁺) chelate drug reducing bioavailability",
            "TPN": "Complex parenteral nutrition formulations cause physical instability",
            "Saline solutions": "Osmolarity mismatch leads to precipitation (particularly Amphotericin B)",
            "Aluminum-containing solutions": "Aluminum binds drug molecules forming insoluble complexes",
            "Alkaline solutions": "High pH causes drug degradation through hydrolysis",
            
            # Drug class incompatibilities
            "Aminoglycosides": "Physical precipitation when mixed + synergistic nephrotoxicity risk",
            "Beta-lactams": "Shared beta-lactam structure increases cross-reactivity risk",
            "Penicillins": "Chemical degradation through nucleophilic reactions",
            "Cephalosporins": "Structural similarity leads to competitive inhibition",
            "Other antifungals": "Therapeutic antagonism of fungal coverage",
            "Vancomycin": "Physical incompatibility when co-infused + additive nephrotoxicity",
            
            # Specific drug incompatibilities
            "Heparin": "Polyanionic nature causes complex precipitation",
            "Loop diuretics": "Synergistic ototoxicity risk (e.g., with aminoglycosides)",
            "NSAIDs": "Additive nephrotoxicity through prostaglandin inhibition",
            "Probenecid": "Competitive renal tubular secretion blockade",
            "Warfarin": "Enhanced anticoagulation through protein binding displacement",
            "Phenytoin": "Hepatic enzyme induction alters metabolism",
            
            # Special cases
            "Calcium": "Ceftriaxone-calcium precipitates in biliary/urinary tract (neonates)",
            "Aminoglycosides (extended)": "Inactivation through pH changes when mixed in same line",
            "Vancomycin (extended)": "Nephrotoxicity potentiation through glomerular injury",
            "Cisplatin": "Additive renal tubular toxicity",
            "Methotrexate": "Competitive renal excretion increases toxicity",
            
            # Pharmacodynamic interactions
            "Neuromuscular blockers": "Enhanced paralysis through presynaptic calcium effects",
            "Statins": "Shared myotoxicity pathways increase rhabdomyolysis risk",
            "QT-prolonging agents": "Additive cardiac repolarization delay",
            
            # Pharmacokinetic interactions
            "Rifampin": "Hepatic enzyme induction accelerates drug metabolism",
            "Azoles": "CYP450 inhibition alters drug clearance",
            "Proton Pump Inhibitors": "Gastric pH changes reduce oral bioavailability"
        }
        
        # Try for exact match first
        if substance in reasons:
            return reasons[substance]
        
        # Partial matches for drug classes
        if "beta-lactam" in substance.lower():
            return reasons["Beta-lactams"]
        if "aminoglycoside" in substance.lower():
            return reasons["Aminoglycosides"]
        
        # Fallback to general mechanism
        return "Physical/Chemical incompatibility - avoid simultaneous administration"
    
    def check_interactions(self):
        """Check for drug-drug interactions and incompatibilities"""
        selected_drugs = [abx["drug_var"].get() for abx in self.antibiotic_frames if abx["drug_var"].get()]
        interactions = []
        incompatibilities = set()
        notes = set()
        
        # Check pairwise interactions
        for i, drug in enumerate(selected_drugs):
            # Add drug-specific notes
            notes.add(f"{drug}: {self.antibiotics_data[drug]['notes']}")
            
            # Check incompatibilities
            incompatibilities.update(self.antibiotics_data[drug]["incompatible"])
            
            for other_drug in selected_drugs[i+1:]:
                # Drug-drug interactions
                common_interactions = set(self.antibiotics_data[drug]["interactions"].keys()) & \
                                    set(self.antibiotics_data[other_drug]["interactions"].keys())
                
                for interaction in common_interactions:
                    interactions.append(f"⚠️ {drug} + {other_drug}: {self.antibiotics_data[drug]['interactions'][interaction]}")

        # Build report
        report = []
        if interactions:
            report.append("🚨 DRUG INTERACTIONS:")
            report.extend(interactions)
            report.append("")
        
        if incompatibilities:
            report.append("🚫 INCOMPATIBILITIES:")
            report.extend(sorted(incompatibilities))
            report.append("")
        
        if notes:
            report.append("📝 IMPORTANT NOTES:")
            report.extend(sorted(notes))
        
        self.interaction_text.delete(1.0, tk.END)
        self.interaction_text.insert(tk.END, "\n".join(report) if report else "✅ No significant interactions detected")

    def clear_antibiotics(self):
        """Clear all inputs and results"""
        self.antibiotics_weight_var.set("")
        self.age_group_var.set("")
        self.interaction_text.delete(1.0, tk.END)
        
        for abx in self.antibiotic_frames:
            abx["drug_var"].set("")
            abx["dose_range_var"].set("")
            abx["frequency_var"].set("")
            abx["dose_var"].set("")
            abx["result_var"].set("")

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

