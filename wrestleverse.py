import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import json
import pandas as pd
import random
import os
import time
import pyodbc
from openai import OpenAI
import datetime
import logging

logging.basicConfig(
    filename='log.txt',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class WrestleverseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wrestleverse")
        self.root.geometry("600x500")
        self.api_key = ""
        self.uid_start = 1
        self.bio_prompt = "Create a biography for a professional wrestler."
        self.access_db_path = ""
        self.pictures_path = ""
        self.client = None
        self.load_settings()
        if self.api_key:
            self.client = OpenAI(api_key=self.api_key)
        else:
            self.client = None
        self.skill_presets = []
        self.load_skill_presets()
        self.wrestlers = []
        self.companies = []
        self.setup_main_menu()
    def add_company_form(self):
        company_frame = ttk.Frame(self.companies_frame, relief="ridge", borderwidth=2)
        company_frame.pack(fill="x", pady=5)
        name_label = ttk.Label(company_frame, text="Name:")
        name_label.grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(company_frame)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        description_label = ttk.Label(company_frame, text="Description:")
        description_label.grid(row=1, column=0, padx=5, pady=5)
        description_entry = ttk.Entry(company_frame, width=40)
        description_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5)
        size_label = ttk.Label(company_frame, text="Size:")
        size_label.grid(row=2, column=0, padx=5, pady=5)
        size_var = tk.StringVar(value="Medium")
        size_dropdown = ttk.Combobox(company_frame, textvariable=size_var, values=["Tiny", "Small", "Medium", "Large"])
        size_dropdown.grid(row=2, column=1, padx=5, pady=5)
        remove_btn = ttk.Button(company_frame, text="❌", command=lambda: self.remove_company_form(company_frame))
        remove_btn.grid(row=0, column=4, padx=5, pady=5)
        self.companies.append({
            "frame": company_frame,
            "name": name_entry,
            "description": description_entry,
            "size": size_var
        })

    def remove_company_form(self, company_frame):
        company_frame.destroy()
        self.companies = [company for company in self.companies if company["frame"] != company_frame]

    def generate_companies(self):
        logging.debug("generate_companies function was invoked.")
        
        if not self.api_key:
            messagebox.showerror("Error", "Please set your API key in settings before generating companies.")
            logging.error("API key not set.")
            return

        try:
            companies_columns = [
                "UID", "Name", "Initials", "URL", "CompanyOpening", "CompanyClosing", "Trading", "Mediagroup",
                "Logo", "Backdrop", "Banner", "Based_In", "Prestige", "Influence", "Money", "Size", "LimitSize",
                "Momentum", "Announce1", "Announce2", "Announce3", "FixBelts", "CompanyNotBefore", "CompanyNotAfter",
                "AlliancePreset", "Ace", "AceLength", "Heir", "HeirLength", "TVFirst", "TVAsc", "EventAsc",
                "TrueBorn", "YoungLion", "HomeArena", "TippyToe", "GeogTag1", "GeogTag2", "GeogTag3", "HQ", "HOF"
            ]

            # Get the next available UID
            uid = self.uid_start
            if self.access_db_path and os.path.exists(self.access_db_path):
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={self.access_db_path};'
                    'PWD=20YearsOfTEW;'
                )
                logging.debug(f"Attempting to connect to Access database at: {self.access_db_path}")
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                cursor.execute("SELECT MAX(UID) FROM tblFed")
                result = cursor.fetchone()
                last_uid = result[0] if result[0] else 0
                uid = last_uid + 1

            companies_data = []
            bio_data = []
            notes_data = []

            # Extract widget values before accessing
            company_data_list = []
            for company in self.companies:
                name = company["name"].get().strip() if company["name"] else ""
                description = company["description"].get().strip() if company["description"] else ""
                size = company["size"].get().strip() if company["size"] else "Medium"
                company_data_list.append({"name": name, "description": description, "size": size})

            total_companies = len(company_data_list)
            for index, company_data in enumerate(company_data_list, start=1):
                self.status_label.config(text=f"Status: Generating company {index}/{total_companies}...")
                self.root.update_idletasks()

                name = company_data["name"]
                description = company_data["description"]
                size = company_data["size"]

                if not name:
                    name = self.generate_company_name(description=description, size=size)
                if not description:
                    description = self.generate_company_description(name=name, size=size)

                # Generate company row data
                company_row = [
                    uid,  # UID
                    name,  # Name
                    self.generate_company_initials(name),  # Initials
                    f"www.{name.replace(' ', '').lower()}.com"[:40],  # URL
                    "1666-01-01",  # CompanyOpening
                    "1666-01-01",  # CompanyClosing
                    -1,  # Trading (True = -1)
                    0,  # Mediagroup
                    f"{name.replace(' ', '').lower()}.jpg"[:35],  # Logo
                    f"{name.replace(' ', '').lower()}BD.jpg"[:35],  # Backdrop
                    f"{name.replace(' ', '').lower()}Banner.jpg"[:30],  # Banner
                    1,  # Based_In
                    random.randint(1, 100),  # Prestige
                    0,  # Influence
                    {"Tiny": 100000, "Small": 1000000, "Medium": 10000000, "Large": 100000000}.get(size, 1000000),  # Money
                    0,  # Size
                    10,  # LimitSize
                    random.randint(1, 100),  # Momentum
                    0,  # Announce1
                    0,  # Announce2
                    0,  # Announce3
                    0,  # FixBelts (False = 0)
                    "1666-01-01",  # CompanyNotBefore
                    "1666-01-01",  # CompanyNotAfter
                    0,  # AlliancePreset
                    0,  # Ace
                    0,  # AceLength
                    0,  # Heir
                    0,  # HeirLength
                    -1,  # TVFirst (True = -1)
                    -1,  # TVAsc (True = -1)
                    -1,  # EventAsc (True = -1)
                    -1,  # TrueBorn (True = -1)
                    0,  # YoungLion (False = 0)
                    0,  # HomeArena
                    0,  # TippyToe (False = 0)
                    "",  # GeogTag1
                    "",  # GeogTag2
                    "",  # GeogTag3
                    0,  # HQ (False = 0)
                    -1,  # HOF (True = -1)
                ]
                companies_data.append(company_row)

                # Generate and store bio
                bio = self.generate_company_bio(name, description, size)
                bio_data.append([uid, bio])

                # Store description in notes with image_generated flag
                notes_data.append({
                    "Name": name,
                    "Description": description,
                    "Size": size,
                    "Logo": f"{name.replace(' ', '').lower()}.jpg"[:35],  # Match the Logo field
                    "Backdrop": f"{name.replace(' ', '').lower()}BD.jpg"[:35],  # Match the Backdrop field
                    "Banner": f"{name.replace(' ', '').lower()}Banner.jpg"[:30],  # Match the Banner field
                    "image_generated": False
                })

                uid += 1

            # Save to Access DB if path exists
            if self.access_db_path and os.path.exists(self.access_db_path):
                logging.debug("Attempting to save to Access database.")
                try:
                    for company_row in companies_data:
                        # Create the SQL statement with bracketed column names
                        sql_insert_company = """
                            INSERT INTO tblFed 
                            ([UID], [Name], [Initials], [URL], [CompanyOpening], [CompanyClosing], [Trading], [Mediagroup],
                            [Logo], [Backdrop], [Banner], [Based_In], [Prestige], [Influence], [Money], [Size], [LimitSize],
                            [Momentum], [Announce1], [Announce2], [Announce3], [FixBelts], [CompanyNotBefore], [CompanyNotAfter],
                            [AlliancePreset], [Ace], [AceLength], [Heir], [HeirLength], [TVFirst], [TVAsc], [EventAsc],
                            [TrueBorn], [YoungLion], [HomeArena], [TippyToe], [GeogTag1], [GeogTag2], [GeogTag3], [HQ], [HOF])
                            VALUES 
                            (?, ?, ?, ?, ?, ?, ?, ?,
                             ?, ?, ?, ?, ?, ?, ?, ?, ?,
                             ?, ?, ?, ?, ?, ?, ?, ?, ?,
                             ?, ?, ?, ?, ?, ?,
                             ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """
                        logging.debug(f"Executing SQL with values: {company_row}")
                        cursor.execute(sql_insert_company, company_row)

                        # Insert into tblFedSchedule
                        sql_insert_schedule = "INSERT INTO tblFedSchedule ([FedUID], [Strategy]) VALUES (?, ?)"
                        schedule_values = (company_row[0], '5')  # UID and Strategy
                        logging.debug(f"Executing Schedule SQL with values: {schedule_values}")
                        cursor.execute(sql_insert_schedule, schedule_values)

                    for bio_row in bio_data:
                        sql_insert_bio = "INSERT INTO tblFedBio ([UID], [Profile]) VALUES (?, ?)"
                        logging.debug(f"Executing Bio SQL with values: {bio_row}")
                        cursor.execute(sql_insert_bio, bio_row)

                    conn.commit()
                    conn.close()
                    logging.debug("Successfully saved to Access database.")
                except Exception as e:
                    logging.error(f"Error saving to Access database: {e}", exc_info=True)
                    logging.debug(f"Last SQL statement attempted: {sql_insert_company}")
                    messagebox.showerror("Error", f"Could not save to Access database: {str(e)}")



            # Inside the try block, just before saving to Excel:
            for note in notes_data:
                try:
                    # Get logo description from GPT
                    logo_prompt = (
                        f"For a professional wrestling company named '{note['Name']}' "
                        f"with the following description: '{note['Description']}', "
                        f"describe in a single sentence what their logo might look like. "
                        f"Focus on style, colors, and iconic elements."
                    )
                    logo_description = self.get_response_from_gpt(logo_prompt)
                    note['logo_description'] = logo_description
                    logging.debug(f"Generated logo description for {note['Name']}: {logo_description}")
                except Exception as e:
                    logging.error(f"Error generating logo description: {e}")
                    note['logo_description'] = ""

            # Save to Excel
            try:
                logging.debug("Attempting to save data to Excel file.")
                companies_df = pd.DataFrame(companies_data, columns=companies_columns)
                bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                notes_df = pd.DataFrame(notes_data)  # Add this line
                
                excel_path = "wrestleverse_companies.xlsx"
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    companies_df.to_excel(writer, sheet_name="Companies", index=False)
                    bio_df.to_excel(writer, sheet_name="Bios", index=False)
                    notes_df.to_excel(writer, sheet_name="Notes", index=False)  # This will now include all columns
                
                logging.debug(f"Excel file saved successfully to {excel_path}")
                messagebox.showinfo("Success", f"Companies saved to {excel_path}")
            except Exception as e:
                logging.error(f"Error saving Excel file: {e}", exc_info=True)
                messagebox.showerror("Error", f"Could not save Excel file: {str(e)}")

        except Exception as e:
            logging.error(f"Unhandled error in generate_companies: {e}", exc_info=True)
            messagebox.showerror("Error", f"Unhandled error: {e}")
        finally:
            self.status_label.config(text="Status: Companies generated successfully!")
            self.root.update_idletasks()



    def generate_company_name(self, description=None, size=None):
        prompt = "Generate a name for a professional wrestling company."
        if description:
            prompt += f" The company's style or theme is: {description}."
        if size:
            prompt += f" The company is of {size.lower()} size."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def generate_company_initials(self, name):
        initials = ''.join([word[0] for word in name.split() if word[0].isalpha()])
        if not initials:
            initials = name[:3].upper()
        return initials[:12]

    def generate_company_description(self, name=None, size=None):
        prompt = "Generate a description for a professional wrestling company."
        if name:
            prompt += f" The company's name is {name}."
        if size:
            prompt += f" The company is of {size.lower()} size."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def generate_company_bio(self, name, description, size):
        prompt = f"Create a detailed profile for a professional wrestling company named {name}."
        if description:
            prompt += f" Description: {description}."
        if size:
            prompt += f" The company is considered {size.lower()} in size."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def setup_main_menu(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Wrestleverse", font=("Helvetica", 20))
        title_label.pack(pady=20)
        generate_wrestlers_btn = ttk.Button(self.root, text="Generate Wrestlers", command=self.open_wrestler_generator)
        generate_wrestlers_btn.pack(pady=10)
        generate_company_btn = ttk.Button(self.root, text="Generate Company", command=self.open_company_generator)
        generate_company_btn.pack(pady=10)
        skill_presets_btn = ttk.Button(self.root, text="Skill Presets", command=self.open_skill_presets)
        skill_presets_btn.pack(pady=10)
        generate_images_btn = ttk.Button(self.root, text="Generate Images", command=self.open_image_generator)
        generate_images_btn.pack(pady=10)
        settings_btn = ttk.Button(self.root, text="⚙ Settings", command=self.open_settings)
        settings_btn.pack(side="bottom", pady=10)

    def open_company_generator(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Company Generator", font=("Helvetica", 16))
        title_label.pack(pady=10)
        add_company_btn = ttk.Button(self.root, text="Add Company", command=self.add_company_form)
        add_company_btn.pack(pady=10)
        self.companies_frame = ttk.Frame(self.root)
        self.companies_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.status_label = ttk.Label(self.root, text="Status: Waiting to generate companies...")
        self.status_label.pack(pady=10)
        generate_btn = ttk.Button(self.root, text="Generate Companies", command=self.generate_companies)
        generate_btn.pack(side="bottom", pady=10)
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def open_wrestler_generator(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Wrestler Generator", font=("Helvetica", 16))
        title_label.pack(pady=10)
        add_wrestler_btn = ttk.Button(self.root, text="Add Wrestler", command=self.add_wrestler_form)
        add_wrestler_btn.pack(pady=10)
        self.wrestlers_frame = ttk.Frame(self.root)
        self.wrestlers_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.status_label = ttk.Label(self.root, text="Status: Waiting to generate wrestlers...")
        self.status_label.pack(pady=10)
        generate_btn = ttk.Button(self.root, text="Generate Wrestlers", command=self.generate_wrestlers)
        generate_btn.pack(side="bottom", pady=10)
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def add_wrestler_form(self):
        wrestler_frame = ttk.Frame(self.wrestlers_frame, relief="ridge", borderwidth=2)
        wrestler_frame.pack(fill="x", pady=5)
        name_label = ttk.Label(wrestler_frame, text="Name:")
        name_label.grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(wrestler_frame)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        gender_label = ttk.Label(wrestler_frame, text="Gender:")
        gender_label.grid(row=0, column=2, padx=5, pady=5)
        gender_var = tk.StringVar(value="Male")
        gender_dropdown = ttk.Combobox(wrestler_frame, textvariable=gender_var, values=["Male", "Female"])
        gender_dropdown.grid(row=0, column=3, padx=5, pady=5)
        company_label = ttk.Label(wrestler_frame, text="Company:")
        company_label.grid(row=1, column=0, padx=5, pady=5)
        
        # Get company names for dropdown (just the names, not the UIDs)
        companies = self.get_companies()
        company_names = ["Random"] + [company[1] for company in companies]  # Only use company names
        
        company_var = tk.StringVar(value="Random")
        company_dropdown = ttk.Combobox(wrestler_frame, textvariable=company_var, values=company_names)
        company_dropdown.grid(row=1, column=1, padx=5, pady=5)
        exclusive_label = ttk.Label(wrestler_frame, text="Exclusive:")
        exclusive_label.grid(row=1, column=2, padx=5, pady=5)
        exclusive_var = tk.StringVar(value="Random")
        exclusive_dropdown = ttk.Combobox(wrestler_frame, textvariable=exclusive_var, values=["Random", "Yes", "No"])
        exclusive_dropdown.grid(row=1, column=3, padx=5, pady=5)
        description_label = ttk.Label(wrestler_frame, text="Description:")
        description_label.grid(row=2, column=0, padx=5, pady=5)
        description_entry = ttk.Entry(wrestler_frame, width=40)
        description_entry.grid(row=2, column=1, columnspan=3, padx=5, pady=5)
        skill_preset_label = ttk.Label(wrestler_frame, text="Skill Preset:")
        skill_preset_label.grid(row=3, column=0, padx=5, pady=5)
        skill_preset_var = tk.StringVar(value="Interpret")
        skill_preset_names = ["Interpret"] + [preset["name"] for preset in self.skill_presets]
        skill_preset_dropdown = ttk.Combobox(wrestler_frame, textvariable=skill_preset_var, values=skill_preset_names)
        skill_preset_dropdown.grid(row=3, column=1, padx=5, pady=5)
        remove_btn = ttk.Button(wrestler_frame, text="❌", command=lambda: self.remove_wrestler_form(wrestler_frame))
        remove_btn.grid(row=0, column=4, padx=5, pady=5)
        self.wrestlers.append({
            "frame": wrestler_frame,
            "name": name_entry,
            "gender": gender_var,
            "company": company_var,
            "exclusive": exclusive_var,
            "description": description_entry,
            "skill_preset": skill_preset_var
        })

    def remove_wrestler_form(self, wrestler_frame):
        wrestler_frame.destroy()
        self.wrestlers = [wrestler for wrestler in self.wrestlers if wrestler["frame"] != wrestler_frame]

    def generate_wrestlers(self):
        if not self.api_key:
            messagebox.showerror("Error", "Please set your API key in settings before generating wrestlers.")
            return

        try:
            workers_data = []
            bio_data = []
            skills_data = []
            contract_data = []
            notes_data = []  # Add this line
            
            # Get starting UIDs from settings
            uid = self.uid_start  # For workers
            contract_uid = self.uid_start  # For contracts - using the same starting point
            
            if self.access_db_path and os.path.exists(self.access_db_path):
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={self.access_db_path};'
                    'PWD=20YearsOfTEW;'
                )
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                
                # Get next worker UID
                cursor.execute("SELECT MAX(UID) FROM tblWorker")
                result = cursor.fetchone()
                last_uid = result[0] if result[0] else 0
                uid = max(last_uid + 1, self.uid_start)
                
                # Get next contract UID - using self.uid_start as minimum
                cursor.execute("SELECT MAX(UID) FROM tblContract")
                result = cursor.fetchone()
                last_contract_uid = result[0] if result[0] else 0
                contract_uid = max(last_contract_uid + 1, self.uid_start)

            # Collect all wrestler data first
            wrestler_data_list = []
            for wrestler in self.wrestlers:
                try:
                    if wrestler["frame"].winfo_exists():  # Check if the frame still exists
                        data = {
                            'name': wrestler["name"].get().strip() if wrestler["name"].winfo_exists() else "",
                            'gender': wrestler["gender"].get().strip() if hasattr(wrestler["gender"], "get") else "Male",
                            'company': wrestler["company"].get().strip() if hasattr(wrestler["company"], "get") else "Random",
                            'exclusive': wrestler["exclusive"].get().strip() if hasattr(wrestler["exclusive"], "get") else "Random",
                            'description': wrestler["description"].get().strip() if wrestler["description"].winfo_exists() else "",
                            'skill_preset': wrestler["skill_preset"].get().strip() if hasattr(wrestler["skill_preset"], "get") else "Default"
                        }
                        wrestler_data_list.append(data)
                        logging.debug(f"Collected data for wrestler: {data['name']}")
                except (tk.TclError, AttributeError) as e:
                    logging.warning(f"Skipping invalid wrestler form: {e}")
                    continue

            total_wrestlers = len(wrestler_data_list)
            for index, wrestler_data in enumerate(wrestler_data_list, start=1):
                self.status_label.config(text=f"Status: Generating wrestler {index}/{total_wrestlers}...")
                self.root.update_idletasks()

                                # Initialize worker data with proper defaults
                name = wrestler["name"].get().strip()
                gender = wrestler["gender"].get().strip()
                description = wrestler["description"].get().strip()
                
                if not name:
                    name = self.generate_name(description=description, gender=gender)
                name = name[:30]
                shortname = name.split()[0][:20] if name else ''
                gender_value = 1 if gender.lower() == 'male' else 5
                weight = random.randint(150, 350)
                min_weight = weight - 20  # Always 20 pounds less than current weight
                max_weight = weight + random.randint(20, 50)

                debut_date = datetime.datetime(
                    random.randint(1980, 2023),  # Year between 1980 and 2023
                    random.randint(1, 12),       # Month
                    random.randint(1, 28)        # Day (avoiding potential month end issues)
                )
                debut_age = random.randint(16, 50) 
                birth_year = debut_date.year - debut_age
                birth_date = datetime.datetime(
                    birth_year,
                    random.randint(1, 12),  # Random birth month
                    random.randint(1, 28)   # Random birth day (avoiding month end issues)
                )
                                # Convert boolean values for Access DB


                logging.debug(f"Generated worker data for {wrestler_data['name']}")

                # Get the style first
                style = style_var.get() if hasattr(self, 'style_var') else "Interpret"

                # Get the bio
                bio_prompt = (
                    f"Create a biography for a professional wrestler. The wrestler's name is {name}. "
                    f"Their gender is {gender}. Description: {description}. "
                    f"Their wrestling style is best described as {style}."
                )
                bio = self.get_response_from_gpt(bio_prompt)

                if not bio:
                    logging.error("Failed to generate biography")
                    return

                race = self.get_race_from_gpt(name, f"{description}\n\nBiography: {bio}")


                # Initialize all required fields with proper defaults
                worker_row = {
                    "UID": int(uid),
                    "User": False,  # Will be converted to 0
                    "Regen": self.ensure_byte(0),
                    "Active": True,  # Will be converted to -1
                    "Name": name[:30],
                    "Shortname": shortname[:20],
                    "Gender": self.ensure_byte(gender_value),
                    "Pronouns": self.ensure_byte(1 if gender_value == 1 else 2),
                    "Sexuality": self.ensure_byte(1),
                    "CompetesAgainst": self.ensure_byte(2 if gender_value == 1 else 3),
                    "Outsiderel": self.ensure_byte(0),
                    "Birthday": birth_date,
                    "DebutDate": debut_date,
                    "DeathDate": "1666-01-01",  # Format as string for Access
                    "BodyType": self.ensure_byte(random.randint(0, 7)),
                    "WorkerHeight": self.ensure_byte(random.randint(20, 42)),
                    "WorkerWeight": weight,
                    "WorkerMinWeight": min_weight,
                    "WorkerMaxWeight": max_weight,
                    "Picture": f"{name.replace(' ', '').lower()}.jpg"[:35],
                    "Nationality": int(1),
                    "Race": self.ensure_byte(race),
                    "Based_In": self.ensure_byte(1),
                    "LeftBusiness": False,  # Will be converted to 0
                    "Dead": False,  # Will be converted to 0
                    "Retired": False,  # Will be converted to 0
                    "NonWrestler": False,  # Will be converted to 0
                    "Celebridad": self.ensure_byte(0),
                    "Style": self.ensure_byte(random.randint(1, 17)),
                    "Freelance": False,  # Will be converted to 0
                    "Loyalty": 0,
                    "TrueBorn": False,  # Will be converted to 0
                    "USA": True,  # Will be converted to -1
                    "Canada": True,  # Will be converted to -1
                    "Mexico": True,  # Will be converted to -1
                    "Japan": True,  # Will be converted to -1
                    "UK": True,  # Will be converted to -1
                    "Europe": True,  # Will be converted to -1
                    "Oz": True,  # Will be converted to -1
                    "India": True,  # Will be converted to -1
                    "Speak_English": int(4),
                    "Speak_Japanese": int(4),
                    "Speak_Spanish": int(4),
                    "Speak_French": int(4),
                    "Speak_Germanic": int(4),
                    "Speak_Med": int(4),
                    "Speak_Slavic": int(4),
                    "Speak_Hindi": int(1),
                    "Moveset": int(0),
                    "Position_Wrestler": True,  # Will be converted to -1
                    "Position_Occasional": False,  # Will be converted to 0
                    "Position_Referee": False,  # Will be converted to 0
                    "Position_Announcer": False,  # Will be converted to 0
                    "Position_Colour": False,  # Will be converted to 0
                    "Position_Manager": False,  # Will be converted to 0
                    "Position_Personality": False,  # Will be converted to 0
                    "Position_Roadagent": False,  # Will be converted to 0
                    "Mask": int(0),
                    "Age_Matures": self.ensure_byte(0),
                    "Age_Declines": self.ensure_byte(0),
                    "Age_TalkDeclines": self.ensure_byte(0),
                    "Age_Retires": self.ensure_byte(0),
                    "OrganicBio": True,  # Will be converted to -1
                    "PlasterCaster_Face": self.generate_gimmick(name, description, gender, "face")[:30],
                    "PlasterCaster_FaceBasis": self.ensure_byte(1),
                    "PlasterCaster_Heel": self.generate_gimmick(name, description, gender, "heel")[:30],
                    "PlasterCaster_HeelBasis": self.ensure_byte(1),
                    "CareerGoal": self.ensure_byte(0)
                }

                worker_row_converted = {
                    key: -1 if isinstance(value, bool) and value else 0 if isinstance(value, bool) else value
                    for key, value in worker_row.items()
                }
                workers_data.append(worker_row_converted)

                
                # Now get race using both description and bio
                
                worker_data = {
                    "UID": uid,
                    "name": name,
                    "shortname": name.split()[0],
                    "picture": f"{name.replace(' ', '').lower()}.jpg",
                    "gender": gender,
                    "gender_id": 1 if gender.lower() == "male" else 2,
                    "Race": race,
                    "style": style,
                    # ... rest of worker_data ...
                }

                bio_data.append([uid, bio])
                logging.debug(f"Generated bio for {wrestler_data['name']}")

                # Generate skills
                if wrestler_data['skill_preset'] == "Interpret":
                    preset_name = self.select_skill_preset_with_chatgpt(
                        wrestler_data['name'],
                        wrestler_data['description'],
                        wrestler_data['gender']
                    )
                else:
                    preset_name = wrestler_data['skill_preset']

                preset = next((p for p in self.skill_presets if p["name"] == preset_name), self.skill_presets[0])
                skills = self.generate_skills(uid, preset)
                skills_data.append(skills)
                logging.debug(f"Generated skills for {wrestler_data['name']} using preset {preset_name}")

                # Generate contract if not freelancer
                if wrestler_data['company'] != "Freelancer":
                    logging.debug(f"Creating contract with company: {wrestler_data['company']}")
                    contract = self.generate_contract(wrestler_data, uid, wrestler_data['company'], contract_uid)
                    contract_data.append(contract)
                    contract_uid += 1  # Increment for next contract
                    logging.debug(f"Generated contract for {wrestler_data['name']} with company {wrestler_data['company']}")

                # Add to notes data with picture path and generation flag
                notes_data.append({
                    "Name": name,
                    "Description": description,
                    "Gender": gender,
                    "Company": wrestler_data.get('company', 'Random'),
                    "Exclusive": wrestler_data.get('exclusive', 'Random'),
                    "Skill_Preset": wrestler_data.get('skill_preset', 'Default'),
                    "Picture": f"{name.replace(' ', '').lower()}.jpg"[:35],  # Match the Picture field
                    "image_generated": False
                })

                uid += 1

            # Save to Access DB if path exists
            if self.access_db_path and os.path.exists(self.access_db_path):
                logging.debug("Attempting to save to Access database.")
                try:
                    for worker_row in workers_data:
                        # Convert boolean values for Access DB (-1 for True, 0 for False)
                        worker_row_converted = {
                            key: -1 if isinstance(value, bool) and value else 0 if isinstance(value, bool) else value
                            for key, value in worker_row.items()
                        }

                        # Ensure text field lengths
                        worker_row_converted["Name"] = worker_row_converted["Name"][:30]
                        worker_row_converted["Shortname"] = worker_row_converted["Shortname"][:20]
                        worker_row_converted["Picture"] = worker_row_converted["Picture"][:35]
                        worker_row_converted["PlasterCaster_Face"] = worker_row_converted["PlasterCaster_Face"][:30]
                        worker_row_converted["PlasterCaster_Heel"] = worker_row_converted["PlasterCaster_Heel"][:30]

                        sql_insert_worker = """
                            INSERT INTO tblWorker (
                                [UID], [User], [Regen], [Active], [Name], [Shortname], [Gender], [Pronouns],
                                [Sexuality], [CompetesAgainst], [Outsiderel], [Birthday], [DebutDate], [DeathDate],
                                [BodyType], [WorkerHeight], [WorkerWeight], [WorkerMinWeight], [WorkerMaxWeight],
                                [Picture], [Nationality], [Race], [Based_In], [LeftBusiness], [Dead], [Retired],
                                [NonWrestler], [Celebridad], [Style], [Freelance], [Loyalty], [TrueBorn], [USA],
                                [Canada], [Mexico], [Japan], [UK], [Europe], [Oz], [India], [Speak_English],
                                [Speak_Japanese], [Speak_Spanish], [Speak_French], [Speak_Germanic], [Speak_Med],
                                [Speak_Slavic], [Speak_Hindi], [Moveset], [Position_Wrestler], [Position_Occasional],
                                [Position_Referee], [Position_Announcer], [Position_Colour], [Position_Manager],
                                [Position_Personality], [Position_Roadagent], [Mask], [Age_Matures], [Age_Declines],
                                [Age_TalkDeclines], [Age_Retires], [OrganicBio], [PlasterCaster_Face],
                                [PlasterCaster_FaceBasis], [PlasterCaster_Heel], [PlasterCaster_HeelBasis],
                                [CareerGoal]
                            ) VALUES (
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                            )
                        """

                        # Convert dates to proper Access format

                        debut_date = datetime.datetime(
                            random.randint(1980, 2023),  # Year between 1980 and 2023
                            random.randint(1, 12),       # Month
                            random.randint(1, 28)        # Day (avoiding potential month end issues)
                        )
                        debut_age = random.randint(16, 50) 
                        birth_year = debut_date.year - debut_age
                        birth_date = datetime.datetime(
                            birth_year,
                            random.randint(1, 12),  # Random birth month
                            random.randint(1, 28)   # Random birth day (avoiding month end issues)
                        )
                        
                        death_date = "1666-01-01"  # Format as string for Access

                        # Convert values to match Access data types
                        worker_values = [
                            int(worker_row_converted["UID"]),                    # Long Integer
                            bool(worker_row_converted["User"]),                  # Yes/No
                            int(worker_row_converted["Regen"]),                  # Byte
                            bool(worker_row_converted["Active"]),                # Yes/No
                            str(worker_row_converted["Name"]),                   # Text(30)
                            str(worker_row_converted["Shortname"]),              # Text(20)
                            int(worker_row_converted["Gender"]),                 # Byte
                            int(worker_row_converted["Pronouns"]),               # Byte
                            int(worker_row_converted["Sexuality"]),              # Byte
                            int(worker_row_converted["CompetesAgainst"]),        # Byte
                            int(worker_row_converted["Outsiderel"]),             # Byte
                            birth_date,                                            # Date/Time
                            debut_date,                                          # Date/Time
                            death_date,                                          # Date/Time
                            int(worker_row_converted["BodyType"]),               # Byte
                            int(worker_row_converted["WorkerHeight"]),           # Byte
                            int(worker_row_converted["WorkerWeight"]),           # Integer
                            int(worker_row_converted["WorkerMinWeight"]),        # Integer
                            int(worker_row_converted["WorkerMaxWeight"]),        # Integer
                            str(worker_row_converted["Picture"]),                # Text(35)
                            int(worker_row_converted["Nationality"]),            # Integer
                            int(worker_row_converted["Race"]),                   # Byte
                            int(worker_row_converted["Based_In"]),               # Byte
                            bool(worker_row_converted["LeftBusiness"]),          # Yes/No
                            bool(worker_row_converted["Dead"]),                  # Yes/No
                            bool(worker_row_converted["Retired"]),               # Yes/No
                            bool(worker_row_converted["NonWrestler"]),           # Yes/No
                            int(worker_row_converted["Celebridad"]),             # Byte
                            int(worker_row_converted["Style"]),                  # Byte
                            bool(worker_row_converted["Freelance"]),             # Yes/No
                            int(worker_row_converted["Loyalty"]),                # Long Integer
                            bool(worker_row_converted["TrueBorn"]),              # Yes/No
                            bool(worker_row_converted["USA"]),                   # Yes/No
                            bool(worker_row_converted["Canada"]),                # Yes/No
                            bool(worker_row_converted["Mexico"]),                # Yes/No
                            bool(worker_row_converted["Japan"]),                 # Yes/No
                            bool(worker_row_converted["UK"]),                    # Yes/No
                            bool(worker_row_converted["Europe"]),                # Yes/No
                            bool(worker_row_converted["Oz"]),                    # Yes/No
                            bool(worker_row_converted["India"]),                 # Yes/No
                            int(worker_row_converted["Speak_English"]),          # Integer
                            int(worker_row_converted["Speak_Japanese"]),         # Integer
                            int(worker_row_converted["Speak_Spanish"]),          # Integer
                            int(worker_row_converted["Speak_French"]),           # Integer
                            int(worker_row_converted["Speak_Germanic"]),         # Integer
                            int(worker_row_converted["Speak_Med"]),              # Integer
                            int(worker_row_converted["Speak_Slavic"]),           # Integer
                            int(worker_row_converted["Speak_Hindi"]),            # Integer
                            int(worker_row_converted["Moveset"]),                # Long Integer
                            bool(worker_row_converted["Position_Wrestler"]),     # Yes/No
                            bool(worker_row_converted["Position_Occasional"]),   # Yes/No
                            bool(worker_row_converted["Position_Referee"]),      # Yes/No
                            bool(worker_row_converted["Position_Announcer"]),    # Yes/No
                            bool(worker_row_converted["Position_Colour"]),       # Yes/No
                            bool(worker_row_converted["Position_Manager"]),      # Yes/No
                            bool(worker_row_converted["Position_Personality"]),  # Yes/No
                            bool(worker_row_converted["Position_Roadagent"]),    # Yes/No
                            int(worker_row_converted["Mask"]),                   # Integer
                            int(worker_row_converted["Age_Matures"]),            # Byte
                            int(worker_row_converted["Age_Declines"]),           # Byte
                            int(worker_row_converted["Age_TalkDeclines"]),       # Byte
                            int(worker_row_converted["Age_Retires"]),            # Byte
                            bool(worker_row_converted["OrganicBio"]),            # Yes/No
                            str(worker_row_converted["PlasterCaster_Face"]),     # Text(30)
                            int(worker_row_converted["PlasterCaster_FaceBasis"]), # Byte
                            str(worker_row_converted["PlasterCaster_Heel"]),     # Text(30)
                            int(worker_row_converted["PlasterCaster_HeelBasis"]), # Byte
                            int(worker_row_converted["CareerGoal"])              # Byte
                        ]

                        cursor.execute(sql_insert_worker, worker_values)

                    for bio_row in bio_data:
                        sql_insert_bio = "INSERT INTO tblWorkerBio ([UID], [Profile]) VALUES (?, ?)"
                        logging.debug(f"Executing Bio SQL with values: {bio_row}")
                        cursor.execute(sql_insert_bio, bio_row)

                    for skills_row in skills_data:
                        sql_insert_skills = """
                            INSERT INTO tblWorkerSkill (
                                WorkerUID, Brawl, Air, Technical, Power, Athletic, Stamina, 
                                Psych, Basics, Tough, Sell, Charisma, Mic, Menace, Respect, 
                                Reputation, Safety, Looks, Star, Consistency, Act, Injury, 
                                Puroresu, Flash, Hardcore, Announcing, Colour, Refereeing, 
                                Experience, PotentialPrimary, PotentialMental, PotentialPerformance, 
                                PotentialFundamental, PotentialPhysical, PotentialAnnouncing, 
                                PotentialColour, PotentialRefereeing, ScoutRing, ScoutPhysical, 
                                ScoutEnt, ScoutBroadcast, ScoutRef
                            ) VALUES (
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                            )
                        """
                        skills_values = [
                            skills_row["WorkerUID"],
                            skills_row.get("Brawl", 0),
                            skills_row.get("Air", 0),
                            skills_row.get("Technical", 0),
                            skills_row.get("Power", 0),
                            skills_row.get("Athletic", 0),
                            skills_row.get("Stamina", 0),
                            skills_row.get("Psych", 0),
                            skills_row.get("Basics", 0),
                            skills_row.get("Tough", 0),
                            skills_row.get("Sell", 0),
                            skills_row.get("Charisma", 0),
                            skills_row.get("Mic", 0),
                            skills_row.get("Menace", 0),
                            skills_row.get("Respect", 0),
                            skills_row.get("Reputation", 0),
                            skills_row.get("Safety", 0),
                            skills_row.get("Looks", 0),
                            skills_row.get("Star", 0),
                            skills_row.get("Consistency", 0),
                            skills_row.get("Act", 0),
                            skills_row.get("Injury", 0),
                            skills_row.get("Puroresu", 0),
                            skills_row.get("Flash", 0),
                            skills_row.get("Hardcore", 0),
                            skills_row.get("Announcing", 0),
                            skills_row.get("Colour", 0),
                            skills_row.get("Refereeing", 0),
                            skills_row.get("Experience", 0),
                            skills_row.get("PotentialPrimary", 0),
                            skills_row.get("PotentialMental", 0),
                            skills_row.get("PotentialPerformance", 0),
                            skills_row.get("PotentialFundamental", 0),
                            skills_row.get("PotentialPhysical", 0),
                            skills_row.get("PotentialAnnouncing", 0),
                            skills_row.get("PotentialColour", 0),
                            skills_row.get("PotentialRefereeing", 0),
                            skills_row.get("ScoutRing", 0),
                            skills_row.get("ScoutPhysical", 0),
                            skills_row.get("ScoutEnt", 0),
                            skills_row.get("ScoutBroadcast", 0),
                            skills_row.get("ScoutRef", 0)
                        ]
                        logging.debug(f"Executing Skills SQL with values: {skills_values}")
                        cursor.execute(sql_insert_skills, skills_values)

                    for contract in contract_data:
                        sql_insert_contract = """
                            INSERT INTO tblContract (
                                [UID], [FedUID], [WorkerUID], [Name], [Shortname], [Picture], 
                                [CompetesIn], [Face], [Division], [Manager], [Moveset], [WrittenContract],
                                [ExclusiveContract], [TouringContract], [PaidMonthly], [OnLoan], [Developmental],
                                [PrimaryUsage], [SecondaryUsage], [ExpectedShows], [BonusAmount], [BonusType],
                                [Creative], [HiringVeto], [WageMatch], [IronClad], [ContractBeganDate], [Daysleft],
                                [Dateslength], [DatesLeft], [ContractDebutDate], [Amount], [Downside], [Brand],
                                [Mask], [ContractMomentum], [Last_Turn], [Travel], [Position_Wrestler],
                                [Position_Occasional], [Position_Referee], [Position_Announcer], [Position_Colour],
                                [Position_Manager], [Position_Personality], [Position_Roadagent], [Merch],
                                [PlasterCaster_Gimmick], [PlasterCaster_Rating], [PlasterCaster_Lifespan],
                                [PlasterCaster_Byte1], [PlasterCaster_Byte2], [PlasterCaster_Byte3],
                                [PlasterCaster_Byte4], [PlasterCaster_Byte5], [PlasterCaster_Byte6],
                                [PlasterCaster_Bool1], [PlasterCaster_Bool2], [PlasterCaster_Bool3],
                                [PlasterCaster_Bool4], [PlasterCaster_Bool5], [PlasterCaster_Bool6],
                                [PlasterCaster_Bool7], [PlasterCaster_Bool8], [PlasterCaster_Bool9],
                                [PlasterCaster_Bool10], [PlasterCaster_Bool11], [PlasterCaster_Bool12],
                                [PlasterCaster_Bool13], [PlasterCaster_Bool14], [PlasterCaster_Bool15],
                                [PlasterCaster_Bool16], [PlasterCaster_Bool17], [PlasterCaster_Bool18],
                                [PlasterCaster_Bool19], [PlasterCaster_Bool20], [PlasterCaster_Bool21],
                                [PlasterCaster_Bool22], [PlasterCaster_Bool23], [PlasterCaster_Bool24],
                                [PlasterCaster_Bool25]
                            ) VALUES (
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                                ?, ?, ?, ?
                            )
                        """
                        # Let's count and verify our values match exactly
                        contract_values = [
                            contract["UID"],                    # 1
                            contract["FedUID"],                 # 2
                            contract["WorkerUID"],              # 3
                            contract["Name"],                   # 4
                            contract["Shortname"],              # 5
                            contract["Picture"],                # 6
                            contract["CompetesIn"],             # 7
                            contract["Face"],                   # 8
                            contract["Division"],               # 9
                            contract["Manager"],                # 10
                            contract["Moveset"],                # 11
                            contract["WrittenContract"],        # 12
                            contract["ExclusiveContract"],      # 13
                            contract["TouringContract"],        # 14
                            contract["PaidMonthly"],            # 15
                            contract["OnLoan"],                 # 16
                            contract["Developmental"],          # 17
                            contract["PrimaryUsage"],           # 18
                            contract["SecondaryUsage"],         # 19
                            contract["ExpectedShows"],          # 20
                            contract["BonusAmount"],            # 21
                            contract["BonusType"],              # 22
                            contract["Creative"],               # 23
                            contract["HiringVeto"],             # 24
                            contract["WageMatch"],              # 25
                            contract["IronClad"],               # 26
                            contract["ContractBeganDate"],      # 27
                            contract["Daysleft"],               # 28
                            contract["Dateslength"],            # 29
                            contract["DatesLeft"],              # 30
                            contract["ContractDebutDate"],      # 31
                            contract["Amount"],                 # 32
                            contract["Downside"],               # 33
                            contract["Brand"],                  # 34
                            contract["Mask"],                   # 35
                            contract["ContractMomentum"],       # 36
                            contract["Last_Turn"],              # 37
                            contract["Travel"],                 # 38
                            contract["Position_Wrestler"],      # 39
                            contract["Position_Occasional"],    # 40
                            contract["Position_Referee"],       # 41
                            contract["Position_Announcer"],     # 42
                            contract["Position_Colour"],        # 43
                            contract["Position_Manager"],       # 44
                            contract["Position_Personality"],   # 45
                            contract["Position_Roadagent"],     # 46
                            contract["Merch"],                  # 47
                            contract["PlasterCaster_Gimmick"],  # 48
                            contract["PlasterCaster_Rating"],   # 49
                            contract["PlasterCaster_Lifespan"], # 50
                            contract["PlasterCaster_Byte1"],    # 51
                            contract["PlasterCaster_Byte2"],    # 52
                            contract["PlasterCaster_Byte3"],    # 53
                            contract["PlasterCaster_Byte4"],    # 54
                            contract["PlasterCaster_Byte5"],    # 55
                            contract["PlasterCaster_Byte6"],    # 56
                            contract["PlasterCaster_Bool1"],    # 57
                            contract["PlasterCaster_Bool2"],    # 58
                            contract["PlasterCaster_Bool3"],    # 59
                            contract["PlasterCaster_Bool4"],    # 60
                            contract["PlasterCaster_Bool5"],    # 61
                            contract["PlasterCaster_Bool6"],    # 62
                            contract["PlasterCaster_Bool7"],    # 63
                            contract["PlasterCaster_Bool8"],    # 64
                            contract["PlasterCaster_Bool9"],    # 65
                            contract["PlasterCaster_Bool10"],   # 66
                            contract["PlasterCaster_Bool11"],   # 67
                            contract["PlasterCaster_Bool12"],   # 68
                            contract["PlasterCaster_Bool13"],   # 69
                            contract["PlasterCaster_Bool14"],   # 70
                            contract["PlasterCaster_Bool15"],   # 71
                            contract["PlasterCaster_Bool16"],   # 72
                            contract["PlasterCaster_Bool17"],   # 73
                            contract["PlasterCaster_Bool18"],   # 74
                            contract["PlasterCaster_Bool19"],   # 75
                            contract["PlasterCaster_Bool20"],   # 76
                            contract["PlasterCaster_Bool21"],   # 77
                            contract["PlasterCaster_Bool22"],   # 78
                            contract["PlasterCaster_Bool23"],   # 79
                            contract["PlasterCaster_Bool24"],   # 80
                            contract["PlasterCaster_Bool25"]    # 81
                        ]
                        logging.debug(f"Executing Contract SQL with values: {contract_values}")
                        cursor.execute(sql_insert_contract, contract_values)

                    conn.commit()
                    logging.debug("Successfully committed all changes to Access database.")
                    conn.close()
                    logging.debug("Successfully closed database connection.")
                except Exception as e:
                    logging.error(f"Error saving to Access database: {e}", exc_info=True)
                    messagebox.showerror("Error", f"Could not save to Access database: {str(e)}")

            # Save to Excel as backup
            try:
                logging.debug("Attempting to save data to Excel file.")
                workers_df = pd.DataFrame(workers_data)
                bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                skills_df = pd.DataFrame(skills_data)
                contracts_df = pd.DataFrame(contract_data)
                notes_df = pd.DataFrame(notes_data)
                
                excel_path = "wrestleverse_workers.xlsx"
                with pd.ExcelWriter(excel_path) as writer:
                    workers_df.to_excel(writer, sheet_name="Workers", index=False)
                    bio_df.to_excel(writer, sheet_name="Bios", index=False)
                    skills_df.to_excel(writer, sheet_name="Skills", index=False)
                    contracts_df.to_excel(writer, sheet_name="Contracts", index=False)
                    notes_df.to_excel(writer, sheet_name="Notes", index=False)
                
                logging.debug(f"Successfully saved data to Excel file at {excel_path}")
                messagebox.showinfo("Success", f"Wrestlers saved to {excel_path}")
            except Exception as e:
                logging.error(f"Error saving Excel file: {e}", exc_info=True)
                messagebox.showerror("Error", f"Could not save Excel file: {str(e)}")
                uid = result[0] + 1 if result[0] else 1
                
                # Get next contract UID
                cursor.execute("SELECT MAX(UID) FROM tblContract")
                result = cursor.fetchone()
                contract_uid = result[0] + 1 if result[0] else 1

            # Inside the try block, just before saving to Excel:
            for note in notes_data:
                try:
                    # Get physical description from GPT
                    physical_prompt = (
                        f"Based on this wrestler's details:\n"
                        f"Name: {note['Name']}\n"
                        f"Description: {note['Description']}\n"
                        f"Gender: {note['Gender']}\n"
                        f"Please provide a single sentence describing their physical appearance. "
                        f"Focus on height, build, and distinctive features."
                    )
                    physical_description = self.get_response_from_gpt(physical_prompt)
                    note['physical_description'] = physical_description
                    logging.debug(f"Generated physical description for {note['Name']}: {physical_description}")
                except Exception as e:
                    logging.error(f"Error generating physical description: {e}")
                    note['physical_description'] = ""

            # Save to Excel
            

        except Exception as e:
            logging.error(f"Unhandled error in generate_wrestlers: {e}", exc_info=True)
            error_message = f"Error generating wrestlers: {str(e)}"
            logging.error(error_message, exc_info=True)
            self.status_label.config(text=f"Status: Error - {str(e)}")
            messagebox.showerror("Error", error_message)
        finally:
            self.status_label.config(text="Status: Generation complete!")
            self.root.update_idletasks()

    def open_company_generator(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Company Generator", font=("Helvetica", 16))
        title_label.pack(pady=10)
        add_company_btn = ttk.Button(self.root, text="Add Company", command=self.add_company_form)
        add_company_btn.pack(pady=10)
        self.companies_frame = ttk.Frame(self.root)
        self.companies_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.status_label = ttk.Label(self.root, text="Status: Waiting to generate companies...")
        self.status_label.pack(pady=10)
        generate_btn = ttk.Button(self.root, text="Generate Companies", command=self.generate_companies)
        generate_btn.pack(side="bottom", pady=10)
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def add_company_form(self):
        company_frame = ttk.Frame(self.companies_frame, relief="ridge", borderwidth=2)
        company_frame.pack(fill="x", pady=5)
        name_label = ttk.Label(company_frame, text="Name:")
        name_label.grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(company_frame)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        description_label = ttk.Label(company_frame, text="Description:")
        description_label.grid(row=1, column=0, padx=5, pady=5)
        description_entry = ttk.Entry(company_frame, width=40)
        description_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5)
        size_label = ttk.Label(company_frame, text="Size:")
        size_label.grid(row=2, column=0, padx=5, pady=5)
        size_var = tk.StringVar(value="Medium")
        size_dropdown = ttk.Combobox(company_frame, textvariable=size_var, values=["Tiny", "Small", "Medium", "Large"])
        size_dropdown.grid(row=2, column=1, padx=5, pady=5)
        remove_btn = ttk.Button(company_frame, text="❌", command=lambda: self.remove_company_form(company_frame))
        remove_btn.grid(row=0, column=4, padx=5, pady=5)
        self.companies.append({
            "frame": company_frame,
            "name": name_entry,
            "description": description_entry,
            "size": size_var
        })

    def remove_company_form(self, company_frame):
        company_frame.destroy()
        self.companies = [company for company in self.companies if company["frame"] != company_frame]



    def generate_company_name(self, description=None, size=None):
        prompt = "Generate a name for a professional wrestling company."
        if description:
            prompt += f" The company's style or theme is: {description}."
        if size:
            prompt += f" The company is of {size.lower()} size."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def generate_company_description(self, name=None, size=None):
        prompt = "Generate a description for a professional wrestling company."
        if name:
            prompt += f" The company's name is {name}."
        if size:
            prompt += f" The company is of {size.lower()} size."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def ensure_byte(self, value):
        return max(0, min(255, int(value)))

    def add_years(self, d, years):
        try:
            return d.replace(year=d.year + years)
        except ValueError:
            return d.replace(month=2, day=28, year=d.year + years)

    def random_date(self, start_date, end_date):
        delta = end_date - start_date
        int_delta = delta.days
        if int_delta <= 0:
            return start_date
        random_day = random.randrange(int_delta)
        return start_date + datetime.timedelta(days=random_day)

    def generate_birthday(self, max_year=2007):
        start_date = datetime.datetime(1970, 1, 1)
        end_date = datetime.datetime(max_year, 12, 31)
        return self.random_date(start_date, end_date)

    def generate_name(self, description=None, gender=None):
        prompt = "Generate a full name for a professional wrestler."
        if description:
            prompt += f" The wrestler's gimmick or description is: {description}."
        if gender:
            prompt += f" The wrestler is {gender}."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def generate_gimmick(self, name, description, gender, alignment):
        prompt = f"Generate a wrestling gimmick for a {alignment} wrestler."
        if name:
            prompt += f" The wrestler's name is {name}."
        if description:
            prompt += f" Description: {description}."
        if gender:
            prompt += f" Gender: {gender}."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def generate_bio(self, name, gender, description, skill_preset_name):
        prompt = self.bio_prompt
        if name:
            prompt += f" The wrestler's name is {name}."
        if gender:
            prompt += f" Their gender is {gender}."
        if description:
            prompt += f" Description: {description}."
        if skill_preset_name:
            prompt += f" Their wrestling style is best described as {skill_preset_name}."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()

    def select_skill_preset_with_chatgpt(self, name, description, gender):
        preset_names = [preset["name"] for preset in self.skill_presets]
        prompt = f"Based on the following wrestler details, select the most appropriate skill preset from the list.\n"
        prompt += f"Wrestler Name: {name}\n"
        if description:
            prompt += f"Description: {description}\n"
        if gender:
            prompt += f"Gender: {gender}\n"
        prompt += "Available Skill Presets: " + ", ".join(preset_names) + "\n"
        prompt += "Provide only the name of the most suitable skill preset."
        response = self.client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        selected_preset_name = response.choices[0].message.content.strip()
        return selected_preset_name

    def generate_skills(self, uid, skill_preset):
        skills_with_defaults = self.get_skills_with_defaults()
        skills = {"WorkerUID": uid}
        for skill in self.get_all_skills():
            min_value = skill_preset["skills"][skill]["min"]
            max_value = skill_preset["skills"][skill]["max"]
            if min_value == max_value:
                value = min_value
            else:
                value = random.randint(min_value, max_value)
            skills[skill] = value
        return skills

    def open_settings(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        settings_title = ttk.Label(self.root, text="Settings", font=("Helvetica", 16))
        settings_title.pack(pady=10)
        
        # API Key
        api_key_label = ttk.Label(self.root, text="ChatGPT API Key:")
        api_key_label.pack(pady=5)
        self.api_key_var = tk.StringVar(value=self.api_key)
        api_key_entry = ttk.Entry(self.root, textvariable=self.api_key_var, width=50)
        api_key_entry.pack(pady=5)
        
        # UID Start
        uid_start_label = ttk.Label(self.root, text="UID Start:")
        uid_start_label.pack(pady=5)
        self.uid_start_var = tk.IntVar(value=self.uid_start)
        uid_start_entry = ttk.Entry(self.root, textvariable=self.uid_start_var, width=10)
        uid_start_entry.pack(pady=5)
        
        # Bio Prompt
        bio_prompt_label = ttk.Label(self.root, text="Bio Prompt:")
        bio_prompt_label.pack(pady=5)
        self.bio_prompt_var = tk.StringVar(value=self.bio_prompt)
        bio_prompt_entry = ttk.Entry(self.root, textvariable=self.bio_prompt_var, width=50)
        bio_prompt_entry.pack(pady=5)
        
        # Access DB Path
        access_db_label = ttk.Label(self.root, text="Access Database Path (Optional):")
        access_db_label.pack(pady=5)
        self.access_db_var = tk.StringVar(value=self.access_db_path)
        access_db_entry = ttk.Entry(self.root, textvariable=self.access_db_var, width=50)
        access_db_entry.pack(pady=5)
        browse_db_btn = ttk.Button(self.root, text="Browse", command=self.browse_access_db)
        browse_db_btn.pack(pady=5)
        
        # Pictures Path
        pictures_label = ttk.Label(self.root, text="Pictures Path (Optional):")
        pictures_label.pack(pady=5)
        self.pictures_var = tk.StringVar(value=self.pictures_path)
        pictures_entry = ttk.Entry(self.root, textvariable=self.pictures_var, width=50)
        pictures_entry.pack(pady=5)
        browse_pics_btn = ttk.Button(self.root, text="Browse", command=self.browse_pictures_path)
        browse_pics_btn.pack(pady=5)
        
        # Save Button
        save_btn = ttk.Button(self.root, text="Save", command=self.save_settings)
        save_btn.pack(pady=10)
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def browse_access_db(self):
        file_path = filedialog.askopenfilename(filetypes=[("Access Database Files", "*.accdb;*.mdb")])
        if file_path:
            self.access_db_var.set(file_path)

    def save_settings(self):
        self.api_key = self.api_key_var.get()
        self.uid_start = self.uid_start_var.get()
        self.bio_prompt = self.bio_prompt_var.get()
        self.access_db_path = self.access_db_var.get()
        self.pictures_path = self.pictures_var.get()
        settings = {
            "api_key": self.api_key,
            "uid_start": self.uid_start,
            "bio_prompt": self.bio_prompt,
            "access_db_path": self.access_db_path,
            "pictures_path": self.pictures_path
        }
        with open("settings.json", "w") as settings_file:
            json.dump(settings, settings_file)
        messagebox.showinfo("Settings", "Settings saved successfully!")
        if self.api_key:
            self.client = OpenAI(api_key=self.api_key)
        else:
            self.client = None

    def load_settings(self):
        try:
            with open("settings.json", "r") as settings_file:
                settings = json.load(settings_file)
                self.api_key = settings.get("api_key", "")
                self.uid_start = settings.get("uid_start", 1)
                self.bio_prompt = settings.get("bio_prompt", "Create a biography for a professional wrestler.")
                self.access_db_path = settings.get("access_db_path", "")
                self.pictures_path = settings.get("pictures_path", "")
        except FileNotFoundError:
            self.api_key = ""
            self.uid_start = 1
            self.bio_prompt = "Create a biography for a professional wrestler."
            self.access_db_path = ""
            self.pictures_path = ""

    def open_skill_presets(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Skill Presets", font=("Helvetica", 16))
        title_label.pack(pady=10)
        presets_frame = ttk.Frame(self.root)
        presets_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.presets_listbox = tk.Listbox(presets_frame)
        self.presets_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(presets_frame, orient="vertical", command=self.presets_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.presets_listbox.config(yscrollcommand=scrollbar.set)
        self.presets_listbox.delete(0, tk.END)
        for preset in self.skill_presets:
            self.presets_listbox.insert("end", preset["name"])
        buttons_frame = ttk.Frame(self.root)
        buttons_frame.pack(pady=10)
        add_btn = ttk.Button(buttons_frame, text="Add Preset", command=self.add_skill_preset)
        add_btn.pack(side="left", padx=5)
        edit_btn = ttk.Button(buttons_frame, text="Edit Preset", command=self.edit_skill_preset)
        edit_btn.pack(side="left", padx=5)
        delete_btn = ttk.Button(buttons_frame, text="Delete Preset", command=self.delete_skill_preset)
        delete_btn.pack(side="left", padx=5)
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def add_skill_preset(self):
        self.preset_window = tk.Toplevel(self.root)
        self.preset_window.title("Add Skill Preset")
        self.preset_window.geometry("400x500")
        name_label = ttk.Label(self.preset_window, text="Preset Name:")
        name_label.pack(pady=5)
        self.preset_name_var = tk.StringVar()
        name_entry = ttk.Entry(self.preset_window, textvariable=self.preset_name_var)
        name_entry.pack(pady=5)
        container = ttk.Frame(self.preset_window)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        skills_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=skills_frame, anchor='nw')
        self.skill_entries = {}
        for i, skill in enumerate(self.get_all_skills()):
            row = ttk.Frame(skills_frame)
            row.pack(fill="x", pady=2)
            skill_label = ttk.Label(row, text=skill, width=15)
            skill_label.pack(side="left")
            min_var = tk.IntVar(value=20)
            max_var = tk.IntVar(value=90)
            if skill in self.get_skills_with_defaults():
                default_value = self.get_skills_with_defaults()[skill]
                min_var.set(default_value)
                max_var.set(default_value)
            min_entry = ttk.Entry(row, textvariable=min_var, width=5)
            min_entry.pack(side="left", padx=5)
            max_entry = ttk.Entry(row, textvariable=max_var, width=5)
            max_entry.pack(side="left", padx=5)
            self.skill_entries[skill] = {"min": min_var, "max": max_var}
        save_btn = ttk.Button(self.preset_window, text="Save Preset", command=self.save_new_preset)
        save_btn.pack(pady=10)

    def save_new_preset(self):
        preset_name = self.preset_name_var.get().strip()
        if not preset_name:
            messagebox.showerror("Error", "Preset name cannot be empty.")
            return
        if any(preset["name"] == preset_name for preset in self.skill_presets):
            messagebox.showerror("Error", "A preset with that name already exists.")
            return
        preset_skills = {}
        for skill, vars in self.skill_entries.items():
            try:
                min_value = int(vars["min"].get())
                max_value = int(vars["max"].get())
                if not (0 <= min_value <= 100) or not (0 <= max_value <= 100):
                    raise ValueError
                if min_value > max_value:
                    messagebox.showerror("Error", f"For skill {skill}, min value cannot be greater than max value.")
                    return
                preset_skills[skill] = {"min": min_value, "max": max_value}
            except ValueError:
                messagebox.showerror("Error", f"Invalid input for skill {skill}. Please enter integers between 0 and 100.")
                return
        new_preset = {
            "name": preset_name,
            "skills": preset_skills
        }
        self.skill_presets.append(new_preset)
        self.save_skill_presets()
        messagebox.showinfo("Success", "Skill preset saved successfully.")
        self.preset_window.destroy()
        self.open_skill_presets()

    def edit_skill_preset(self):
        selected_indices = self.presets_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Please select a preset to edit.")
            return
        index = selected_indices[0]
        preset = self.skill_presets[index]
        self.preset_window = tk.Toplevel(self.root)
        self.preset_window.title("Edit Skill Preset")
        self.preset_window.geometry("400x500")
        name_label = ttk.Label(self.preset_window, text="Preset Name:")
        name_label.pack(pady=5)
        self.preset_name_var = tk.StringVar(value=preset["name"])
        name_entry = ttk.Entry(self.preset_window, textvariable=self.preset_name_var)
        name_entry.pack(pady=5)
        container = ttk.Frame(self.preset_window)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        skills_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=skills_frame, anchor='nw')
        self.skill_entries = {}
        for i, skill in enumerate(self.get_all_skills()):
            row = ttk.Frame(skills_frame)
            row.pack(fill="x", pady=2)
            skill_label = ttk.Label(row, text=skill, width=15)
            skill_label.pack(side="left")
            min_var = tk.IntVar()
            max_var = tk.IntVar()
            min_var.set(preset["skills"][skill]["min"])
            max_var.set(preset["skills"][skill]["max"])
            min_entry = ttk.Entry(row, textvariable=min_var, width=5)
            min_entry.pack(side="left", padx=5)
            max_entry = ttk.Entry(row, textvariable=max_var, width=5)
            max_entry.pack(side="left", padx=5)
            self.skill_entries[skill] = {"min": min_var, "max": max_var}
        save_btn = ttk.Button(self.preset_window, text="Save Preset", command=lambda: self.save_edited_preset(index))
        save_btn.pack(pady=10)

    def save_edited_preset(self, index):
        preset_name = self.preset_name_var.get().strip()
        if not preset_name:
            messagebox.showerror("Error", "Preset name cannot be empty.")
            return
        if any(i != index and preset["name"] == preset_name for i, preset in enumerate(self.skill_presets)):
            messagebox.showerror("Error", "A preset with that name already exists.")
            return
        preset_skills = {}
        for skill, vars in self.skill_entries.items():
            try:
                min_value = int(vars["min"].get())
                max_value = int(vars["max"].get())
                if not (0 <= min_value <= 100) or not (0 <= max_value <= 100):
                    raise ValueError
                if min_value > max_value:
                    messagebox.showerror("Error", f"For skill {skill}, min value cannot be greater than max value.")
                    return
                preset_skills[skill] = {"min": min_value, "max": max_value}
            except ValueError:
                messagebox.showerror("Error", f"Invalid input for skill {skill}. Please enter integers between 0 and 100.")
                return
        self.skill_presets[index]["name"] = preset_name
        self.skill_presets[index]["skills"] = preset_skills
        self.save_skill_presets()
        messagebox.showinfo("Success", "Skill preset updated successfully.")
        self.preset_window.destroy()
        self.open_skill_presets()

    def delete_skill_preset(self):
        selected_indices = self.presets_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Please select a preset to delete.")
            return
        index = selected_indices[0]
        preset = self.skill_presets[index]
        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the preset '{preset['name']}'?")
        if confirm:
            del self.skill_presets[index]
            self.save_skill_presets()
            messagebox.showinfo("Success", "Skill preset deleted successfully.")
            self.open_skill_presets()

    def load_skill_presets(self):
        try:
            with open("skill_presets.json", "r") as f:
                self.skill_presets = json.load(f)
        except FileNotFoundError:
            skills_list = self.get_all_skills()
            skills_with_defaults = self.get_skills_with_defaults()
            default_preset = {
                "name": "Default",
                "skills": {}
            }
            for skill in skills_list:
                if skill in skills_with_defaults:
                    default_value = skills_with_defaults[skill]
                    default_preset["skills"][skill] = {"min": default_value, "max": default_value}
                else:
                    default_preset["skills"][skill] = {"min": 20, "max": 90}
            self.skill_presets = [default_preset]
            self.save_skill_presets()

    def save_skill_presets(self):
        with open("skill_presets.json", "w") as f:
            json.dump(self.skill_presets, f, indent=4)

    def get_all_skills(self):
        return [
            "Brawl", "Air", "Technical", "Power", "Athletic", "Stamina", "Psych", "Basics", "Tough",
            "Sell", "Charisma", "Mic", "Menace", "Respect", "Reputation", "Safety", "Looks", "Star",
            "Consistency", "Act", "Injury", "Puroresu", "Flash", "Hardcore", "Announcing", "Colour",
            "Refereeing", "Experience", "PotentialPrimary", "PotentialMental", "PotentialPerformance",
            "PotentialFundamental", "PotentialPhysical", "PotentialAnnouncing", "PotentialColour",
            "PotentialRefereeing", "ScoutRing", "ScoutPhysical", "ScoutEnt", "ScoutBroadcast", "ScoutRef"
        ]

    def get_skills_with_defaults(self):
        return {
            "Respect": 100,
            "Reputation": 100,
            "Announcing": 0,
            "Colour": 0,
            "Refereeing": 0,
            "Experience": 100,
            "PotentialPrimary": 0,
            "PotentialMental": 0,
            "PotentialPerformance": 0,
            "PotentialFundamental": 0,
            "PotentialPhysical": 0,
            "PotentialAnnouncing": 0,
            "PotentialColour": 0,
            "PotentialRefereeing": 0,
            "ScoutRing": 6,
            "ScoutPhysical": 6,
            "ScoutEnt": 6,
            "ScoutBroadcast": 6,
            "ScoutRef": 6
        }

    def get_companies(self):
        """Get list of companies from the database"""
        if self.access_db_path and os.path.exists(self.access_db_path):
            try:
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={self.access_db_path};'
                    'PWD=20YearsOfTEW;'
                )
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                
                cursor.execute("SELECT UID, Name FROM tblFed")
                companies = cursor.fetchall()
                conn.close()
                
                # Log what we got from the database
                logging.debug(f"Retrieved companies from database: {companies}")
                
                # Convert to list of tuples and ensure UIDs are integers
                company_list = [(int(company[0]), company[1]) for company in companies]
                logging.debug(f"Converted company list: {company_list}")
                
                return company_list
            except Exception as e:
                logging.error(f"Error getting companies: {e}", exc_info=True)
                return []
        return []

    def generate_contract(self, worker_data, worker_uid, company_choice, contract_uid):
        """Generate contract data for a worker"""
        logging.debug(f"Generating contract for worker {worker_uid} with company choice: '{company_choice}'")
        
        # Determine company UID
        companies = self.get_companies()
        fed_uid = None
        
        if companies:
            if company_choice == "Random":
                company = random.choice(companies)
                fed_uid = company[0]
                logging.debug(f"Random company selected: {company[1]} (UID: {fed_uid})")
            elif company_choice == "Freelancer":
                logging.debug("Freelancer selected - no contract needed")
                return None
            else:
                # Find exact matching company by name
                for company in companies:
                    if company[1] == company_choice:  # Exact match
                        fed_uid = company[0]
                        logging.debug(f"Found exact matching company: {company[1]} (UID: {fed_uid})")
                        break
                
                if fed_uid is None:
                    logging.error(f"No matching company found for: {company_choice}")
                    return None

        # Determine face/heel alignment
        alignment_prompt = f"For a wrestler named {worker_data['name']}, should they be a face (good guy) or heel (bad guy)? Answer with only the word 'face' or 'heel'."
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": alignment_prompt}]
            )
            alignment = response.choices[0].message.content.strip().lower()
            is_face = alignment == "face"
        except Exception as e:
            logging.error(f"Error getting alignment from GPT: {e}", exc_info=True)
            is_face = random.choice([True, False])

        # Generate contract dates
        try:
            # Contract began date (between 2010 and 2022)
            contract_began = datetime.datetime(
                random.randint(2010, 2022),
                random.randint(1, 12),
                random.randint(1, 28)  # Using 28 to avoid month end issues
            )
            
            # Contract debut date (using 1900 instead of 1666)
            contract_debut = datetime.datetime(1900, 1, 1)
            
            logging.debug(f"Generated dates - Began: {contract_began}, Debut: {contract_debut}")
        except ValueError as e:
            logging.error(f"Error generating dates: {e}")
            # Fallback dates if there's an error
            contract_began = datetime.datetime(2020, 1, 1)
            contract_debut = datetime.datetime(1900, 1, 1)

        # Determine exclusive contract status
        exclusive_choice = worker_data.get('exclusive', 'Random')
        is_exclusive = random.choice([True, False]) if exclusive_choice == 'Random' else exclusive_choice == 'Yes'

        contract_data = {
            "UID": contract_uid,
            "FedUID": fed_uid,
            "WorkerUID": worker_uid,
            "Name": worker_data['name'][:30],
            "Shortname": worker_data['name'].split()[0][:20],
            "Picture": f"{worker_data['name'].replace(' ', '').lower()}.jpg"[:35],
            "CompetesIn": 1 if worker_data['gender'].lower() == 'male' else 2,
            "Face": is_face,
            "Division": 0,
            "Manager": 0,
            "Moveset": 0,
            "WrittenContract": True,
            "ExclusiveContract": is_exclusive,
            "TouringContract": False,
            "PaidMonthly": True,
            "OnLoan": False,
            "Developmental": False,
            "PrimaryUsage": 0,
            "SecondaryUsage": 0,
            "ExpectedShows": 0,
            "BonusAmount": 0,
            "BonusType": 0,
            "Creative": False,
            "HiringVeto": False,
            "WageMatch": False,
            "IronClad": is_exclusive,
            "ContractBeganDate": contract_began,
            "Daysleft": random.randint(7, 300),
            "Dateslength": 0,
            "DatesLeft": 0,
            "ContractDebutDate": contract_debut,
            "Amount": -1,
            "Downside": -1,
            "Brand": 0,
            "Mask": 0,
            "ContractMomentum": random.randint(1, 5),
            "Last_Turn": 0,
            "Travel": True,
            "Position_Wrestler": False,
            "Position_Occasional": False,
            "Position_Referee": False,
            "Position_Announcer": False,
            "Position_Colour": False,
            "Position_Manager": False,
            "Position_Personality": False,
            "Position_Roadagent": False,
            "Merch": 200,
            "PlasterCaster_Gimmick": worker_data.get('face_gimmick', '')[:30],
            "PlasterCaster_Rating": random.choice([0, 3, 4, 5, 6]),
            "PlasterCaster_Lifespan": random.randint(0, 4)
        }

        # Add all PlasterCaster_Byte fields (1-6)
        for i in range(1, 7):
            contract_data[f"PlasterCaster_Byte{i}"] = 0

        # Add all PlasterCaster_Bool fields (1-25)
        for i in range(1, 26):
            contract_data[f"PlasterCaster_Bool{i}"] = False

        return contract_data

    def get_race_from_gpt(self, name, description):
        """Get race index from GPT based on name and description"""
        prompt = (
            f"Based on the name '{name}' and description '{description}', select the most appropriate "
            "race from this list and respond with ONLY the corresponding number:\n"
            "1: White\n2: Black\n3: Asian\n4: Hispanic\n5: American Indian\n"
            "6: Middle Eastern\n7: South Asian\n8: Pacific\n9: Other\n\n"
            "Respond with only the number."
        )
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )
            race_str = response.choices[0].message.content.strip()
            logging.info(f"Race response from GPT: {race_str}")
            # Try to convert to int and validate range
            try:
                race = int(race_str)
                if 1 <= race <= 9:
                    return race
            except ValueError:
                pass
            
            # If we get here, either conversion failed or number was out of range
            logging.warning(f"Invalid race response from GPT: {race_str}, defaulting to 9")
            return 9
            
        except Exception as e:
            logging.error(f"Error getting race from GPT: {e}", exc_info=True)
            return 9

    def save_to_access_db(self, worker_data, bio_data, skills_data, contract_data):
        """Save generated data to Access database"""
        logging.debug("Attempting to save to Access database.")
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.access_db_path};'
                'PWD=20YearsOfTEW;'
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Insert worker data
            worker_sql = """
                INSERT INTO tblWorker (
                    UID, Name, ShortName, Picture, Gender, Race, Nationality, Height, Weight, 
                    BirthDate, DebutDate, Face_Gimmick, Heel_Gimmick, Style
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            worker_values = [
                worker_data['UID'], worker_data['name'], worker_data['shortname'], 
                worker_data['picture'], worker_data['gender_id'], worker_data['Race'], 
                worker_data['nationality'], worker_data['height'], worker_data['weight'],
                worker_data['Birthday'], worker_data['DebutDate'], 
                worker_data['face_gimmick'], worker_data['heel_gimmick'], 
                worker_data['style']
            ]
            cursor.execute(worker_sql, worker_values)
            
            # ... rest of database saving code ...

            conn.commit()
            logging.info(f"Created wrestler: {worker_data['name']} (UID: {worker_data['UID']}) - " 
                        f"Race: {worker_data['Race']}, Nationality: {worker_data['nationality']}")
            
            # Print a summary of the created wrestler
            print(f"\nCreated new wrestler:")
            print(f"Name: {worker_data['name']}")
            print(f"Race: {worker_data['Race']}")
            print(f"Nationality: {worker_data['nationality']}")
            print(f"Style: {worker_data['style']}")
            if contract_data:
                print(f"Company: {contract_data['Name']}")
            else:
                print("Status: Freelancer")
            print("-" * 50)

            return True
        except Exception as e:
            logging.error(f"Error saving to Access database: {e}", exc_info=True)
            return False

    def get_response_from_gpt(self, prompt):
        """Helper method to get responses from GPT"""
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logging.error(f"Error getting response from GPT: {e}", exc_info=True)
            return ""

    def open_image_generator(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        title_label = ttk.Label(self.root, text="Image Generator", font=("Helvetica", 16))
        title_label.pack(pady=20)
        
        generate_wrestler_images_btn = ttk.Button(self.root, text="Generate Wrestler Images", command=self.generate_wrestler_images)
        generate_wrestler_images_btn.pack(pady=10)
        
        generate_company_images_btn = ttk.Button(self.root, text="Generate Company Images", command=self.generate_company_images)
        generate_company_images_btn.pack(pady=10)
        
        self.status_label = ttk.Label(self.root, text="")
        self.status_label.pack(pady=10)
        
        back_btn = ttk.Button(self.root, text="Back", command=self.setup_main_menu)
        back_btn.pack(side="bottom", pady=10)

    def generate_wrestler_images(self):
        if not self.api_key:
            messagebox.showerror("Error", "Please set your API key in settings first.")
            return
        
        if not self.pictures_path:
            messagebox.showerror("Error", "Please set your pictures path in settings first.")
            return
        
        try:
            # Read the Excel file
            excel_path = "wrestleverse_workers.xlsx"
            if not os.path.exists(excel_path):
                messagebox.showerror("Error", "Wrestlers Excel file not found.")
                return
            
            # Create People directory if it doesn't exist
            people_dir = os.path.join(self.pictures_path, "People")
            os.makedirs(people_dir, exist_ok=True)
            
            # Read the Notes sheet
            df = pd.read_excel(excel_path, sheet_name="Notes")
            
            # Filter for rows where images haven't been generated
            pending_images = df[df['image_generated'] == False]
            
            if pending_images.empty:
                messagebox.showinfo("Info", "No pending images to generate.")
                return
            
            total_images = len(pending_images)
            generated_count = 0
            
            # Update status
            self.status_label.config(text=f"Generating images: 0/{total_images}")
            self.root.update_idletasks()
            
            # Process each pending image
            for index, row in pending_images.iterrows():
                try:
                    # Construct the prompt
                    prompt = (
                        f"Professional wrestling promotional photo. {row['physical_description']} "
                        "The image should be a high-quality, professional headshot style photo "
                        "with good lighting and a neutral background. The subject should be "
                        "looking directly at the camera with a confident expression."
                    )
                    
                    # Generate the image
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    # Get the image URL
                    image_url = response.data[0].url
                    
                    # Download and save the image
                    import requests
                    image_name = row['Picture']
                    image_path = os.path.join(people_dir, image_name)
                    
                    response = requests.get(image_url)
                    response.raise_for_status()
                    
                    with open(image_path, 'wb') as f:
                        f.write(response.content)
                    
                    # Update the Excel file to mark this image as generated
                    df.at[index, 'image_generated'] = True
                    generated_count += 1
                    
                    # Update status
                    self.status_label.config(text=f"Generating images: {generated_count}/{total_images}")
                    self.root.update_idletasks()
                    
                    # Add a small delay to avoid rate limiting
                    time.sleep(1)
                    
                except Exception as e:
                    logging.error(f"Error generating image for {row['Name']}: {str(e)}")
                    messagebox.showerror("Error", f"Failed to generate image for {row['Name']}: {str(e)}")
                    continue
            
            # Save the updated Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Notes", index=False)
            
            messagebox.showinfo("Success", f"Generated {generated_count} images successfully!")
            
        except Exception as e:
            logging.error(f"Error in generate_wrestler_images: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.status_label.config("")
            self.root.update_idletasks()

    def generate_company_images(self):
        if not self.api_key:
            messagebox.showerror("Error", "Please set your API key in settings first.")
            return
            
        if not self.pictures_path:
            messagebox.showerror("Error", "Please set your pictures path in settings first.")
            return
            
        try:
            # Read the Excel file
            excel_path = "wrestleverse_companies.xlsx"
            if not os.path.exists(excel_path):
                messagebox.showerror("Error", "Companies Excel file not found.")
                return
                
            # Create necessary directories if they don't exist
            logos_dir = os.path.join(self.pictures_path, "Logos")
            banners_dir = os.path.join(self.pictures_path, "Banners")
            backdrops_dir = os.path.join(self.pictures_path, "Logo_Backdrops")
            
            for directory in [logos_dir, banners_dir, backdrops_dir]:
                os.makedirs(directory, exist_ok=True)
            
            # Read the Notes sheet
            df = pd.read_excel(excel_path, sheet_name="Notes")
            
            # Filter for rows where images haven't been generated
            pending_images = df[df['image_generated'] == False]
            
            if pending_images.empty:
                messagebox.showinfo("Info", "No pending images to generate.")
                return
                
            total_companies = len(pending_images)
            generated_count = 0
            
            # Update status
            self.status_label.config(text=f"Generating images: 0/{total_companies}")
            self.root.update_idletasks()
            
            # Process each pending company
            for index, row in pending_images.iterrows():
                try:
                    company_name = row['Name']
                    description = row['Description']
                    
                    # Clean up file names by removing quotes and invalid characters
                    logo_filename = row['Logo'].replace('"', '').replace('\\', '')
                    banner_filename = row['Banner'].replace('"', '').replace('\\', '')
                    backdrop_filename = row['Backdrop'].replace('"', '').replace('\\', '')
                    
                    # Generate Logo
                    logo_prompt = (
                        f"Professional wrestling company logo for \"{company_name}\". "
                        f"Company description: {description}. "
                        "The logo should be professional, memorable, and include the company name. "
                        "Use a transparent or solid background. Make it bold and striking."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=logo_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    # Save Logo
                    image_url = response.data[0].url
                    import requests
                    response = requests.get(image_url)
                    response.raise_for_status()
                    with open(os.path.join(logos_dir, logo_filename), 'wb') as f:
                        f.write(response.content)
                    
                    time.sleep(1)  # Delay to avoid rate limiting
                    
                    # Generate Banner
                    banner_prompt = (
                        f"Professional wrestling company banner for \"{company_name}\". "
                        f"Company description: {description}. "
                        "Create a wide promotional banner (1:5 aspect ratio) with dynamic wrestling imagery. "
                        "Include the company name prominently. Make it exciting and eye-catching."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=banner_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    # Save Banner
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    with open(os.path.join(banners_dir, banner_filename), 'wb') as f:
                        f.write(response.content)
                    
                    time.sleep(1)  # Delay to avoid rate limiting
                    
                    # Generate Backdrop
                    backdrop_prompt = (
                        f"Professional wrestling backdrop for \"{company_name}\". "
                        f"Company description: {description}. "
                        "Create a dramatic arena backdrop with the company's branding. "
                        "Should be suitable for a wrestling event, with atmospheric lighting and professional design."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=backdrop_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    # Save Backdrop
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    with open(os.path.join(backdrops_dir, backdrop_filename), 'wb') as f:
                        f.write(response.content)
                    
                    # Update the Excel file to mark this company's images as generated
                    df.at[index, 'image_generated'] = True
                    generated_count += 1
                    
                    # Update status
                    self.status_label.config(text=f"Generating images: {generated_count}/{total_companies}")
                    self.root.update_idletasks()
                    
                    time.sleep(1)  # Delay to avoid rate limiting
                    
                except Exception as e:
                    logging.error(f"Error generating images for {company_name}: {str(e)}")
                    messagebox.showerror("Error", f"Failed to generate images for {company_name}: {str(e)}")
                    continue
            
            # Save the updated Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Notes", index=False)
            
            messagebox.showinfo("Success", f"Generated images for {generated_count} companies successfully!")
            
        except Exception as e:
            logging.error(f"Error in generate_company_images: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.status_label.config(text="")
            self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = WrestleverseApp(root)
    root.mainloop()