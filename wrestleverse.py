import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import json
import pandas as pd
import random
import os
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

                # Store description in notes
                notes_data.append({
                    "Name": name,
                    "Description": description,
                    "Size": size
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
                             ?, ?, ?, ?, ?, ?, ?,
                             ?, ?, ?, ?, ?, ?, ?, ?,
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

            # Save to Excel
            try:
                logging.debug("Attempting to save data to Excel file.")
                companies_df = pd.DataFrame(companies_data, columns=companies_columns)
                bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                notes_df = pd.DataFrame(notes_data)
                
                excel_path = "wrestleverse_companies.xlsx"
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    companies_df.to_excel(writer, sheet_name="Companies", index=False)
                    bio_df.to_excel(writer, sheet_name="Bios", index=False)
                    notes_df.to_excel(writer, sheet_name="Notes", index=False)
                
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
        description_label = ttk.Label(wrestler_frame, text="Description:")
        description_label.grid(row=1, column=0, padx=5, pady=5)
        description_entry = ttk.Entry(wrestler_frame, width=40)
        description_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5)
        skill_preset_label = ttk.Label(wrestler_frame, text="Skill Preset:")
        skill_preset_label.grid(row=2, column=0, padx=5, pady=5)
        skill_preset_var = tk.StringVar(value="Interpret")
        skill_preset_names = ["Interpret"] + [preset["name"] for preset in self.skill_presets]
        skill_preset_dropdown = ttk.Combobox(wrestler_frame, textvariable=skill_preset_var, values=skill_preset_names)
        skill_preset_dropdown.grid(row=2, column=1, padx=5, pady=5)
        remove_btn = ttk.Button(wrestler_frame, text="❌", command=lambda: self.remove_wrestler_form(wrestler_frame))
        remove_btn.grid(row=0, column=4, padx=5, pady=5)
        self.wrestlers.append({
            "frame": wrestler_frame,
            "name": name_entry,
            "gender": gender_var,
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
        self.status_label.config(text="Status: Generating wrestlers...")
        self.root.update_idletasks()
        try:
            workers_data = []
            bio_data = []
            skills_data = []
            notes_data = []
            uid = self.uid_start
            if self.access_db_path and os.path.exists(self.access_db_path):
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={self.access_db_path};'
                    'PWD=20YearsOfTEW;'
                )
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                cursor.execute("SELECT MAX(UID) FROM tblWorker")
                result = cursor.fetchone()
                last_uid = result[0] if result[0] else 0
                uid = last_uid + 1
            else:
                existing_uids = []
                if os.path.exists("wrestleverse_workers.xlsx"):
                    existing_df = pd.read_excel("wrestleverse_workers.xlsx", sheet_name="worker")
                    if not existing_df.empty:
                        existing_uids = existing_df["UID"].tolist()
                        last_uid = max(existing_uids)
                        uid = max(uid, last_uid + 1)
                        workers_data.extend(existing_df.values.tolist())
                        existing_bio_df = pd.read_excel("wrestleverse_workers.xlsx", sheet_name="bio")
                        bio_data.extend(existing_bio_df.values.tolist())
                        existing_skills_df = pd.read_excel("wrestleverse_workers.xlsx", sheet_name="skills")
                        skills_data.extend(existing_skills_df.to_dict('records'))
                        existing_notes_df = pd.read_excel("wrestleverse_workers.xlsx", sheet_name="notes")
                        notes_data.extend(existing_notes_df.values.tolist())
            workers_columns = [
                "UID",
                "User",
                "Regen",
                "Active",
                "Name",
                "Shortname",
                "Gender",
                "Pronouns",
                "Sexuality",
                "CompetesAgainst",
                "Outsiderel",
                "Birthday",
                "DebutDate",
                "DeathDate",
                "BodyType",
                "WorkerHeight",
                "WorkerWeight",
                "WorkerMinWeight",
                "WorkerMaxWeight",
                "Picture",
                "Nationality",
                "Race",
                "Based_In",
                "LeftBusiness",
                "Dead",
                "Retired",
                "NonWrestler",
                "Celebridad",
                "Style",
                "Freelance",
                "Loyalty",
                "TrueBorn",
                "USA",
                "Canada",
                "Mexico",
                "Japan",
                "UK",
                "Europe",
                "Oz",
                "India",
                "Speak_English",
                "Speak_Japanese",
                "Speak_Spanish",
                "Speak_French",
                "Speak_Germanic",
                "Speak_Med",
                "Speak_Slavic",
                "Speak_Hindi",
                "Moveset",
                "Position_Wrestler",
                "Position_Occasional",
                "Position_Referee",
                "Position_Announcer",
                "Position_Colour",
                "Position_Manager",
                "Position_Personality",
                "Position_Roadagent",
                "Mask",
                "Age_Matures",
                "Age_Declines",
                "Age_TalkDeclines",
                "Age_Retires",
                "OrganicBio",
                "PlasterCaster_Face",
                "PlasterCaster_FaceBasis",
                "PlasterCaster_Heel",
                "PlasterCaster_HeelBasis",
                "CareerGoal"
            ]
            total_wrestlers = len(self.wrestlers)
            for index, wrestler in enumerate(self.wrestlers, start=1):
                self.status_label.config(text=f"Status: Generating wrestler {index}/{total_wrestlers}...")
                self.root.update_idletasks()
                name = wrestler["name"].get().strip()
                gender = wrestler["gender"].get().strip()
                description = wrestler["description"].get().strip()
                if not name:
                    name = self.generate_name(description=description, gender=gender)
                name = name[:30]
                shortname = name.split()[0][:20] if name else ''
                gender_value = 1 if gender.lower() == 'male' else 5
                pronouns = 1 if gender_value == 1 else 2
                body_type = random.randint(0, 7)
                worker_height = random.randint(20, 42)
                worker_weight = random.randint(150, 350)
                worker_min_weight = max(0, worker_weight - random.randint(20, 50))
                worker_max_weight = worker_weight + random.randint(20, 50)
                race = random.randint(1, 9)
                style = random.randint(1, 17)
                body_type = self.ensure_byte(body_type)
                worker_height = self.ensure_byte(worker_height)
                race = self.ensure_byte(race)
                style = self.ensure_byte(style)
                pronouns = self.ensure_byte(pronouns)
                gender_value = self.ensure_byte(gender_value)
                sexuality = self.ensure_byte(1)
                competes_against = self.ensure_byte(2 if gender_value == 1 else 3)
                outsiderel = self.ensure_byte(0)
                based_in = self.ensure_byte(1)
                celebridad = self.ensure_byte(0)
                loyalty = int(0)
                picture = f"{name.replace(' ', '').lower()}.jpg"[:35]
                plastercaster_face = self.generate_gimmick(name, description, gender, "face")[:30]
                plastercaster_heel = self.generate_gimmick(name, description, gender, "heel")[:30]
                birthday = self.generate_birthday()
                earliest_debut_date = self.add_years(birthday, 16)
                latest_debut_date = datetime.datetime(2024, 1, 1)
                if earliest_debut_date >= latest_debut_date:
                    birthday = self.generate_birthday(max_year=2007 - 16)
                    earliest_debut_date = self.add_years(birthday, 16)
                debut_date = self.random_date(earliest_debut_date, latest_debut_date)
                worker_row = [
                    int(uid),
                    False,
                    self.ensure_byte(0),
                    True,
                    name,
                    shortname,
                    gender_value,
                    pronouns,
                    sexuality,
                    competes_against,
                    outsiderel,
                    birthday,
                    debut_date,
                    "1666-01-01",
                    body_type,
                    self.ensure_byte(worker_height),
                    int(worker_weight),
                    int(worker_min_weight),
                    int(worker_max_weight),
                    picture,
                    int(1),
                    race,
                    based_in,
                    False,
                    False,
                    False,
                    False,
                    celebridad,
                    style,
                    False,
                    loyalty,
                    False,
                    True,
                    True,
                    True,
                    True,
                    True,
                    True,
                    True,
                    True,
                    int(4),
                    int(4),
                    int(4),
                    int(4),
                    int(4),
                    int(4),
                    int(4),
                    int(1),
                    int(0),
                    True,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    False,
                    int(0),
                    self.ensure_byte(0),
                    self.ensure_byte(0),
                    self.ensure_byte(0),
                    self.ensure_byte(0),
                    True,
                    plastercaster_face,
                    self.ensure_byte(1),
                    plastercaster_heel,
                    self.ensure_byte(1),
                    self.ensure_byte(0)
                ]
                skill_preset_name = wrestler["skill_preset"].get()
                if skill_preset_name == "Interpret":
                    skill_preset_name = self.select_skill_preset_with_chatgpt(name, description, gender)
                    if skill_preset_name not in [preset["name"] for preset in self.skill_presets]:
                        skill_preset = next((preset for preset in self.skill_presets if preset["name"] == "Default"), None)
                    else:
                        skill_preset = next((preset for preset in self.skill_presets if preset["name"] == skill_preset_name), None)
                else:
                    skill_preset = next((preset for preset in self.skill_presets if preset["name"] == skill_preset_name), None)
                    if not skill_preset:
                        skill_preset = next((preset for preset in self.skill_presets if preset["name"] == "Default"), None)
                skills = self.generate_skills(uid, skill_preset)
                skills_data.append(skills)
                bio = self.generate_bio(name, gender, description, skill_preset_name)
                bio_data.append([int(uid), bio])
                notes_data.append([uid, name, skill_preset_name, description])
                workers_data.append(worker_row)
                uid += 1
            if self.access_db_path and os.path.exists(self.access_db_path):
                self.status_label.config(text="Status: Inserting data into Access database...")
                self.root.update_idletasks()
                for worker_row in workers_data:
                    worker_row_converted = [
                        -1 if value is True else 0 if value is False else value
                        for value in worker_row
                    ]
                    placeholders = ", ".join(["?"] * len(worker_row_converted))
                    sql_insert_worker = f"INSERT INTO tblWorker ({', '.join(workers_columns)}) VALUES ({placeholders})"
                    cursor.execute(sql_insert_worker, worker_row_converted)
                for bio_row in bio_data:
                    sql_insert_bio = "INSERT INTO tblWorkerBio (UID, Profile) VALUES (?, ?)"
                    cursor.execute(sql_insert_bio, bio_row)
                for skills_row in skills_data:
                    columns = list(skills_row.keys())
                    values = [skills_row[col] for col in columns]
                    placeholders = ", ".join(["?"] * len(values))
                    sql_insert_skills = f"INSERT INTO tblWorkerSkill ({', '.join(columns)}) VALUES ({placeholders})"
                    cursor.execute(sql_insert_skills, values)
                conn.commit()
                conn.close()
                self.status_label.config(text="Status: Wrestlers generated and inserted into Access database successfully!")
            else:
                self.status_label.config(text="Status: Saving data to Excel file...")
                self.root.update_idletasks()
                workers_df = pd.DataFrame(workers_data, columns=workers_columns)
                bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                skills_df = pd.DataFrame(skills_data)
                notes_df = pd.DataFrame(notes_data, columns=["UID", "Name", "Skill Preset", "Description"])
                workers_df = workers_df.astype({
                    "UID": "int64",
                    "User": "bool",
                    "Regen": "uint8",
                    "Active": "bool",
                    "Name": "string",
                    "Shortname": "string",
                    "Gender": "uint8",
                    "Pronouns": "uint8",
                    "Sexuality": "uint8",
                    "CompetesAgainst": "uint8",
                    "Outsiderel": "uint8",
                    "Birthday": "datetime64[ns]",
                    "DebutDate": "datetime64[ns]",
                    "DeathDate": "string",
                    "BodyType": "uint8",
                    "WorkerHeight": "uint8",
                    "WorkerWeight": "int32",
                    "WorkerMinWeight": "int32",
                    "WorkerMaxWeight": "int32",
                    "Picture": "string",
                    "Nationality": "int32",
                    "Race": "uint8",
                    "Based_In": "uint8",
                    "LeftBusiness": "bool",
                    "Dead": "bool",
                    "Retired": "bool",
                    "NonWrestler": "bool",
                    "Celebridad": "uint8",
                    "Style": "uint8",
                    "Freelance": "bool",
                    "Loyalty": "int64",
                    "TrueBorn": "bool",
                    "USA": "bool",
                    "Canada": "bool",
                    "Mexico": "bool",
                    "Japan": "bool",
                    "UK": "bool",
                    "Europe": "bool",
                    "Oz": "bool",
                    "India": "bool",
                    "Speak_English": "int32",
                    "Speak_Japanese": "int32",
                    "Speak_Spanish": "int32",
                    "Speak_French": "int32",
                    "Speak_Germanic": "int32",
                    "Speak_Med": "int32",
                    "Speak_Slavic": "int32",
                    "Speak_Hindi": "int32",
                    "Moveset": "int64",
                    "Position_Wrestler": "bool",
                    "Position_Occasional": "bool",
                    "Position_Referee": "bool",
                    "Position_Announcer": "bool",
                    "Position_Colour": "bool",
                    "Position_Manager": "bool",
                    "Position_Personality": "bool",
                    "Position_Roadagent": "bool",
                    "Mask": "int32",
                    "Age_Matures": "uint8",
                    "Age_Declines": "uint8",
                    "Age_TalkDeclines": "uint8",
                    "Age_Retires": "uint8",
                    "OrganicBio": "bool",
                    "PlasterCaster_Face": "string",
                    "PlasterCaster_FaceBasis": "uint8",
                    "PlasterCaster_Heel": "string",
                    "PlasterCaster_HeelBasis": "uint8",
                    "CareerGoal": "uint8"
                })
                bio_df = bio_df.astype({"UID": "int64", "Bio": "string"})
                skills_df = skills_df.astype({skill: "int32" for skill in skills_df.columns if skill != "WorkerUID"})
                skills_df = skills_df.astype({"WorkerUID": "int64"})
                notes_df = notes_df.astype({"UID": "int64", "Name": "string", "Skill Preset": "string", "Description": "string"})
                with pd.ExcelWriter("wrestleverse_workers.xlsx") as writer:
                    workers_df.to_excel(writer, sheet_name="worker", index=False)
                    bio_df.to_excel(writer, sheet_name="bio", index=False)
                    skills_df.to_excel(writer, sheet_name="skills", index=False)
                    notes_df.to_excel(writer, sheet_name="notes", index=False)
                self.status_label.config(text="Status: Wrestlers generated and saved to Excel successfully!")
        except Exception as e:
            error_message = f"Status: Error - {str(e)}"
            self.status_label.config(text=error_message)
            logging.error(error_message, exc_info=True)
        finally:
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
        api_key_label = ttk.Label(self.root, text="ChatGPT API Key:")
        api_key_label.pack(pady=5)
        self.api_key_var = tk.StringVar(value=self.api_key)
        api_key_entry = ttk.Entry(self.root, textvariable=self.api_key_var, width=50)
        api_key_entry.pack(pady=5)
        uid_start_label = ttk.Label(self.root, text="UID Start:")
        uid_start_label.pack(pady=5)
        self.uid_start_var = tk.IntVar(value=self.uid_start)
        uid_start_entry = ttk.Entry(self.root, textvariable=self.uid_start_var, width=10)
        uid_start_entry.pack(pady=5)
        bio_prompt_label = ttk.Label(self.root, text="Bio Prompt:")
        bio_prompt_label.pack(pady=5)
        self.bio_prompt_var = tk.StringVar(value=self.bio_prompt)
        bio_prompt_entry = ttk.Entry(self.root, textvariable=self.bio_prompt_var, width=50)
        bio_prompt_entry.pack(pady=5)
        access_db_label = ttk.Label(self.root, text="Access Database Path (Optional):")
        access_db_label.pack(pady=5)
        self.access_db_var = tk.StringVar(value=self.access_db_path)
        access_db_entry = ttk.Entry(self.root, textvariable=self.access_db_var, width=50)
        access_db_entry.pack(pady=5)
        browse_btn = ttk.Button(self.root, text="Browse", command=self.browse_access_db)
        browse_btn.pack(pady=5)
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
        settings = {
            "api_key": self.api_key,
            "uid_start": self.uid_start,
            "bio_prompt": self.bio_prompt,
            "access_db_path": self.access_db_path
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
        except FileNotFoundError:
            self.api_key = ""
            self.uid_start = 1
            self.bio_prompt = "Create a biography for a professional wrestler."
            self.access_db_path = ""

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

if __name__ == "__main__":
    root = tk.Tk()
    app = WrestleverseApp(root)
    root.mainloop()
