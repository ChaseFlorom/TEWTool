import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import json
import pandas as pd
import random
import os
import time
import requests
import pyodbc
from openai import OpenAI
import datetime
import logging
from PIL import Image
import io

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
        self.start_date_str = ""
        self.start_date = datetime.datetime(2020,1,1)
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

                if not name and not description:
                    prompt = f"Generate a name and description for a {size.lower()} professional wrestling company."
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":prompt}]
                    )
                    text = resp.choices[0].message.content.strip()
                    lines = text.split('\n')
                    name = lines[0].strip() if lines else "Default Company"
                    description = ' '.join(lines[1:]).strip() if len(lines)>1 else "A professional wrestling company."
                elif not name:
                    name_prompt = f"Generate a name for a {size.lower()} professional wrestling company with the following description: {description}"
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":name_prompt}]
                    )
                    name = resp.choices[0].message.content.strip()
                elif not description:
                    desc_prompt = f"Generate a description for a {size.lower()} professional wrestling company named {name}."
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":desc_prompt}]
                    )
                    description = resp.choices[0].message.content.strip()

                base_name = name.replace(' ', '').replace('.', '').lower()
                logo_name = f"{base_name[:26]}.jpg"
                backdrop_name = f"{base_name[:24]}BD.jpg"
                banner_name = f"{base_name[:24]}BN.jpg"

                company_row = [
                    uid,
                    name,
                    self.generate_company_initials(name),
                    f"www.{name.replace(' ', '').lower()}.com"[:40],
                    "1666-01-01",
                    "1666-01-01",
                    -1,
                    0,
                    logo_name,
                    backdrop_name,
                    banner_name,
                    1,
                    random.randint(1, 100),
                    0,
                    {"Tiny": 100000, "Small": 1000000, "Medium": 10000000, "Large": 100000000}.get(size, 1000000),
                    0,
                    10,
                    random.randint(1, 100),
                    0,
                    0,
                    0,
                    0,
                    "1666-01-01",
                    "1666-01-01",
                    0,
                    0,
                    0,
                    0,
                    0,
                    -1,
                    -1,
                    -1,
                    -1,
                    0,
                    0,
                    0,
                    "",
                    "",
                    "",
                    0,
                    -1,
                ]
                companies_data.append(company_row)

                bio = self.generate_company_bio(name, description, size)
                bio_data.append([uid, bio])

                notes_data.append({
                    "Name": name,
                    "Description": description,
                    "Size": size,
                    "Logo": f"{name.replace(' ', '').lower()}.jpg"[:35],
                    "Backdrop": f"{name.replace(' ', '').lower()}BD.jpg"[:35],
                    "Banner": f"{name.replace(' ', '').lower()}Banner.jpg"[:30],
                    "image_generated": False
                })

                uid += 1

            if self.access_db_path and os.path.exists(self.access_db_path):
                logging.debug("Attempting to save to Access database.")
                try:
                    conn_str = (
                        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                        f'DBQ={self.access_db_path};'
                        'PWD=20YearsOfTEW;'
                    )
                    for company_row in companies_data:
                        conn = pyodbc.connect(conn_str)
                        cursor = conn.cursor()
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
                        cursor.execute(sql_insert_company, company_row)
                        sql_insert_schedule = "INSERT INTO tblFedSchedule ([FedUID], [Strategy]) VALUES (?, ?)"
                        schedule_values = (company_row[0], '5')
                        cursor.execute(sql_insert_schedule, schedule_values)
                        conn.commit()
                        conn.close()

                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()
                    for bio_row in bio_data:
                        sql_insert_bio = "INSERT INTO tblFedBio ([UID], [Profile]) VALUES (?, ?)"
                        cursor.execute(sql_insert_bio, bio_row)
                    conn.commit()
                    conn.close()
                    logging.debug("Successfully saved to Access database.")
                except Exception as e:
                    logging.error(f"Error saving to Access database: {e}", exc_info=True)
                    messagebox.showerror("Error", f"Could not save to Access database: {str(e)}")

            for note in notes_data:
                try:
                    logo_prompt = (
                        f"For a professional wrestling company named '{note['Name']}' "
                        f"with the following description: '{note['Description']}', "
                        f"describe in a single sentence what their logo might look like."
                    )
                    logo_description = self.get_response_from_gpt(logo_prompt)
                    note['logo_description'] = logo_description
                except Exception as e:
                    logging.error(f"Error generating logo description: {e}")
                    note['logo_description'] = ""

            try:
                companies_df = pd.DataFrame(companies_data, columns=companies_columns)
                bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                notes_df = pd.DataFrame(notes_data)
                
                excel_path = "wrestleverse_companies.xlsx"
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    companies_df.to_excel(writer, sheet_name="Companies", index=False)
                    bio_df.to_excel(writer, sheet_name="Bios", index=False)
                    notes_df.to_excel(writer, sheet_name="Notes", index=False)
                
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
        
        companies = self.get_companies()
        company_names = ["Random"] + [company[1] for company in companies]
        
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
            notes_data = []

            try:
                if self.start_date_str:
                    self.start_date = datetime.datetime.strptime(self.start_date_str, "%Y-%m-%d")
                else:
                    self.start_date = datetime.datetime(2020,1,1)
            except:
                self.start_date = datetime.datetime(2020,1,1)

            uid = self.uid_start
            contract_uid = self.uid_start
            
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
                uid = max(last_uid + 1, self.uid_start)
                
                cursor.execute("SELECT MAX(UID) FROM tblContract")
                result = cursor.fetchone()
                last_contract_uid = result[0] if result[0] else 0
                contract_uid = max(last_contract_uid + 1, self.uid_start)

            wrestler_data_list = []
            for wrestler in self.wrestlers:
                try:
                    if wrestler["frame"].winfo_exists():
                        data = {
                            'name': wrestler["name"].get().strip() if wrestler["name"].winfo_exists() else "",
                            'gender': wrestler["gender"].get().strip() if hasattr(wrestler["gender"], "get") else "Male",
                            'company': wrestler["company"].get().strip() if hasattr(wrestler["company"], "get") else "Random",
                            'exclusive': wrestler["exclusive"].get().strip() if hasattr(wrestler["exclusive"], "get") else "Random",
                            'description': wrestler["description"].get().strip() if wrestler["description"].winfo_exists() else "",
                            'skill_preset': wrestler["skill_preset"].get().strip() if hasattr(wrestler["skill_preset"], "get") else "Default"
                        }
                        wrestler_data_list.append(data)
                except (tk.TclError, AttributeError):
                    continue

            total_wrestlers = len(wrestler_data_list)
            for index, wrestler_data in enumerate(wrestler_data_list, start=1):
                self.status_label.config(text=f"Status: Generating wrestler {index}/{total_wrestlers}...")
                self.root.update_idletasks()

                name = wrestler_data['name']
                gender = wrestler_data['gender']
                description = wrestler_data['description']

                # Use player-provided description if available
                player_description = description if description else ""

                if not name and not description:
                    prompt = f"Generate a name and description for a professional wrestler. The wrestler's gender is {gender}."
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":prompt}]
                    )
                    text = resp.choices[0].message.content.strip()
                    lines = text.split('\n')
                    name = lines[0].strip() if lines else "Default Wrestler"
                    description = ' '.join(lines[1:]).strip() if len(lines)>1 else "A professional wrestler."
                elif not name:
                    prompt = f"Generate a name for a professional wrestler. The wrestler's gender is {gender}. Description: {description}"
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":prompt}]
                    )
                    name = resp.choices[0].message.content.strip()
                elif not description:
                    prompt = f"Generate a description for a professional wrestler named {name}. The wrestler's gender is {gender}."
                    resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":prompt}]
                    )
                    description = resp.choices[0].message.content.strip()

                name = name.replace('.', '').strip()
                name = name[:30]

                shortname = name.split()[0][:20] if name else ''
                gender_value = 1 if gender.lower() == 'male' else 5

                age_prompt = (
                    f"Given this wrestler's description: {description}\n"
                    f"and gender {gender}, estimate their age at the start date. If the description suggests 'old', choose an older age (40-50). If 'young' choose younger (16-20). Otherwise pick an age between 16 and 50.\n"
                    "Respond with just a number."
                )
                age = 30
                try:
                    age_resp = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role":"user","content":age_prompt}]
                    )
                    age_str = age_resp.choices[0].message.content.strip()
                    age_val = int(age_str)
                    if 16 <= age_val <= 50:
                        age = age_val
                except:
                    pass

                birth_year = self.start_date.year - age
                birth_month = random.randint(1,12)
                birth_day = random.randint(1,28)
                birth_date = datetime.datetime(birth_year, birth_month, birth_day)

                debut_age = random.randint(16, age)
                debut_year = birth_year + debut_age
                if debut_year >= self.start_date.year:
                    debut_year = self.start_date.year - 1
                debut_month = random.randint(1,12)
                debut_day = random.randint(1,28)
                debut_date = datetime.datetime(debut_year, debut_month, debut_day)
                if debut_date >= self.start_date:
                    debut_date = self.start_date - datetime.timedelta(days=30)

                style = wrestler_data['skill_preset'] if wrestler_data['skill_preset'] != "Interpret" else "Interpret"
                bio_prompt = (
                    f"{self.bio_prompt} The wrestler's name is {name}. "
                    f"Their gender is {gender}. Description: {description}. "
                    f"Their wrestling style is best described as {style}."
                )
                bio = self.get_response_from_gpt(bio_prompt)
                if not bio:
                    bio = "A professional wrestler."

                style_num = self.get_style_from_gpt(bio) 
                race = self.get_race_from_gpt(name, f"{description}\n\nBiography: {bio}")
                picture_name = f"{name.replace(' ', '').lower()[:26]}.jpg"

                roles_lang_body_prompt = (
                    f"Given the wrestler's name: {name}, description: {player_description if player_description else description}, and bio: {bio}, "
                    "provide a JSON response with the following:\n"
                    "- Boolean values for: Wrestler, OccasionalWrestler, Manager, OnScreenPersonality, PlayByPlayCommentator, ColourCommentator, Referee, RoadAgent\n"
                    "- Language fluencies (1-4) for English (always 4), Japanese, Spanish, French, Germanic, Mediterranean, Slavic, Hindi\n"
                    "- A body type number (1-7) based on their description.\n\n"
                    "Return JSON only."
                )

                # Try multiple attempts if parsing fails, for more bulletproof approach
                # We'll loop a few times to try and get a valid JSON
                attempts = 0
                roles_lang_body_data = None
                while attempts < 3 and roles_lang_body_data is None:
                    attempts += 1
                    try:
                        roles_lang_body_response = self.client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role":"user","content":roles_lang_body_prompt}]
                        )
                        roles_lang_body_content = roles_lang_body_response.choices[0].message.content.strip()
                        roles_lang_body_data = json.loads(roles_lang_body_content)
                    except Exception as e:
                        logging.error(f"Error getting roles/lang/body from GPT attempt {attempts}: {e}")
                        roles_lang_body_data = None

                # Default fallback if still None
                if roles_lang_body_data is None:
                    roles_lang_body_data = {
                        "Wrestler": True,
                        "OccasionalWrestler": False,
                        "Manager": False,
                        "OnScreenPersonality": False,
                        "PlayByPlayCommentator": False,
                        "ColourCommentator": False,
                        "Referee": False,
                        "RoadAgent": False,
                        "Languages": {
                            "English": 4,
                            "Japanese": 1,
                            "Spanish": 2,
                            "French": 2,
                            "Germanic": 2,
                            "Mediterranean": 2,
                            "Slavic": 1,
                            "Hindi": 1
                        },
                        "BodyType": 1
                    }

                def bool_to_access(val):
                    return -1 if val else 0

                position_wrestler = bool_to_access(roles_lang_body_data.get("Wrestler", False))
                position_occasional = bool_to_access(roles_lang_body_data.get("OccasionalWrestler", False))
                position_manager = bool_to_access(roles_lang_body_data.get("Manager", False))
                position_personality = bool_to_access(roles_lang_body_data.get("OnScreenPersonality", False))
                position_announcer = bool_to_access(roles_lang_body_data.get("PlayByPlayCommentator", False))
                position_colour = bool_to_access(roles_lang_body_data.get("ColourCommentator", False))
                position_referee = bool_to_access(roles_lang_body_data.get("Referee", False))
                position_roadagent = bool_to_access(roles_lang_body_data.get("RoadAgent", False))

                languages = roles_lang_body_data.get("Languages", {})
                speak_english = languages.get("English", 4)
                speak_japanese = languages.get("Japanese", 1)
                speak_spanish = languages.get("Spanish", 2)
                speak_french = languages.get("French", 2)
                speak_germanic = languages.get("Germanic", 2)
                speak_med = languages.get("Mediterranean", 2)
                speak_slavic = languages.get("Slavic", 1)
                speak_hindi = languages.get("Hindi", 1)

                body_type_code = roles_lang_body_data.get("BodyType", 1)
                body_type_map = {
                    1: "Average",
                    2: "Skinny",
                    3: "Toned",
                    4: "Muscular",
                    5: "Ripped",
                    6: "Flabby",
                    7: "Obese"
                }
                body_type_text = body_type_map.get(body_type_code, "Average")

                worker_row = {
                    "UID": int(uid),
                    "User": False,
                    "Regen": self.ensure_byte(0),
                    "Active": True,
                    "Name": name[:30],
                    "Shortname": shortname[:20],
                    "Gender": self.ensure_byte(gender_value),
                    "Pronouns": self.ensure_byte(1 if gender_value == 1 else 2),
                    "Sexuality": self.ensure_byte(1),
                    "CompetesAgainst": self.ensure_byte(2 if gender_value == 1 else 3),
                    "Outsiderel": self.ensure_byte(0),
                    "Birthday": birth_date,
                    "DebutDate": debut_date,
                    "DeathDate": "1666-01-01",
                    "BodyType": self.ensure_byte(body_type_code),
                    "WorkerHeight": self.ensure_byte(random.randint(20, 42)),
                    "WorkerWeight": random.randint(150,350),
                    "WorkerMinWeight": 150,
                    "WorkerMaxWeight": 350,
                    "Picture": picture_name,
                    "Nationality": int(1),
                    "Race": self.ensure_byte(race),
                    "Based_In": self.ensure_byte(1),
                    "LeftBusiness": False,
                    "Dead": False,
                    "Retired": False,
                    "NonWrestler": False,
                    "Celebridad": self.ensure_byte(0),
                    "Style": style_num,
                    "Freelance": False,
                    "Loyalty": 0,
                    "TrueBorn": False,
                    "USA": True,
                    "Canada": True,
                    "Mexico": True,
                    "Japan": True,
                    "UK": True,
                    "Europe": True,
                    "Oz": True,
                    "India": True,
                    "Speak_English": int(speak_english),
                    "Speak_Japanese": int(speak_japanese),
                    "Speak_Spanish": int(speak_spanish),
                    "Speak_French": int(speak_french),
                    "Speak_Germanic": int(speak_germanic),
                    "Speak_Med": int(speak_med),
                    "Speak_Slavic": int(speak_slavic),
                    "Speak_Hindi": int(speak_hindi),
                    "Moveset": int(0),
                    "Position_Wrestler": position_wrestler,
                    "Position_Occasional": position_occasional,
                    "Position_Referee": position_referee,
                    "Position_Announcer": position_announcer,
                    "Position_Colour": position_colour,
                    "Position_Manager": position_manager,
                    "Position_Personality": position_personality,
                    "Position_Roadagent": position_roadagent,
                    "Mask": int(0),
                    "Age_Matures": self.ensure_byte(0),
                    "Age_Declines": self.ensure_byte(0),
                    "Age_TalkDeclines": self.ensure_byte(0),
                    "Age_Retires": self.ensure_byte(0),
                    "OrganicBio": True,
                    "PlasterCaster_Face": self.generate_gimmick(name, description, gender, "face")[:30],
                    "PlasterCaster_FaceBasis": self.ensure_byte(1),
                    "PlasterCaster_Heel": self.generate_gimmick(name, description, gender, "heel")[:30],
                    "PlasterCaster_HeelBasis": self.ensure_byte(1),
                    "CareerGoal": self.ensure_byte(0)
                }

                worker_row_converted = {
                    key: (-1 if (isinstance(value, bool) and value) else (0 if isinstance(value, bool) else value))
                    for key, value in worker_row.items()
                }
                workers_data.append(worker_row_converted)
                bio_data.append([uid, bio])

                if wrestler_data['skill_preset'] == "Interpret":
                    preset_name = self.select_skill_preset_with_chatgpt(
                        name,
                        description,
                        gender
                    )
                else:
                    preset_name = wrestler_data['skill_preset']

                preset = next((p for p in self.skill_presets if p["name"] == preset_name), self.skill_presets[0])
                skills = self.generate_skills(uid, preset)
                skills_data.append(skills)

                if wrestler_data['company'] != "Freelancer":
                    contract_began_year = self.start_date.year - random.randint(1,5)
                    contract_began_month = random.randint(1,12)
                    contract_began_day = random.randint(1,28)
                    contract_began = datetime.datetime(contract_began_year, contract_began_month, contract_began_day)
                    if contract_began >= self.start_date:
                        contract_began = self.start_date - datetime.timedelta(days=60)

                    contract = self.generate_contract(wrestler_data, uid, wrestler_data['company'], contract_uid)
                    if contract:
                        contract["ContractBeganDate"] = contract_began
                        contract["ContractDebutDate"] = datetime.datetime(1900,1,1)
                        contract_data.append(contract)
                        contract_uid += 1

                    # Use body_type_text and race in physical prompt, and also the player-provided description if exists
                    physical_prompt = (
                        f"Based on this wrestler's details:\n"
                        f"Name: {name}\n"
                        f"Description: {player_description if player_description else description}\n"
                        f"Gender: {gender}\n"
                        f"Race: {self.get_race_name(race)}\n"
                        f"Body Type: {body_type_text}\n"
                        f"Please provide a single sentence describing their physical appearance. "
                        f"Focus on height, build, and distinctive features."
                    )
                    physical_description = self.get_response_from_gpt(physical_prompt)
                else:
                    physical_description = ""

                notes_data.append({
                    "Name": name,
                    "Description": player_description if player_description else description,
                    "Gender": gender,
                    "Company": wrestler_data.get('company', 'Random'),
                    "Exclusive": wrestler_data.get('exclusive', 'Random'),
                    "Skill_Preset": wrestler_data.get('skill_preset', 'Default'),
                    "Picture": f"{name.replace(' ', '').lower()}.jpg"[:35],
                    "physical_description": physical_description,
                    "image_generated": False,
                    "Race": race
                })

                uid += 1

            if self.access_db_path and os.path.exists(self.access_db_path):
                try:
                    conn_str = (
                        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                        f'DBQ={self.access_db_path};'
                        'PWD=20YearsOfTEW;'
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    for worker_row in workers_data:
                        worker_row_converted = {
                            key: (-1 if (isinstance(value, bool) and value) else (0 if isinstance(value, bool) else value))
                            for key, value in worker_row.items()
                        }

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

                        birth_date = worker_row["Birthday"]
                        debut_date = worker_row["DebutDate"]
                        death_date = "1666-01-01"
                        worker_values = [
                            int(worker_row_converted["UID"]),
                            bool(worker_row_converted["User"]),
                            int(worker_row_converted["Regen"]),
                            bool(worker_row_converted["Active"]),
                            str(worker_row_converted["Name"]),
                            str(worker_row_converted["Shortname"]),
                            int(worker_row_converted["Gender"]),
                            int(worker_row_converted["Pronouns"]),
                            int(worker_row_converted["Sexuality"]),
                            int(worker_row_converted["CompetesAgainst"]),
                            int(worker_row_converted["Outsiderel"]),
                            birth_date,
                            debut_date,
                            death_date,
                            int(worker_row_converted["BodyType"]),
                            int(worker_row_converted["WorkerHeight"]),
                            int(worker_row_converted["WorkerWeight"]),
                            int(worker_row_converted["WorkerMinWeight"]),
                            int(worker_row_converted["WorkerMaxWeight"]),
                            str(worker_row_converted["Picture"]),
                            int(worker_row_converted["Nationality"]),
                            int(worker_row_converted["Race"]),
                            int(worker_row_converted["Based_In"]),
                            bool(worker_row_converted["LeftBusiness"]),
                            bool(worker_row_converted["Dead"]),
                            bool(worker_row_converted["Retired"]),
                            bool(worker_row_converted["NonWrestler"]),
                            int(worker_row_converted["Celebridad"]),
                            int(worker_row_converted["Style"]),
                            bool(worker_row_converted["Freelance"]),
                            int(worker_row_converted["Loyalty"]),
                            bool(worker_row_converted["TrueBorn"]),
                            bool(worker_row_converted["USA"]),
                            bool(worker_row_converted["Canada"]),
                            bool(worker_row_converted["Mexico"]),
                            bool(worker_row_converted["Japan"]),
                            bool(worker_row_converted["UK"]),
                            bool(worker_row_converted["Europe"]),
                            bool(worker_row_converted["Oz"]),
                            bool(worker_row_converted["India"]),
                            int(worker_row_converted["Speak_English"]),
                            int(worker_row_converted["Speak_Japanese"]),
                            int(worker_row_converted["Speak_Spanish"]),
                            int(worker_row_converted["Speak_French"]),
                            int(worker_row_converted["Speak_Germanic"]),
                            int(worker_row_converted["Speak_Med"]),
                            int(worker_row_converted["Speak_Slavic"]),
                            int(worker_row_converted["Speak_Hindi"]),
                            int(worker_row_converted["Moveset"]),
                            bool(worker_row_converted["Position_Wrestler"]),
                            bool(worker_row_converted["Position_Occasional"]),
                            bool(worker_row_converted["Position_Referee"]),
                            bool(worker_row_converted["Position_Announcer"]),
                            bool(worker_row_converted["Position_Colour"]),
                            bool(worker_row_converted["Position_Manager"]),
                            bool(worker_row_converted["Position_Personality"]),
                            bool(worker_row_converted["Position_Roadagent"]),
                            int(worker_row_converted["Mask"]),
                            int(worker_row_converted["Age_Matures"]),
                            int(worker_row_converted["Age_Declines"]),
                            int(worker_row_converted["Age_TalkDeclines"]),
                            int(worker_row_converted["Age_Retires"]),
                            bool(worker_row_converted["OrganicBio"]),
                            str(worker_row_converted["PlasterCaster_Face"]),
                            int(worker_row_converted["PlasterCaster_FaceBasis"]),
                            str(worker_row_converted["PlasterCaster_Heel"]),
                            int(worker_row_converted["PlasterCaster_HeelBasis"]),
                            int(worker_row_converted["CareerGoal"])
                        ]
                        cursor.execute(sql_insert_worker, worker_values)

                    for bio_row in bio_data:
                        sql_insert_bio = "INSERT INTO tblWorkerBio ([UID], [Profile]) VALUES (?, ?)"
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
                        contract_values = [
                            contract["UID"],
                            contract["FedUID"],
                            contract["WorkerUID"],
                            contract["Name"],
                            contract["Shortname"],
                            contract["Picture"],
                            contract["CompetesIn"],
                            contract["Face"],
                            contract["Division"],
                            contract["Manager"],
                            contract["Moveset"],
                            contract["WrittenContract"],
                            contract["ExclusiveContract"],
                            contract["TouringContract"],
                            contract["PaidMonthly"],
                            contract["OnLoan"],
                            contract["Developmental"],
                            contract["PrimaryUsage"],
                            contract["SecondaryUsage"],
                            contract["ExpectedShows"],
                            contract["BonusAmount"],
                            contract["BonusType"],
                            contract["Creative"],
                            contract["HiringVeto"],
                            contract["WageMatch"],
                            contract["IronClad"],
                            contract["ContractBeganDate"],
                            contract["Daysleft"],
                            contract["Dateslength"],
                            contract["DatesLeft"],
                            contract["ContractDebutDate"],
                            contract["Amount"],
                            contract["Downside"],
                            contract["Brand"],
                            contract["Mask"],
                            contract["ContractMomentum"],
                            contract["Last_Turn"],
                            contract["Travel"],
                            contract["Position_Wrestler"],
                            contract["Position_Occasional"],
                            contract["Position_Referee"],
                            contract["Position_Announcer"],
                            contract["Position_Colour"],
                            contract["Position_Manager"],
                            contract["Position_Personality"],
                            contract["Position_Roadagent"],
                            contract["Merch"],
                            contract["PlasterCaster_Gimmick"],
                            contract["PlasterCaster_Rating"],
                            contract["PlasterCaster_Lifespan"],
                            contract["PlasterCaster_Byte1"],
                            contract["PlasterCaster_Byte2"],
                            contract["PlasterCaster_Byte3"],
                            contract["PlasterCaster_Byte4"],
                            contract["PlasterCaster_Byte5"],
                            contract["PlasterCaster_Byte6"]
                        ]
                        for i in range(1, 26):
                            contract_values.append(contract[f"PlasterCaster_Bool{i}"])
                        cursor.execute(sql_insert_contract, contract_values)

                    # Enhanced popularity section: multiple attempts to get valid JSON
                    for worker_row in workers_data:
                        worker_uid = worker_row["UID"]
                        bio_for_pop = ""
                        name_for_pop = ""
                        for b in bio_data:
                            if b[0] == worker_uid:
                                bio_for_pop = b[1]
                                break
                        for n in notes_data:
                            if n["Name"][:30] == worker_row["Name"]:
                                name_for_pop = n["Name"]
                                # We'll also pass player description if it was provided
                                # but the function currently only uses name and bio.
                                # Let's try to make popularity prompt robust by including description:
                                player_desc_for_pop = n["Description"] if n["Description"] else ""
                                break

                        pop_prompt = (
                            f"Provide popularity categories for a wrestler named {name_for_pop} with the bio {bio_for_pop}.\n"
                            "Consider the wrestler's description if available: " + player_desc_for_pop + "\n\n"
                            "Regions:\n"
                            "America\nCanada\nMexico\nBritish Isles\nJapan\nEurope\nOceania\nIndia\n\n"
                            "The categories are: Unknown, Insignificant, Indie Popularity, Recognized, Well Known, Very Popular, Superstar.\n"
                            "Return JSON only with keys as region names and values as categories.\n"
                            "If unsure, use 'Unknown'."
                        )

                        popularity_data = None
                        attempts = 0
                        while attempts < 3 and popularity_data is None:
                            attempts += 1
                            try:
                                response = self.client.chat.completions.create(
                                    model="gpt-3.5-turbo",
                                    messages=[{"role": "user", "content": pop_prompt}]
                                )
                                content = response.choices[0].message.content.strip()
                                p_data = json.loads(content)
                                required_keys = ["America","Canada","Mexico","British Isles","Japan","Europe","Oceania","India"]
                                if all(key in p_data for key in required_keys):
                                    popularity_data = p_data
                                else:
                                    raise ValueError("Missing required keys in popularity JSON.")
                            except Exception as e:
                                logging.error(f"Error getting popularity from GPT attempt {attempts}: {e}")
                                popularity_data = None

                        if popularity_data is None:
                            popularity_data = {
                                "America": "Unknown",
                                "Canada": "Unknown",
                                "Mexico": "Unknown",
                                "British Isles": "Unknown",
                                "Japan": "Unknown",
                                "Europe": "Unknown",
                                "Oceania": "Unknown",
                                "India": "Unknown"
                            }

                        popularity_values = self.convert_popularity_categories_to_values(popularity_data)
                        columns = ["WorkerUID"] + [f"Over{i}" for i in range(1,58)]
                        placeholders = ", ".join(["?"] * 58)
                        sql_insert_over = f"INSERT INTO tblWorkerOver ({', '.join(columns)}) VALUES ({placeholders})"
                        over_values = [worker_uid] + popularity_values
                        cursor.execute(sql_insert_over, over_values)

                    conn.commit()
                    conn.close()
                except Exception as e:
                    logging.error(f"Error saving to Access database: {e}", exc_info=True)
                    messagebox.showerror("Error", f"Could not save to Access database: {str(e)}")

            try:
                excel_path = "wrestleverse_workers.xlsx"
                if os.path.exists(excel_path):
                    existing_workers = pd.read_excel(excel_path, sheet_name="Workers")
                    existing_bios = pd.read_excel(excel_path, sheet_name="Bios")
                    existing_skills = pd.read_excel(excel_path, sheet_name="Skills")
                    existing_contracts = pd.read_excel(excel_path, sheet_name="Contracts")
                    existing_notes = pd.read_excel(excel_path, sheet_name="Notes")
                    
                    new_workers = pd.DataFrame(workers_data)
                    new_bios = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                    new_skills = pd.DataFrame(skills_data)
                    new_contracts = pd.DataFrame(contract_data)
                    new_notes = pd.DataFrame(notes_data)
                    
                    workers_df = pd.concat([existing_workers, new_workers], ignore_index=True)
                    bio_df = pd.concat([existing_bios, new_bios], ignore_index=True)
                    skills_df = pd.concat([existing_skills, new_skills], ignore_index=True)
                    contracts_df = pd.concat([existing_contracts, new_contracts], ignore_index=True)
                    notes_df = pd.concat([existing_notes, new_notes], ignore_index=True)
                else:
                    workers_df = pd.DataFrame(workers_data)
                    bio_df = pd.DataFrame(bio_data, columns=["UID", "Bio"])
                    skills_df = pd.DataFrame(skills_data)
                    contracts_df = pd.DataFrame(contract_data)
                    notes_df = pd.DataFrame(notes_data)

                if 'physical_description' not in notes_df.columns:
                    notes_df['physical_description'] = ''
                
                with pd.ExcelWriter(excel_path) as writer:
                    workers_df.to_excel(writer, sheet_name="Workers", index=False)
                    bio_df.to_excel(writer, sheet_name="Bios", index=False)
                    skills_df.to_excel(writer, sheet_name="Skills", index=False)
                    contracts_df.to_excel(writer, sheet_name="Contracts", index=False)
                    notes_df.to_excel(writer, sheet_name="Notes", index=False)
                
                messagebox.showinfo("Success", f"Wrestlers saved to {excel_path}")
            except Exception as e:
                logging.error(f"Error saving Excel file: {e}", exc_info=True)
                messagebox.showerror("Error", f"Could not save Excel file: {str(e)}")

            self.status_label.config(text="Status: Generation complete!")
            self.root.update_idletasks()

        except Exception as e:
            logging.error(f"Unhandled error in generate_wrestlers: {e}", exc_info=True)
            error_message = f"Error generating wrestlers: {str(e)}"
            self.status_label.config(text=f"Status: Error - {str(e)}")
            messagebox.showerror("Error", error_message)

    def get_region_popularity_from_gpt(self, name, bio):
        # This function is still called but we now implement a more robust approach inline above.
        # We'll keep this for compatibility, but we'll rely on the robust approach.
        return {
            "America": "Unknown",
            "Canada": "Unknown",
            "Mexico": "Unknown",
            "British Isles": "Unknown",
            "Japan": "Unknown",
            "Europe": "Unknown",
            "Oceania": "Unknown",
            "India": "Unknown"
        }

    def convert_popularity_categories_to_values(self, categories):
        def range_for_category(cat):
            if cat == "Unknown":
                return (0, 0)
            elif cat == "Insignificant":
                return (0, 15)
            elif cat == "Indie Popularity":
                return (10, 35)
            elif cat == "Recognized":
                return (20, 49)
            elif cat == "Well Known":
                return (35, 65)
            elif cat == "Very Popular":
                return (50, 85)
            elif cat == "Superstar":
                return (75, 99)
            else:
                return (0,0)

        def assign_values_for_region(region_name, count):
            cat = categories.get(region_name, "Unknown")
            min_val, max_val = range_for_category(cat)
            return [random.randint(min_val, max_val) for _ in range(count)]

        values = []
        values += assign_values_for_region("America", 11)
        values += assign_values_for_region("Canada", 7)
        values += assign_values_for_region("Mexico", 6)
        values += assign_values_for_region("British Isles", 6)
        values += assign_values_for_region("Japan", 8)
        values += assign_values_for_region("Europe", 8)
        values += assign_values_for_region("Oceania", 7)
        values += assign_values_for_region("India", 4)
        return values

    def ensure_byte(self, value):
        return max(0, min(255, int(value)))

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
        prompt = f"Generate a wrestling gimmick for a {alignment} wrestler. Return only a couple of words for the gimmick and not other text."
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
        browse_db_btn = ttk.Button(self.root, text="Browse", command=self.browse_access_db)
        browse_db_btn.pack(pady=5)
        
        pictures_label = ttk.Label(self.root, text="Pictures Path (Optional):")
        pictures_label.pack(pady=5)
        self.pictures_var = tk.StringVar(value=self.pictures_path)
        pictures_entry = ttk.Entry(self.root, textvariable=self.pictures_var, width=50)
        pictures_entry.pack(pady=5)
        browse_pics_btn = ttk.Button(self.root, text="Browse", command=self.browse_pictures_path)
        browse_pics_btn.pack(pady=5)

        start_date_label = ttk.Label(self.root, text="Start Date (YYYY-MM-DD):")
        start_date_label.pack(pady=5)
        self.start_date_var = tk.StringVar(value=self.start_date_str)
        start_date_entry = ttk.Entry(self.root, textvariable=self.start_date_var, width=15)
        start_date_entry.pack(pady=5)
        
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
        self.start_date_str = self.start_date_var.get()
        settings = {
            "api_key": self.api_key,
            "uid_start": self.uid_start,
            "bio_prompt": self.bio_prompt,
            "access_db_path": self.access_db_path,
            "pictures_path": self.pictures_path,
            "start_date": self.start_date_str
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
                self.start_date_str = settings.get("start_date", "")
        except FileNotFoundError:
            self.api_key = ""
            self.uid_start = 1
            self.bio_prompt = "Create a biography for a professional wrestler."
            self.access_db_path = ""
            self.pictures_path = ""
            self.start_date_str = ""

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

    def get_style_from_gpt(self, bio):
        prompt = (
            "Based on this wrestler's biography, select the most appropriate wrestling style number from this list:\n"
            "1-Regular\n2-Entertainer\n3-Comedy\n4-Powerhouse\n5-Impactful\n6-Striker\n"
            "7-Brawler\n8-Hardcore\n9-Psychopath\n10-Luchador\n11-High Flyer\n"
            "12-Technician\n13-Technician Flyer\n14-Technician Striker\n15-Daredevil\n"
            "16-MMA Crossover\n17-Never Wrestles\n\n"
            f"Biography: {bio}\n\n"
            "Respond with ONLY the number (1-17)."
        )
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )
            style_text = response.choices[0].message.content.strip()
            try:
                style = int(style_text)
                if 1 <= style <= 17:
                    return style
            except ValueError:
                pass
            return 1
        except Exception as e:
            logging.error(f"Error getting style from GPT: {e}")
            return 1

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
                messagebox.showerror("Error", f"Invalid input for skill {skill}.")
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
                messagebox.showerror("Error", f"Invalid input for skill {skill}.")
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
                
                company_list = [(int(company[0]), company[1]) for company in companies]
                return company_list
            except Exception as e:
                logging.error(f"Error getting companies: {e}", exc_info=True)
                return []
        return []

    def generate_contract(self, worker_data, worker_uid, company_choice, contract_uid):
        companies = self.get_companies()
        fed_uid = None
        
        if companies:
            if company_choice == "Random":
                company = random.choice(companies)
                fed_uid = company[0]
            elif company_choice == "Freelancer":
                return None
            else:
                for company in companies:
                    if company[1] == company_choice:
                        fed_uid = company[0]
                        break
                if fed_uid is None:
                    return None

        alignment_prompt = f"For a wrestler named {worker_data['name']}, should they be face or heel? Answer with 'face' or 'heel'."
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": alignment_prompt}]
            )
            alignment = response.choices[0].message.content.strip().lower()
            is_face = (alignment == "face")
        except:
            is_face = random.choice([True, False])

        contract_began_year = self.start_date.year - random.randint(1,5)
        contract_began_month = random.randint(1,12)
        contract_began_day = random.randint(1,28)
        contract_began = datetime.datetime(contract_began_year, contract_began_month, contract_began_day)
        if contract_began >= self.start_date:
            contract_began = self.start_date - datetime.timedelta(days=60)

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
            "ContractDebutDate": datetime.datetime(1900, 1, 1),
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

        for i in range(1, 7):
            contract_data[f"PlasterCaster_Byte{i}"] = 0

        for i in range(1, 26):
            contract_data[f"PlasterCaster_Bool{i}"] = False

        return contract_data

    def get_race_from_gpt(self, name, description):
        prompt = (
            f"Based on the name '{name}' and description '{description}', select the most appropriate "
            "race from this list and respond with ONLY the corresponding number:\n"
            "1: White\n2: Black\n3: Asian\n4: Hispanic\n5: American Indian\n"
            "6: Middle Eastern\n7: South Asian\n8: Pacific\n9: Other\n"
            "Respond with only the number."
        )
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}]
            )
            race_str = response.choices[0].message.content.strip()
            race = int(race_str)
            if 1 <= race <= 9:
                return race
            return 9
        except:
            return 9

    def get_response_from_gpt(self, prompt):
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
            excel_path = "wrestleverse_workers.xlsx"
            if not os.path.exists(excel_path):
                messagebox.showerror("Error", "Wrestlers Excel file not found.")
                return
            
            people_dir = os.path.join(self.pictures_path, "People")
            os.makedirs(people_dir, exist_ok=True)
            
            df = pd.read_excel(excel_path, sheet_name="Notes")
            pending_images = df[df['image_generated'] == False]
            
            if pending_images.empty:
                messagebox.showinfo("Info", "No pending images to generate.")
                return
            
            total_images = len(pending_images)
            generated_count = 0
            
            self.status_label.config(text=f"Generating images: 0/{total_images}")
            self.root.update_idletasks()
            
            for index, row in pending_images.iterrows():
                try:
                    prompt = (
                        f"Professional wrestling promotional photo. {row['physical_description']} "
                        "The image should be a high-quality, professional headshot style photo "
                        "with good lighting and a neutral background. The subject should be "
                        "looking directly at the camera with a confident expression."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    
                    resized_image = self.resize_image(response.content, (150, 150))
                    
                    image_name = row['Picture'][:26]
                    image_path = os.path.join(people_dir, image_name)
                    
                    with open(image_path, 'wb') as f:
                        f.write(resized_image)
                    
                    df.at[index, 'image_generated'] = True
                    generated_count += 1
                    
                    self.status_label.config(text=f"Generating images: {generated_count}/{total_images}")
                    self.root.update_idletasks()
                    
                    time.sleep(1)
                    
                except Exception as e:
                    logging.error(f"Error generating image for {row['Name']}: {str(e)}")
                    messagebox.showerror("Error", f"Failed to generate image for {row['Name']}: {str(e)}")
                    continue
            
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
            excel_path = "wrestleverse_companies.xlsx"
            if not os.path.exists(excel_path):
                messagebox.showerror("Error", "Companies Excel file not found.")
                return
                
            logos_dir = os.path.join(self.pictures_path, "Logos")
            banners_dir = os.path.join(self.pictures_path, "Banners")
            backdrops_dir = os.path.join(self.pictures_path, "Logo_Backdrops")
            
            for directory in [logos_dir, banners_dir, backdrops_dir]:
                os.makedirs(directory, exist_ok=True)
            
            df = pd.read_excel(excel_path, sheet_name="Notes")
            pending_images = df[df['image_generated'] == False]
            
            if pending_images.empty:
                messagebox.showinfo("Info", "No pending images to generate.")
                return
                
            total_companies = len(pending_images)
            generated_count = 0
            
            self.status_label.config(text=f"Generating images: 0/{total_companies}")
            self.root.update_idletasks()
            
            for index, row in pending_images.iterrows():
                try:
                    company_name = row['Name']
                    description = row['Description']
                    
                    logo_filename = (row['Logo'].replace('"', '').replace('\\', '')[:26])
                    banner_filename = (row['Banner'].replace('"', '').replace('\\', '')[:26] + '.jpg')
                    backdrop_filename = (row['Backdrop'].replace('"', '').replace('\\', '')[:26])
                    
                    logo_prompt = (
                        f"Professional wrestling company logo for \"{company_name}\". "
                        f"Company description: {description}. "
                        "The logo should be professional, memorable, and include the company name. "
                        "Use a transparent or solid background."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=logo_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    
                    resized_logo = self.resize_image(response.content, (150, 150))
                    with open(os.path.join(logos_dir, logo_filename), 'wb') as f:
                        f.write(resized_logo)
                    
                    time.sleep(1)
                    
                    banner_prompt = (
                        f"Professional wrestling company banner for \"{company_name}\". "
                        f"Company description: {description}. "
                        "Create a wide promotional banner (1:5 aspect ratio) with dynamic wrestling imagery. "
                        "Include the company name prominently."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=banner_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    
                    resized_banner = self.resize_image(response.content, (500, 40))
                    with open(os.path.join(banners_dir, banner_filename), 'wb') as f:
                        f.write(resized_banner)
                    
                    time.sleep(1)
                    
                    backdrop_prompt = (
                        f"Professional wrestling backdrop for \"{company_name}\". "
                        f"Company description: {description}. "
                        "Create a dramatic arena backdrop with the company's branding."
                    )
                    
                    response = self.client.images.generate(
                        model="dall-e-3",
                        prompt=backdrop_prompt,
                        size="1024x1024",
                        quality="standard",
                        n=1,
                    )
                    
                    image_url = response.data[0].url
                    response = requests.get(image_url)
                    response.raise_for_status()
                    
                    resized_backdrop = self.resize_image(response.content, (150, 150))
                    with open(os.path.join(backdrops_dir, backdrop_filename), 'wb') as f:
                        f.write(resized_backdrop)
                    
                    df.at[index, 'image_generated'] = True
                    generated_count += 1
                    
                    self.status_label.config(text=f"Generating images: {generated_count}/{total_companies}")
                    self.root.update_idletasks()
                    
                    time.sleep(1)
                    
                except Exception as e:
                    logging.error(f"Error generating images for {company_name}: {str(e)}")
                    messagebox.showerror("Error", f"Failed to generate images for {company_name}: {str(e)}")
                    continue
            
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Notes", index=False)
            
            messagebox.showinfo("Success", f"Generated images for {generated_count} companies successfully!")
            
        except Exception as e:
            logging.error(f"Error in generate_company_images: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.status_label.config(text="")
            self.root.update_idletasks()

    def get_race_name(self, race_number):
        race_map = {
            1: "Caucasian",
            2: "African American",
            3: "Asian",
            4: "Hispanic",
            5: "Indian",
            6: "Middle Eastern",
            7: "Pacific Islander",
            8: "Mixed Race"
        }
        return race_map.get(race_number, "Unknown")

    def resize_image(self, image_data, size):
        image = Image.open(io.BytesIO(image_data))
        image = image.resize(size, Image.Resampling.LANCZOS)
        output = io.BytesIO()
        image.save(output, format='JPEG')
        return output.getvalue()

    def browse_pictures_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self.pictures_var.set(directory)

if __name__ == "__main__":
    root = tk.Tk()
    app = WrestleverseApp(root)
    root.mainloop()
