import tkinter as tk
from tkinter import ttk, messagebox
import requests

RALPH_URL = 'http://192.168.20.41'
API_TOKEN = '81ec9623bf5a0def010def81d66468e532293ecd'
HEADERS = {
    'Authorization': f'Token {API_TOKEN}',
    'Content-Type': 'application/json'
}

# Predefined regions and warehouses
REGIONS = {
    "Melbourne Airport": 6,
    "Sydney": 5,
    "Townsville": 4,
    "Melbourne": 3,
    "Mackay": 2,
    "Brisbane": 1
}

WAREHOUSES = {
    "Townsville": 6,
    "Sydney": 7,
    "Melbourne Airport": 8,
    "Melbourne": 5,
    "Mackay": 4,
    "Caboolture": 3,
    "Brisbane": 2
}

MANUFACTURERS = {
    "Microsoft": 5,
    "HP": 1,
    "Metabox": 4,
    "Lenovo": 3
}

class RalphGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ralph Manager")

        self.laptop_model_map = {}  # Initialize the laptop_model_map

        # Create tab control
        self.tabControl = ttk.Notebook(self)
        self.tabControl.pack(expand=1, fill="both")

        # Create tabs
        self.add_user_tab = ttk.Frame(self.tabControl)
        self.add_laptop_tab = ttk.Frame(self.tabControl)
        self.add_model_tab = ttk.Frame(self.tabControl)
        self.assign_laptop_tab = ttk.Frame(self.tabControl)
        self.user_assets_tab = ttk.Frame(self.tabControl)

        self.tabControl.add(self.add_user_tab, text='Add User')
        self.tabControl.add(self.add_laptop_tab, text='Add Laptop')
        self.tabControl.add(self.add_model_tab, text='Add Model')
        self.tabControl.add(self.assign_laptop_tab, text='Assign Laptop')
        self.tabControl.add(self.user_assets_tab, text='User Assets')

        # Add content to tabs
        self.create_add_user_tab()
        self.create_add_laptop_tab()
        self.create_add_model_tab()
        self.create_assign_laptop_tab()
        self.create_user_assets_tab()

        # Bind tab change event to refresh data
        self.tabControl.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def create_add_user_tab(self):
        tk.Label(self.add_user_tab, text="Username:").grid(row=0, column=0, padx=10, pady=10)
        self.username_entry = tk.Entry(self.add_user_tab)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.add_user_tab, text="First Name:").grid(row=1, column=0, padx=10, pady=10)
        self.firstname_entry = tk.Entry(self.add_user_tab)
        self.firstname_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.add_user_tab, text="Last Name:").grid(row=2, column=0, padx=10, pady=10)
        self.lastname_entry = tk.Entry(self.add_user_tab)
        self.lastname_entry.grid(row=2, column=1, padx=10, pady=10)

        tk.Label(self.add_user_tab, text="Email:").grid(row=3, column=0, padx=10, pady=10)
        self.email_entry = tk.Entry(self.add_user_tab)
        self.email_entry.grid(row=3, column=1, padx=10, pady=10)

        tk.Button(self.add_user_tab, text="Add User", command=self.add_user).grid(row=4, column=0, columnspan=2, pady=10)

    def create_add_laptop_tab(self):
        tk.Label(self.add_laptop_tab, text="Model Name:").grid(row=0, column=0, padx=10, pady=10)
        self.model_combo = ttk.Combobox(self.add_laptop_tab, state="readonly", width=50)
        self.model_combo.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Serial Number:").grid(row=1, column=0, padx=10, pady=10)
        self.sn_entry = tk.Entry(self.add_laptop_tab)
        self.sn_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Barcode:").grid(row=2, column=0, padx=10, pady=10)
        self.barcode_entry = tk.Entry(self.add_laptop_tab)
        self.barcode_entry.grid(row=2, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Region:").grid(row=3, column=0, padx=10, pady=10)
        self.region_combo = ttk.Combobox(self.add_laptop_tab, state="readonly", width=50)
        self.region_combo['values'] = list(REGIONS.keys())
        self.region_combo.grid(row=3, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Warehouse:").grid(row=4, column=0, padx=10, pady=10)
        self.warehouse_combo = ttk.Combobox(self.add_laptop_tab, state="readonly", width=50)
        self.warehouse_combo['values'] = list(WAREHOUSES.keys())
        self.warehouse_combo.grid(row=4, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Manufacturer:").grid(row=5, column=0, padx=10, pady=10)
        self.manufacturer_combo = ttk.Combobox(self.add_laptop_tab, state="readonly", width=50)
        self.manufacturer_combo['values'] = list(MANUFACTURERS.keys())
        self.manufacturer_combo.grid(row=5, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Price:").grid(row=6, column=0, padx=10, pady=10)
        self.price_entry = tk.Entry(self.add_laptop_tab)
        self.price_entry.grid(row=6, column=1, padx=10, pady=10)

        tk.Label(self.add_laptop_tab, text="Price Currency:").grid(row=7, column=0, padx=10, pady=10)
        self.price_currency_entry = tk.Entry(self.add_laptop_tab)
        self.price_currency_entry.grid(row=7, column=1, padx=10, pady=10)

        tk.Button(self.add_laptop_tab, text="Add Laptop", command=self.add_laptop).grid(row=8, column=0, columnspan=2, pady=10)

    def create_add_model_tab(self):
        tk.Label(self.add_model_tab, text="Model Name:").grid(row=0, column=0, padx=10, pady=10)
        self.model_name_entry = tk.Entry(self.add_model_tab)
        self.model_name_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.add_model_tab, text="Manufacturer:").grid(row=1, column=0, padx=10, pady=10)
        self.manufacturer_combo_model = ttk.Combobox(self.add_model_tab, state="readonly", width=50)
        self.manufacturer_combo_model['values'] = list(MANUFACTURERS.keys())
        self.manufacturer_combo_model.grid(row=1, column=1, padx=10, pady=10)

        tk.Button(self.add_model_tab, text="Add Model", command=self.add_model).grid(row=2, column=0, columnspan=2, pady=10)

    def create_assign_laptop_tab(self):
        tk.Label(self.assign_laptop_tab, text="Select User:").grid(row=0, column=0, padx=10, pady=10)
        self.user_combo = ttk.Combobox(self.assign_laptop_tab, state="readonly", width=50)
        self.user_combo.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.assign_laptop_tab, text="Select Laptop Serial Number:").grid(row=1, column=0, padx=10, pady=10)
        self.laptop_model_combo = ttk.Combobox(self.assign_laptop_tab, state="readonly", width=50)
        self.laptop_model_combo.grid(row=1, column=1, padx=10, pady=10)

        tk.Button(self.assign_laptop_tab, text="Assign Laptop", command=self.assign_laptop).grid(row=2, column=0, pady=10)
       
    def create_user_assets_tab(self):
        tk.Label(self.user_assets_tab, text="Search User:").grid(row=0, column=0, padx=10, pady=10)
        self.search_user_entry = tk.Entry(self.user_assets_tab)
        self.search_user_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.user_assets_tab, text="Search", command=self.search_user_assets).grid(row=0, column=2, padx=10, pady=10)

        self.assets_listbox = tk.Listbox(self.user_assets_tab)
        self.assets_listbox.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        self.user_assets_tab.grid_columnconfigure(0, weight=1)
        self.user_assets_tab.grid_rowconfigure(1, weight=1)

        tk.Button(self.user_assets_tab, text="Remove Selected Laptop", command=self.remove_selected_laptop).grid(row=2, column=0, columnspan=3, pady=10)

    def on_tab_change(self, event):
        selected_tab = event.widget.tab(event.widget.select(), "text")
        if selected_tab == 'Assign Laptop':
            self.update_users()
            self.update_laptop_models()
        elif selected_tab == 'Add Laptop':
            self.update_models()

    def update_laptop_models(self):
        laptop_models = self.fetch_available_laptop_models()
        if laptop_models:
            serial_numbers = [model['sn'] for model in laptop_models]
            self.laptop_model_combo['values'] = serial_numbers
            self.laptop_model_map = {model['sn']: model['id'] for model in laptop_models}
        else:
            self.laptop_model_combo['values'] = []
            messagebox.showinfo("Info", "No available laptops found.")

    def fetch_available_laptop_models(self):
        try:
            response = requests.get(f'{RALPH_URL}/api/back-office-assets/', headers=HEADERS, params={'limit': 200})
            if response.status_code == 200:
                assets = response.json().get('results', [])
                available_statuses = {'return in progress', 'new'}
                return [asset for asset in assets if asset['status'] in available_statuses]
            else:
                messagebox.showerror("Error", f"Failed to fetch laptop models: {response.text}")
                return []
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            return []

    def update_models(self):
        try:
            response = requests.get(f'{RALPH_URL}/api/assetmodels/', headers=HEADERS, params={'limit': 200})
            if response.status_code == 200:
                models = response.json().get('results', [])
                model_names = [model['name'] for model in models]
                self.model_combo['values'] = model_names
                self.model_map = {model['name']: model['id'] for model in models}
            else:
                messagebox.showerror("Error", f"Failed to fetch models: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def update_users(self):
        try:
            response = requests.get(f'{RALPH_URL}/api/users/', headers=HEADERS, params={'limit': 500})
            if response.status_code == 200:
                users = response.json().get('results', [])
                usernames = [user['username'] for user in users]
                self.user_combo['values'] = usernames
                self.user_map = {user['username']: user['id'] for user in users}
            else:
                messagebox.showerror("Error", f"Failed to fetch users: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def add_user(self):
        username = self.username_entry.get()
        first_name = self.firstname_entry.get()
        last_name = self.lastname_entry.get()
        email = self.email_entry.get()

        if not username or not first_name or not last_name or not email:
            messagebox.showerror("Error", "All fields are required")
            return

        user_data = {
            "username": username,
            "first_name": first_name,
            "last_name": last_name,
            "email": email
        }

        try:
            response = requests.post(f'{RALPH_URL}/api/users/', headers=HEADERS, json=user_data)
            if response.status_code == 201:
                messagebox.showinfo("Success", "User added successfully")
            else:
                messagebox.showerror("Error", f"Failed to add user: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def assign_laptop(self):
        username = self.user_combo.get()
        selected_sn = self.laptop_model_combo.get()

        if not username or not selected_sn:
            messagebox.showerror("Error", "All fields are required")
            return

        user_id = self.user_map[username]
        asset_id = self.laptop_model_map[selected_sn]

        data = {
            "user": user_id,
            "status": "in use"
        }

        try:
            response = requests.patch(f'{RALPH_URL}/api/back-office-assets/{asset_id}/', headers=HEADERS, json=data)
            if response.status_code == 200:
                messagebox.showinfo("Success", "Laptop assigned successfully")
            else:
                messagebox.showerror("Error", f"Failed to assign laptop: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def remove_laptop(self):
        username = self.user_combo.get()
        selected_sn = self.laptop_model_combo.get()

        if not username or not selected_sn:
            messagebox.showerror("Error", "All fields are required")
            return

        asset_id = self.laptop_model_map[selected_sn]

        data = {
            "user": None,
            "status": "free"
        }

        try:
            response = requests.patch(f'{RALPH_URL}/api/back-office-assets/{asset_id}/', headers=HEADERS, json=data)
            if response.status_code == 200:
                messagebox.showinfo("Success", "Laptop removed successfully")
            else:
                messagebox.showerror("Error", f"Failed to remove laptop: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def remove_selected_laptop(self):
        selected_asset = self.assets_listbox.get(tk.ACTIVE)
        if not selected_asset:
            messagebox.showerror("Error", "No laptop selected")
            return

        asset_sn = selected_asset.split(' - ')[0]

        if asset_sn not in self.laptop_model_map:
            messagebox.showerror("Error", "Selected laptop not found")
            return

        asset_id = self.laptop_model_map[asset_sn]

        data = {
            "user": None,
            "status": "free"
        }

        try:
            response = requests.patch(f'{RALPH_URL}/api/back-office-assets/{asset_id}/', headers=HEADERS, json=data)
            if response.status_code == 200:
                messagebox.showinfo("Success", "Laptop removed successfully")
                self.search_user_assets()
            else:
                messagebox.showerror("Error", f"Failed to remove laptop: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def search_user_assets(self):
        username = self.search_user_entry.get()

        if not username:
            messagebox.showerror("Error", "Username is required")
            return

        try:
            response = requests.get(f'{RALPH_URL}/api/users/', headers=HEADERS, params={'username': username})
            if response.status_code == 200:
                users = response.json().get('results', [])
                if users:
                    user_id = users[0]['id']
                    self.fetch_user_assets(user_id)
                else:
                    messagebox.showinfo("Info", "No user found")
            else:
                messagebox.showerror("Error", f"Failed to fetch user: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def fetch_user_assets(self, user_id):
        try:
            response = requests.get(f'{RALPH_URL}/api/back-office-assets/', headers=HEADERS, params={'user': user_id})
            if response.status_code == 200:
                assets = response.json().get('results', [])
                self.assets_listbox.delete(0, tk.END)
                for asset in assets:
                    self.assets_listbox.insert(tk.END, f"{asset['sn']} - {asset['model']['name']}")
                self.laptop_model_map = {asset['sn']: asset['id'] for asset in assets}  # Update the laptop_model_map
            else:
                messagebox.showerror("Error", f"Failed to fetch user assets: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def add_laptop(self):
        model_name = self.model_combo.get()
        sn = self.sn_entry.get()
        barcode = self.barcode_entry.get()
        region_name = self.region_combo.get()
        warehouse_name = self.warehouse_combo.get()
        manufacturer_name = self.manufacturer_combo.get()
        price = self.price_entry.get()
        price_currency = self.price_currency_entry.get()

        if not model_name or not sn or not barcode or not region_name or not warehouse_name or not manufacturer_name or not price or not price_currency:
            messagebox.showerror("Error", "All fields are required")
            return

        model_id = self.model_map[model_name]
        region_id = REGIONS[region_name]
        warehouse_id = WAREHOUSES[warehouse_name]
        manufacturer_id = MANUFACTURERS[manufacturer_name]

        data = {
            "model": model_id,
            "sn": sn,
            "barcode": barcode,
            "type": "back_office",
            "status": "new",
            "region": region_id,
            "warehouse": warehouse_id,
            "price": price,
            "price_currency": price_currency,
            "manufacturer": manufacturer_id
        }

        try:
            response = requests.post(f'{RALPH_URL}/api/back-office-assets/', headers=HEADERS, json=data)
            if response.status_code == 201:
                messagebox.showinfo("Success", "Laptop added successfully")
            else:
                messagebox.showerror("Error", f"Failed to add laptop: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def add_model(self):
        model_name = self.model_name_entry.get()
        manufacturer_name = self.manufacturer_combo_model.get()

        if not model_name or not manufacturer_name:
            messagebox.showerror("Error", "All fields are required")
            return

        manufacturer_id = MANUFACTURERS[manufacturer_name]

        data = {
            "name": model_name,
            "manufacturer": manufacturer_id,
            "power_consumption": 60,
            "height_of_device": 1,
            "cores_count": 1,
            "type": "back_office"
        }

        try:
            response = requests.post(f'{RALPH_URL}/api/assetmodels/', headers=HEADERS, json=data)
            if response.status_code == 201:
                messagebox.showinfo("Success", "Model added successfully")
                self.update_models()
            else:
                messagebox.showerror("Error", f"Failed to add model: {response.text}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = RalphGUI()
    app.mainloop()

