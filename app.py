import os
import sys
import json
import warnings
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import openpyxl
# Global Variables
file_paths = []
dynamic_entry_fields = []
dynamic_header_entry_fields = []
agent_profile = []
agent_ids = []
agent_ids_loaded = []
agent_names_loaded = []
template_names_loaded = []
header_names_loaded = []
assigned_templates = {}
assigned_headers = {}
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# Global Paths
agents_json = resource_path('data/agents.json')
templates_json = resource_path('data/templates.json')
header_json = resource_path('data/header.json')
forest_light = resource_path("assets/forest-light.tcl")
forest_dark = resource_path("assets/forest-dark.tcl")
icon_path =  resource_path("assets/icon.ico")
# File Selection
def select_files():
    file_path = tk.filedialog.askopenfilenames(filetype=[("Excel Files", "*.xlsx")])
    for path in file_path:
    #     file_name = os.path.basename(path)
        file_listbox.insert(tk.END, path)
        width_entry = max(file_listbox.get(0, tk.END), key=len)
        file_listbox.config(width=len(width_entry))
    #     file_paths.append(file_path)
# Remove File From Selected
def remove_file():
    selected_file = file_listbox.curselection()
    if selected_file:
        file_listbox.delete(selected_file)
# Dynamic Entry's
def dynamic_template_entry():
    new_entry = ttk.Entry(template_creation)
    new_entry.grid(padx=5, pady=5)
    add_placeholder(new_entry, "Column Name")
    dynamic_entry_fields.append(new_entry)
def dynamic_agent_entry():
    new_entry = ttk.Entry(agent_creation)
    new_entry.grid(padx=5, pady=5)
    add_placeholder(new_entry, "Agent ID")
    agent_ids.append(new_entry)
def dynamic_header_entry():
    new_entry = ttk.Entry(header_creation)
    new_entry.grid(padx=5, pady=5)
    add_placeholder(new_entry, "Header Title")
    dynamic_header_entry_fields.append(new_entry)
def remove_dynamic_entry():
    for entry in dynamic_entry_fields:
        entry.destroy()
    for entry in agent_ids:
        entry.destroy()
    for entry in dynamic_header_entry_fields:
        entry.destroy()
    dynamic_entry_fields.clear()
    dynamic_header_entry_fields.clear()
    agent_ids.clear()
    load_data()
# Change Theme
def change_theme():
    if theme_switch.instate(["selected"]):
        style.theme_use("forest-light")
        theme_switch.config(text="Light")
    else:
        style.theme_use("forest-dark")
        theme_switch.config(text="Dark")
# Entry Placeholder Text
def add_placeholder(entry, text):
    def on_entry_focusin():
        if entry.get() == text:
            entry.delete(0, tk.END)
            entry.configure(show='')
    def on_entry_focusout():
        if entry.get() == '':
            entry.insert(0, text)
    entry.insert(0, text)
    entry.bind("<FocusIn>", on_entry_focusin)
    entry.bind("<FocusOut>", on_entry_focusout)
# Add Agent To Listbox / Profile File
def add_agent():
    agent_profile.append(name_agent_entry.get())
    current_agent_listbox.insert(tk.END, name_agent_entry.get())
    try:
        with open(agents_json, 'r') as file:
            agents = json.load(file)
    except FileNotFoundError:
        agents = []
    agents_data = {
        "name": name_agent_entry.get(),
        "identifiers": ", ".join([entry.get() for entry in agent_ids])
    }
    agents.append(agents_data)
    with open(agents_json, 'w') as file:
        json.dump(agents, file, indent=4)
        file.write('\n')
    name_agent_entry.delete(0, tk.END)
    add_placeholder(name_agent_entry, "Agent Name")
    remove_dynamic_entry()
# Remove Agent From Listbox / Profile File
def remove_agent():
    selected_indices = current_agent_listbox.curselection()
    for index in selected_indices:
        selected_agent = current_agent_listbox.get(index)
        with open(agents_json, 'r') as file:
            data = json.load(file)
            remove_indices = [i for i, entry in enumerate(data) if entry["name"] == selected_agent]
            for i in remove_indices:
                data.pop(i)
            with open(agents_json, 'w') as file:
                json.dump(data, file, indent=4)
        current_agent_listbox.delete(index)
# Adds Templates To Listbox / Templates File
def add_template():
    template_listbox.insert(tk.END, name_template_entry.get())
    try:
        with open(templates_json, 'r') as file:
            template_f = json.load(file)
    except FileNotFoundError:
        template_f = []
    template_data = {
        "name": name_template_entry.get(),
        "file": "",
        "sheet": sheet_template_entry.get(),
        "header": header_template_entry.get(),
        "id_column": identifier_template_entry.get(),
        "columnscopy": ", ".join([entry.get() for entry in dynamic_entry_fields])
    }
    template_f.append(template_data)
    with open(templates_json, 'w') as file:
        json.dump(template_f, file, indent=4)
        file.write('\n')
    name_template_entry.delete(0, tk.END)
    add_placeholder(name_template_entry, "Template Name")
    sheet_template_entry.delete(0, tk.END)
    add_placeholder(sheet_template_entry, "Sheet Name")
    header_template_entry.set(0)
    identifier_template_entry.delete(0, tk.END)
    add_placeholder(identifier_template_entry, "Identifier Column")
    remove_dynamic_entry()
# Remove Template From Listbox / Templates File
def remove_template():
    selected_indices = template_listbox.curselection()
    for index in selected_indices:
        selected_name = template_listbox.get(index)
        with open(templates_json, 'r') as file:
            data = json.load(file)
            remove_indices = [i for i, entry in enumerate(data) if entry["name"] == selected_name]
            for i in remove_indices:
                data.pop(i)
            with open(templates_json, 'w') as file:
                json.dump(data, file, indent=4)
        template_listbox.delete(index)
# Add Header To Listbox / Header File
def add_header():
    header_listbox.insert(tk.END, name_header_entry.get())
    try:
        with open(header_json, 'r') as file:
            headers = json.load(file)
    except FileNotFoundError:
        headers = []
    header_data = {
        "name": name_header_entry.get(),
        "file": "",
        "sheet": sheet_header_entry.get(),
        "headers": ", ".join([entry.get() for entry in dynamic_header_entry_fields])
    }
    headers.append(header_data)
    with open(header_json, 'w') as file:
        json.dump(headers, file, indent=4)
        file.write('\n')
    name_header_entry.delete(0, tk.END)
    add_placeholder(name_header_entry, "Header Template Name")
    sheet_header_entry.delete(0, tk.END)
    add_placeholder(sheet_header_entry, "Sheet Name")
    remove_dynamic_entry()
# Remove Header From Listbox / Header File
def remove_header():
    selected_indices = header_listbox.curselection()
    for index in selected_indices:
        selected_name = header_listbox.get(index)
        with open(header_json, 'r') as file:
            data = json.load(file)
            remove_indices = [i for i, entry in enumerate(data) if entry["name"] == selected_name]
            for i in remove_indices:
                data.pop(i)
            with open(header_json, 'w') as file:
                json.dump(data, file, indent=4)
        header_listbox.delete(index)
# Show Header Info
def show_header_info(event):
    index = header_listbox.curselection()
    if index:
        selected_header = header_listbox.get(index)
        with open(header_json, 'r') as file:
            data = json.load(file)
        for item in data:
            if item['name'] == selected_header:
                sheet = item['sheet']
                headers = item['headers']
                info = f"Sheet: {sheet}", f"Headers: {headers}"
                header_info_listbox.delete(0, tk.END)
                for i in info:
                    header_info_listbox.insert(tk.END, i)
                    entry_width = max(header_info_listbox.get(0, tk.END), key=len)
                    header_info_listbox.config(width=len(entry_width))
# Load JSON Files TODO:Fix loading agent ids
def load_data():
    global agent_ids_loaded, agent_names_loaded, template_names_loaded, header_names_loaded
    with open(agents_json, 'r') as file:
        data_a = json.load(file)
        agent_ids_loaded = [agent['identifiers'] for agent in data_a if 'identifiers' in agent]
        agent_names_loaded = [name['name'] for name in data_a]
    with open(templates_json, 'r') as file:
        data_t = json.load(file)
        template_names_loaded = [template['name'] for template in data_t]
    with open(header_json, 'r') as file:
        data_h = json.load(file)
        header_names_loaded = [header['name'] for header in data_h]
# Show IDs
def show_ids(event):
    index = current_agent_listbox.curselection()
    if index:
        selected_name = current_agent_listbox.get(index)
        with open(agents_json, 'r') as file:
            data = json.load(file)
            identifiers = []
            for ids in data:
                if ids['name'] == selected_name:
                    identifiers = ids['identifiers'].split(", ")
                agent_id_listbox.delete(0, tk.END)
                for identifier in identifiers:
                    agent_id_listbox.insert(tk.END, identifier.strip())
# Remove Id From Agent TODO:Does not do anything yet
def remove_identifier():
    index = agent_id_listbox.curselection()
    if index:
        selected_identifer = agent_id_listbox.get(index)
        selected_agent_index = current_agent_listbox.curselection()
        if selected_agent_index:
            selected_agent = current_agent_listbox.get(selected_agent_index)
            with open(agents_json, 'r+') as file:
                data = json.load(file)
                for agent in data:
                    if agent['name'] == selected_agent:
                        identifiers = agent['identifiers'].split(", ")
                        identifiers.remove(selected_identifer)
                        agent['identifiers'] = ", ".join(identifiers)
                with open(agents_json, 'w') as file:
                    json.dump(data, file, indent=4)
                agent_id_listbox.delete(index)
# Show Template Info
def show_info(event):
    index = template_listbox.curselection()
    if index:
        selected_template = template_listbox.get(index)
        with open(templates_json, 'r') as file:
            data = json.load(file)
        for item in data:
            if item['name'] == selected_template:
                sheet = item['sheet']
                header = item['header']
                id_column = item['id_column']
                columns = item['columnscopy']
                info = f"Sheet: {sheet}", f"Header: {header}", f"Identifying Column: {id_column}", f"Wanted Columns: {columns}"
                template_info_listbox.delete(0, tk.END)
                for i in info:
                    template_info_listbox.insert(tk.END, i)
                    entry_width = max(template_info_listbox.get(0, tk.END), key=len)
                    template_info_listbox.config(width=len(entry_width))
# Assign Template To File
def assign_template():
    selected_file = file_listbox.get(file_listbox.curselection())
    selected_template = template_selection_cb.get()
    with open(templates_json) as file:
        data = json.load(file)
    modify = next((entry for entry in data if entry['name'] == selected_template), None)
    if modify is not None:
        modify['file'] = selected_file
    with open(templates_json, 'w') as file:
        json.dump(data, file, indent=4)
    global assigned_templates
    template = template_selection_cb.get()
    assigned_templates[selected_file] = template
    assigned_listbox.delete(0, tk.END)
    for assigned, template in assigned_templates.items():
        file_name = os.path.basename(assigned)
        assigned_listbox.insert(tk.END, f"{file_name}-{template}")
        entry_width = max(assigned_listbox.get(0, tk.END), key=len)
        assigned_listbox.config(width=len(entry_width))
# Assign Header To File
def assign_header():
    selected_file = file_listbox.get(file_listbox.curselection())
    selected_header = header_selection_cb.get()
    with open(header_json) as file:
        data = json.load(file)
    modify = next((entry for entry in data if entry['name'] == selected_header), None)
    if modify is not None:
        modify['file'] = selected_file
    with open(header_json, 'w') as file:
        json.dump(data, file, indent=4)
    global assigned_headers
    header = header_selection_cb.get()
    assigned_headers[selected_file] = header
    assigned_listbox.delete(0, tk.END)
    for assigned, header in assigned_headers.items():
        file_name = os.path.basename(assigned)
        assigned_listbox.insert(tk.END, f"{file_name}-H{header}")
        entry_width = max(assigned_listbox.get(0, tk.END), key=len)
        assigned_listbox.config(width=len(entry_width))
# Reset File Dir
def reset_file_dir():
    with open(templates_json) as file:
        data = json.load(file)
    for entry in data:
        entry['file'] = ""
    with open(templates_json, 'w') as file:
        json.dump(data, file, indent=4)
    with open(header_json) as file:
        data = json.load(file)
    for entry in data:
        entry['file'] = ""
    with open(header_json, 'w') as file:
        json.dump(data, file, indent=4)
# Move Wanted Data
def data_move():
    add_header_to_excel()
    print("header added")
    output_file = 'Commission Data.xlsx'
    warnings.filterwarnings("ignore", category=FutureWarning)
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    print("fil, warning, writer")
    with open(agents_json, 'r') as agent_file:
        agent_data = json.load(agent_file)
    print("opened agent json")
    for agent in agent_data:
        agent_name = agent['name']
        agent_identifier = agent['identifiers'].split(", ")
        print("agent_name, Identifiers")
        with open(templates_json) as file:
            file_data = json.load(file)
        print("Loaded template json")
        for index, data in enumerate(file_data):
            print(data)
            if data['file'] == "":
                # messagebox.showerror("Error", "No Files Selected")
                continue
            name = data['name']
            file = data['file']
            sheet = data['sheet']
            header = data['header']
            id_column = data['id_column']
            columnscopy = data['columnscopy'].split(", ")
            print("loaded data")
            df = pd.read_excel(file, sheet_name=sheet, header=int(header))
            data = df.loc[df[id_column].isin(agent_identifier)]
            print("searched and copied excel")
            if not data.empty:
                data = data[columnscopy]
                print(agent_name, name)
                print(data)
                if os.path.exists(output_file):
                    sheet_name = f"{agent_name}_{name}"
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                print(f"{agent_name}: No Data Found In {name}")
    print("saved to excel")
    writer._save()
    messagebox.showinfo("Success", "Your Data Has Been Saved.")

# Adds Header To Excel File
def add_header_to_excel():
    with open(header_json) as file:
        file_data = json.load(file)
    for i, data in enumerate(file_data, start=1):
        if 'file' not in data or data['file'] == "":
            continue
        file = data['file']
        file_sheet = data['sheet']
        header_col = data['headers'].split(", ")
        if file != "":
            header_workbook = openpyxl.load_workbook(file)
            sheet = header_workbook[file_sheet]
            sheet.insert_rows(1)
            for i, header in enumerate(header_col, start=1):
                sheet.cell(row=1, column=i).value = header
            header_workbook.save(file)
            break
# Load Data In App
load_data()
reset_file_dir()
# Initialize App
root = tk.Tk()
root.title("Commissions Expeditor")
root.iconbitmap(icon_path)
tab_control = ttk.Notebook(root)
tab_control.pack(expand=1, fill="both")
# root.maxsize(width=800, height=850)
# Sets Style/Theme
style = ttk.Style(root)
root.tk.call("source", forest_light)
root.tk.call("source", forest_dark)
style.theme_use("forest-dark")
# Home Tab
tab0 = ttk.Frame(tab_control)
tab_control.add(tab0, text="Home")
home_frame = ttk.Frame(tab0)
home_frame.pack()
# Frames
file_selection_frame = ttk.Labelframe(home_frame, text="File Selection")
file_selection_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
file_display_frame = ttk.Labelframe(home_frame, text="Files Selected")
file_display_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
# Buttons
select_file_button = ttk.Button(file_selection_frame, text="Select Files", style='Accent.TButton', command=select_files)
select_file_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
assign_template_button = ttk.Button(file_selection_frame, text="Assign Template", command=assign_template)
assign_template_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
assign_header_button = ttk.Button(file_selection_frame, text="Assign Header", command=assign_header)
assign_header_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
remove_file_button = ttk.Button(file_display_frame, text="Remove File", command=remove_file)
remove_file_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
confirm_button = ttk.Button(file_selection_frame, text="Confirm and Run", command=data_move)
confirm_button.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
# Combobox / Listbox
template_selection_cb = ttk.Combobox(file_selection_frame, values=template_names_loaded)
template_selection_cb.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
header_selection_cb = ttk.Combobox(file_selection_frame, values=header_names_loaded)
header_selection_cb.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
file_listbox = tk.Listbox(file_display_frame)
file_listbox.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
file_display_frame.grid_rowconfigure(0, weight=1)
assigned_listbox = tk.Listbox(file_display_frame)
assigned_listbox.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
# Template Tab
tab1 = ttk.Frame(tab_control)
tab_control.add(tab1, text="Templates")
template_frame = ttk.Frame(tab1)
template_frame.pack()
# Frames
template_creation = ttk.Labelframe(template_frame, text="Template Creation")
template_creation.grid(row=0, column=0, padx=5, pady=5)
template_display = ttk.Labelframe(template_frame, text="Current Templates")
template_display.grid(row=0, column=1, padx=5, pady=5)
# Buttons
create_template_button = ttk.Button(template_creation, text="Create Template", style='Accent.TButton', command=add_template)
create_template_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
remove_template_button = ttk.Button(template_display, text="Remove Template", command=remove_template)
remove_template_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
add_column_button = ttk.Button(template_creation, text="Add Column", command=dynamic_template_entry)
add_column_button.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
# Entry's / Spinbox / Listbox
name_template_entry = ttk.Entry(template_creation)
name_template_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(name_template_entry, "Template Name")
sheet_template_entry = ttk.Entry(template_creation)
sheet_template_entry.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(sheet_template_entry, "Sheet Name")
identifier_template_entry = ttk.Entry(template_creation)
identifier_template_entry.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(identifier_template_entry, "Identifier Column")
header_template_entry = ttk.Spinbox(template_creation)
header_template_entry.set(0)
header_template_entry.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
template_listbox = tk.Listbox(template_display)
template_listbox.grid(row=0, column=0, padx=5, pady=5)
for template_item in template_names_loaded:
    template_listbox.insert(tk.END, template_item)
template_info_listbox = tk.Listbox(template_display)
template_info_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
template_listbox.bind('<<ListboxSelect>>', show_info)
# Headers Tab
tab4 = ttk.Frame(tab_control)
tab_control.add(tab4, text="Headers")
header_frame = ttk.Frame(tab4)
header_frame.pack()
# Frames
header_creation = ttk.Labelframe(header_frame, text="Header Creation")
header_creation.grid(row=0, column=0, padx=5, pady=5)
header_display = ttk.Labelframe(header_frame, text="Header Display")
header_display.grid(row=0, column=1, padx=5, pady=5)
# Buttons
create_header_button = ttk.Button(header_creation, text="Create Header", style='Accent.TButton', command=add_header)
create_header_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
remove_header_button = ttk.Button(header_display, text="Remove Header", command=remove_header)
remove_header_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
add_header_button = ttk.Button(header_creation, text="Add Header", command=dynamic_header_entry)
add_header_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
# Entry's / Listbox's
name_header_entry = ttk.Entry(header_creation)
name_header_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(name_header_entry, "Header Template Name")
sheet_header_entry = ttk.Entry(header_creation)
sheet_header_entry.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(sheet_header_entry, "Sheet Name")
header_listbox = tk.Listbox(header_display)
header_listbox.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
header_listbox.bind('<<ListboxSelect>>', show_header_info)
for header_item in header_names_loaded:
    header_listbox.insert(tk.END, header_item)
header_info_listbox = tk.Listbox(header_display)
header_info_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
# Agents Tab
tab2 = ttk.Frame(tab_control)
tab_control.add(tab2, text="Agents")
agent_frame = ttk.Frame(tab2)
agent_frame.pack()
# Frames
agent_creation = ttk.Labelframe(agent_frame, text="Agent Creation")
agent_creation.grid(row=0, column=0, padx=5, pady=5)
agent_display = ttk.Labelframe(agent_frame, text="Agent Display")
agent_display.grid(row=0, column=1, padx=5, pady=5)
# Buttons
create_agent_button = ttk.Button(agent_creation, text="Create Agent", style='Accent.TButton', command=add_agent)
create_agent_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
remove_agent_button = ttk.Button(agent_display, text="Remove Agent", command=remove_agent)
remove_agent_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
remove_agent_id_button = ttk.Button(agent_display, text="Remove ID", command=remove_identifier)
remove_agent_id_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
add_id_agent_button = ttk.Button(agent_creation, text="Add Agent ID", command=dynamic_agent_entry)
add_id_agent_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
# Entry's / Listbox
name_agent_entry = ttk.Entry(agent_creation)
name_agent_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
add_placeholder(name_agent_entry, "Agents Name")
current_agent_listbox = tk.Listbox(agent_display)
current_agent_listbox.grid(row=0, column=0, padx=5, pady=5)
current_agent_listbox.bind('<<ListboxSelect>>', show_ids)
for agent_item in agent_names_loaded:
    current_agent_listbox.insert(tk.END, agent_item)
agent_id_listbox = tk.Listbox(agent_display)
agent_id_listbox.grid(row=0, column=1, padx=5, pady=5)
# Settings Tab
tab3 = ttk.Frame(tab_control)
tab_control.add(tab3, text="Settings")
setting_frame = ttk.Frame(tab3)
setting_frame.pack()
# Frames
theme_change = ttk.Labelframe(setting_frame, text="Change Theme")
theme_change.grid(row=0, column=0, padx=5, pady=5)
# Checkbox's
theme_switch = ttk.Checkbutton(theme_change, text="Dark", style="Switch", command=change_theme)
theme_switch.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
# Runs App
root.mainloop()
