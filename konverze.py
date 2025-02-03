#Tento program je poskytován „tak, jak je“, bez jakýchkoliv záruk,
# #výslovných či předpokládaných, včetně (mimo jiné) předpokládaných
#záruk prodejnosti nebo vhodnosti pro určitý účel. Autor nenese
# #odpovědnost za jakékoli přímé, nepřímé nebo následné škody vyplývající
#z použití tohoto programu. Používáte jej výhradně na vlastní nebezpečí.
#Program je šířen zdarma.

import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import time
import threading

def convert_doc_to_docx(input_path, output_path):
    try:
        subprocess.run(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'docx', '--outdir', output_path, input_path], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f'Failed to convert {input_path}: {e}')
        return False

def convert_all_docs_in_directory(directory):
    docx_directory = os.path.join(directory, "DOCX")
    os.makedirs(docx_directory, exist_ok=True)

    doc_files = [f for f in os.listdir(directory) if f.endswith(".doc")]
    total_files = len(doc_files)
    
    progress_bar["maximum"] = total_files
    start_time = time.time()

    converted_files = 0
    for index, filename in enumerate(doc_files):
        input_file = os.path.join(directory, filename)
        if convert_doc_to_docx(input_file, docx_directory):
            converted_files += 1

        # Aktualizace průběhu a zbývajícího času
        progress_bar["value"] = index + 1
        progress_label.config(text=f"Převádím: {filename} ({index + 1}/{total_files})")
        elapsed_time = time.time() - start_time
        estimated_total_time = (elapsed_time / (index + 1)) * total_files
        remaining_time = estimated_total_time - elapsed_time
        time_label.config(text=f"Zbývající čas: {int(remaining_time)} sekund")
        
        root.update_idletasks()  # Ujistí se, že GUI se aktualizuje během běhu

    if converted_files > 0:
        messagebox.showinfo("Hotovo", f"Všechny soubory .doc byly úspěšně převedeny a uloženy do {docx_directory}.")
    else:
        messagebox.showwarning("Upozornění", "Žádné soubory nebyly převedeny.")
    
    # Po dokončení odemkni tlačítko a resetuj ukazatele
    select_button.config(state=tk.NORMAL)
    progress_label.config(text="Čekání na výběr složky...")
    progress_bar["value"] = 0
    time_label.config(text="")

def start_conversion_thread(directory):
    select_button.config(state=tk.DISABLED)
    conversion_thread = threading.Thread(target=convert_all_docs_in_directory, args=(directory,))
    conversion_thread.start()

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        start_conversion_thread(directory)

# Vytvoření hlavního okna
root = tk.Tk()
root.title("Hromadný převod DOC souborů na DOCX")
root.geometry("500x300")
root.resizable(False, False)

# Nastavení stylu pomocí ttk
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12), padding=10)

# Přidání titulku
title_label = tk.Label(root, text="DOC na DOCX Konvertor", font=('Helvetica', 16, 'bold'), pady=20)
title_label.pack()

# Přidání tlačítka pro výběr složky
select_button = ttk.Button(root, text="Vyber složku", command=select_directory)
select_button.pack(pady=10)

# Přidání ukazatele průběhu
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=20)

# Přidání textového labelu pro zobrazení aktuálně zpracovávaného souboru
progress_label = tk.Label(root, text="Čekání na výběr složky...", font=('Helvetica', 12))
progress_label.pack()

# Přidání textového labelu pro zobrazení zbývajícího času
time_label = tk.Label(root, text="", font=('Helvetica', 12))
time_label.pack()

# Přidání poděkování nebo kontaktní informace
footer_label = tk.Label(root, text="www.dtpstudio.eu", font=('Helvetica', 10), pady=10)
footer_label.pack(side=tk.BOTTOM)

# Spuštění hlavní smyčky Tkinteru
root.mainloop()

# Ukončení skriptu po zavření hlavního okna
root.destroy()
