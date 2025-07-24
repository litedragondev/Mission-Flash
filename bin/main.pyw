import pyautogui
import keyboard
import time
import pygetwindow as gw
import tkinter as tk
from tkinter import messagebox, ttk
import json
import os
import threading
import pyperclip

class MacroCalibration:
    def __init__(self):
        self.config_file = "calibration_config.json"
        self.coordinates = {}
        self.load_config()
        
    def load_config(self):
        """Charge la configuration sauvegardée ou crée un fichier vide"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    self.coordinates = json.load(f)
            except:
                self.coordinates = {}
        else:
            # Fichier inexistant : création
            self.coordinates = {}
            self.save_config()
            messagebox.showinfo("Configuration", "Fichier de calibration créé automatiquement.")

    def save_config(self):
        """Sauvegarde la configuration"""
        with open(self.config_file, 'w') as f:
            json.dump(self.coordinates, f, indent=2)
    
    def capture_coordinate(self, name, description):
        """Capture une coordonnée avec le curseur"""
        def on_click():
            x, y = pyautogui.position()
            self.coordinates[name] = {"x": x, "y": y, "description": description}
            print(f"Coordonnée '{name}' capturée: ({x}, {y})")
            return False
        
        print(f"Positionnez votre curseur sur: {description}")
        print("Appuyez sur CTRL pour capturer la position...")
        
        keyboard.wait('ctrl')
        on_click()

class MacroApp:
    def __init__(self):
        self.calibration = MacroCalibration()
        self.calibrating = False
        self.calibration_index = 0
        self.setup_gui()
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Saisie Automatique Devis Mission")
        self.root.geometry("600x800")
        
        # Version container
        version_frame = tk.Frame(self.root, bg="#f0f0f0")
        version_frame.pack(fill='x', padx=5, pady=2)
        tk.Label(version_frame, text="v0.4.0", bg="#f0f0f0", fg="#666", font=("Arial", 8)).pack(side='right')
        
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.setup_devis_tab(notebook)
        
    def validate_input(self, char):
        return char in '0123456789-/'

    def validate_num(self, char):
        return char in '0123456789-/azertyuiopqsdfghjklmwxcvbnAZERTYUIOPQSDFGHJKLMWXCVBN'
    
    def validate_montant(self, char):
        return char in '0123456789-/,'
    
    def setup_devis_tab(self, notebook):
        self.devis_frame = ttk.Frame(notebook)
        notebook.add(self.devis_frame, text="Devis")

        vcmd_num = (self.root.register(self.validate_num), '%S')
        vcmd = (self.root.register(self.validate_input), '%S')
        vcmd_montant = (self.root.register(self.validate_montant), '%S')

        windowName = tk.StringVar(value="Alteva.MissionOne")
        
        # Calibration button
        self.calibrer_button = tk.Button(self.devis_frame, text="Calibrer", command=self.toggle_calibration,
                 bg="orange", fg="white", font=("Arial", 10, "bold"))
        self.calibrer_button.pack(pady=5)
        
        # Calibration section (initially hidden)
        self.calibration_frame = tk.Frame(self.devis_frame)
        self.calibration_visible = False
        
        instructions = tk.Text(self.calibration_frame, height=6, width=70, wrap=tk.WORD)
        instructions.pack(padx=10, pady=5)
        instructions.insert(tk.END, """CALIBRATION: Ouvrez Mission One dans l'onglet Devis, cliquez 'Démarrer la calibration', positionnez le curseur sur chaque élément et appuyez CTRL.""")
        instructions.config(state=tk.DISABLED)
        
        tk.Button(self.calibration_frame, text="Démarrer la calibration", 
                 command=self.start_calibration, bg="red", fg="white").pack(pady=5)
        
        self.calibration_status = tk.Label(self.calibration_frame, text="", fg="red")
        self.calibration_status.pack(pady=2)
        
        self.coord_listbox = tk.Listbox(self.calibration_frame, height=8, width=70)
        self.coord_listbox.pack(padx=10, pady=5)
        
        button_frame = tk.Frame(self.calibration_frame)
        button_frame.pack(pady=5)
        
        tk.Button(button_frame, text="Reset", command=self.reset_calibration).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="Supprimer", command=self.delete_selected_coord).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="Test", command=self.quick_test).pack(side=tk.LEFT, padx=2)
        
        self.update_coord_list()
        
        # Main form (in a separate frame for easy hide/show)
        self.main_form_frame = tk.Frame(self.devis_frame)
        self.main_form_frame.pack(fill='both', expand=True)
        
        ttk.Separator(self.main_form_frame, orient='horizontal').pack(fill='x', pady=10)
        
        tk.Label(self.main_form_frame, text="Nom de la fenêtre :").pack(pady=2)
        self.entry_fenetre = tk.Entry(self.main_form_frame, textvariable=windowName, width=50)
        self.entry_fenetre.pack(padx=10, pady=5)
        
        tk.Label(self.main_form_frame, text="Numéro de commande :").pack(pady=2)
        self.entry_num_commande = tk.Entry(self.main_form_frame, width=50, validate='key', validatecommand=vcmd_num)
        self.entry_num_commande.pack(padx=10, pady=5)
        
        tk.Label(self.main_form_frame, text="Date (JJ/MM/AAAA slashs optionnels) :").pack(pady=2)
        self.entry_date = tk.Entry(self.main_form_frame, width=50, validate='key', validatecommand=vcmd_montant)
        self.entry_date.pack(padx=10, pady=5)
        
        tk.Label(self.main_form_frame, text="Remarque (multi-ligne) :").pack(pady=2)
        self.text_remarque = tk.Text(self.main_form_frame, height=5, width=50)
        self.text_remarque.pack(padx=10, pady=5)
        
        self.cocher_case_var = tk.BooleanVar()
        chk1 = tk.Checkbutton(self.main_form_frame, text="Cocher une case dans le formulaire", 
                             variable=self.cocher_case_var)
        chk1.pack(pady=5)
        
        self.montant_var = tk.BooleanVar()
        chk2 = tk.Checkbutton(self.main_form_frame, text="Ajouter le champ 'Montant de la commande'", 
                             variable=self.montant_var, command=self.toggle_montant_field)
        chk2.pack(pady=5)
        
        self.entry_montant = tk.Entry(self.main_form_frame, width=50, validate='key', validatecommand=vcmd_montant)
        
        tk.Button(self.main_form_frame, text="Lancer la macro", command=self.lancement_macro,
                 bg="green", fg="white", font=("Arial", 12, "bold")).pack(pady=15)
        
        self.status_label = tk.Label(self.main_form_frame, text="", fg="blue")
        self.status_label.pack(pady=5)
    
    def toggle_calibration(self):
        if self.calibrating:
            self.calibrating = False
            self.calibration_status.config(text="Calibration annulée", fg="red")
            self.calibrer_button.config(text="Retour", state="normal")
            return
            
        if self.calibration_visible:
            self.calibration_frame.pack_forget()
            self.main_form_frame.pack(fill='both', expand=True)
            self.calibrer_button.config(text="Calibrer")
            self.calibration_visible = False
        else:
            self.main_form_frame.pack_forget()
            self.calibration_frame.pack(fill='x', padx=10, pady=5, after=self.devis_frame.winfo_children()[0])
            self.calibrer_button.config(text="Retour")
            self.calibration_visible = True
    
    def start_calibration(self):
        def calibration_thread():
            elements_to_calibrate = [
                ("commandes_recues", "Bouton 'Commandes reçues'"),
                ("nouvelle_commande", "Bouton 'Ajouter nouvelle commande'"),
                ("num_commande", "Champ 'Numéro de commande'"),
                ("date", "Champ 'Date'"),
                ("case_a_cocher", "Case à cocher (optionnelle)"),
                ("montant", "Champ 'Montant' (optionnel)"),
                ("remarque", "Champ 'Remarque'"),
                ("validation", "Bouton 'Validation'")
            ]
            
            self.calibrating = True
            self.calibrer_button.config(text="Annuler", state="normal")
            
            for i in range(self.calibration_index, len(elements_to_calibrate)):
                if not self.calibrating:
                    break
                name, description = elements_to_calibrate[i]
                try:
                    self.calibration_status.config(text=f"Calibration en cours... Positionnez sur {description} puis pressez CTRL", fg="orange")
                    self.root.update()
                    
                    # Capture coordinate with interruption check
                    print(f"Positionnez votre curseur sur: {description}")
                    print("Appuyez sur CTRL pour capturer la position...")
                    
                    while self.calibrating:
                        if keyboard.is_pressed('ctrl'):
                            x, y = pyautogui.position()
                            self.calibration.coordinates[name] = {"x": x, "y": y, "description": description}
                            print(f"Coordonnée '{name}' capturée: ({x}, {y})")
                            break
                        time.sleep(0.1)
                    
                    if not self.calibrating:
                        break
                        
                    self.update_coord_list()
                    self.calibration_index = i + 1
                    time.sleep(0.25)
                except KeyboardInterrupt:
                    break
            
            if self.calibrating and self.calibration_index >= len(elements_to_calibrate):
                self.calibration.save_config()
                self.calibration_status.config(text="Calibration terminée!", fg="green")
                self.update_coord_list()
                self.calibration_index = 0
            
            self.calibrating = False
            self.calibrer_button.config(text="Retour", state="normal")
        
        thread = threading.Thread(target=calibration_thread)
        thread.daemon = True
        thread.start()
    
    def reset_calibration(self):
        self.calibrating = False
        self.calibration_index = 0
        self.calibration.coordinates = {}
        self.calibration.save_config()
        self.update_coord_list()
        self.calibration_status.config(text="Configuration supprimée. Prêt pour une nouvelle calibration.", fg="blue")
        self.calibrer_button.config(text="Retour", state="normal")
    
    def update_coord_list(self):
        self.coord_listbox.delete(0, tk.END)
        for name, data in self.calibration.coordinates.items():
            self.coord_listbox.insert(tk.END, 
                f"{name}: ({data['x']}, {data['y']}) - {data['description']}")
    
    def delete_selected_coord(self):
        selection = self.coord_listbox.curselection()
        if selection:
            index = selection[0]
            name = list(self.calibration.coordinates.keys())[index]
            del self.calibration.coordinates[name]
            self.calibration.save_config()
            self.update_coord_list()
    
    def quick_test(self):
        def test_thread():
            for name, data in self.calibration.coordinates.items():
                pyautogui.moveTo(data['x'], data['y'])
                time.sleep(1)
        
        thread = threading.Thread(target=test_thread)
        thread.daemon = True
        thread.start()
    
    def toggle_montant_field(self):
        if self.montant_var.get():
            self.entry_montant.pack(padx=10, pady=5, before=self.main_form_frame.winfo_children()[-2])
        else:
            self.entry_montant.pack_forget()
    
    def focus_fenetre(self, nom_fenetre):
        fenetres = gw.getWindowsWithTitle(nom_fenetre)
        if fenetres:
            fenetre = fenetres[0]
            fenetre.activate()
            time.sleep(1)
            return True
        else:
            return False
    
    def safe_click(self, coord_name):
        if coord_name in self.calibration.coordinates:
            coord = self.calibration.coordinates[coord_name]
            pyautogui.click(x=coord['x'], y=coord['y'])
            return True
        else:
            messagebox.showerror("Erreur de calibration", 
                               f"Coordonnée '{coord_name}' non configurée. "
                               "Veuillez refaire la calibration.")
            return False

    def lancement_macro(self):
        numCommande = self.entry_num_commande.get()
        date = self.entry_date.get()
        remarque = self.text_remarque.get("1.0", tk.END).strip()
        fenetre = self.entry_fenetre.get()
        montant = self.entry_montant.get() if self.montant_var.get() else None

        if not numCommande or not date or not remarque or not fenetre:
            messagebox.showwarning("Champs manquants",
                                   "Veuillez remplir tous les champs requis.")
            return

        if self.montant_var.get() and not montant:
            messagebox.showwarning("Champ manquant",
                                   "Veuillez saisir le montant de la commande.")
            return

        required_coords = ["commandes_recues", "nouvelle_commande", "num_commande",
                           "date", "remarque", "validation"]
        missing_coords = [coord for coord in required_coords
                          if coord not in self.calibration.coordinates]

        if missing_coords:
            messagebox.showerror("Calibration incomplète",
                                 f"Coordonnées manquantes: {', '.join(missing_coords)}\n"
                                 "Veuillez compléter la calibration.")
            return

        if not self.focus_fenetre(fenetre):
            messagebox.showerror("Erreur", f"Fenêtre contenant '{fenetre}' non trouvée.")
            return

        self.status_label.config(text="Exécution de la macro en cours...", fg="blue")
        self.root.update()

        try:
            time.sleep(0.5)

            if not self.safe_click("commandes_recues"):
                return
            time.sleep(0.25)

            if not self.safe_click("nouvelle_commande"):
                return
            time.sleep(0.25)

            if not self.safe_click("num_commande"):
                return
            time.sleep(0.25)
            print(f"Saisie du numéro de commande : {numCommande}")
            pyautogui.write(numCommande, interval=0.01)

            if not self.safe_click("date"):
                return
            time.sleep(0.25)
            print(f"Saisie de la date : {date}")
            pyautogui.write(date, interval=0.01)

            if self.cocher_case_var.get():
                if "case_a_cocher" in self.calibration.coordinates:
                    self.safe_click("case_a_cocher")
                    time.sleep(0.25)

            if self.montant_var.get() and montant:
                if "montant" in self.calibration.coordinates:
                    self.safe_click("montant")
                    time.sleep(0.25)
                    print(f"Saisie du montant : {montant}")
                    pyautogui.write(montant, interval=0.01)

            if not self.safe_click("remarque"):
                return
            time.sleep(0.25)
            print(f"Saisie de la remarque : {remarque}")
            pyautogui.write(remarque, interval=0.01)

            if not self.safe_click("validation"):
                return

            self.status_label.config(text="Macro exécutée avec succès!", fg="green")
            print("Fin de la Macro")

        except Exception as e:
            messagebox.showerror("Erreur d'exécution", f"Erreur lors de l'exécution: {str(e)}")
            self.status_label.config(text="Erreur lors de l'exécution", fg="red")

    def run(self):
        self.root.mainloop()

    def hideWidget(widget):
        widget.pack_forget()

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1
    
    app = MacroApp()
    app.run()
