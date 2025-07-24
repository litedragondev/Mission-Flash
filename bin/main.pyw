import pyautogui
import keyboard
import time
import pygetwindow as gw
import tkinter as tk
from tkinter import messagebox, ttk
import json
import os
import threading

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
        print("Appuyez sur CTRL+C pour capturer la position...")
        
        keyboard.wait('ctrl+c')
        on_click()

class MacroApp:
    def __init__(self):
        self.calibration = MacroCalibration()
        self.setup_gui()
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Saisie Automatique Devis Mission")
        self.root.geometry("600x800")
        
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.setup_devis_tab(notebook)
        
    def setup_devis_tab(self, notebook):
        self.devis_frame = ttk.Frame(notebook)
        notebook.add(self.devis_frame, text="Devis")
        
        # Calibration button
        tk.Button(self.devis_frame, text="Calibrer", command=self.toggle_calibration,
                 bg="orange", fg="white", font=("Arial", 10, "bold")).pack(pady=5)
        
        # Calibration section (initially hidden)
        self.calibration_frame = tk.Frame(self.devis_frame)
        self.calibration_visible = False
        
        instructions = tk.Text(self.calibration_frame, height=6, width=70, wrap=tk.WORD)
        instructions.pack(padx=10, pady=5)
        instructions.insert(tk.END, """CALIBRATION: Ouvrez l'application cible, cliquez 'Démarrer', positionnez le curseur sur chaque élément et appuyez CTRL+C.""")
        instructions.config(state=tk.DISABLED)
        
        tk.Button(self.calibration_frame, text="Démarrer la calibration", 
                 command=self.start_calibration, bg="red", fg="white").pack(pady=5)
        
        self.calibration_status = tk.Label(self.calibration_frame, text="", fg="red")
        self.calibration_status.pack(pady=2)
        
        self.coord_listbox = tk.Listbox(self.calibration_frame, height=8, width=70)
        self.coord_listbox.pack(padx=10, pady=5)
        
        button_frame = tk.Frame(self.calibration_frame)
        button_frame.pack(pady=5)
        
        tk.Button(button_frame, text="Actualiser", command=self.update_coord_list).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="Supprimer", command=self.delete_selected_coord).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="Test", command=self.quick_test).pack(side=tk.LEFT, padx=2)
        
        self.update_coord_list()
        
        # Separator
        ttk.Separator(self.devis_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Main form
        tk.Label(self.devis_frame, text="Nom de la fenêtre :").pack(pady=2)
        self.entry_fenetre = tk.Entry(self.devis_frame, width=50)
        self.entry_fenetre.pack(padx=10, pady=5)
        
        tk.Label(self.devis_frame, text="Numéro de commande :").pack(pady=2)
        self.entry_num_commande = tk.Entry(self.devis_frame, width=50)
        self.entry_num_commande.pack(padx=10, pady=5)
        
        tk.Label(self.devis_frame, text="Date :").pack(pady=2)
        self.entry_date = tk.Entry(self.devis_frame, width=50)
        self.entry_date.pack(padx=10, pady=5)
        
        tk.Label(self.devis_frame, text="Remarque (multi-ligne) :").pack(pady=2)
        self.text_remarque = tk.Text(self.devis_frame, height=5, width=50)
        self.text_remarque.pack(padx=10, pady=5)
        
        self.cocher_case_var = tk.BooleanVar()
        chk1 = tk.Checkbutton(self.devis_frame, text="Cocher une case dans le formulaire", 
                             variable=self.cocher_case_var)
        chk1.pack(pady=5)
        
        self.montant_var = tk.BooleanVar()
        chk2 = tk.Checkbutton(self.devis_frame, text="Ajouter le champ 'Montant de la commande'", 
                             variable=self.montant_var, command=self.toggle_montant_field)
        chk2.pack(pady=5)
        
        self.entry_montant = tk.Entry(self.devis_frame, width=50)
        
        tk.Button(self.devis_frame, text="Lancer la macro", command=self.lancement_macro,
                 bg="green", fg="white", font=("Arial", 12, "bold")).pack(pady=15)
        
        self.status_label = tk.Label(self.devis_frame, text="", fg="blue")
        self.status_label.pack(pady=5)
    
    def toggle_calibration(self):
        if self.calibration_visible:
            self.calibration_frame.pack_forget()
            self.calibration_visible = False
        else:
            self.calibration_frame.pack(fill='x', padx=10, pady=5, after=self.devis_frame.winfo_children()[0])
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
            
            self.calibration_status.config(text="Calibration en cours...", fg="orange")
            self.root.update()
            
            for name, description in elements_to_calibrate:
                try:
                    self.calibration.capture_coordinate(name, description)
                    self.update_coord_list()
                    time.sleep(0.5)
                except KeyboardInterrupt:
                    break
            
            self.calibration.save_config()
            self.calibration_status.config(text="Calibration terminée!", fg="green")
            self.update_coord_list()
        
        thread = threading.Thread(target=calibration_thread)
        thread.daemon = True
        thread.start()
    
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
            self.entry_montant.pack(padx=10, pady=5)
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
            pyautogui.write(numCommande)
            
            if not self.safe_click("date"):
                return
            time.sleep(0.25)
            pyautogui.write(date)
            
            if self.cocher_case_var.get():
                if "case_a_cocher" in self.calibration.coordinates:
                    self.safe_click("case_a_cocher")
                    time.sleep(0.25)
            
            if self.montant_var.get() and montant:
                if "montant" in self.calibration.coordinates:
                    self.safe_click("montant")
                    time.sleep(0.25)
                    pyautogui.write(montant)
            
            if not self.safe_click("remarque"):
                return
            time.sleep(0.25)
            for i, ligne in enumerate(remarque.splitlines()):
                pyautogui.write(ligne)
                if i < len(remarque.splitlines()) - 1:
                    pyautogui.press('enter')
            
            if not self.safe_click("validation"):
                return
            
            self.status_label.config(text="Macro exécutée avec succès!", fg="green")
            print("Fin de la Macro")
            
        except Exception as e:
            messagebox.showerror("Erreur d'exécution", f"Erreur lors de l'exécution: {str(e)}")
            self.status_label.config(text="Erreur lors de l'exécution", fg="red")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1
    
    app = MacroApp()
    app.run()
