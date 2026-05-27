"""
Interface graphique pour le contrôle d'exhaustivité SMR.
Construite avec Tkinter (standard Python) et utilise un Threading basique pour ne pas freezer l'UI
lors du traitement des fichiers Excel.

Même style et même structure que l'interface de la mission 1, mais simplifiée :
seulement 2 fichiers d'entrée (Orbis SMR + Hexagone SMR) au lieu de 7.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path
import controle_smr_deepseek


class AppControleSMR:
    """
    Classe englobant toute l'interface. On utilise une classe pour partager 
    l'état (les variables de chemins de fichiers, les labels d'état) entre toutes les fonctions.
    """
    def __init__(self, root):
        self.root = root  # Fenêtre principale (racine de l'arbre Tkinter)
        self.root.title("Controle d'exhaustivite SMR - DIM")

        # Dimensions par défaut et minimales pour éviter un layout cassé
        self.root.geometry("1000x500")
        self.root.minsize(800, 420)
        self.root.resizable(True, True)

        # Couleurs (identiques à la mission 1 pour la cohérence visuelle)
        bg_color = "#eef3f9"
        panel_color = "#ffffff"
        header_color = "#2F5496"
        btn_color = "#2F5496"
        btn_hover = "#1a3a6e"
        muted_text = "#5b6470"

        self.root.configure(bg=bg_color)

        # --- Personnalisation du style (Thème) ---
        # On utilise ttk (Themed Tkinter) pour avoir des composants un peu plus modernes.
        style = ttk.Style(self.root)
        try:
            # "clam" est un thème intégré qui fait moins "Windows 95" que le style par défaut.
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Configuration des couleurs pour nos composants
        style.configure("App.TFrame", background=bg_color)
        style.configure("Panel.TFrame", background=panel_color)
        style.configure("App.TLabel", background=bg_color, foreground="#17212b")
        style.configure("Muted.TLabel", background=bg_color, foreground=muted_text)
        style.configure("Panel.TLabel", background=panel_color, foreground="#17212b")
        style.configure("App.Horizontal.TProgressbar",
                        troughcolor="#d9e2ef", background=btn_color,
                        bordercolor="#d9e2ef", lightcolor=btn_color, darkcolor=btn_color)

        self.default_entry_width = 72
        self.browse_width = 4

        # --- Variables d'état Tkinter (StringVar) ---
        # Les StringVar sont liées aux widgets : quand le code modifie la variable,
        # l'interface se met à jour automatiquement.
        self.fichier_orbis = tk.StringVar()
        self.fichier_hexa = tk.StringVar()

        # Dossier d'export par défaut
        self.dossier_export = tk.StringVar(value=str(Path('./data_test/export_test').resolve()))

        # ==========================================
        # EN-TÊTE
        # ==========================================
        header_frame = tk.Frame(root, bg=header_color, height=72)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame, text="Controle d'exhaustivite SMR - DIM",
            font=("Segoe UI Semibold", 18), fg="white", bg=header_color
        ).pack(pady=18)

        # ==========================================
        # CORPS PRINCIPAL
        # ==========================================
        content_host = tk.Frame(root, bg=bg_color)
        content_host.pack(fill=tk.BOTH, expand=True)

        # --- Système de défilement (Scrollbar) ---
        # Tkinter ne sait pas faire défiler une Frame nativement.
        # L'astuce est de créer un Canvas (qui sait défiler) et d'y attacher notre Frame.
        self.canvas = tk.Canvas(content_host, bg=bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_host, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.main_frame = tk.Frame(self.canvas, bg=bg_color, padx=20, pady=18)
        self.main_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._ajuster_largeur_canvas)
        self.root.bind("<MouseWheel>", self._defiler_souris)

        # --- Barre d'info ---
        info_bar = tk.Frame(self.main_frame, bg=bg_color)
        info_bar.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        tk.Label(
            info_bar,
            text="Sélectionnez le fichier Orbis SMR et le fichier Hexagone SMR, puis lancez le traitement.",
            font=("Segoe UI", 9),
            bg=bg_color,
            fg=muted_text,
            anchor="w",
            justify="left"
        ).pack(fill=tk.X)

        # --- Panneau principal des fichiers ---
        files_panel = tk.Frame(self.main_frame, bg=panel_color, padx=16, pady=16, bd=1, relief="solid")
        files_panel.grid(row=1, column=0, sticky="nsew")

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)

        # Section Orbis SMR
        orbis_section = tk.LabelFrame(files_panel, text="Fichier Orbis SMR",
                                      bg=panel_color, fg="#17212b", padx=12, pady=12)
        orbis_section.grid(row=0, column=0, sticky="ew")
        orbis_section.columnconfigure(1, weight=1)
        self._creer_section_fichier(orbis_section, "Orbis SMR", self.fichier_orbis, 0)

        # Section Hexagone SMR
        hexa_section = tk.LabelFrame(files_panel, text="Fichier Hexagone SMR",
                                     bg=panel_color, fg="#17212b", padx=12, pady=12)
        hexa_section.grid(row=1, column=0, sticky="ew", pady=(14, 10))
        hexa_section.columnconfigure(1, weight=1)
        self._creer_section_fichier(hexa_section, "Hexagone SMR", self.fichier_hexa, 0)

        # Section Export
        export_section = tk.LabelFrame(files_panel, text="Export",
                                       bg=panel_color, fg="#17212b", padx=12, pady=12)
        export_section.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        export_section.columnconfigure(1, weight=1)
        self._creer_section_export(export_section, 0)

        files_panel.columnconfigure(0, weight=1)

        # ==========================================
        # BARRE DE PROGRESSION + BOUTON
        # ==========================================
        bottom_frame = tk.Frame(root, bg=bg_color, padx=20, pady=12)
        bottom_frame.pack(fill=tk.X)

        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate', length=360,
                                        style="App.Horizontal.TProgressbar")
        self.progress.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True)

        self.status_label = tk.Label(bottom_frame, text="Pret.", font=("Segoe UI", 9),
                                     bg=bg_color, fg="#555")
        self.status_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

        self.btn_lancer = tk.Button(
            bottom_frame, text="Lancer le traitement", font=("Segoe UI", 11, "bold"),
            bg=btn_color, fg="white", activebackground=btn_hover, activeforeground="white",
            relief="flat", padx=15, pady=5, cursor="hand2",
            command=self._lancer_traitement
        )
        self.btn_lancer.pack(side=tk.RIGHT)

    def _creer_section_fichier(self, parent, label, var, row):
        """
        Fonction utilitaire pour créer une ligne de sélection de fichier :
        Label + champ texte (readonly) + bouton '...'
        """
        bg_color = parent.cget("bg")
        tk.Label(parent, text=label, font=("Segoe UI", 9, "bold"),
                 bg=bg_color, anchor="w").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))

        # L'Entry est en "readonly" pour forcer l'utilisateur à utiliser le bouton
        tk.Entry(parent, textvariable=var, width=self.default_entry_width,
                 font=("Segoe UI", 9), state="readonly").grid(row=row, column=1, sticky="ew", padx=5, pady=6)

        tk.Button(parent, text="...", width=self.browse_width,
                  command=lambda: self._choisir_fichier(var)).grid(row=row, column=2, pady=6, padx=(6, 0))

    def _creer_section_export(self, parent, row):
        """Crée la ligne de sélection du dossier d'export."""
        bg_color = parent.cget("bg")
        tk.Label(parent, text="Dossier export", font=("Segoe UI", 9, "bold"),
                 bg=bg_color, anchor="w").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
        tk.Entry(parent, textvariable=self.dossier_export, width=self.default_entry_width,
                 font=("Segoe UI", 9), state="readonly").grid(row=row, column=1, sticky="ew", padx=5, pady=6)
        tk.Button(parent, text="...", width=self.browse_width,
                  command=lambda: self._choisir_dossier(self.dossier_export)).grid(row=row, column=2, pady=6, padx=(6, 0))

    def _choisir_fichier(self, var):
        """Ouvre une boîte de dialogue pour sélectionner un fichier Excel."""
        fichier = filedialog.askopenfilename(
            title="Selectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if fichier:
            var.set(fichier)

    def _choisir_dossier(self, var):
        """Ouvre une boîte de dialogue pour sélectionner un dossier."""
        dossier = filedialog.askdirectory(title="Selectionner le dossier d'export")
        if dossier:
            var.set(dossier)

    def _ajuster_largeur_canvas(self, event):
        """Ajuste la largeur du contenu au Canvas quand on redimensionne la fenêtre."""
        self.canvas.itemconfigure(self.main_window, width=event.width)

    def _defiler_souris(self, event):
        """Gère le scroll à la molette de souris."""
        if self.canvas.winfo_height() < self.main_frame.winfo_reqheight():
            self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def _lancer_traitement(self):
        """
        Vérifie que les fichiers sont bien sélectionnés, puis lance le traitement
        dans un thread séparé pour ne pas bloquer l'interface.
        """
        # Vérifications
        if not self.fichier_orbis.get():
            messagebox.showwarning("Attention", "Veuillez selectionner le fichier Orbis SMR.")
            return

        if not self.fichier_hexa.get():
            messagebox.showwarning("Attention", "Veuillez selectionner le fichier Hexagone SMR.")
            return

        # Désactivation du bouton et démarrage de la barre de progression
        self.btn_lancer.config(state="disabled")
        self.progress.start(15)
        self.status_label.config(text="Traitement en cours...")
        self.root.update_idletasks()

        # --- THREADING ---
        # On délègue l'exécution à un thread secondaire pour ne pas bloquer l'UI.
        # daemon=True permet au thread de s'arrêter proprement si on ferme la fenêtre.
        thread = threading.Thread(target=self._executer_script, daemon=True)
        thread.start()

    def _executer_script(self):
        """
        Méthode exécutée dans un thread secondaire.
        Fait l'interface entre l'UI et la logique métier (controle_smr_deepseek).
        """
        try:
            import controle_smr_deepseek

            orbis_path = self.fichier_orbis.get()
            hexa_path = self.fichier_hexa.get()
            export_dir = self.dossier_export.get()

            # Appel de la fonction principale de la logique métier
            controle_smr_deepseek.lancer_controle_smr(
                orbis_path=orbis_path,
                hexa_path=hexa_path,
                export_dir=export_dir
            )

            # Succès → notification sur le thread principal
            # IMPORTANT : on ne modifie jamais l'UI Tkinter depuis un thread secondaire.
            # root.after(0, func) met la fonction dans la file d'attente du thread principal.
            self.root.after(0, self._traitement_termine, True, "")
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            self.root.after(0, self._traitement_termine, False, error_msg)

    def _traitement_termine(self, success, error_msg):
        """Callback appelé sur le thread principal quand le traitement est fini."""
        self.progress.stop()
        self.btn_lancer.config(state="normal")

        if success:
            self.status_label.config(text="Traitement termine avec succes !", fg="#2e7d32")
            messagebox.showinfo("Succes",
                                f"Traitement termine !\n\nFichiers generes dans :\n{self.dossier_export.get()}")
        else:
            self.status_label.config(text="Erreur lors du traitement.", fg="#c62828")
            short_error = error_msg[-500:] if len(error_msg) > 500 else error_msg
            messagebox.showerror("Erreur", f"Le traitement a echoue :\n\n{short_error}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppControleSMR(root)
    root.mainloop()
