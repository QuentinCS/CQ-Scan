# -*- coding: utf-8 -*-
"""
Created on Mon Mar 17 15:22:21 2025

@author: 147032
"""

# Installations package: pip install pydicom ttkthemes

import matplotlib.pyplot as plt
from datetime import datetime
import numpy as np
import pandas as pd
import pydicom
import math
import os
from tkinter import ttk, filedialog, scrolledtext, Frame
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from ttkthemes import ThemedTk
import matplotlib.patches as mpatches


# Classe pour l'analyse de la qualité image du CQI scanner 
class CT_quality:
    def __init__(self, root):
        self.root = root
        self.root.title("Qualité image CT") # Titre de l'interface 
        self.root.state('zoomed') # Taille de l'interface (fullscreen)
        
        # Dictionnaire pour les marges internes du fantôme par défaut pour les différents constructeurs  
        self.phantom_dict = {
            "Philips" : 6,
            "GE MEDICAL SYSTEMS" : 7,
            "Canon Medical Systems": 5,
            "SIEMENS": 5,
            }
        
        # Initialisation des variables 
        self.num_image = 0
        self.slice = 0
        self.total_slice_number = 0
        self.seuil = -500 # Seuillage en UH pour la détection du contour 
        self.internal_margin = 0 # Epaisseur de la paroi du fantôme
        self.external_roi_offset = 20 # Décalage des ROIs externes par rapport au contour interne du fantôme
        self.dicom_data = None  # Stocke les données DICOM chargées
        self.dicom_files = []  # Liste pour enregistrer les fichiers Dicom
        self.canvas = None  # Stocke le canevas d'affichage de l'image
        self.dicom_data = None  # Stocke les données DICOM chargées
        self.central_roi_s = 0.4 # Taille de la roi centrale 
        self.external_roi_s = 0.1 # Taille des rois latérales 
        
        # Variables pour stocker les valeurs des mesures 
        self.n_ct_center = None
        self.n_ct_lateral_N = None
        self.n_ct_lateral_S = None
        self.n_ct_lateral_E = None 
        self.n_ct_lateral_W = None 
        self.std_ct_center = None
        self.std_ct_lateral_N = None
        self.std_ct_lateral_S = None
        self.std_ct_lateral_E = None 
        self.std_ct_lateral_W = None 
        self.unif = None    
        
        # Dictionnaires et listes pour organiser les données
        self.dicom_files_by_directory = {}
        self.directories_with_dicom = []
        self.dicom_images = {}
        self.dicom_metadata = {}
        self.images_names = []
        self.central_roi = {}
        self.external_roi_N = {}
        self.external_roi_S = {}
        self.external_roi_E = {}
        self.external_roi_W = {}
        
        # Définir la fenêtre de visualisation
        self.WW = 80  # Largeur de la fenêtre
        self.WL = 0   # Niveau de la fenêtre

        # Appliquer la fenêtre de contraste
        self.min_val = self.WL - (self.WW / 2)
        self.max_val = self.WL + (self.WW / 2)
     
        ########################## Création de l'interface Tkinter ########################################   
     
        self.root.columnconfigure(0, weight=0)  # Colonne des boutons (fixe)
        self.root.columnconfigure(1, weight=1)  # Zone de texte prend plus d'espace
        self.root.columnconfigure(2, weight=1)  # Zone de texte prend plus d'espace
        self.root.rowconfigure(1, weight=1)      # Figure prend plus d'espace
        
        # ---- FRAME pour contenir les zones de texte ----
        self.text_frame = ttk.Frame(root)
        self.text_frame.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Configuration pour que les deux colonnes du frame restent équilibrées
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.columnconfigure(1, weight=1)
     
        # Création d'un Frame pour les boutons 
        self.button_frame = ttk.Frame(root)
        self.button_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nw") 
        
        # Création d'une zone de texte pour afficher les métadonnées
        self.text_info_area = scrolledtext.ScrolledText(self.text_frame, width=15, height=16)
        self.text_info_area.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        # Création d'une zone de texte pour afficher les résultats
        self.result_area = scrolledtext.ScrolledText(self.text_frame, width=35, height=16)
        self.result_area.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
        
        # ---- Label au-dessus de la première zone de texte ----
        self.label1 = ttk.Label(self.text_frame, text="Données")
        self.label1.grid(row=0, column=0, padx=10, pady=(0,5), sticky="w")

        # ---- Label au-dessus de la deuxième zone de texte ----
        self.label2 = ttk.Label(self.text_frame, text="Résultats")
        self.label2.grid(row=0, column=1, padx=10, pady=(0,5), sticky="w")
        
        # Frame pour afficher l'image DICOM
        self.image_frame = Frame(root)
        self.image_frame.grid(row=1, column=0, columnspan=2, pady=5)
     
        # Création d'un bouton pour sélectionner les fichiers DICOM avec les images 
        self.load_button = ttk.Button(self.button_frame, text="Charger le dossier", command=self.load_dicom)
        self.load_button.grid(pady=5)
        
        # ---- Champ de saisie pour entrer une valeur ----
        self.entry_label_slice = ttk.Label(self.button_frame, text="Coupe (facultatif, par défaut centrale):")
        self.entry_label_slice.grid(row=1, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # Création d'un champ de saisie (Entry) pour l'utilisateur
        self.entry_value_slice = ttk.Entry(self.button_frame)
        self.entry_value_slice.grid(row=2, column=0, padx=10, pady=(5, 0), sticky="ew")
        
        # Création d'un bouton pour récupérer la valeur entrée par l'utilisateur
        self.get_value_button_slice = ttk.Button(self.button_frame, text="Valider", command=self.get_slice)
        self.get_value_button_slice.grid(row=2, column=1, padx=10, pady=5)
             
        # ---- Champ de saisie pour entrer une valeur ----
        self.entry_label_internal_margin = ttk.Label(self.button_frame, text="Epaisseur paroi (mm, facultatif):")
        self.entry_label_internal_margin.grid(row=3, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # Création d'un champ de saisie (Entry) pour l'utilisateur
        self.entry_value_internal_margin = ttk.Entry(self.button_frame)
        self.entry_value_internal_margin.grid(row=4, column=0, padx=10, pady=(5, 0), sticky="ew")
        
        # Création d'un bouton pour récupérer la valeur entrée par l'utilisateur
        self.get_value_button_internal_margin = ttk.Button(self.button_frame, text="Valider", command=self.get_internal_margin)
        self.get_value_button_internal_margin.grid(row=4, column=1, padx=10, pady=5)
               
        # ---- Champ de saisie pour entrer une valeur ----
        self.entry_label = ttk.Label(self.button_frame, text="Offset ROIs (mm, par défaut 20 mm):")
        self.entry_label.grid(row=5, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # Création d'un champ de saisie (Entry) pour l'utilisateur
        self.entry_value_offset = ttk.Entry(self.button_frame)
        self.entry_value_offset.grid(row=6, column=0, padx=10, pady=(5, 0), sticky="ew")
        
        # Création d'un bouton pour récupérer la valeur entrée par l'utilisateur
        self.get_value_button_offset = ttk.Button(self.button_frame, text="Valider", command=self.get_external_roi_offset)
        self.get_value_button_offset.grid(row=6, column=1, padx=10, pady=5)
               
        # Création d'un bouton pour lancer la détection des contours 
        self.analyze_button = ttk.Button(self.button_frame, text="Mesurer", command=self.analyze_all_images)
        self.analyze_button.grid(row=7, column=0, pady=5)
        
        # Création d'un bouton pour sauvegarder les données
        self.save_button = ttk.Button(self.button_frame, text="Sauvegarder", command=self.save_results)
        self.save_button.grid(row=7, column=1, pady=5)
        
        # Création d'un bouton pour réinitialiser l'interface
        self.reinitialize_button = ttk.Button(self.button_frame, text="Réinitialiser", command=self.reinitialize)
        self.reinitialize_button.grid(row=8, column=1, pady=5)
           

    ############################# Déclaration des fonctions de la classe #################################

    # Fonction pour sélectionner l'image Dicom de la coupe centrale (de l'image pas du fantôme) des images Dicom 
    # En assumant un fichier par coupe
    # Condition permettant de recharger l'image en sélectionnant le numéro d'une coupe 
    def load_dicom(self):
        if self.slice == 0:
            folder_path = filedialog.askdirectory()
            
            if folder_path:
                for root, dirs, files in os.walk(folder_path):
                    self.dicom_files = [file for file in files]# if file.endswith('.dcm')]
                    if self.dicom_files:
                        self.directory_name = os.path.basename(root)
                        self.dicom_files_by_directory[self.directory_name] = self.dicom_files
                        self.directories_with_dicom.append(root)
                        self.images_names.append(self.directory_name)
            
                # Ouvre et stocke chaque image DICOM dans une variable distincte
                for self.directory in self.directories_with_dicom:
                    self.dicom_files = self.dicom_files_by_directory[os.path.basename(self.directory)]
                    for self.dicom_file in self.dicom_files:
                        self.dicom_path = os.path.join(self.directory, self.dicom_file)
                        try:
                            self.dicom_data = pydicom.dcmread(self.dicom_path)
                            self.total_slice_number = int(len(self.dicom_files))
                            if self.dicom_data[0x0020, 0x0013].value == len(self.dicom_files) // 2:    
                                # Stocke l'image dans un dictionnaire avec un nom unique
                                self.image_name = f"image_{self.dicom_file}"
                                self.set_dicom_tag()
                                self.dicom_images[self.image_name] = (
                                    self.dicom_data.pixel_array * self.dicom_data.RescaleSlope + self.dicom_data.RescaleIntercept
                                    )
                                if self.num_image == 0: 
                                    self.device = self.dicom_metadata.get(self.image_name, {}).get('Device', 'Inconnu')
                                    self.institution = self.dicom_metadata.get(self.image_name, {}).get('Institution', 'Inconnu')
                                    self.display_info()
                                    self.find_contour()

                                self.num_image += 1
                                break
                        except Exception as e:
                            self.text_info_area.insert("end", f"Erreur lors de la lecture du fichier {self.dicom_path}: {e}")

        else:
                # Réinitialisation des variables 
                self.dicom_images.clear()
                self.dicom_metadata.clear()
                self.num_image = 0
                # Ouvre et stocke chaque image DICOM dans une variable distincte
                for self.directory in self.directories_with_dicom:
                    self.dicom_files = self.dicom_files_by_directory[os.path.basename(self.directory)]
                    for self.dicom_file in self.dicom_files:
                        self.dicom_path = os.path.join(self.directory, self.dicom_file)
                        try:
                            self.dicom_data = pydicom.dcmread(self.dicom_path)
                            slice_n = self.dicom_data[0x0020, 0x0013].value
                            if slice_n == self.slice:  
                                # Stocke l'image dans un dictionnaire avec un nom unique
                                self.image_name = f"image_{self.dicom_file}"
                                self.set_dicom_tag()
                                self.dicom_images[self.image_name] = (
                                    self.dicom_data.pixel_array * self.dicom_data.RescaleSlope + self.dicom_data.RescaleIntercept
                                    )
                                if self.num_image == 0: 
                                    self.device = self.dicom_metadata.get(self.image_name, {}).get('Device', 'Inconnu')
                                    self.institution = self.dicom_metadata.get(self.image_name, {}).get('Institution', 'Inconnu')
                                    self.display_info()
                                    self.find_contour()

                                self.num_image += 1
                                break
                        except Exception as e:
                            self.text_info_area.insert("end", f"Erreur lors de la lecture du fichier {self.dicom_path}: {e}")

        # Appliquer les ROis détecté sur la première images sur toutes les images 
        self.apply_rois_to_all_images()
        
        # Afficher les images avec les ROIs 
        self.display_image_rois()

    # Fonction pour afficher les informations Dicom
    def display_info(self):
        self.text_info_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.text_info_area.insert("end", f'{len(self.images_names)-1} dossiers trouvés: {self.images_names[1:]} \n\n')
        self.text_info_area.insert("end", f'Site: {self.dicom_metadata[self.image_name]["Institution"]} \nConstructeur: {self.dicom_metadata[self.image_name]["Manufacturer"]} \nModèle: {self.dicom_metadata[self.image_name]["Device"]} \nNom: {self.dicom_metadata[self.image_name]["Name"]} \nCoupe: {self.dicom_metadata[self.image_name]["Slice_number"]} \nDate de mesure: {self.dicom_metadata[self.image_name]["Date"]} \n\nEpaisseur paroi: {self.internal_margin} mm \nOffset ROI latérales: {self.external_roi_offset} mm \n')

    # Fonction pour afficher les données Dicom
    def display_dicom_tag(self): 
        self.text_area.delete("1.0", "end")  # Efface le texte précédent
        for elem in self.dicom_data:
            self.text_area.insert("end", f"{elem}\n")

    # Fonction pour lire et récupérer les informations importantes des tags Dicom 
    def set_dicom_tag(self):
        # Extrait les champs DICOM souhaités
        metadata = {
            "Device": self.dicom_data[0x0008, 0x1090].value,
            "Manufacturer": self.dicom_data[0x0008, 0x0070].value,
            "Institution": self.dicom_data[0x0008, 0x0080].value,
            "Name": self.dicom_data[0x0008, 0x1010].value,
            "Tension": self.dicom_data[0x0018, 0x0060].value,
            "Serie": self.dicom_data[0x008,0x103E].value,
            "Total Collimation": self.dicom_data[0x0018, 0x9307].value,
            "Single Collimation": self.dicom_data[0x0018, 0x9306].value,
            "Date": datetime.strptime(self.dicom_data[0x0008, 0x0020].value, "%Y%m%d").strftime("%d/%m/%Y"),
            "Size_x": self.dicom_data[0x0028, 0x0011].value,
            "Size_y": self.dicom_data[0x0028, 0x0010].value,
            "Pixel_spacing": self.dicom_data[0x0028, 0x0030].value,
            "Slice_thickness": self.dicom_data[0x0018, 0x0050].value,
            "Slice_spacing": self.dicom_data[0x0018, 0x0088].value,
            "Slice_number": self.dicom_data[0x0020, 0x0013].value,
        }
        
        self.pixel_spacing = metadata["Pixel_spacing"]
        self.size_x = metadata["Size_x"]
        self.size_y = metadata["Size_y"]
        
        # Stocke les métadonnées dans un dictionnaire avec un nom unique et un dataframe
        self.dicom_metadata[self.image_name] = metadata       
        self.df_data = pd.DataFrame(self.dicom_metadata)
        
        # Utiliser la marge présente dans le dictionnaire si le constructeur existe 
        if str(self.dicom_metadata[self.image_name]["Manufacturer"]) in self.phantom_dict.keys(): 
            self.internal_margin = self.phantom_dict[str(self.dicom_metadata[self.image_name]["Manufacturer"])]
        else:
            self.internal_margin = 0
    # Fonction pour trier les images par ordre de tension
    def image_sort(self):
        image_tension_list = [
            (image_name, self.dicom_metadata.get(image_name, {}).get('Tension', 'Inconnu'))
            for image_name in self.dicom_images.keys()
            ]

        # Convertir les tensions en float pour le tri, en ignorant les valeurs non numériques
        image_tension_list = [
            (image_name, float(tension)) if tension != 'Inconnu' else (image_name, float('inf'))
            for image_name, tension in image_tension_list
            ]

        # Trier par tension croissante
        image_tension_list.sort(key=lambda x: x[1])
        return image_tension_list
    
    # Fonction pour afficher l'image sélectionnée 
    def display_image_rois(self):
        if self.dicom_data is not None and hasattr(self.dicom_data, 'pixel_array'):
            # Créer une figure Matplotlib
            if not self.dicom_images:
                self.text_info_area.insert("end", "Aucune image DICOM à afficher.")
                return
                
            num_images = len(self.dicom_images)
            
            # Adaptation automatique des figures en fonction du nombre d'images à afficher 
            num_cols = min(4, num_images)
            num_rows = math.ceil(num_images / num_cols)
            
            fig_width = 5 * num_cols
            fig_height = 5 * num_rows
            
            # S'assurer qu'il y a au moins une image avant de créer la figure
            if num_images == 0:
                self.text_info_area.insert("end", "Erreur : aucune image à afficher.")
                return

            fig, axes = plt.subplots(num_rows, num_cols, figsize=(fig_width, fig_height))
            plt.rc('font', size=10)
            self.legend_patches = [
                mpatches.Patch(color='red', label='Contour fantôme détecté'),
                mpatches.Patch(color='orange', label='Contour interne'),
                mpatches.Patch(color='blue', label='Contours ROIs'),
                plt.Line2D([], [], color='green', marker='o', linestyle='None', markersize=3, label='Centre')
            ]
            
            image_list = self.image_sort()

            # Gérer le cas où il n'y a qu'une seule image
            if num_images == 1:
                axes = [axes]   
            else:
                axes = axes.flatten()
            
            # Affichage des images
            for i, (image_name, _) in enumerate(image_list):
                if i >= len(axes):
                    break  # Sécurité si trop d'images vs axes
                serie = self.dicom_metadata.get(image_name, {}).get('Serie', 'Inconnu')
                
                ax = axes[i]
                pixel_array = self.dicom_images[image_name]
                ax.imshow(pixel_array, cmap='gray', vmin=self.min_val, vmax=self.max_val)
                ax.scatter(self.phantom_center[1], self.phantom_center[0], color='green', label='Centre')
                ax.contour(self.phantom, colors='red')
                ax.contour(self.internal_phantom, colors='orange')
                ax.contour(self.mask_central_roi, colors='blue', linewidths=1)
                ax.contour(self.mask_external_N_roi, colors='blue', linewidths=1)
                ax.contour(self.mask_external_S_roi, colors='blue', linewidths=1)
                ax.contour(self.mask_external_E_roi, colors='blue', linewidths=1)
                ax.contour(self.mask_external_W_roi, colors='blue', linewidths=1)
                ax.text(0.05, 0.05, f'{serie}', color='white', fontsize=7, fontweight='bold', ha='left', va='top')
                ax.legend(handles=self.legend_patches, fontsize=5)
                ax.axis('off')

            # Masquer les axes vides restants
            for j in range(i + 1, len(axes)):
                axes[j].axis('off')

            plt.tight_layout()
            
            # Nettoyer l'affichage précédent
            for widget in self.image_frame.winfo_children():
                widget.destroy()

            self.canvas = FigureCanvasTkAgg(fig, master=self.image_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill="both", expand=True)

        else:
            self.text_info_area.insert("end", "Aucune image trouvée dans ce fichier DICOM.\n")
           
    # Fonction pour identifier les contours du fantôme dans l'image à partir d'un seuillage sur les UH "self.seuil" 
    def find_contour(self):
        # Créer un masque binaire à partir des seuils
        self.mask = self.dicom_images[self.image_name] >= self.seuil

        # Récupérer les pixels à l'intérieur du mask
        self.active_pixels = np.argwhere(self.mask)
        
        # Calcul des distances entre les pixels du mask et le centre du mask
        self.phantom_center = self.active_pixels.mean(axis=0) # Attention à l'ordre des axes, [?, ?] => [y, x] [horizontal, vertical]
        
        # Calculer la distance entre chaque pixel du mask et le centre
        self.distances = np.sqrt((self.active_pixels[:, 0] - self.phantom_center[0])**2 + (self.active_pixels[:, 1] - self.phantom_center[1])**2)
        
        # Calculer le rayon max => Rayon du fantôme
        self.max_radius = self.distances.max()
        self.internal_radius = self.max_radius - (self.internal_margin / self.pixel_spacing[0])
        self.diameter_mes_mm = self.max_radius*2*self.pixel_spacing[0]
        self.internal_diameter_mm = 2*self.internal_radius*self.pixel_spacing[0]
        
        # Création du mask d'une ROI circulaire de la taille du fantôme et du mask d'une ROI circulaire correspondant au fantôme interne en utilisant le rayon interne
        self.phantom = self.create_circular_roi([self.size_x, self.size_y], self.phantom_center, self.max_radius)
        self.internal_phantom = self.create_circular_roi([self.size_x, self.size_y], self.phantom_center, self.internal_radius)
        
        # Définition des tailles des ROIs centrale et latérales 
        self.central_roi_size = self.central_roi_s * self.internal_radius
        self.external_roi_size = self.external_roi_s * self.internal_radius
        
        # Distance des rois externes du bord interne du fantôme
        self.radius = self.internal_radius - (self.external_roi_offset / self.pixel_spacing[0])
        
        # Fixe la distance minimum des ROIs latérales à la taille des ROIs afin de maintenir les ROIs dans le fantômes interne 
        if self.radius >= (self.internal_radius - self.external_roi_size):
            self.radius = self.internal_radius - self.external_roi_size
        # Création des maks des ROIs centrale et latérales 
        self.mask_central_roi = self.create_circular_roi([self.size_x, self.size_y], self.phantom_center , self.central_roi_size)
        self.mask_external_N_roi = self.create_circular_roi([self.size_x, self.size_y], [self.phantom_center[0] - self.radius, self.phantom_center[1]] , self.external_roi_size)
        self.mask_external_S_roi = self.create_circular_roi([self.size_x, self.size_y], [self.phantom_center[0] + self.radius, self.phantom_center[1]] , self.external_roi_size)
        self.mask_external_W_roi = self.create_circular_roi([self.size_x, self.size_y], [self.phantom_center[0], self.phantom_center[1] - self.radius] , self.external_roi_size)
        self.mask_external_E_roi = self.create_circular_roi([self.size_x, self.size_y], [self.phantom_center[0], self.phantom_center[1] + self.radius] , self.external_roi_size)

        self.text_info_area.insert("end", f'Diamètre fantôme: {self.diameter_mes_mm:.1f} mm')
        self.text_info_area.insert("end", f'\nDiamètre interne fantôme: {self.internal_diameter_mm:.1f} mm \n')
      
    # Fonction qui applique les ROIs sur toutes les images 
    def apply_rois_to_all_images(self):
        if not self.dicom_images:
            self.text_info_area.insert("end", "Aucune image disponible pour appliquer les ROIs.")
            return

        # Dictionnaires pour stocker les ROIs de chaque image
        self.central_roi = {}
        self.external_roi_N = {}
        self.external_roi_S = {}
        self.external_roi_E = {}
        self.external_roi_W = {}

        for image_name, image_data in self.dicom_images.items():
            try:
                # Vérifie si les masques sont bien définis
                if not hasattr(self, "mask_central_roi") or self.mask_central_roi is None:
                    self.text_info_area.insert("end", f"Masques non définis pour {image_name}. Exécutez find_contour() en premier.")
                    continue
            
                # Appliquer les masques sur chaque image
                self.central_roi[image_name] = image_data[self.mask_central_roi]
                self.external_roi_N[image_name] = image_data[self.mask_external_N_roi]
                self.external_roi_S[image_name] = image_data[self.mask_external_S_roi]
                self.external_roi_E[image_name] = image_data[self.mask_external_E_roi]
                self.external_roi_W[image_name] = image_data[self.mask_external_W_roi]

            except Exception as e:
                self.text_info_area.insert("end", f"Erreur lors de l'application des ROIs sur {image_name}: {e}")

    def analyze_all_images(self):
        if not self.dicom_images:
            self.text_info_area.insert("end", "Aucune image à analyser.")
            return

        # Initialiser une liste pour stocker les résultats de chaque image
        results_list = []
        
        # Initialiser les variables globales pour éviter de référencer None
        self.n_ct_center = self.sigma_ct_center = None
        self.n_ct_lateral_N = self.sigma_ct_lateral_N = None
        self.n_ct_lateral_S = self.sigma_ct_lateral_S = None
        self.n_ct_lateral_E = self.sigma_ct_lateral_E = None
        self.n_ct_lateral_W = self.sigma_ct_lateral_W = None
        self.unif = None
        
        for image_name in self.dicom_images.keys():
            try:
                # Vérifier si les ROIs existent pour cette image
                if image_name not in self.central_roi:
                    self.text_info_area.insert("end", f"ROIs non disponibles pour {image_name}. Assurez-vous que apply_rois_to_all_images() a été exécutée.")
                    continue
                
                # Calculer les statistiques pour cette image
                self.n_ct_center = np.mean(self.central_roi[image_name])
                self.sigma_ct_center = np.std(self.central_roi[image_name])
                
                self.n_ct_lateral_N = np.mean(self.external_roi_N[image_name])
                self.sigma_ct_lateral_N = np.std(self.external_roi_N[image_name])
                
                self.n_ct_lateral_S = np.mean(self.external_roi_S[image_name])
                self.sigma_ct_lateral_S = np.std(self.external_roi_S[image_name])
                
                self.n_ct_lateral_E = np.mean(self.external_roi_E[image_name])
                self.sigma_ct_lateral_E = np.std(self.external_roi_E[image_name])
                
                self.n_ct_lateral_W = np.mean(self.external_roi_W[image_name])
                self.sigma_ct_lateral_W = np.std(self.external_roi_W[image_name])
                
                self.unif = abs(max([
                    abs(self.n_ct_center - self.n_ct_lateral_N),
                    abs(self.n_ct_center - self.n_ct_lateral_S),
                    abs(self.n_ct_center - self.n_ct_lateral_E),
                    abs(self.n_ct_center - self.n_ct_lateral_W)
                    ]))
                
                # Ajouter les résultats de cette image à la liste
                results_list.append({
                    "Image": image_name,
                    "Serie": f"{self.dicom_metadata.get(image_name, {}).get('Serie', 'Inconnu')}",
                    "Tension": f"{int(self.dicom_metadata.get(image_name, {}).get('Tension', 'Inconnu'))}",
                    "NCT Centre": self.n_ct_center,
                    "Bruit Centre": self.sigma_ct_center,
                    "NCT Haut": self.n_ct_lateral_N,
                    "NCT Droit": self.n_ct_lateral_E,
                    "NCT Bas": self.n_ct_lateral_S,
                    "NCT Gauche": self.n_ct_lateral_W,
                    "Uniformité": self.unif,
                    "Bruit Haut": self.sigma_ct_lateral_N,
                    "Bruit Droit": self.sigma_ct_lateral_E,
                    "Bruit Bas": self.sigma_ct_lateral_S,
                    "Bruit Gauche": self.sigma_ct_lateral_W,
                    })
                
            except Exception as e:
                self.text_info_area.insert("end", f"Erreur lors de l'analyse de {image_name}: {e}")
                
            # Ajout des paramètres d'analyse dans le dataframe de données 
            self.df_data.loc["Date analyse"] = datetime.now().date().strftime("%d/%m/%Y")
            self.df_data.loc["Diamètre fantôme (mm)"] =  f'{self.diameter_mes_mm:.1f}'
            self.df_data.loc["Diamètre interne fantôme (mm)"] = f'{self.internal_diameter_mm:.1f}'
            self.df_data.loc["Epaisseur paroi (mm)"] = self.internal_margin
            self.df_data.loc["Offset ROIs externes (mm)"] = self.external_roi_offset
                
        # Convertir la liste en DataFrame
        self.df_results = pd.DataFrame(results_list)
        self.df_results = self.df_results.T
        
        # Tri des colonnes du dataframe selon la tension 
        self.df_results.loc["Tension"] = pd.to_numeric(self.df_results.loc["Tension"], errors='coerce')
        sorted_columns = self.df_results.loc["Tension"].sort_values(ascending=True).index
        self.df_results = self.df_results[sorted_columns]
        #self.df_results.columns = self.df_results.loc['Tension'].values
        self.df_results.columns = self.df_results.loc['Serie'].values

        # Fixe le format des données dans les dataframes 
        pd.options.display.float_format = "{:,.2f}".format
        
        # Afficher les résultats dans la zone de texte 
        self.display_results()
        
    # Fonction qui affiche les résultats dans la zone de texte                        
    def display_results(self):
        self.result_area.delete("1.0", "end")
        if None in [self.n_ct_center, self.sigma_ct_center, self.n_ct_lateral_N, self.sigma_ct_lateral_N, 
                self.n_ct_lateral_S, self.sigma_ct_lateral_S, self.n_ct_lateral_E, self.sigma_ct_lateral_E, 
                self.n_ct_lateral_W, self.sigma_ct_lateral_W, self.unif]:
            self.result_area.insert("end", "Certaines valeurs sont manquantes, veuillez vérifier l'analyse.")
            return
        
        self.result_area.insert("end", self.df_results)

    # Fonction pour créer une ROI circulaire avec une taille de matrice, la position du centre et le rayon 
    def create_circular_roi(self, size, center, radius):
        y, x = np.ogrid[:size[0], :size[1]]
        roi = (x - center[1])**2 + (y - center[0])**2 <= radius**2
        return roi

    # Fonction pour récupérer la valeur donnée par l'utilisateur pour la coupe à utiliser
    def get_slice(self):
        self.text_info_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        if int(self.entry_value_slice.get()) > self.total_slice_number:
            self.text_info_area.insert("end", f'Numéro de coupe supérieur au nombre de coupes, nombre de coupes: {self.total_slice_number}')
        else: 
            self.slice = self.entry_value_slice.get()
            self.load_dicom()

    # Fonction pour récupérer la valeur donnée par l'utilisateur pour la marge interne 
    def get_internal_margin(self):
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.internal_margin = float(self.entry_value_internal_margin.get())
        self.display_info()
        self.find_contour()
        self.apply_rois_to_all_images()
        self.display_image_rois()
    
    # Fonction pour récupérer la valeur donnée par l'utilisateur pour la marge externe 
    def get_external_roi_offset(self):
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.external_roi_offset = float(self.entry_value_offset.get())
        self.display_info()
        self.find_contour()
        self.apply_rois_to_all_images()
        self.display_image_rois()
            
    # Sauvegarde des images seule (artéfacts) et des images avec les ROIs 
    def save_image_rois(self):
        if self.dicom_data is not None and hasattr(self.dicom_data, 'pixel_array'):
            # Créer une figure Matplotlib
            if not self.dicom_images:
                self.text_info_area.insert("end", "Aucune image DICOM à afficher")
                return
                
            self.legend_patches = [
                mpatches.Patch(color='red', label='Contour fantôme détecté'),
                mpatches.Patch(color='orange', label='Contour interne'),
                mpatches.Patch(color='blue', label='Contours ROIs'),
                plt.Line2D([], [], color='green', marker='o', linestyle='None', markersize=3, label='Centre')
            ]
            
            # Sauvegarde des images individuellement 
            for image_name, pixel_array in self.dicom_images.items():
                try:
                    serie = self.dicom_metadata.get(image_name, {}).get('Serie', 'Inconnu')

                    self.df = self.df_results.T
                    
                    # Récupération des infos et résultats des ROIs pour les afficher sur les images
                    if serie in self.df.index:
                        results = self.df.loc[serie]
                        # Récupéraion des ROIs
                        rois = {
                            'Central': self.mask_central_roi,
                            'External N': self.mask_external_N_roi,
                            'External S': self.mask_external_S_roi,
                            'External E': self.mask_external_E_roi,
                            'External W': self.mask_external_W_roi
                            }
                            
                        # Text à afficher au niveau des ROIs
                        roi_texts = {
                            'Central': f"NCT: {results['NCT Centre']:.2f} \nBruit: {results['Bruit Centre']:.2f}",
                            'External N': f"NCT: {results['NCT Haut']:.2f} \nBruit: {results['Bruit Haut']:.2f}",
                            'External S': f"NCT: {results['NCT Bas']:.2f} \nBruit: {results['Bruit Bas']:.2f}",
                            'External E': f"NCT: {results['NCT Droit']:.2f} \nBruit: {results['Bruit Droit']:.2f}",
                            'External W': f"NCT: {results['NCT Gauche']:.2f} \nBruit: {results['Bruit Gauche']:.2f}"
                            }
                        
                        # Positions décalée des texts par rapport aux ROIs
                        roi_pos_x = {'Central': 20,'External N': 50, 'External S': 50, 'External E': 0, 'External W': 0 }
                        roi_pos_y = {'Central': 20,'External N': 0, 'External S': 0, 'External E': 35, 'External W': 35 }
                
                        fig, ax = plt.subplots(figsize=(20, 20))
                        plt.rc('font', size=40)
                        ax.imshow(pixel_array, cmap='gray', vmin=self.min_val, vmax=self.max_val)
                        ax.set_title(f"{serie}")
                        ax.text(0.99, 0.01,
                            f'{self.dicom_metadata[image_name]["Name"]}\n'
                            f'{self.dicom_metadata[image_name]["Manufacturer"]}\n'
                            f'{self.dicom_metadata[image_name]["Device"]}\n'
                            f'{self.dicom_metadata[image_name]["Institution"]}\n'
                            f'{self.dicom_metadata[image_name]["Tension"]} kV\n'
                            f'{self.dicom_metadata[image_name]["Date"]}',
                            color='white', fontsize=20, ha='right', va='bottom',
                            bbox=dict(facecolor='black', alpha=0.5),
                            transform=ax.transAxes)
                        #ax.legend(handles=self.legend_patches, fontsize=15)
                        ax.axis('off')                  
                    
                        fig1, ax1 = plt.subplots(figsize=(20, 20))
                        ax1.imshow(pixel_array, cmap='gray', vmin=self.min_val, vmax=self.max_val)
                        ax1.scatter(self.phantom_center[1], self.phantom_center[0], color='green', label='Centre')
                        ax1.contour(self.phantom, colors='red', linewidths=2)
                        ax1.contour(self.internal_phantom, colors='orange', linewidths=2)
                        ax1.contour(self.mask_central_roi, colors='blue', linewidths=2)
                        ax1.contour(self.mask_external_N_roi, colors='blue', linewidths=2)
                        ax1.contour(self.mask_external_S_roi, colors='blue', linewidths=2)
                        ax1.contour(self.mask_external_E_roi, colors='blue', linewidths=2)
                        ax1.contour(self.mask_external_W_roi, colors='blue', linewidths=2)
                    
                        # Afficher les résultats des ROIs seulement si elles existent
                        # Calculer la position moyenne de la ROI pour placer le texte
                        for roi_name, roi_mask in rois.items():
                            if np.any(roi_mask):  
                                y, x = np.where(roi_mask)
                                center_y, center_x = np.mean(y) + roi_pos_y[roi_name], np.mean(x) + roi_pos_x[roi_name]
                                ax1.text(center_x, center_y, roi_texts[roi_name], color='white', fontsize=20, ha='center', va='center', bbox=dict(facecolor='black', alpha=0.5))
                    
                        ax1.set_title(f"{serie}")
                        ax1.text(0.99, 0.01,
                            f'{self.dicom_metadata[image_name]["Name"]}\n'
                            f'{self.dicom_metadata[image_name]["Manufacturer"]}\n'
                            f'{self.dicom_metadata[image_name]["Device"]}\n'
                            f'{self.dicom_metadata[image_name]["Institution"]}\n'
                            f'{self.dicom_metadata[image_name]["Tension"]} kV\n'
                            f'{self.dicom_metadata[image_name]["Date"]}',
                            color='white', fontsize=20, ha='right', va='bottom',
                            bbox=dict(facecolor='black', alpha=0.5),
                            transform=ax.transAxes)
                        ax1.legend(handles=self.legend_patches, fontsize=15)
                        ax1.axis('off')                    
                        
                        self.png_file = f"CT_image_quality_{self.device}_{self.institution}_{serie}_{datetime.now().date()}_Artefacts.png"
                        self.png_path = os.path.join(self.path, self.png_file)
                        self.png_file1 = f"CT_image_quality_{self.device}_{self.institution}_{serie}_{datetime.now().date()}_ROIs.png"
                        self.png_path1 = os.path.join(self.path, self.png_file1)
                        
                        fig.savefig(f"{self.png_path}")
                        fig1.savefig(f"{self.png_path1}")
                        
                        plt.close(fig)
                        plt.close(fig1)
                except Exception as e:
                    self.text_info_area.insert("end", f"Erreur lors de la sauvegarde de l'image': {str(e)}\n")
                        
        else:
            self.text_info_area.insert("end", "Aucune image trouvée dans ce fichier DICOM.\n")
    
    # Fonction pour sauvegarder les résultats dans un fichier excel et les images 
    def save_results(self):
        self.path = filedialog.askdirectory()
        self.text_info_area.insert("end", "\nSauvegarde en cours ...\n")
        self.root.update_idletasks()

        self.excel_file = f"CT_image_quality_{self.device}_{self.institution}_{datetime.now().date()}.xlsx"
        self.save_path = os.path.join(self.path, self.excel_file)
        
        # 2 chiffres significatifs pour les données 
        self.df_results1 = self.df_results.round(2)
        
        try:
            # Sauvegarde des DataFrames dans un fichier Excel avec plusieurs feuilles
            with pd.ExcelWriter(self.save_path, engine="openpyxl") as writer:
                self.df_results1.head(9).to_excel(writer, sheet_name="Résultats", float_format="%.2f")
                #self.df_data.iloc[:, [1]].to_excel(writer, sheet_name="Données")            
                self.df_data.to_excel(writer, sheet_name="Données")            
            self.save_image_rois()
            self.text_info_area.insert("end", "Sauvegarde terminée.\n")
            
        except Exception as e:
            self.text_info_area.insert("end", f"Erreur lors de la sauvegarde : {str(e)}\n")
        
        self.text_info_area.yview("end")
        
    # Fonction pour réinitialiser l'interface, supprimer les infos, données et canvas pour sélectionner un nouveau dossier 
    def reinitialize(self):
        self.text_info_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        
        # Réinitialisation des variables 
        self.dicom_files_by_directory = {}
        self.directories_with_dicom = []
        self.dicom_images = {}
        self.dicom_metadata = {}
        self.images_names = []
        self.central_roi = {}
        self.external_roi_N = {}
        self.external_roi_S = {}
        self.external_roi_E = {}
        self.external_roi_W = {}
        
        # Suppression du texte rentré manuellement par l'utilisateur 
        self.entry_value_slice.delete(0, 'end')
        self.entry_value_internal_margin.delete(0, 'end')
        self.entry_value_offset.delete(0, 'end')
        
        # Suppression de tous les canvas
        for item in self.canvas.get_tk_widget().find_all():
            self.canvas.get_tk_widget().delete(item) 
        
        # Réinitialisation pour le choix de la coupe 
        self.slice = 0
        self.num_image = 0

# Application principale qui lance le programme 
if __name__ == "__main__":

    root = ThemedTk(theme="plastik")
    app = CT_quality(root)
    root.mainloop()
