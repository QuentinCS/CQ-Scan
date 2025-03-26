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
import time 
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
        self.root.geometry("1400x1000") # Taille de l'interface 

        # Dictionnaire pour les marges internes du fantôme par défaut pour les différents constructeurs  
        self.phantom_dict = {
            "Philips" : 8,
            "GE MEDICAL SYSTEMS" : 7,
            "Canon Medical Systems": 4,
            "SIEMENS": 5,
            }
        
        # Initialisation des variables 
        self.slice = 0
        self.total_slice_number = 0
        self.seuil = -500 # Seuillage en UH pour la détection du contour 
        self.mean_radius = None # Rayon moyen du contour => rayon du fantôme détecté
        self.roi_centrale_size = None # Taille de la ROI centrale 
        self.roi_laterale_size = None # Taille des ROIs latérales
        self.internal_margin = 0 # Marge interne du fantôme vis àa vis du contour extérieur du fantôme
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
        self.text_info_area = scrolledtext.ScrolledText(self.text_frame, width=20, height=15)
        self.text_info_area.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        
        # Création d'une zone de texte pour afficher les métadonnées
        self.result_area = scrolledtext.ScrolledText(self.text_frame, width=20, height=15)
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
        self.load_button = ttk.Button(self.button_frame, text="Charger un dossier avec les images Dicom", command=self.load_dicom)
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
        self.entry_label_internal_margin = ttk.Label(self.button_frame, text="Marge interne (mm, facultatif):")
        self.entry_label_internal_margin.grid(row=3, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # Création d'un champ de saisie (Entry) pour l'utilisateur
        self.entry_value_internal_margin = ttk.Entry(self.button_frame)
        self.entry_value_internal_margin.grid(row=4, column=0, padx=10, pady=(5, 0), sticky="ew")
        
        # Création d'un bouton pour récupérer la valeur entrée par l'utilisateur
        self.get_value_button_internal_margin = ttk.Button(self.button_frame, text="Valider", command=self.get_internal_margin)
        self.get_value_button_internal_margin.grid(row=4, column=1, padx=10, pady=5)
               
        # ---- Champ de saisie pour entrer une valeur ----
        self.entry_label = ttk.Label(self.button_frame, text="Offset ROIs latérales (mm, facultatif, par défaut 20 mm):")
        self.entry_label.grid(row=5, column=0, padx=10, pady=(5, 0), sticky="w")
        
        # Création d'un champ de saisie (Entry) pour l'utilisateur
        self.entry_value_offset = ttk.Entry(self.button_frame)
        self.entry_value_offset.grid(row=6, column=0, padx=10, pady=(5, 0), sticky="ew")
        
        # Création d'un bouton pour récupérer la valeur entrée par l'utilisateur
        self.get_value_button_offset = ttk.Button(self.button_frame, text="Valider", command=self.get_external_roi_offset)
        self.get_value_button_offset.grid(row=6, column=1, padx=10, pady=5)
               
        # Création d'un bouton pour lancer la détection des contours 
        self.analyze_button = ttk.Button(self.button_frame, text="Mesurer", command=self.analyze)
        self.analyze_button.grid(pady=5)
        
        # Création d'un bouton pour sauvegarder les données
        self.save_button = ttk.Button(self.button_frame, text="Sauvegarder", command=self.save_results)
        self.save_button.grid(pady=5)
        
        # Création d'un bouton pour réinitialiser l'interface
        self.reinitialize_button = ttk.Button(self.button_frame, text="Réinitialiser", command=self.reinitialize)
        self.reinitialize_button.grid(pady=5)
           
    ############################# Déclaration des fonctions de la classe #################################
    
    # Fonction pour sélectionner l'image Dicom de la coupe centrale (de l'image pas du fantôme) des images Dicom 
    # En assumant un fichier par coupe
    # Condition permettant de recharger l'image en sélectionnant le numéro d'une coupe 
    def load_dicom(self):
        if self.slice == 0:
            folder_path = filedialog.askdirectory()
            if folder_path:
               self.dicom_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)]# if f.endswith(".dcm")]
               self.total_slice_number = int(len(self.dicom_files))
               for i in range(len(self.dicom_files)):
                   self.dicom_data = pydicom.dcmread(self.dicom_files[i])
                   if self.dicom_data[0x0020, 0x0013].value == len(self.dicom_files) // 2:       
                       # Extraire l'image de la donnée DICOM, en utilisant bien les UH
                       self.image = self.dicom_data.pixel_array * self.dicom_data.RescaleSlope + self.dicom_data.RescaleIntercept
                       self.read_dicom_tag()
                       self.display_info()
                       self.find_contour()
                       self.display_image_rois()
                       
                       break
        else: 
            if self.dicom_files:
                for i in range(0, len(self.dicom_files)):
                    self.dicom_data = pydicom.dcmread(self.dicom_files[i])
                    slice_n = self.dicom_data[0x0020, 0x0013].value
                    if slice_n == self.slice:
                        self.dicom_data = pydicom.dcmread(self.dicom_files[i])
                        # Extraire l'image de la donnée DICOM, en utilisant bien les UH
                        self.image = self.dicom_data.pixel_array * self.dicom_data.RescaleSlope + self.dicom_data.RescaleIntercept
                        self.read_dicom_tag()
                        self.display_info()
                        self.find_contour()
                        self.display_image_rois()
                        break

    # Fonction pour afficher les données Dicom
    def display_dicom_tag(self):
        self.text_area.delete("1.0", "end")  # Efface le texte précédent
        for elem in self.dicom_data:
            self.text_area.insert("end", f"{elem}\n")
            
    # Fonction pour lire et récupérer les informations importantes des tags Dicom 
    def read_dicom_tag(self):
        self.device = self.dicom_data[0x0008, 0x1090].value
        self.manufacturer = self.dicom_data[0x0008, 0x0070].value
        self.institution =self.dicom_data[0x0008, 0x0080].value
        self.name = self.dicom_data[0x0008, 0x1010].value
        self.tension = self.dicom_data[0x0018, 0x0060].value
        self.date = datetime.strptime(self.dicom_data[0x0008, 0x0020].value, "%Y%m%d")
        self.date = self.date.strftime("%d/%m/%Y")
        self.size_x = self.dicom_data[0x0028, 0x0011].value
        self.size_y = self.dicom_data[0x0028, 0x0010].value
        self.pixel_spacing = self.dicom_data[0x0028, 0x0030].value
        self.slice_number = self.dicom_data[0x0020, 0x0013].value
        
        # Utiliser la marge présente dans le dictionnaire si le constructeur existe 
        if str(self.manufacturer) in self.phantom_dict.keys(): 
            self.internal_margin = self.phantom_dict[str(self.manufacturer)]
        else:
            self.internal_margin = 0
                    
    # Fonction pour afficher l'image sélectionnée 
    def display_image(self):
        if self.dicom_data is not None and hasattr(self.dicom_data, 'pixel_array'):
            # Créer une figure Matplotlib
            fig, ax = plt.subplots(figsize=(20, 20))
            plt.rc('font', size=10)
            ax.imshow(self.image, cmap=plt.cm.gray)
            ax.set_title(f'{self.name},\n {self.manufacturer}, {self.device}, {self.tension} kV')
            ax.axis("off")

            # Nettoyer l'affichage précédent
            for widget in self.image_frame.winfo_children():
                widget.destroy()

            # Intégrer l'image dans Tkinter
            self.canvas = FigureCanvasTkAgg(fig, master=self.image_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            self.text_info_area.insert("end", "Aucune image trouvée dans ce fichier DICOM.\n")
            
    # Fonction pour afficher l'image sélectionnée 
    def display_image_rois(self):
        if self.dicom_data is not None and hasattr(self.dicom_data, 'pixel_array'):
            # Créer une figure Matplotlib
            fig, axs = plt.subplots(1, 2, figsize=(20, 20))
            plt.rc('font', size=20)
            
            # Mise en forme de la légende par Patch
            self.legend_patches = [
                mpatches.Patch(color='red', label='Contour fantôme détecté'),
                mpatches.Patch(color='orange', label='Contour interne'),
                mpatches.Patch(color='blue', label='Contours ROIs'),
                plt.Line2D([], [], color='green', marker='o', linestyle='None', markersize=3, label='Centre')
            ]
            
            axs[0].imshow(self.image, cmap=plt.cm.gray)
            axs[0].set_title(f'{self.name},\n {self.manufacturer}, {self.device}, {self.tension} kV', size=10)
            axs[0].axis("off")
                       
            axs[1].imshow(self.image, cmap=plt.cm.gray)
            axs[1].contour(self.phantom, colors='red')#, labels='Contour fantôme détecté')#, linewidth=1)
            axs[1].contour(self.internal_phantom, colors='orange')#, labels='Contour fantôme détecté')#, linewidth=1)
            axs[1].contour(self.mask_central_roi, colors='blue')#, linewidth=1)
            axs[1].contour(self.mask_external_N_roi, colors='blue')#, linewidth=1)
            axs[1].contour(self.mask_external_S_roi, colors='blue')#, linewidth=1)
            axs[1].contour(self.mask_external_E_roi, colors='blue')#, linewidth=1)
            axs[1].contour(self.mask_external_W_roi, colors='blue')#, linewidth=1)
            axs[1].scatter(self.phantom_center[1], self.phantom_center[0], color='green', label='Centre')
            axs[1].set_title("Détection du contour et ROIs", size=10)
            axs[1].axis("off")
            axs[1].legend(handles=self.legend_patches, fontsize=5)
            
            # Nettoyer l'affichage précédent
            for widget in self.image_frame.winfo_children():
                widget.destroy()

            # Intégrer l'image dans Tkinter
            self.canvas = FigureCanvasTkAgg(fig, master=self.image_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            self.text_info_area.insert("end", "Aucune image trouvée dans ce fichier DICOM.\n")

    # Fonction pour identifier les contours du fantôme dans l'image à partir d'un seuillage sur les UH "self.seuil" 
    def find_contour(self):

        # Créer un masque binaire à partir des seuils
        self.mask = self.image >= self.seuil

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

        # Création des ROIs en utilisant les mask sur l'image d'origine 
        self.central_roi = self.image[self.mask_central_roi]
        self.external_roi_N = self.image[self.mask_external_N_roi]
        self.external_roi_S = self.image[self.mask_external_S_roi]
        self.external_roi_E = self.image[self.mask_external_E_roi]
        self.external_roi_W = self.image[self.mask_external_W_roi]
        
        # Affichage de l'image et des différentes ROIs
        self.display_image_rois()
        
        self.text_info_area.insert("end", f'\nDiamètre fantôme mesuré: {self.diameter_mes_mm:.1f} mm \n')
        
    # Fonction pour réaliser la mesure des ROIs et enregistrer les données sous forme de dataframe
    def analyze(self):
        self.n_ct_center = np.mean(self.central_roi)
        self.sigma_ct_center = np.std(self.central_roi)

        self.n_ct_lateral_N = np.mean(self.external_roi_N)
        self.sigma_ct_lateral_N = np.std(self.external_roi_N)

        self.n_ct_lateral_S = np.mean(self.external_roi_S)
        self.sigma_ct_lateral_S = np.std(self.external_roi_S)

        self.n_ct_lateral_E = np.mean(self.external_roi_E)
        self.sigma_ct_lateral_E = np.std(self.external_roi_E)

        self.n_ct_lateral_W = np.mean(self.external_roi_W)
        self.sigma_ct_lateral_W = np.std(self.external_roi_W)
        
        self.unif = abs(max([self.n_ct_center-self.n_ct_lateral_N, self.n_ct_center-self.n_ct_lateral_S, self.n_ct_center-self.n_ct_lateral_E, self.n_ct_center-self.n_ct_lateral_W]))

        self.display_results()
        
        # Dictionnaire pour créer un dataframe pour les résultats complets
        results = {
            "Region": ["Centrale", "Lateral_N", "Lateral_S", "Lateral_E", "Lateral_W"],
            "Moyenne": [
                self.n_ct_center, self.n_ct_lateral_N, self.n_ct_lateral_S,
                self.n_ct_lateral_E, self.n_ct_lateral_W
            ],
            "Écart-type": [
                self.sigma_ct_center, self.sigma_ct_lateral_N, self.sigma_ct_lateral_S,
                self.sigma_ct_lateral_E, self.sigma_ct_lateral_W
            ]
        }
        self.df_results = pd.DataFrame(results)
        self.df_results["Uniformité"] = [self.unif] + [None] * 4  # Ajouter la valeur de l'uniformité à la première ligne seulement
        
        # Dictionnaire pour créer un dataframe pour les résultats à copier 
        short_results = {
            "Region": ["NCT centre", "Bruit Centre", "NCT Haut", "NCT Droit", "NCT Bas", "NCT Gauche"],
            "Value": [
                self.n_ct_center, self.sigma_ct_center, self.n_ct_lateral_N, self.n_ct_lateral_E, self.n_ct_lateral_S, self.n_ct_lateral_W
            ],
        }
        self.df_short_results = pd.DataFrame(short_results)
        # Dictionnaire pour créer un dataframe pour les données Dicom et d'analyse
        data = {
            "Appareil": [self.device],
            "Fabricant": [self.manufacturer],
            "Institution": [self.institution],
            "Nom": [self.name],
            "Tension (kV)": [self.tension],
            "Taille X": [self.size_x],
            "Taille Y": [self.size_y],
            "Pixel Spacing": [str(self.pixel_spacing)],
            "Numéro Slice": [self.slice_number],
            "Date Acquisition": [self.date],
            "Date Analyse": [datetime.now().date().strftime("%d/%m/%Y")],
            "Marge interne (mm)": [self.internal_margin],
            "Marge ROIs externes (mm)": [self.external_roi_offset],
        }
        self.df_data = pd.DataFrame(data)
        self.df_data = self.df_data.T
        self.df_data.reset_index(inplace=True)  # Transformer l'index en colonne
        self.df_data.columns = ["Paramètre", "Valeur"]  # Renommer les colonnes

    # Fonction pour afficher les informations Dicom
    def display_info(self):
        self.text_info_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.text_info_area.insert("end", f'Site: {self.institution} \nConstructeur: {self.manufacturer} \nModèle: {self.device} \nNom: {self.name}\nTension: {self.tension} kV \nMatrice: {self.size_x} X {self.size_y} \nCoupe: {self.slice_number} \nDate de mesure: {self.date} \n\nMarge interne: {self.internal_margin} mm \nOffset ROI latérales: {self.external_roi_offset} mm \n')

    # Fonction pour afficher les résultats après la mesure 
    def display_results(self):
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.insert("end", f'N CT Centre:  {self.n_ct_center:.3f},  STD:  {self.sigma_ct_center:.3f} \nN CT Nord:     {self.n_ct_lateral_N:.3f},  STD:  {self.sigma_ct_lateral_N:.3f} \nN CT Sud:      {self.n_ct_lateral_S:.3f},  STD:  {self.sigma_ct_lateral_S:.3f} \nN CT Est:      {self.n_ct_lateral_E:.3f},  STD:  {self.sigma_ct_lateral_E:.3f} \nN CT Ouest:    {self.n_ct_lateral_W:.3f},  STD:  {self.sigma_ct_lateral_W:.3f} \nUniformité: {self.unif:.2f}\n')
        
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
    
    # Fonction pour récupérer la valeur donnée par l'utilisateur pour la marge externe 
    def get_external_roi_offset(self):
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        self.external_roi_offset = float(self.entry_value_offset.get())
        self.display_info()
        self.find_contour()
 
    # Fonction pour sauvegarder l'image et les ROIs sous forme de png 
    def save_figure(self):
        fig, axs = plt.subplots(1, 2, figsize=(40, 20))
        plt.rc('font', size=30)
        
        self.legend_patches = [
            mpatches.Patch(color='red', label='Contour fantôme détecté'),
            mpatches.Patch(color='orange', label='Contour interne'),
            mpatches.Patch(color='blue', label='Contours ROIs'),
            plt.Line2D([], [], color='green', marker='o', linestyle='None', markersize=3, label='Centre')
        ]
        
        axs[0].imshow(self.image, cmap=plt.cm.gray)
        axs[0].set_title(f'{self.name},\n {self.manufacturer}, {self.device}, {self.tension} kV', size=25)
        axs[0].axis("off")
                   
        axs[1].imshow(self.image, cmap=plt.cm.gray)
        axs[1].contour(self.phantom, colors='red')#, labels='Contour fantôme détecté')#, linewidth=1)
        axs[1].contour(self.internal_phantom, colors='orange')#, labels='Contour fantôme détecté')#, linewidth=1)
        axs[1].contour(self.mask_central_roi, colors='blue')#, linewidth=1)
        axs[1].contour(self.mask_external_N_roi, colors='blue')#, linewidth=1)
        axs[1].contour(self.mask_external_S_roi, colors='blue')#, linewidth=1)
        axs[1].contour(self.mask_external_E_roi, colors='blue')#, linewidth=1)
        axs[1].contour(self.mask_external_W_roi, colors='blue')#, linewidth=1)
        axs[1].scatter(self.phantom_center[1], self.phantom_center[0], color='green', label='Centre')
        axs[1].set_title("Détection du contour et ROIs", size=20)
        axs[1].axis("off")
        axs[1].legend(handles=self.legend_patches, fontsize=18)
        
        self.png_file = f"CT_image_quality_{self.device}_{self.institution}_{datetime.now().date()}_ROIs.png"
        self.png_path = os.path.join(self.path, self.png_file)
        
        fig.savefig(f"{self.png_path}")
        
    # Fonction pour sauvegarder les résultats dans un fichier excel et les images 
    def save_results(self):
        self.path = filedialog.askdirectory()
        self.excel_file = f"CT_image_quality_{self.device}_{self.institution}_{datetime.now().date()}.xlsx"
        self.save_path = os.path.join(self.path, self.excel_file)
                
        try:
            # Sauvegarde des DataFrames dans un fichier Excel avec plusieurs feuilles
            with pd.ExcelWriter(self.save_path, engine="openpyxl") as writer:
                self.df_short_results.to_excel(writer, sheet_name="Résultats à copier", index=False)
                self.df_data.to_excel(writer, sheet_name="Données", index=False)
                self.df_results.to_excel(writer, sheet_name="Résultats complets", index=False)
            self.save_figure()
            self.text_info_area.insert("end", "Données sauvegardées\n")
        except Exception as e:
            self.text_info_area.insert("end", f"Erreur lors de la sauvegarde : {str(e)}\n")
        
    # Fonction pour supprimer les infos et canvas pour sélectionner une autre image 
    def reinitialize(self):
        # Suppression du texte dans les zones d'affichage de text
        self.text_info_area.delete("1.0", "end")  # Efface le texte précédent
        self.result_area.delete("1.0", "end")  # Efface le texte précédent
        
        # Suppression du texte rentré manuellement par l'utilisateur 
        self.entry_value_slice.delete(0, 'end')
        self.entry_value_internal_margin.delete(0, 'end')
        self.entry_value_offset.delete(0, 'end')
        
        # Suppression de tous les canvas
        for item in self.canvas.get_tk_widget().find_all():
            self.canvas.get_tk_widget().delete(item) 
        
        # Réinitialisation pour le choix de la coupe 
        self.slice = 0

# Application principale qui lance le programme 
if __name__ == "__main__":
    start_time = time.time()
    
    root = ThemedTk(theme="plastik")
    app = CT_quality(root)
    root.mainloop()
    
    duree = time.time()- start_time
