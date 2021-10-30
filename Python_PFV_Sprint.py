# -*- coding: utf-8 -*-

"""
Created on Tue Apr 20 16:37:10 2021

@author: gae0a
"""
################# Import des librairies utiles au code : ######################
#module tkinter + file dialog pour aller cherche le fichier à analyser :
#from tkinter import filedialog as fd
import tkinter as tk
from tkinter import filedialog
# from tkinter import simpledialog
# from tkinter import *
#librairie pour ouvrir un fichier excel et y travailler dedans, le contrôler depuis python --> automation python.
# import xlwings
import xlwings as xw #Xlwings is a module to allow Excel to be automated with Python instead of VBA.
#Import scipy et scip pour tets normalité 
#from scipy.stats import shapiro
#
from scipy.optimize import leastsq
#from scipy.optimize import curve_fit
#from scipy.optimize import least_squares

#import pandas
import pandas as pd  
# from pandas import DataFrame
#import pandas as pd

#Import numpy
import numpy as np

# importing the required matplotlib module
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
from matplotlib.widgets import Cursor, Button
#import datetime object :
from datetime import datetime

#plotly librairie 
# from plotly.offline import plot
# import plotly.graph_objs as go



################ différentes class défini ######################## Principalement pour les GUI interactifs.
class Plot_Embed:
    def __init__(self, master, DataX, DataY, X1, X2):
        """
        class : créer une fentre tkinter avec un plot interactif embarqué dedans avec retour d'information affiché'
        
        Parameters
        ----------
        master : root = tk.Tk() --> une racine de fentre tkinter 
        DataX : float (float of time series --> in seconds) --> Timeline en secondes
        DataY : float --> Speed value en (m/s) 
        
        __Init__ --> Build GUI with plot and buttons embedded.
        
        connect --> fig.canvas.mpl_connec=  method returns a connection id (an integer), set plot interactive.
        
        onclick(self, event) --> handle event of mouse click on plot. return x,y coordinates and draw axvline + text annotation
        
        find_nearest_points --> empty attributes = ?
        
        return_coordinates --> retourne les coordonnées des points x et y sélectionnés/modifiés pour le début et la fin du sprint.
        
        -------
        None.

        """
        
        #Init Boutton selection Gestion START ou STOP Sprint :
        self.phase = None
        #Init Coordonnées :
        #Init des valeurs des coordonnées marquant le début du sprint :
        self.x1 = X1
        self.y1 = None
        #Init des valeurs des coordonnées marquant la fin du sprint :
        self.x2 = X2
        self.y2 = None
        #Init des handles des lignes verticals
        self.AxVline_START = None
        self.AxVline_STOP = None
        
        # print(dir(self.START_label.get()))
        
        #Init tkinter Windows :
        self.master = master
        master.title("START : STOP Sprint Validation ?")
        master.attributes("-topmost", True) #mets le fenetre tkinter root devant toutes les autres fenetres
        master.geometry("900x900")

        #titre de la fenetre tkinter
        self.label = tk.Label(master, text= "Plot Speed + START : STOP Sprint Validation ?")
        
        #Bouton de selection du début ou fin de Sprint :
        self.START_button = tk.Button(master, text="START", command=lambda: self.phase_to_set("START"))
        self.START_label_Widget = tk.Label(master, text = "Début du Sprint à = %1.2f secondes du temps total " %self.x1)
                
        self.STOP_button = tk.Button(master, text="STOP", command=lambda: self.phase_to_set("STOP"))
        self.STOP_label_Widget = tk.Label(master, text= "Début du Sprint à = %1.2f secondes du temps total " %self.x2)
        
        self.QUIT_button = tk.Button(master, text="QUIT", command= self.quit_app)
               
        #Configure Plot windows :  
        self.fig = plt.Figure() 
        
        #
        self.canvas = FigureCanvasTkAgg(self.fig, master = self.master)  # A tk.DrawingArea.
        self.canvas.draw()
        
        self.toolbarFrame = tk.Frame(self.master)
        self.toolbarFrame.grid(row=10,column=1)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbarFrame)
        #self.toolbar = NavigationToolbar2Tk(self.canvas, self.master)
        #self.toolbar.update()
         
        self.DataX = DataX
        self.DataY = DataY
         
        self.ax = self.fig.add_subplot() 
        self.line = self.ax.plot(self.DataX, self.DataY)
        
        self.ax.set_xlabel('time [s]')
        self.ax.set_ylabel('Speed [m/s]')
        
        # Vertical ligne of Sprint Start
        self.AxVline_START = self.ax.axvline(self.x1, ymin=0, ymax=max(self.DataY),
                                                     color = "red",
                                                     linestyle = '--',
                                                     label = 'Début Sprint')
        #Annotation Sprint Start :
        self.AxVline_START_Text =self.ax.text(self.x1, max(self.DataY)/2, 'Début Sprint', fontsize=10,
                                rotation=90, rotation_mode='anchor',
                                transform_rotates_text=True)
        
        # Vertical ligne of Sprint Stop
        self.AxVline_STOP = self.ax.axvline(self.x2, ymin=0, ymax=max(self.DataY),
                                                     color = "black",
                                                     linestyle = '--',
                                                     label = 'Fin Sprint')
        #Annotation Sprint Stop :
        self.AxVline_STOP_Text = self.ax.text(self.x2, max(self.DataY)/2, 'Fin Sprint', fontsize=10,
                                rotation=90, rotation_mode='anchor',
                                transform_rotates_text=True)
        
        # LAYOUT (Disposition) Tkinter Windows with plot :
        self.label.grid(row=0, column=1, columnspan=3, sticky= tk.W + tk.E)
        self.START_button.grid(row=2, column=1, sticky=  "nesw")
        self.START_label_Widget.grid(row=3, column=1, sticky=  "nesw")
        self.STOP_button.grid(row=4, column=1, sticky= "nesw")
        self.STOP_label_Widget.grid(row=5, column=1, sticky=  "nesw")
        self.QUIT_button.grid(row=6, column=1, sticky = "nesw")
        self.canvas.get_tk_widget().grid(row=1, column=1)
        #self.toolbar.get_tk_widget().grid(row=8, column=1)

         
        # vcmd = master.register(self.validate) # we have to wrap the command
        # self.entry = tk.Entry(master, validate="key", validatecommand=(vcmd, '%P'))
        
    #Besoin ? vue qu'on a deja tkinter et que le plot va être dedans en faites ? --> oui besoin !
        #de plus, attention de toujours placer le connecteur avant de lancer 
    def connect(self):
        """
        

        Returns
        -------
        make plot interactive in the tkinter windows
        Must be called before Plot_Embed.display_plot().

        """
        self.cid = self.fig.canvas.mpl_connect('button_press_event', self.onclick)

    def onclick(self, event):
        """Attribut de la classe qui permet de cliquer sur le graph, de récuperer les coordonnées du point sur lequel,
        le click a été effectué et de stocké dans "self.x1" et autres.
        + draw sur le graph d'une ligne verticale marquant la phase déterminé/modifié (Début ou fin de sprint)"""
        
        if self.phase == None :
            tk.messagebox.showwarning("Attention !",
                                      "Avant de sélectionner un point, veuillez choisir quel phase (via les boutons de droite : START/STOP), vous souhaitez modifier !")
        
        elif self.phase == 1 :
            self.x1 = event.xdata
            self.y1 = event.ydata
            #print ("START :",  self.x1, self.y1)
        
            self.START_label_Widget.configure(text = "Début du Sprint à = %1.2f secondes du temps total " %self.x1)
            
                      
            #on suprime la ligne vertical existante et on en crée une autre (ou on change les coordonnées x de cette ligne vertical.)
            self.AxVline_START.remove()
            self.AxVline_START_Text.remove()

            self.AxVline_START = self.ax.axvline(self.x1, ymin=0, ymax=max(self.DataY),
                                                     color = "red",
                                                     linestyle = '--',
                                                     label = "Début Sprint")
                
            #annotation sur le graph le long de la ligne vertical
            self.AxVline_START_Text = self.ax.text(self.x1, max(self.DataY)/2, 'Début Sprint', fontsize=10,
                                rotation=90, rotation_mode='anchor',
                                transform_rotates_text=True)
            self.canvas.draw()
        
        elif self.phase == 2 :
            self.x2 = event.xdata
            self.y2 = event.ydata
            #print ("STOP :",  self.x2, self.y2)
            
            self.STOP_label_Widget.configure(text ="Fin du Sprint à = %1.2f secondes du temps total" %self.x2)
            
            #on suprime la ligne vertical existante et on en crée une autre (ou on change les coordonnées x de cette ligne vertical.)
            self.AxVline_STOP.remove()
            self.AxVline_STOP_Text.remove()
            
            #display a vertical line on plot to markdown STOP 
            self.AxVline_STOP = self.ax.axvline(self.x2, ymin=0, ymax=max(self.DataY),
                                                     color = "black",
                                                     linestyle = '--',
                                                     label = 'Fin Sprint')
                
            #annotation sur le graph le long de la ligne vertical
            self.AxVline_STOP_Text = self.ax.text(self.x2, max(self.DataY)/2, 'Fin Sprint', fontsize=10,
                                rotation=90, rotation_mode='anchor',
                                transform_rotates_text=True)
            self.canvas.draw()

        
        #save result of click in a variable to display in a part of windows :

            
    def phase_to_set(self, method):
        """attribut pour permettre de selectionner la forme ou sera affiché les coordonnées selectioné 
        + lié a un label ou autre afin de l'associé au début ou à la fin du sprint."""
        
        if method == "START":
            self.phase = 1
        elif method == "STOP":
            self.phase = 2

  
    def find_nearest_points(self):
        
        #return the nearest point 
        pass

    def return_coordinates(self):
        """
        get back coordinates choose by clicking on the plot in tkinter windows 
        
        Returns
        -------
        TYPE
            xvalue --> value of time (start sprint).
        TYPE
            yvalue --> value of speed (m/s) (start sprint)
        TYPE
            xvalue --> value of time (stop sprint)
        TYPE
            yvalue --> value of speed (m/s) (stop sprint)

        """
        return self.x1, self.y1, self.x2, self.y2 #self.phase

    def display_plot(self):
        """
        Attriut qui permet de créer la mainloop de la windows tkinter

        Returns
        -------
        None.

        """
        self.master.mainloop()
        # plt.ioff()
        # plt.show()       
    
    def quit_app(self):
        """
        attribut qui permet de quitter tkinter windows.

        Returns
        -------
        None.

        """
        if self.x1 != None and self.x2 != None :
           self.master.quit()
           self.master.destroy()
        elif self.x1 == None and self.x2 != None:
            tk.messagebox.showwarning("Attention !",
                                      "Données de Départ de Sprint manquantes --> veuillez selectionner le départ du sprint, avant de quitter !")
        elif self.x1 != None and self.x2 == None:
            tk.messagebox.showwarning("Attention !",
                                      "Données de Fin de Sprint manquantes --> veuillez selectionner le point de fin de sprint, avant de quitter !")
        else :
            tk.messagebox.showwarning("Attention !",
                                      "Wesh, tu dors !! !")
        # plt.ioff()
        # plt.show()   
            
#*****************************************************************************


class Enter_Data_GUI :
    def __init__(self, master):
        
        """
        class : créer une fenetre tkinter avec des zones d'entrée pour ajouter des valeurs manuellement :'
        
        Parameters
        ----------
        master : root = tk.Tk()

        __Init__ --> Build GUI with Label and Entry embedded.
        
        display_GUI --> mainloop du tkinter 
        
        getvalue --> return entry value
        
        quit_app --> quit tkinter windows application.
        
        -------
        None.

        """
        
        self.master = master
        self.master.title("Input Dialog Windows")
        self.master.attributes("-topmost", True) #mets le fenetre tkinter root devant toutes les autres fenetres
        self.master.geometry("400x300")#largeur x Longueur de la fenetre 
        self.master.eval('tk::PlaceWindow . center')
        #self.master.overrideredirect(1)
        
        self.event_counter = 0
        
        #Données Anthropo :
        # self.Taille
        self.Taille = tk.StringVar(self.master, '1.70')
        # self.Masse
        self.Masse = tk.StringVar(self.master, '60.0')
        
        #Données Envirronemental :
        self.Temp = tk.StringVar(self.master, '20.0')
        self.Vent = tk.StringVar(self.master, '2.0')
        self.Pression_Atmo = tk.StringVar(self.master, '750.0')
        
        #composants :
        self.Label1 = tk.Label(master, text="Taille (m) : ") #label
        self.Entry1 = tk.Entry(master, textvariable = self.Taille )
        
        self.Label2 =  tk.Label(master, text="Masse (kg) :") #label
        self.Entry2 =  tk.Entry(master, textvariable = self.Masse )
        
        self.Label3 = tk.Label(master, text="Temperature (°c) :") #label
        self.Entry3 = tk.Entry(master, textvariable = self.Temp)
        
        self.Label4 =  tk.Label(master, text="Vent (m/s) :") #label
        self.Entry4 =  tk.Entry(master, textvariable = self.Vent)
        
        self.Label5 =  tk.Label(master, text="Pression_Atmo (mmHg) :") #label
        self.Entry5 =  tk.Entry(master, textvariable = self.Pression_Atmo)
        

        self.Quit_Button = tk.Button(master, text="Quitter", command= self.quit_app)
        
        #Layout
        self.Label1.grid(row=0, sticky=tk.W)
        self.Entry1.grid(row=0, column=1, sticky=tk.E)
        
        self.Label2.grid(row=1, sticky=tk.W)  
        self.Entry2.grid(row=1, column=1, sticky=tk.E) 
        
        self.Label3.grid(row=3, sticky=tk.W) 
        self.Entry3.grid(row=3, column=1, sticky=tk.E) 
        
        self.Label4.grid(row=4, sticky=tk.W)  
        self.Entry4.grid(row=4, column=1, sticky=tk.E) 
        
        self.Label5.grid(row=5, sticky=tk.W)  
        self.Entry5.grid(row=5, column=1, sticky=tk.E) 
        
        self.Quit_Button.grid(row=7, column=3, sticky=tk.W)
        
        
        # self.Entry1.bind('<Enter>', self.count_event)
        # self.Entry2.bind('<Enter>', self.count_event)
        # self.Entry3.bind('<Enter>', self.count_event)
        # self.Entry4.bind('<Enter>', self.count_event)
        # self.Entry5.bind('<Enter>', self.count_event)
        
    # def count_event(self, event):
        
    #     self.event_counter += 1
    #     print(self.event_counter)
    
    
    def display_GUI(self):
        self.master.mainloop()
   
    
    def getvalue(self):
        return self.Masse.get(), self.Taille.get(), self.Temp.get(), self.Vent.get(), self.Pression_Atmo.get()
        
    def quit_app(self):
        self.msg_quit = tk.messagebox.askokcancel("Données saisies correct ?",
                                        "Continuer ? ")
       
        if self.msg_quit == 1 :
            self.master.quit()
            self.master.destroy()
 
###############################################################################

################ fonctions définies  ########################
            
#>fonction de calcul d'une moyenne glissante centrée 
def moving_average(x, w): #moving average with convolution of value and zero arrays
    size_MA = w
    h = np.ones(size_MA) / size_MA
    data_convolved = np.convolve(x, h, mode='same')
    return data_convolved


# Fonction de la vitesse modélisé lors du sprint : avec les 3 coefficients exprimés de manières séparés dans une variable chacun "a,b,c"

def VitesseModélisé (x,a,b,c) : #--> Necessaire pour l'utilisation avec curvefit fonction --> curvefit (voir plus bas)
    return a*(1-np.exp(-((x-b)/c))) 

##### Fonction de la vitesse modélisé lors du sprint : avec les 3 coefficients dans une même variable de type numpy.array.float
def VitesseModélisé2 (x,coeffs) :
    return coeffs[0]*(1-np.exp(-((x-coeffs[1])/coeffs[2]))) 

    #avec leastsq de la bibliothèque de scipy --> 1) définit la fonction à minimiser en définissant la fonction objective (dont les paramètres sont a déterminés)
def minimisation(coeffs, y, t) :
    return y - VitesseModélisé2(t,coeffs)

#fonction to caracterise file to open/ file select with tkinter dialog windows --> retrieve file format 
def Open_File(PATH):
    XX = PATH.find('.',-10,-1)
    return PATH[XX:]

########################## Début du Script ###################################
#Mimics "uigetfile" from MATLAB --> Open a windows to search and find file to retrieve file path. 
root = tk.Tk()#racine de la fentre tkinter 
root.attributes("-topmost", True) #mets le fenetre tkinter root devant toutes les autres fenetres
root.withdraw() #cache la fenetre 
file_path = filedialog.askopenfilename(title="Selectionner le fichier RAD (.rad) à traiter ",
                                       filetypes=[('RAD', '.rad'), ('all files', '.*')], 
                                       multiple = True) #ouvre nouvelle fenetre pour selectionner les chemins d'accés + les noms du fichiers selectionnés, avec possibilité de sélectioner quasi seulement les fichiers en .rad
root.destroy()#ferme et détruit la racine de la fenetre tkinter --> a changer par une class

for filename in file_path:
    print(filename)
######## Condition if pour récuperer les datas en fonctions du format du fichier d'import
file = open(filename, "r", encoding='utf-8')
DATA = file.read()
file.close()


#find some informations about testing trial in METADATA:
TRIAL_NAME = DATA[DATA.find(':',DATA.find('TRIAL NAME'))+2:DATA.find('\n',DATA.find('TRIAL NAME'))]
#Date et heure du test de sprint
DATE_HEURE = datetime.strptime(DATA[DATA.find('\n',DATA.find('TRIAL NAME'))+1:DATA.find('(',DATA.find('TRIAL NAME'))-1], '%m/%d/%Y %H:%M:%S')#DATE + HEURE au format datetime
#Fréquence d'acquisition du radar :
SAMPLE_RATE = np.float(DATA[DATA.find(':',DATA.find('SAMPLE RATE'))+3:DATA.find(':',DATA.find('SAMPLE RATE'))+9].replace(',', '.', 1)) #prend les valeurs de DATA contenue à l'index +3 après l'index du premier ":" apres l'index du mot : "SAMPLE RATE" + 3 jusqu'a l'index +9 après l'index du premier ":" toujours après l'index du mot "SAMPLE RATE"
#Nombre de données
SAMPLES = int(DATA[DATA.find(':',DATA.find('SAMPLES'))+3 : DATA.find(':',DATA.find('SAMPLES'))+9].replace(',', '.', 1)) 

#Data type (=Range setting du radar lors du test )
DATA_TYPE =int(DATA[DATA.find(':',DATA.find('DATA TYPE'))+7: DATA.find(':',DATA.find('DATA TYPE'))+9])

#UNIT
UNIT =int(DATA[DATA.find(':',DATA.find('UNITS'))+7: DATA.find(':',DATA.find('UNITS'))+9])

#Speed_Units
Speed_Units = DATA[DATA.find(':',DATA.find('Speed Units'))+2: DATA.find('\n',DATA.find('Speed Units'))]

#Accel Units
Accel_Units = DATA[DATA.find(':',DATA.find('Accel Units')) +2 : DATA.find('\n',DATA.find('Accel Units'))]

#Dist  Units
Dist_Units = DATA[DATA.find(':',DATA.find('Dist  Units')) +2 : DATA.find('\n',DATA.find('Dist  Units'))]


# --> Text to Dataframe : import des datas en format text, en format dataframe avec seulement les données numériques qui nous interresse
col_names = ('Sample','Time','Speed','Accel','Dist')
comment_lines = 17 #ligne de commentaire dans le fichier txt à zappé
header = 2 #idem avec le nom des colonnes + lignes vierges dans le fichier brut 

# METADATA To DataFrames :
df = pd.read_csv(file_path, skiprows=comment_lines + header,
                 header=None, names=col_names, delimiter=' ', skipinitialspace=True, error_bad_lines=False, engine = 'python', skipfooter = 1, decimal=',')

#Extrait du df, les series en tableau de variables numpy à 1D :
Sample = df[df.columns[0]].to_numpy(dtype ='float32')
Temps = df[df.columns[1]].to_numpy(dtype ='float32')
Vitesse_KMH = df[df.columns[2]].to_numpy(dtype ='float32')
Accel = df[df.columns[3]].to_numpy(dtype ='float32')
Dist = df[df.columns[4]].to_numpy(dtype ='float32')

#Conversion de la vitesse en km/h en m/s :
Vitesse_MS = Vitesse_KMH /3.6

#définit le style utilisé sur (tout) les graphiques suivants :
plt.style.use('seaborn')

#verification de la fréquence d'échantillonage : Nb Données Total/ Temps max d'enregistrement 
Fs_Radar = len(Sample)/Temps[-1]
print('la fréquence d\'echantillon recalculé avec les données serait de : %2.3f Hz' %Fs_Radar)


### Ouverture d'une fenetre tkinter pour demander les info Anthropo et envirro du sujet et du test (anthropo, conditions climatiques) :
GUI_root = tk.Tk()
Entry_GUI = Enter_Data_GUI(GUI_root)
Entry_GUI.display_GUI() #display tkinter windows
value_entry = Entry_GUI.getvalue()


#########################################   maintenant que les données sont importés, on ouvre le fichier Excel d'Export ###################################
### Fichier excel --> fiche retour de test avec toute les données et les mises en formes ... --> avec xlwings
#1) Choisit le directory ou sauvegarder le fichier excel avec un GUI tkinter:
root = tk.Tk()
root.attributes("-topmost", True) #mets le fenetre tkinter root devant toutes les autres fenetres
root.withdraw()
folder_Import = filedialog.askdirectory(parent=root,
                                          title="Choisir le dossier d'export des fiches Excels de résultats aux tests",
                                          initialdir= file_path[:file_path.find('Manip OM') + len('Manip OM')],
                                          mustexist= True)

#2) Choisit le directory ou est contenue la fiche modèle de retour de test avec un GUI tkinter:
root = tk.Tk()
root.attributes("-topmost", True) #mets le fenetre tkinter root devant toutes les autres fenetres
root.withdraw()
folder_Fiche_Retour = filedialog.askdirectory(parent=root,
                                          title="Choisir le dossier où se situe la fiche modèle de retour de test",
                                          initialdir= file_path[:file_path.find('Documents') + len('Documents')],
                                          mustexist= True)

#Ouvre un nouveau workbook excel --> Workbook d'export avec xlwings :
#ouverture du workbook d'export des calculs et des données --> "Modèle_Fiche_Retour_Sprint_Force_Vitesse"
Excel_App = xw.App(visible=True, add_book=False) #Ouvre l'application COM Excel, sans ouvrir directement un nvx Workbook( ligne suivante) + rend visible/invisible 
Excel_App.display_alerts = False #pour supprimer tous messages d'alertes avec ouverture de fenetre d'alerte lors de l'ouverture d'un workbook

wb = Excel_App.books.open(folder_Fiche_Retour + '/Modèle_Fiche_Retour_Sprint_Force_Vitesse.xlsx') #ouverture du modèle de fiche pour le rendu du test (export)
#wb = handle du workbook globale 

#Une fois que la fihe modèle est importer --> On sauvegarde en changeant le nom du workbook + sauvegarde dans le dossier d'export des fiches Excels de résultats aux tests
wb.save(folder_Import + '/Fiche_Retour_PFV_Sprint_' + TRIAL_NAME + '.xlsx')

#Handle de la 1ere page = Fiche classique de retour de test :
Feuil_FICHE_RETOUR = wb.sheets(wb.sheets[0].name) # = wb.sheets(1)

#Add nouveaux feuillets excel au workbook --> Données FV-Sprint : Feuil_DATA_BRUT
wb.sheets.add(name= 'DATA_Brut', after = wb.sheets[0]) #Open new sheet and make this sheet :active sheet.
Feuil_DATA_BRUT = wb.sheets(wb.sheets[1].name) # = wb.sheets(2)

#ajout des métadonnées dans le classeur d'export :
Feuil_DATA_BRUT.range('A1').value = 'TRIAL_NAME :'
Feuil_DATA_BRUT.range('A2').value = 'DATE_HEURE :'
Feuil_DATA_BRUT.range('A3').value = 'SAMPLE_RATE (Hz):'
Feuil_DATA_BRUT.range('A4').value = 'SAMPLES :'
Feuil_DATA_BRUT.range('A5').value = 'UNIT :'
Feuil_DATA_BRUT.range('A6').value = 'Speed_Units (%s)' %Speed_Units
Feuil_DATA_BRUT.range('A7').value = 'Accel_Units (%s)' %Accel_Units
Feuil_DATA_BRUT.range('A8').value = 'Dist_Units (%s)' %Dist_Units
Feuil_DATA_BRUT.range('A9').value = 'Fs réel (Hz)'

Feuil_DATA_BRUT.range('B1').value = TRIAL_NAME
Feuil_DATA_BRUT.range('B2').value = DATE_HEURE
Feuil_DATA_BRUT.range('B3').value = SAMPLE_RATE
Feuil_DATA_BRUT.range('B4').value = SAMPLES
Feuil_DATA_BRUT.range('B5').value = UNIT
Feuil_DATA_BRUT.range('B6').value = Speed_Units
Feuil_DATA_BRUT.range('B7').value = Accel_Units
Feuil_DATA_BRUT.range('B8').value = Dist_Units
Feuil_DATA_BRUT.range('B9').value = Fs_Radar

#ajout des données 'Samples' :
Feuil_DATA_BRUT.range('D1').value = 'Sample'
Feuil_DATA_BRUT.range('D2:D%d' %int(len(Sample)+1)).options(transpose=True).value = Sample

#ajout des données 'Temps' :
Feuil_DATA_BRUT.range('E1').value = 'Temps'
Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).options(transpose=True).value = Temps

#ajout des données 'Vitesse' :
Feuil_DATA_BRUT.range('F1').value = 'Vitesse (%s)' %Speed_Units
Feuil_DATA_BRUT.range('F2:F%d' %int(len(Sample)+1)).options(transpose=True).value = Vitesse_KMH

#ajout des données 'Accel' :
Feuil_DATA_BRUT.range('G1').value = 'Accel (%s)' %Accel_Units
Feuil_DATA_BRUT.range('G2:G%d' %int(len(Sample)+1)).options(transpose=True).value = Accel

#ajout des données 'Dist' :
Feuil_DATA_BRUT.range('H1').value = 'Dist (%s)' %Dist_Units
Feuil_DATA_BRUT.range('H2:H%d' %int(len(Sample)+1)).options(transpose=True).value = Dist

#ajout des données 'Vitesse m/s' :
Feuil_DATA_BRUT.range('I1').value = 'Vitesse (m/s)'
Feuil_DATA_BRUT.range('I2:I%d' %int(len(Sample)+1)).options(transpose=True).value = Vitesse_MS

#changer le format d'affichage des valeurs/données dans les différentes colonnes du feuillet DATA_BRUT :
Feuil_DATA_BRUT.range('E2:I%d' %int(len(Sample)+1)).number_format = '0,00'

#Commande pour rechercher des plots existants ou nombre de plot existants:
#wb.sheets[1].charts #wb.sheets[1].charts.count

#Plot directement dans Excel des valeurs de  : Vitesses (unités = brut), Accel (g), distance.
Chart_Brut = Feuil_DATA_BRUT.charts #Handle pour ajouter des objet chart sur le feuillet : DATA_BRUT
Chart_Brut.add(700, 50, 600, 400)
Chart_Brut[0].name = 'Chart_Brut'
Chart_Brut[0].chart_type = 'line'

# Ajout de la série VitesseBrut :
Chart_Brut[0].api[1].SeriesCollection().NewSeries()  
Chart_Brut[0].api[1].SeriesCollection(1).Name= 'Vitesse Brut'
Chart_Brut[0].api[1].SeriesCollection(1).XValues = Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).api
Chart_Brut[0].api[1].SeriesCollection(1).Values = Feuil_DATA_BRUT.range('F2:F%d' %int(len(Sample)+1)).api

# Ajout de la série Accel_Brut :
Chart_Brut[0].api[1].SeriesCollection().NewSeries()  
Chart_Brut[0].api[1].SeriesCollection(2).Name= 'Accel_Brut'
Chart_Brut[0].api[1].SeriesCollection(2).XValues = Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).api
Chart_Brut[0].api[1].SeriesCollection(2).Values = Feuil_DATA_BRUT.range('G2:G%d' %int(len(Sample)+1)).api

# Ajout de la série Distance_Brut :
Chart_Brut[0].api[1].SeriesCollection().NewSeries()  
Chart_Brut[0].api[1].SeriesCollection(3).Name= 'Distance_Brut'
Chart_Brut[0].api[1].SeriesCollection(3).XValues = Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).api
Chart_Brut[0].api[1].SeriesCollection(3).Values = Feuil_DATA_BRUT.range('H2:H%d' %int(len(Sample)+1)).api

# wb.save()

############### Recherche du départ du sprint :
############ Automatiquement : --> 1) Determine une moyenne glissante, 2)recherche max sur la moyenne glissante, 3) recherche via récursivité inverse la v1ère valeur en dessous d'un seuil fixé en dur...

#### Il faudrait mettre en place un GUI pour vérifier le point de départ du sprint et pouvoir le valider et/ou le changer si pas ok #### 

#Determination de la moyenne glissante de la vitesse sur 50 points de données (Pourquoi ?) :
Nb_Points = 50
MA_Vitesse = moving_average(Vitesse_MS, Nb_Points)

#ajout des données 'Vitesse m/s' :
Feuil_DATA_BRUT.range('J1').value = 'MG_Vitesse (m/s)'
Feuil_DATA_BRUT.range('J2:J%d' %int(len(Sample)+1)).options(transpose=True).value = MA_Vitesse

#plot sur Excel de la vitesse m/s + Moyenne glissante sur x points de ces données :
# Chart_Brut = Feuil_DATA_BRUT.charts #Handle pour gerer les objets charts sur le feuillet : DATA_BRUT
Chart_Brut.add(700, 550, 600, 400)
Chart_Brut[1].name = 'Chart_V_m/s_MA_V_m/s'
Chart_Brut[1].chart_type = 'line'

# Ajout de la série VitesseBrut en m/s :
Chart_Brut[1].api[1].SeriesCollection().NewSeries()  
Chart_Brut[1].api[1].SeriesCollection(1).Name= 'Vitesse_m/s_Brut'
Chart_Brut[1].api[1].SeriesCollection(1).XValues = Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).api
Chart_Brut[1].api[1].SeriesCollection(1).Values = Feuil_DATA_BRUT.range('I2:I%d' %int(len(Sample)+1)).api

# Ajout de la série moyenne glissante de la vitesse Brut :
Chart_Brut[1].api[1].SeriesCollection().NewSeries()  
Chart_Brut[1].api[1].SeriesCollection(2).Name= 'Moyenne glissante \n Vitesse_m/s_Brut'
Chart_Brut[1].api[1].SeriesCollection(2).XValues = Feuil_DATA_BRUT.range('E2:E%d' %int(len(Sample)+1)).api
Chart_Brut[1].api[1].SeriesCollection(2).Values = Feuil_DATA_BRUT.range('J2:J%d' %int(len(Sample)+1)).api

# 1) Recherche du max de vitesse et Index de la valeur Max + Max de la Moyenne Glissante de la vitesse avec l'index  :
index_MAX_Vitesse = np.argmax(Vitesse_MS )
index_MAX_MA_Vitesse = np.argmax(MA_Vitesse)

# #Visu moyenne glissante :
# plt.figure()
# plt.plot(Temps,Vitesse_MS, 'r')
# plt.plot(Temps,MA_Vitesse, 'b')
# plt.axvline(Temps[index_MAX_MA_Vitesse], color = 'black', linestyle = '--')
# plt.legend(['Vitesse Réel (m/s)', 'Mean Average Speed', 'Ligne Vert. Vmax Sprint'])
# plt.xlabel('Temps (secondes)')
# plt.ylabel('Vitesse (m/s)')
# plt.title('Vitesse (m/s) réelle et moyenne glissante (%d points) \n en fonction du Temps (sec) de Début_Sprint : Vmax Sprint' %Nb_Points)
# plt.show()

# #Via Plotly sur le navigateur web par dféaut ou ouvert :
# fig = go.Figure(data=[{'type': 'scatter', 'y': Vitesse_MS}])
# fig.add_trace(go.Scatter(y=MA_Vitesse))
# plot(fig)


#find 1st point above a limit : on fixe le seuil de départ de sprint à 0.2 --> la première valeur avant le max qui est inf à 0.2 m/s = valeur de Dbt de sprint : 
i = index_MAX_MA_Vitesse
while Vitesse_MS[i] > 0.5 :
    # print(i)
    i -= 1

#Index de début de Sprint :
DBT = i
#Index d'atteinte de Vmax (--> determiné à partir de la moyenne glissante sur x points de la vitesse en m/s... Autre solution ?? )
FIN = index_MAX_MA_Vitesse

######################## GUI Pour selectionner et/ou valider à la main le début du sprint : #######################
Master = tk.Tk()
ex = Plot_Embed(Master,Temps, Vitesse_MS, Temps[DBT], Temps[FIN])
ex.connect() #appel la connection au canvas.mpl_connect()
ex.display_plot() #display tkinter windows
coor_xy_START_STOP = ex.return_coordinates()

#visualisation des bornes déterminées automatiquement + Vitesse (m/s)  et loyenne glissante vitesse (m/s) 
# fig1 = plt.figure()
# plt.plot(Temps,Vitesse_MS, 'r')
# plt.plot(Temps,MA_Vitesse, 'b')
# plt.axvline(Temps[i], color = 'black', linestyle = '-.', label = 'Dbt Sprint')
# plt.axvline(Temps[index_MAX_MA_Vitesse], color = 'black', linestyle = '--', label = 'Vmax Sprint')
# plt.legend(['Vitesse Réel (m/s)', 'Mean Average Speed','Ligne V Départ Sprint', 'Ligne V Vmax Sprint'])
# plt.xlabel('Temps (secondes)')
# plt.ylabel('Vitesse (m/s)')
# plt.title('Vitesse (m/s) réelle et moyenne glissante (%d points) \n en fonction du Temps (sec) de Début_Sprint : Vmax Sprint' %Nb_Points)
# plt.show()

#Copy d'une figure/plot python dans le feuillet excel 
# Feuil_DATA_BRUT.pictures.add(fig1, name='Plot_Borne_Sprint', update=True,
#                      left=Feuil_DATA_BRUT.range('V36').left, top=Feuil_DATA_BRUT.range('V36').top)


#Add nouveaux feuillets excel au workbook --> Données FV-Sprint :
wb.sheets.add(name= 'Profil Force Vitesse', after = wb.sheets[1]) #Open new sheet and make this sheet :active sheet.
Feuil_DATA_Acc = wb.sheets('Profil Force Vitesse')

#Fréquence d'échantillon du radar :
Freq_Echantillon = 48.785 #SAMPLE_RATE
#--> Attention, on a eu des bizarreries sur la fréquence d'échantillon ............!!!!!
#--> Normalement Fs =48.785 sur ce radar ATS Pro 2

np.where(coor_xy_START_STOP[0] <= Temps)[0][0]
np.where(coor_xy_START_STOP[2] <= Temps)[0][0]

#Vitesse et Acceleration --> Seulement sur la phase d'acceleration du sprint jusqu'à Vmax pour le calcul du FV Sprint :
Acc_Vitesse_MS = Vitesse_MS[np.where(coor_xy_START_STOP[0] <= Temps)[0][0] : np.where(coor_xy_START_STOP[2] <= Temps)[0][0]]
#On recrée le vecteur Temps sur la phase d'accéleration :
Acc_Temps  = np.linspace(0, (1/Freq_Echantillon)*len(Acc_Vitesse_MS),len(Acc_Vitesse_MS), endpoint=True) 
#Attention , pb entre le temps réel mesuré par le radar et le temps recrée avec linspace : 
#--> Semble que le radar ou les données du fichier Rad, ne soit pas à une fréquence d'acquisition = à 46.875 comme définit pas le constructeur...
#Du au fait que l'on supprime des valeurs de vitesse ??? 

#Write data of start to Vmax --> to the sheet :
#ajout des données 'Temps :
Feuil_DATA_Acc.range('A1').value = 'Temps_Acc'
Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).options(transpose=True).value = Acc_Temps 

#ajout des données 'Acc_Vitesse_MS' :
Feuil_DATA_Acc.range('B1').value = 'Vitesse Réelle (m/s)'
Feuil_DATA_Acc.range('B2:B%d' %int(len(Acc_Vitesse_MS)+1)).options(transpose=True).value = Acc_Vitesse_MS

#plot/graph on excel --> phase acceleration :
Chart_Acc = Feuil_DATA_Acc.charts #Handle pour ajouter des objet chart sur le feuillet : DATA_BRUT
Chart_Acc.add(Feuil_DATA_Acc.range('R2').left, Feuil_DATA_Acc.range('R2').height, 600, 400)
Chart_Acc[0].name = 'Chart_Acc'
Chart_Acc[0].chart_type = 'line'

# Ajout de la série VitesseBrut :
Chart_Acc[0].api[1].SeriesCollection().NewSeries()  
Chart_Acc[0].api[1].SeriesCollection(1).Name= 'Vitesse Phase Acc'
Chart_Acc[0].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[0].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('B2:B%d' %int(len(Acc_Vitesse_MS)+1)).api


#plot du sprint --> Seulement sur la partie Accéleration :
plt.figure()
plt.plot(Acc_Temps,Acc_Vitesse_MS, 'r')
plt.legend(['Vitesse Réel/Mesuré (m/s) '])
plt.xlabel('Temps (secondes)')
plt.ylabel('Vitesse (m/s)')
plt.title('Vitesse (m/s) réelle mesuré en fonction du Temps (sec) sur la phase d\'accéleration jusqu\'a Vmax ')
plt.show()


#Initial solution = les valeurs vers lesquelles nous souhaitons que les coeficients se rapprochent.
Val_Initiale = np.array([8, 0, 1], dtype=float) 

COEFFS, flag = leastsq(minimisation, Val_Initiale,args=(Acc_Vitesse_MS,Acc_Temps)) #obj de leastsq sera de minimiser les écarts entre y (vitesse réel) et les valeurs de VitesseModélisé2 (via la détermination des 3 paramètres), et cela en partant des valeurs initiales spécifié précèdemment
print("les coefficients définis par la minimisation (leastsq) sont : Vmax = %.2f, t = %.2f, Tau = %.2f" % (COEFFS[0],COEFFS[1],COEFFS[2])) 

    #avec curve_fit de la bibliothèque de scipy --> Use non-linear least squares to fit a function, f, to data.
#Fit for the parameters a, b, c of the function func: 
# popt, pcov = curve_fit(VitesseModélisé, Acc_Temps, Acc_Vitesse_MS)
# print("les coefficients définis par la minimisation (curve_fit) sont : Vmax = %.2f, t = %.2f, Tau = %.2f" % (popt[0],popt[1],popt[2])) 


#ajout des données  valeurs de la modélisation dans le classeur excel d'export :
Feuil_DATA_Acc.range('M1').value = 'Vmax'
Feuil_DATA_Acc.range('N1').value = COEFFS[0]

Feuil_DATA_Acc.range('M2').value = 'Delay'
Feuil_DATA_Acc.range('N2').value = COEFFS[1]

Feuil_DATA_Acc.range('M3').value = 'Tau'
Feuil_DATA_Acc.range('N3').value = COEFFS[2]


################### Vitesse_Modélisé avec les bon coefficients : ###############
VITESSE_MODELISE1 = VitesseModélisé(Acc_Temps,COEFFS[0],COEFFS[1],COEFFS[2])
# VITESSE_MODELISE2 = VitesseModélisé2(Acc_Temps,popt)

#Write data of start to Vmax --> to the sheet :
#ajout des données 'Acc_Vitesse_MS' :
Feuil_DATA_Acc.range('C1').value = 'VITESSE_MODELISE (m/s)'
Feuil_DATA_Acc.range('C2').value = ['=$N$1 * (1-EXP(-(A2-$N$2)/$N$3))'] ;#Inscrit la formule pour le calcul de la vitesse modélisé 
Feuil_DATA_Acc.range('C2').api.AutoFill(Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api,0) ;#Autofill pour étendre la formule à chaque ligne 

#Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).options(transpose=True).value = VITESSE_MODELISE1 #--> Si on veut ecrire directement les valeurs de la vitesse modélisé dans excel depuis python sans la formules, en brut. 
#coeffs[0]*(1-np.exp(-((x-coeffs[1])/coeffs[2]))) 

#calcul de la distance parcourue à chaque instant t du début du sprint à Vmax : Pb est que on calcule la distance parcourure que jusqu'à Vmax (que l'on a determiné comme valeur max de la Vitesse moyenne glissante...)
distance1 = COEFFS[0] * ((Acc_Temps-COEFFS[1])+COEFFS[2]*np.exp(-((Acc_Temps-COEFFS[1])))) - COEFFS[0]*COEFFS[2] #Vmax*((Time)+Tau*exp(-(Time/Tau)))-Vmax*Tau ;

Feuil_DATA_Acc.range('D1').value = 'Distance (m)'
Feuil_DATA_Acc.range('D2:D%d' %int(len(distance1)+1)).options(transpose=True).value =  distance1
Feuil_DATA_Acc.range('D2').value = ['=$N$1 * (($A2-$N$2) + $N$3 * EXP(-((A2-$N$2)))) - $N$1 * $N$3'] #Inscrit la formule pour le calcul de la vitesse modélisé 
Feuil_DATA_Acc.range('D2').api.AutoFill(Feuil_DATA_Acc.range('D2:D%d' %int(len(distance1)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 

#changer le format d'affichage des valeurs/données dans les différentes colonnes du feuillet Profil Force Vitesse :
Feuil_DATA_Acc.range('A2:D%d' %int(len(Acc_Vitesse_MS)+1)).number_format = '0,00'

#plot 
plt.figure()
plt.plot(Acc_Temps,distance1)
# plt.plot(Acc_Temps,distance2)
plt.legend(['distance1 (Leastsq)','distance2 (curve_fit)'])
plt.xlabel('Temps (secondes)')
plt.ylabel('Distance (m)')
plt.title('Distance parcourue en mètre en fonction du temps, du début du Sprint jusqu\'à Vmax')
plt.show()
print("théoriquement à Vmax au bout de %d mètres " % distance1[-1])

#plot/graph on excel --> phase acceleration :
# Ajout de la série VITESSE_MODELISE1 au grahique de la vitesse brut sur la phase d'accéleration :
Chart_Acc[0].api[1].SeriesCollection().NewSeries()  
Chart_Acc[0].api[1].SeriesCollection(2).Name= 'VITESSE_MODELISE1'
Chart_Acc[0].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(VITESSE_MODELISE1)+1)).api
Chart_Acc[0].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api

# Ajout de la série  :
Chart_Acc[0].api[1].SeriesCollection().NewSeries()  
Chart_Acc[0].api[1].SeriesCollection(3).Name= 'Distance (m)'
Chart_Acc[0].api[1].SeriesCollection(3).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[0].api[1].SeriesCollection(3).Values = Feuil_DATA_Acc.range('D2:D%d' %int(len(distance1)+1)).api
Chart_Acc[0].api[1].SeriesCollection(3).AxisGroup = 2

################### Visualisation de la minimisation pour trouver les paramètres de la vitesse modélisée : #################################################################################################################################
#Matplotlib :
#fitting on acceleration phase :
plt.figure()
plt.plot(Acc_Temps,Acc_Vitesse_MS, 'b')
#visualisation avec les bonnes valeurs de coefficients "a,b,c" :
plt.plot(Acc_Temps,VITESSE_MODELISE1, 'k--')
# #visualisation avec les bonnes valeurs de coefficients "a,b,c" :
# plt.plot(Acc_Temps,VITESSE_MODELISE2, 'r--')
plt.legend(['Vitesse Réel', 'VitesseModélisé1 (Leastsq)','VitesseModélisé2 (curve_fit)'])
plt.xlabel('Temps (secondes)')
plt.ylabel('Vitesse (m/s)')
plt.title('Vitesse (m/s) réelle et modélisée en fonction du Temps (sec) de Début_Sprint jusqu\'à Vmax Sprint')
plt.show()

################### Détermination de la fin du sprint (cherche DBT phase de deceleration/fin de sprint) #################
#Donnée en gardant du début du sprint jusqu'a la dernière donnée ... 
#(genre faudrer trover un moyen de prendre jusqu'à la dernière donnée ou il pousse, soit de couper dès lors que le mec est en deceleration)
Vitesse_Sprint_ALL = Vitesse_MS[DBT:]
MA_Vitesse_ALL = moving_average(Vitesse_Sprint_ALL, 50)
Temps_ALL  = np.linspace(0, (1/Freq_Echantillon)*len(Vitesse_Sprint_ALL),len(Vitesse_Sprint_ALL), endpoint=True)
VitesseMod_Sprint_ALL = VitesseModélisé(Temps_ALL,COEFFS[0],COEFFS[1],COEFFS[2])

# #Calcul de la différence entre MG_V et Mod_V  après Vmax ... :
Vmax_MA = np.where(MA_Vitesse_ALL == np.amax(MA_Vitesse_ALL))
Diff_VMG_Vmod = abs(VitesseModélisé(Temps_ALL,COEFFS[0],COEFFS[1],COEFFS[2]) - MA_Vitesse_ALL) #différence entre vitesse modélisé et vitesse MA

#Add nouveaux feuillets excel au workbook --> Données Sprint (Dbt jusqu'à deceleration) :
wb.sheets.add(name= 'DATA_Sprint_ALL', after = wb.sheets[2]) #Open new sheet and make this sheet :active sheet.
Feuil_DATA_Sprint_All = wb.sheets('DATA_Sprint_ALL')

#Write data of start to end of data --> to the sheet 'DATA_Sprint_ALL'
#ajout des données 'Temps_ALL :
Feuil_DATA_Sprint_All.range('A1').value = 'Temps_Sprint_All'
Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).options(transpose=True).value = Temps_ALL

#ajout des données 'Vitesse_Sprint_ALL ' :
Feuil_DATA_Sprint_All.range('B1').value = 'Vitesse_Sprint_All'
Feuil_DATA_Sprint_All.range('B2:B%d' %int(len(Vitesse_Sprint_ALL )+1)).options(transpose=True).value = Vitesse_Sprint_ALL 

#ajout des données 'Vitesse_Sprint_ALL_Moyenne glissante  ' :
Feuil_DATA_Sprint_All.range('C1').value = 'MG_Vitesse_Sprint_All'
Feuil_DATA_Sprint_All.range('C2:C%d' %int(len(MA_Vitesse_ALL )+1)).options(transpose=True).value = MA_Vitesse_ALL

#ajout des données 'Vitesse_Sprint_ALL_Moyenne glissante  ' :
Feuil_DATA_Sprint_All.range('D1').value = 'VitesseMod_Sprint_ALL'
Feuil_DATA_Sprint_All.range('D2:D%d' %int(len(VitesseMod_Sprint_ALL)+1)).options(transpose=True).value = VitesseMod_Sprint_ALL

#ajout des données 'Vitesse_Sprint_ALL_Moyenne glissante  ' :
Feuil_DATA_Sprint_All.range('E1').value = 'Diff_Vit._MoyG'
Feuil_DATA_Sprint_All.range('E2:E%d' %int(len(Diff_VMG_Vmod)+1)).options(transpose=True).value = Diff_VMG_Vmod


#idem mais calcul distance jusqu'à la dernière donnée enregistrée ... attention, tjrs selon la vitesse modélisé :
distance3 = COEFFS[0] * ((Temps_ALL-COEFFS[1])+COEFFS[2]*np.exp(-((Temps_ALL-COEFFS[1])))) - COEFFS[0]*COEFFS[2] #Vmax*((Time)+Tau*exp(-(Time/Tau)))-Vmax*Tau ;


plt.figure()
plt.plot(Temps_ALL,distance3,'b')
# plt.plot(Temps_ALL,distance4,'g')
plt.legend(['distance3 (Leastsq)'])
plt.xlabel('Temps (secondes)')
plt.ylabel('Distance (m)')
plt.title('Distance parcourue en mètre en fonction du temps, \n du début du Sprint jusqu\'à la fin dernière donnée enregistrée (selon le modèle de la vitesse...)')
plt.show()

#Ajout de la distance total du départ jusqu'à la dernière donnée enregistrée :
Feuil_DATA_Sprint_All.range('F1').value = 'Distance'
Feuil_DATA_Sprint_All.range('F2:F%d' %int(len(distance3)+1)).options(transpose=True).value =  distance3


#changer le format d'affichage des valeurs/données dans les différentes colonnes du feuillet DATA_Sprint_ALL :
Feuil_DATA_Sprint_All.range('A2:F%d' %int(len(Temps_ALL)+1)).number_format = '0,00'

#plot/graph on excel --> phase acceleration jusqu'à end data:
Chart_Sprint_All = Feuil_DATA_Sprint_All.charts #Handle pour ajouter des objet chart sur le feuillet : Feuil_DATA_Sprint_All
Chart_Sprint_All .add(Feuil_DATA_Sprint_All.range('K2').left, Feuil_DATA_Sprint_All.range('K2').height, 600, 400)
Chart_Sprint_All [0].name = 'Chart_Acc'
Chart_Sprint_All [0].chart_type = 'line'

# Ajout de la série VitesseBrut :
Chart_Sprint_All [0].api[1].SeriesCollection().NewSeries()  
Chart_Sprint_All [0].api[1].SeriesCollection(1).Name= 'Vitesse All Sprint '
Chart_Sprint_All [0].api[1].SeriesCollection(1).XValues = Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).api
Chart_Sprint_All [0].api[1].SeriesCollection(1).Values = Feuil_DATA_Sprint_All.range('B2:B%d' %int(len(Vitesse_Sprint_ALL)+1)).api

# Ajout de la série "MA_Vitesse_ALL" :
Chart_Sprint_All [0].api[1].SeriesCollection().NewSeries()  
Chart_Sprint_All [0].api[1].SeriesCollection(2).Name= 'Vitesse MG All '
Chart_Sprint_All [0].api[1].SeriesCollection(2).XValues = Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).api
Chart_Sprint_All [0].api[1].SeriesCollection(2).Values = Feuil_DATA_Sprint_All.range('C2:C%d' %int(len(MA_Vitesse_ALL)+1)).api

# Ajout de la série "MA_Vitesse_ALL" :
Chart_Sprint_All [0].api[1].SeriesCollection().NewSeries()  
Chart_Sprint_All [0].api[1].SeriesCollection(3).Name= 'Vitesse Mod All '
Chart_Sprint_All [0].api[1].SeriesCollection(3).XValues = Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).api
Chart_Sprint_All [0].api[1].SeriesCollection(3).Values = Feuil_DATA_Sprint_All.range('D2:D%d' %int(len(VitesseMod_Sprint_ALL)+1)).api


#plot/graph on excel --> phase acceleration jusqu'à end data:
Chart_Sprint_All.add(Feuil_DATA_Sprint_All.range('U2').left, Feuil_DATA_Sprint_All.range('U2').top, 600, 400)
Chart_Sprint_All [1].name = 'Chart_Diff'
Chart_Sprint_All [1].chart_type = 'line'
# Ajout de la série Diff :
Chart_Sprint_All [1].api[1].SeriesCollection().NewSeries()  
Chart_Sprint_All [1].api[1].SeriesCollection(1).Name= 'Diff Vit_Mod_All / Vit_MG'
Chart_Sprint_All [1].api[1].SeriesCollection(1).XValues = Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).api
Chart_Sprint_All [1].api[1].SeriesCollection(1).Values = Feuil_DATA_Sprint_All.range('E2:E%d' %int(len(Vitesse_Sprint_ALL)+1)).api

Seuil_Diff_MA = np.argmax(Diff_VMG_Vmod[Vmax_MA[0][0]:] >= np.mean(Diff_VMG_Vmod)) #on cherche la valeur de diff entre VMod et V_MG qui sera superieur à la moyenne des différences sur tout le graph
# np.mean(Diff_VMG_Vmod)

#plot pour voir l'index dont la valeur de y et ymodélisé sont différents de +1 m/s
fig2 = plt.figure()
plt.plot(Temps_ALL,Vitesse_Sprint_ALL, 'b')
#visualisation avec les bonnes valeurs de coefficients "a,b,c" :
plt.plot(Temps_ALL,VitesseModélisé(Temps_ALL,COEFFS[0],COEFFS[1],COEFFS[2]), 'r--')
plt.plot(Temps_ALL,MA_Vitesse_ALL, 'k-.')
plt.axvline(Temps_ALL[np.where(MA_Vitesse_ALL == np.amax(MA_Vitesse_ALL))], color ='black' , alpha = 0.6, linestyle = ':')
plt.axvline(Temps_ALL[Vmax_MA[0][0] + Seuil_Diff_MA], color ='black', alpha = 0.6, linestyle = ':')
#visualisation avec les bonnes valeurs de coefficients "a,b,c" :
plt.legend(['Vitesse Réel', 'VitesseModélisé', 'Moyenne glissante Vitesse', 'Verticale Line Vmax', 'Verticale Line End Sprint (Supposed)'])
plt.xlabel('Temps (secondes)')
plt.ylabel('Vitesse (m/s)')
plt.title('Vitesse (m/s) réelle et modélisée en fonction du Temps (sec) de Début_Sprint : Last_Data')
plt.show()

Feuil_DATA_Sprint_All.pictures.add(fig2, name='Plot_VR_VM_Diff', update=True,
                     left=Feuil_DATA_Sprint_All.range('U4').left, top=Feuil_DATA_Sprint_All.range('U4').top) ;

#plot/graph on excel --> phase acceleration jusqu'à end data:
Chart_Sprint_All .add(Feuil_DATA_Sprint_All.range('K30').left, Feuil_DATA_Sprint_All.range('K30').top, 600, 400)
Chart_Sprint_All [2].name = 'Chart_Distance'
Chart_Sprint_All [2].chart_type = 'line'
# Ajout de la série Distance :
Chart_Sprint_All [2].api[1].SeriesCollection().NewSeries()  
Chart_Sprint_All [2].api[1].SeriesCollection(1).Name= 'Distance'
Chart_Sprint_All [2].api[1].SeriesCollection(1).XValues = Feuil_DATA_Sprint_All.range('A2:A%d' %int(len(Temps_ALL)+1)).api
Chart_Sprint_All [2].api[1].SeriesCollection(1).Values = Feuil_DATA_Sprint_All.range('F2:F%d' %int(len(distance3)+1)).api

#Notation des temps à chaque dizaine de mètre :
if max(distance1) > 30 :
    Temps10m = Temps_ALL[np.where(distance3 <= 10)][-1]
    Temps20m = Temps_ALL[np.where(distance3 <= 20)][-1]
    Temps30m = Temps_ALL[np.where(distance3 <= 30)][-1]
    Temps40m = Temps_ALL[np.where(distance3 <= 40)][-1]
    Temps50m = Temps_ALL[np.where(distance3 <= 50)][-1]
    Temps60m = Temps_ALL[np.where(distance3 <= 60)][-1]

elif max(distance1) < 30:
    Temps10m = Temps_ALL[np.where(distance3 <= 10)][-1]
    Temps20m = Temps_ALL[np.where(distance3 <= 20)][-1]
    Temps30m = Temps_ALL[np.where(distance3 <= 30)][-1]
        
    
#Val_Anthro = [62.9, 1.75, 19, 1, 750]

#Si pas de boite de dialogue --> faut rentrer les valeurs en dur en dessous... pas idéal pour enchainer automatiquement les traitements :
#Valeurs anthroppo du sujet :
MasseSujet = value_entry[0]
TailleSujet = value_entry[1] #attention, TailleSujet en (m) 
Temperature = value_entry[2]
Vent = value_entry[3]
Pb = value_entry[4]

# on ajoute nos valeur pour tous ceux qui est Force de resistance à l'air (temperature, pression barométrique)
Cd = 0.9
Af = (0.2025*float(TailleSujet)**0.725*float(MasseSujet)**0.425)*0.266 
rho = 1.293*float(Pb)/760*273/(273+float(Temperature)) 
K = 0.5*rho*Af*Cd 
 

#Ajout de ces données dans le excel d'export : Anthropo /envirronement
Feuil_DATA_Acc.range('M5').value = 'Taille (cm)'
Feuil_DATA_Acc.range('N5').value = TailleSujet 

Feuil_DATA_Acc.range('M6').value = 'Masse (Kg)'
Feuil_DATA_Acc.range('N6').value = MasseSujet

Feuil_DATA_Acc.range('M7').value = 'Temperature (°C)'
Feuil_DATA_Acc.range('N7').value = Temperature

Feuil_DATA_Acc.range('M8').value = 'Pression (mmHg)'
Feuil_DATA_Acc.range('N8').value = Pb 

Feuil_DATA_Acc.range('M9').value = 'Coefficient de penetration'
Feuil_DATA_Acc.range('N9').value = Cd

Feuil_DATA_Acc.range('M10').value = 'Af :'
Feuil_DATA_Acc.range('N10').value = ['=(0.2025 * ($N$5)^0.725 * $N$6^0.425) * 0.266']

Feuil_DATA_Acc.range('M11').value = 'Rho :'
Feuil_DATA_Acc.range('N11').value = ['=1.293 * $N$8 / 760 * 273 / (273 + $N$7)']

Feuil_DATA_Acc.range('M12').value = 'K:'
Feuil_DATA_Acc.range('N12').value = ['= 0.5 * $N$11 * $N10 * $N$9']

Feuil_DATA_Acc.range('M13').value = 'Vent (m/s)'
Feuil_DATA_Acc.range('N13').value = Vent


# calcul de l'accélération horizontale réel 
Acc_H = np.zeros(len(Acc_Vitesse_MS), dtype=float)
for i in range(1,len(Acc_Vitesse_MS)-1) :
    Acc_H [i] = (Acc_Vitesse_MS[i+1]-Acc_Vitesse_MS[i])/(Acc_Temps[i+1]-Acc_Temps[i]);
#plot Acc horizontale réelle via python :
plt.plot(Acc_Temps,Acc_H)


#calcul accélération horizontale modélisé
Acc_H_Mod = (COEFFS[0]/COEFFS[2])*np.exp(-((Acc_Temps-COEFFS[1])/COEFFS[2]));


# calcul force horizontale à partir Acc_H_Mod
F_Hor = float(MasseSujet) * Acc_H_Mod + K * (VITESSE_MODELISE1-float(Vent))**2;

#Calcul de la force normalisée au poids du sujet.
F_Hor_Poids = F_Hor/float(MasseSujet)

# Calcule de la puissance avec P = V_Mod*F
Puissance = VITESSE_MODELISE1 * F_Hor_Poids;

#ecriture dans excel :
#ajout des données 'Acc_Horizontale' :
Feuil_DATA_Acc.range('E1').value = 'Acc_H_Mod (m/s²)'
# Feuil_DATA_Acc.range('E2:E%d' %int(len(Acc_H_Mod)+1)).options(transpose=True).value = Acc_H_Mod
Feuil_DATA_Acc.range('E2').value = ['=$N$1/$N$3 * EXP(-(($A2-$N$2)/$N$3))'] #formule calcul en mou dans excel de l'acceleration horizontale 
Feuil_DATA_Acc.range('E2').api.AutoFill(Feuil_DATA_Acc.range('E2:E%d' %int(len(Acc_H_Mod)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 

#ajout des données 'Force_Horizontale' : Total qui inclue également les forces de friction de l'air + vent associé.
Feuil_DATA_Acc.range('F1').value = 'Force_Horizontale (N)'
# Feuil_DATA_Acc.range('F2:F%d' %int(len(F_Hor)+1)).options(transpose=True).value = F_Hor
Feuil_DATA_Acc.range('F2').value = ['=$N$6 * $E2 + $N$12 * ($C2-$N$13)^2'] #formule calcul en mou dans excel de l'acceleration horizontale 
Feuil_DATA_Acc.range('F2').api.AutoFill(Feuil_DATA_Acc.range('F2:F%d' %int(len(F_Hor)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 


#ajout des données 'Force_Horizontale_PDC' :
Feuil_DATA_Acc.range('G1').value = 'Force_Horizontale Relative (N/Kg)'
# Feuil_DATA_Acc.range('G2:G%d' %int(len(F_Hor_Poids)+1)).options(transpose=True).value = F_Hor_Poids
Feuil_DATA_Acc.range('G2').value = ['=$F2 / $N$6'] #formule calcul en mou dans excel de l'acceleration horizontale 
Feuil_DATA_Acc.range('G2').api.AutoFill(Feuil_DATA_Acc.range('G2:G%d' %int(len(F_Hor_Poids)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 

#ajout des données 'Puissance' :
Feuil_DATA_Acc.range('H1').value = 'Puissance Horizontale (W/kg)'
# Feuil_DATA_Acc.range('H2:H%d' %int(len(Puissance)+1)).options(transpose=True).value = Puissance
Feuil_DATA_Acc.range('H2').value = ['=$C2 * $G2'] #formule calcul en mou dans excel de la puissance 
Feuil_DATA_Acc.range('H2').api.AutoFill(Feuil_DATA_Acc.range('H2:H%d' %int(len(Puissance)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 


#changer le format d'affichage des valeurs/données dans les différentes colonnes du feuillet Feuil_DATA_Acc :
Feuil_DATA_Acc.range('E2:H%d' %int(len(Puissance)+1)).number_format = '0,00'

#plot dans execl :

# Ajout de la série 'Acc_H_Mod' au grahique de la vitesse brut sur la phase d'accéleration :
Chart_Acc[0].api[1].SeriesCollection().NewSeries()  
Chart_Acc[0].api[1].SeriesCollection(4).Name= 'Acc_H_Mod'
Chart_Acc[0].api[1].SeriesCollection(4).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_H_Mod)+1)).api
Chart_Acc[0].api[1].SeriesCollection(4).Values = Feuil_DATA_Acc.range('E2:E%d' %int(len(Acc_H_Mod)+1)).api

#plot Acc horizontal réel et Acc horizontal mod:
plt.figure()
plt.plot(Acc_Temps,Acc_H)
plt.plot(Acc_Temps,Acc_H_Mod )
plt.legend(['Acc-Horizontal Réel','Acc-Horizontal Modélisé'])
plt.xlabel('Temps (secondes)')
plt.ylabel('Acc (m/s²)')
plt.title('Acceleration Horizontal au cours du sprint jusqu\'à Vmax')
plt.show()

#plot/graph on excel --> phase acceleration :
Chart_Acc.add(Feuil_DATA_Acc.range('AB2').left, Feuil_DATA_Acc.range('AB2').top, 600, 400)
Chart_Acc[1].name = 'Paramètres sprint modélisés'
Chart_Acc[1].chart_type = 'line'

# Ajout de la série For horizontale relative au pdc :
Chart_Acc[1].api[1].SeriesCollection().NewSeries()  
Chart_Acc[1].api[1].SeriesCollection(1).Name= 'Force Horizontale (N/kg)'
Chart_Acc[1].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[1].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('G2:G%d' %int(len(F_Hor_Poids)+1)).api

# Ajout de la série Vitesse Modélisée :
Chart_Acc[1].api[1].SeriesCollection().NewSeries()  
Chart_Acc[1].api[1].SeriesCollection(2).Name= 'Vitesse modélisée (m/s)'
Chart_Acc[1].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[1].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api

# Ajout de la série  Puissance Horizontale :
Chart_Acc[1].api[1].SeriesCollection().NewSeries()  
Chart_Acc[1].api[1].SeriesCollection(3).Name= 'Puissance (W/Kg)'
Chart_Acc[1].api[1].SeriesCollection(3).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[1].api[1].SeriesCollection(3).Values = Feuil_DATA_Acc.range('H2:H%d' %int(len(Puissance)+1)).api


#Visualisation matplotlib :
plt.figure()
plt.plot(Acc_Temps, F_Hor_Poids)
plt.plot(Acc_Temps, VITESSE_MODELISE1)
plt.plot(Acc_Temps, Puissance)
plt.legend(['Force Horizontale Réelle','Vitesse modélisée','Puissance-Horizontal Modélisé'])
plt.xlabel('Temps (secondes)')
plt.ylabel('N/Kg, m/s, m/s²') #pas sur des Unités 
plt.title('Force, Vitesse et Acceleration horizontale')
plt.show()

#calcul V0 (m/s), F0 (N), Sfv, R² 
#[polyFunction] = polyfitT(VitesseModélisé,F_Hor,1);

PFV_Sprint = np.polyfit(VITESSE_MODELISE1, F_Hor, 1)
PFV_Sprint_Relative = np.polyfit(VITESSE_MODELISE1, F_Hor_Poids, 1)

F0 = PFV_Sprint[1]
Sfv = PFV_Sprint[0]
V0 = -F0/Sfv
P0 = F0*V0/4
Vopt = VITESSE_MODELISE1[np.argmax(Puissance, axis=0)] #Vopt = Vitesse modélisé à Puissance max 


#ajout des données valeurs de pentes + extreme (F0/V0 du PFV) :
Feuil_DATA_Acc.range('M15').value = 'F0 (N)'
Feuil_DATA_Acc.range('N15').value = F0 #attention F0 poids en valeurs absolu

Feuil_DATA_Acc.range('M16').value = 'F0 (N/Kg)'
Feuil_DATA_Acc.range('N16').value = F0/float(MasseSujet)

Feuil_DATA_Acc.range('M17').value = 'V0 (m/s)'
Feuil_DATA_Acc.range('N17').value = V0

Feuil_DATA_Acc.range('M18').value = 'P0 (W)'
Feuil_DATA_Acc.range('N18').value = P0

Feuil_DATA_Acc.range('M19').value = 'P0 (W/kg)'
Feuil_DATA_Acc.range('N19').value = P0/float(MasseSujet)

Feuil_DATA_Acc.range('M20').value = 'Sfv'
Feuil_DATA_Acc.range('N20').value = Sfv

Feuil_DATA_Acc.range('M23').value = 'Vopt'
Feuil_DATA_Acc.range('N23').value = Vopt

# Feuil_DATA_Acc.range('N24').value = ["=MAX(R1:R171)"]
# Feuil_DATA_Acc.range('N25').value = ["=MATCH(N24;R1:R171;0)"]
# Feuil_DATA_Acc.range('N26').value = ["=INDEX(P1:P171;N25)"]
#les valeurs de Sfv et F0 (forcement)) seront différentes selon que l'on calcul le profil en prenant les valeurs de force en absolu ou relative au poids de corps.
Feuil_DATA_Acc.range('M15:N21').number_format = '0,00'


#plot/graph on excel --> phase acceleration :
Chart_Acc.add(Feuil_DATA_Acc.range('R29').left, Feuil_DATA_Acc.range('R29').top, 600, 400)
Chart_Acc[2].name = 'PFV during Sprint'
Chart_Acc[2].chart_type = 'xy_scatter'

# Ajout de la série "Force-Vitesse" :
Chart_Acc[2].api[1].SeriesCollection().NewSeries()  
Chart_Acc[2].api[1].SeriesCollection(1).Name= 'Force-Velocity Profil'
Chart_Acc[2].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('C2:C%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[2].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('G2:G%d' %int(len(F_Hor_Poids)+1)).api
#--> Information sur le graphique à complèter (titre, noms des axes, unités...)
#--> Voir : https://docs.xlwings.org/en/stable/api.html


#Calcul de la Force total en N
Ftot = np.sqrt((F_Hor)**2 + (float(MasseSujet)*9.81)**2);

#ajout des données 'Ftot' :
Feuil_DATA_Acc.range('I1').value = 'Force Totale (N)'
# Feuil_DATA_Acc.range('L2:L%d' %int(len(Ftot)+1)).options(transpose=True).value = Ftot
Feuil_DATA_Acc.range('I2').value = ['= SQRT(($F2)^2 + ($N$6 *9.81)^2)']
Feuil_DATA_Acc.range('I2').api.AutoFill(Feuil_DATA_Acc.range('I2:I%d' %int(len(Ftot)+1)).api,0); #Autofill pour étendre la formule à chaque ligne 

#Calcul de RF, et cela lorsque le temps (t)>= 0.3s
Find_t = np.where(Acc_Temps >= 0.3 ) #cherche les index où les valeurs de la variables Temps sont supérieurs à 0.3s :

Début_RF = Find_t[0][0] #Garde la première valeur d'index superieur à 0.3s de la variable Temps 

#Calcul RF après 0,3 sec du temps de sprint (Pourquoi ? à,3s ? c'est ce qui est fait dans le classeur témoin de J.B Morin + programme scilab M2 EOPS Saint etienne)
RF = F_Hor[Début_RF:]/Ftot[Début_RF:]
#RF = F_Hor/Ftot
plt.plot(Acc_Temps[Début_RF:],RF)

#ajout des données 'Ftot' :
Feuil_DATA_Acc.range('J1').value = 'RF (%)'
Feuil_DATA_Acc.range('J%d' %(Début_RF+2)).value = '=F%d / I%d' %(Début_RF+2,Début_RF+2)
Feuil_DATA_Acc.range('J%d' %(Début_RF+2)).api.AutoFill(Feuil_DATA_Acc.range('J%d:J%d' %(Début_RF+2, int(len(RF) + Début_RF +1))).api,0); #Autofill pour étendre la formule à chaque ligne 
# Feuil_DATA_Acc.range('J%d:J%d' %(Début_RF, int(len(RF) + Début_RF +1))).options(transpose=True).value = RF

#changer le format d'affichage des valeurs/données dans les différentes colonnes du feuillet Feuil_DATA_Acc :
Feuil_DATA_Acc.range('M%d:M%d' %(Début_RF+2, int(len(RF) + Début_RF +1))).number_format = '0,00%'

#plot/graph on excel --> Plot RF au cours du sprint :
Chart_Acc.add(Feuil_DATA_Acc.range('AB29').left, Feuil_DATA_Acc.range('AB29').top, 600, 400)
Chart_Acc[3].name = 'RF during Sprint'
Chart_Acc[3].chart_type = 'xy_scatter'

# Ajout de la série "Ratio de force (horizontale) durant le sprint" :
Chart_Acc[3].api[1].SeriesCollection().NewSeries()  
Chart_Acc[3].api[1].SeriesCollection(1).Name= 'RF during sprint'
Chart_Acc[3].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api
Chart_Acc[3].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('J2:J%d' %int(len(VITESSE_MODELISE1)+1)).api
# mettre les valeurs en pourcentage --> soit directement sur python, soit ensuite sur excel 

Feuil_DATA_Acc.range('M21').value = 'RF max'
Feuil_DATA_Acc.range('N21').value =  ['=MAX(J%d:J%d)' %(Début_RF+2, int(len(RF) + Début_RF +1))]

Drf = np.polyfit(VITESSE_MODELISE1[Début_RF:], RF, 1)

Feuil_DATA_Acc.range('M22').value = 'Drf'
Feuil_DATA_Acc.range('N22').value = Drf[0]
Feuil_DATA_Acc.range('N22').number_format = '0,00%'


# # Trentem = np.where(distance1 >= 30)[0][0] #Temps au dessus de 30m selon le temps (Temps durant phase acceleration)
# print("Temps au 30m (selon Vecteur Temps sur la phase d'acceleration) = %.2f  sec" % Temps_ALL[np.where(distance1 >= 30)[0][0]]) #Temps a 30m 
# #Trentem = np.where(distance3 >= 30)[0][0] #Temps au dessus de 30m 
# print("Temps au 30m (selon Vecteur Temps sur tout le sprint) = %.2f  sec" % Temps_ALL[np.where(distance3 >= 30)[0][0]]) #Temps a 30m  
# # Soixantem = np.where(distance3 >= 60)[0][0] #Temps au dessus de 60m 
# print("Temps au 30m (selon Vecteur Temps sur tout le sprint) = %.2f  sec" % Temps_ALL[np.where(distance3 >= 60)[0][0]]) #Temps a 30m  


#calcul RF sur 10m et moyenne 
# Find_10m = np.where(distance1 <= 10 ) 
# Mean_RF_10m = np.mean(RF[0:Find_10m[0][-1]])

#calcul des valeurs de la relation puissance-vitesse --> Puissance absolu mais faire aussi puissance relative au PDC ?!
Relation_Pui_V = np.polyfit(VITESSE_MODELISE1,Puissance,2)

#plot vitesse modélisé vs puissance :
plt.figure()
plt.plot(VITESSE_MODELISE1,Puissance)

# Calcul Pmax_Théorique
Vopt = -Relation_Pui_V[1]/(2*Relation_Pui_V[2])
Pmax_Théorique = Relation_Pui_V[2]* (Vopt)**2 + Relation_Pui_V[1] * Vopt + Relation_Pui_V[0]

# Ajout de la série "Puissance-Vitesse" au graphique : Chart_Acc[2].name = 'PFV during Sprint'
Chart_Acc[2].api[1].SeriesCollection().NewSeries()  
Chart_Acc[2].api[1].SeriesCollection(2).Name= 'Power-Velocity Profil'
Chart_Acc[2].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('C2:C%d' %int(len(Acc_Temps)+1)).api
Chart_Acc[2].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('H2:H%d' %int(len(F_Hor_Poids)+1)).api
#--> Information sur le graphique à complèter (titre, noms des axes, unités...)
#--> Voir : https://docs.xlwings.org/en/stable/api.html

#calcul Puissance_Modélisé
Puissance_Modélisé = Relation_Pui_V[2] * (VITESSE_MODELISE1)**2 + Relation_Pui_V[1] * VITESSE_MODELISE1 + Relation_Pui_V[0]

#reconstruit le spectre de force-Vitesse et Puissance-Vitesse de F0 à V0 et de V = 0 jusqu'à V0 --> pour la visualisation
Feuil_DATA_Acc.range('P1').value = 0
Feuil_DATA_Acc.range('Q1').value =  Feuil_DATA_Acc.range('N16').value
Feuil_DATA_Acc.range('R1').value =  0

Feuil_DATA_Acc.range('P2').value = ['=C2'] ;#Inscrit la formule pour le calcul de la vitesse modélisé 
Feuil_DATA_Acc.range('P2').api.AutoFill(Feuil_DATA_Acc.range('P2:P%d' %int(len(VITESSE_MODELISE1)+1)).api,0) ;#Autofill pour étendre la formule à chaque ligne 

Feuil_DATA_Acc.range('Q2').value = ['=G2'] ;#Inscrit la formule pour le calcul de la vitesse modélisé 
Feuil_DATA_Acc.range('Q2').api.AutoFill(Feuil_DATA_Acc.range('Q2:Q%d' %int(len(VITESSE_MODELISE1)+1)).api,0) ;#Autofill pour étendre la formule à chaque ligne 

Feuil_DATA_Acc.range('R2').value = ['=H2'] ;#Inscrit la formule pour le calcul de la vitesse modélisé 
Feuil_DATA_Acc.range('R2').api.AutoFill(Feuil_DATA_Acc.range('R2:R%d' %int(len(VITESSE_MODELISE1)+1)).api,0) ;#Autofill pour étendre la formule à chaque ligne 

Feuil_DATA_Acc.range('P%d' %int(len(VITESSE_MODELISE1)+2)).value = Feuil_DATA_Acc.range('N17').value
Feuil_DATA_Acc.range('Q%d' %int(len(VITESSE_MODELISE1)+2)).value =  0
Feuil_DATA_Acc.range('R%d' %int(len(VITESSE_MODELISE1)+2)).value =  0

#Mise en forme des charts de la fiche principale de retour de test qui sera sorti ensuite en PDF :
Fiche_Retour_Charts = Feuil_FICHE_RETOUR.charts

# Ajout des series au Chart 1 + Mise en forme Chart 1 :
Fiche_Retour_Charts[0].name ='ForceVelocity & PowerVelocity Profile'
Fiche_Retour_Charts[0].chart_type = 'xy_scatter'
Fiche_Retour_Charts[0].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Name = 'Profil Force-Vitesse'
Fiche_Retour_Charts[0].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('P1:P%d' %int(len(Acc_Temps)+2)).api
Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('Q1:Q%d' %int(len(F_Hor_Poids)+2)).api
# Fiche_Retour_Charts[0].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('P1:P172').api
# Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('Q1:Q172').api
Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Markersize = 3 ;
Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Markerstyle = 8;


################################ 
#--> But de ce bout de Script = mettre en place des courbes de tendances sur le plot de Force-Vitesse + Puissance-Vitesse

#Fiche_Retour_Charts[0].api[1].SeriesCollection(1).HasTrendlines = 1;
# xxx = Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Trendlines.Add
# a = Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Trendlines
# Fiche_Retour_Charts[0].api[1].add_series({
#     'xvalues':    '=Profil Force Vitesse!$P$1:$P$314',
#     'values':    '=Profil Force Vitesse!$Q$1:$Q$314',
#     'trendline': {'type': 'linear'},
# })
# print(dir(Fiche_Retour_Charts[0].api[1]))
# print(getattr(Fiche_Retour_Charts[0].api[1].SeriesCollection(1), 'Trendlines'))
# print(getattr(Fiche_Retour_Charts[0].api[1].SeriesCollection(1).Trendlines, 'Add'))
# type(Fiche_Retour_Charts[0].api[1].SeriesCollection(1))
# Fiche_Retour_Charts[0].api[1].FullSeriesCollection(1).Name
# Fiche_Retour_Charts[0].api[1].FullSeriesCollection(1).Trendlines().NewSeries() 
#('Type'=xlLinear, 'Forward'= 0, 'Backward' = 0, 'DisplayEquation' = 0, 'DisplayRSquared' = 1, 'Name' = 'Linéaire (Profil Force-Vitesse)')

#print(getattr(Fiche_Retour_Charts[0].api[1].FullSeriesCollection(1).Trendlines(), '__getattr__')) #− to access the attribute of object.

################################

Fiche_Retour_Charts[0].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).Name= 'Profil Puissance-Vitesse'
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('P1:P%d' %int(len(Acc_Temps)+2)).api
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('R1:R%d' %int(len(F_Hor_Poids)+2)).api

Fiche_Retour_Charts[0].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('P1:P172').api
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('R1:R172').api

Fiche_Retour_Charts[0].api[1].SeriesCollection(2).AxisGroup = 2
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).Markersize = 3 ;
Fiche_Retour_Charts[0].api[1].SeriesCollection(2).Markerstyle = 8;

Fiche_Retour_Charts[0].api[1].Axes(1).HasTitle = True # This line creates the x axis label.
Fiche_Retour_Charts[0].api[1].Axes(2).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[0].api[1].Axes(2, 2).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[0].api[1].Axes(1).AxisTitle.Text = "Vitesse (m/s)" 
Fiche_Retour_Charts[0].api[1].Axes(2).AxisTitle.Text = "Force Horizontale (N/Kg)"
Fiche_Retour_Charts[0].api[1].Axes(2, 2).AxisTitle.Text = "Puissance Horizontale (W/Kg)"
Fiche_Retour_Charts[0].api[1].Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 0

# Ajout des series au Chart 2 + Mise en forme Chart 2 :
Fiche_Retour_Charts[1].name ='Model des caractéristiques mécaniques durant le sprint '
Fiche_Retour_Charts[1].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[1].chart_type = 'line'
Fiche_Retour_Charts[1].api[1].SeriesCollection(1).Name= 'Vitesse modélisé (m/s)'
Fiche_Retour_Charts[1].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Fiche_Retour_Charts[1].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api
Fiche_Retour_Charts[1].api[1].Axes(1).TickLabels.NumberFormat = "0.00"

Fiche_Retour_Charts[1].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[1].api[1].SeriesCollection(2).Name= 'Force Horizontale (N/Kg)'
Fiche_Retour_Charts[1].api[1].SeriesCollection(2).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Fiche_Retour_Charts[1].api[1].SeriesCollection(2).Values = Feuil_DATA_Acc.range('G2:G%d' %int(len(F_Hor_Poids)+1)).api
Fiche_Retour_Charts[1].api[1].Axes(2).TickLabels.NumberFormat = "0"
Fiche_Retour_Charts[1].api[1].Axes(2).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[1].api[1].Axes(2).AxisTitle.Text = "Force (N/kg), Vitesse (m/s)"

Fiche_Retour_Charts[1].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[1].api[1].SeriesCollection(3).Name= 'Puissance (W/Kg)'
Fiche_Retour_Charts[1].api[1].SeriesCollection(3).XValues = Feuil_DATA_Acc.range('A2:A%d' %int(len(Acc_Temps)+1)).api
Fiche_Retour_Charts[1].api[1].SeriesCollection(3).Values = Feuil_DATA_Acc.range('H2:H%d' %int(len(Puissance)+1)).api
Fiche_Retour_Charts[1].api[1].SeriesCollection(3).AxisGroup = 2
Fiche_Retour_Charts[1].api[1].Axes(2,2).TickLabels.NumberFormat = "0"
Fiche_Retour_Charts[1].api[1].Axes(2, 2).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[1].api[1].Axes(2, 2).AxisTitle.Text = "Puissance Horizontale (W/Kg)"
Fiche_Retour_Charts[1].api[1].Axes(2, 2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = 0

# Ajout des series au chart 3 + Mise en forme Chart 3 :
Fiche_Retour_Charts[2].name ='Orientation de la force (RF et Drf) '
Fiche_Retour_Charts[2].api[1].SeriesCollection().NewSeries()  
Fiche_Retour_Charts[2].api[1].SeriesCollection(1).Name= 'Ratio de Force (%)'
Fiche_Retour_Charts[2].api[1].SeriesCollection(1).XValues = Feuil_DATA_Acc.range('C2:C%d' %int(len(VITESSE_MODELISE1)+1)).api
Fiche_Retour_Charts[2].api[1].SeriesCollection(1).Values = Feuil_DATA_Acc.range('J2:J%d' %int(len(VITESSE_MODELISE1)+1)).api
Fiche_Retour_Charts[2].api[1].SeriesCollection(1).Markersize = 3 ;
Fiche_Retour_Charts[2].api[1].SeriesCollection(1).Markerstyle = 8;
Fiche_Retour_Charts[2].api[1].Axes(1).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[2].api[1].Axes(1).AxisTitle.Text = "Vitesse (m/s)"
Fiche_Retour_Charts[2].api[1].Axes(2).TickLabels.NumberFormat = "0%"
Fiche_Retour_Charts[2].api[1].Axes(2).HasTitle = True # This line creates the Y axis label.
Fiche_Retour_Charts[2].api[1].Axes(2).AxisTitle.Text = "Ratio de Force (%)"
Fiche_Retour_Charts[2].api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10
Fiche_Retour_Charts[2].api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10
Fiche_Retour_Charts[2].api[1].Legend.Position = -4160

#copie des datas FV dans la fiche résumé :
Feuil_FICHE_RETOUR.range('I18').value = MasseSujet #Masse sujet 


Feuil_FICHE_RETOUR.range('C42').value = ["='Profil Force Vitesse'!N15"] #F0 (N)
Feuil_FICHE_RETOUR.range('C43').value = ["='Profil Force Vitesse'!N16"] #FO (N/Kg)
Feuil_FICHE_RETOUR.range('C44').value = ["='Profil Force Vitesse'!N17"] #VO (m/s)
Feuil_FICHE_RETOUR.range('C45').value = ["='Profil Force Vitesse'!N23"] #Vopt
Feuil_FICHE_RETOUR.range('C46').value = ["='Profil Force Vitesse'!N18"] #P0 (W)
Feuil_FICHE_RETOUR.range('C47').value = ["='Profil Force Vitesse'!N19"] #P0 (W/Kg)


Feuil_FICHE_RETOUR.range('C64').value = Feuil_DATA_Acc.range('N21').value #RF max (%)
Feuil_FICHE_RETOUR.range('C65').value = Feuil_DATA_Acc.range('N22').value # Drf
Feuil_FICHE_RETOUR.range('C66').value = max(VITESSE_MODELISE1)
Feuil_FICHE_RETOUR.range('C64:C65').number_format = '0,00%'

Feuil_FICHE_RETOUR.range('C67').value = Temps30m
if 'Temps60m' in globals() :
    Feuil_FICHE_RETOUR.range('C68').value = Temps60m


#exit(0)
#quitter le workbook et close excel application :
wb.save()
wb.app.quit()
Excel_App.quit()
Excel_App.kill()
del Excel_App


# getattr(obj, name[, default]) − to access the attribute of object.

#******************************************************************************

## Python tips --> Comprehensive list :
# x = [i+1 for i, ltr in enumerate(PATH) if ltr == '.'] #
#print(x)


#Via Plotly sur le navigateur web par dféaut ou ouvert :
# fig = go.Figure(data=[{'type': 'scatter', 'x' : Acc_Temps ,'y': Acc_Vitesse_MS}])
# fig.add_trace(go.Scatter({'x' : Temps[DBT:FIN] ,'y' : Vitesse_MS[DBT:FIN]}))
# plot(fig)