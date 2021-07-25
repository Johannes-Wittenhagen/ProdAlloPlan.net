import webbrowser
import os
from tkinter import *
from PIL import ImageTk, Image
from tkinter import filedialog
import pandas as pd
import PAPModel_Sensitivity_Hour
import PAPModel_Sensitivity_Transport
import PAPModel_Sensitivity_Demand
import PAPModel_Sensitivity_Supplier
import PAPModel_Sensitivity_Relocation
import ProdAlloPlanNet_Model




# Erstellung des Hauptfensters, in dem die Excel übergeben wird
main_window = Tk()
main_window.title("ProdAlloPlan.net Optimierungsmodell")
main_window.resizable(False, False)
main_window.iconbitmap("C:/GUI/PycharmProjects/ProdAlloPlan.net/images/BVL123_RZ_Logo_rgb.ico")



# Einfügen der BVL und ProdAlloPlan-Logos
image_prodalloplan = ImageTk.PhotoImage(Image.open("images/ProdAlloPlan.net_Logo.png"))
image_bvl = ImageTk.PhotoImage(Image.open("images/BVL123_RZ_Logo_rgb.png"))
image_labelbvl = Label(image=image_bvl)
image_labelbvl.grid(row=0, column=0, sticky="nw")
image_labelpap = Label(image=image_prodalloplan)
image_labelpap.grid(row=0, column=2, columnspan=2, sticky="ne")



# Erstellung der einzelnen Frames des Hauptfensters


# Begrüßungslabel
welcome_label = Label(main_window,
            text="\n\nDas Optimierungsmodell wurde am wbk Institut für Produktionstechnik des Karlsruher Institut"
                 " für Technologie (KIT)\n im Rahmen des Forschungsprojekts ProdAlloPlan.net entwickelt. Das "
                 "IGF-Vorhaben 20467 N der Bundesvereinigung\n Logistik (BVL) e. V., Schlachte 31, 28195 Bremen wurde "
                 "über die AiF im Rahmen des Programms zur Förderung der\n industriellen Gemeinschaftsforschung (IGF) "
                 "vom Bundesministerium für Wirtschaft und Energie (BMWi) aufgrund eines\n Beschlusses des Deutschen "
                 "Bundestages gefördert.\n\n")
welcome_label.grid(row=1, column=0, columnspan=3)


# Frame zur Eingabe der Datei
uebergabe_frame = LabelFrame(main_window, text="Übergabe der Exceldatei", padx=50, pady=50)
uebergabe_frame.grid(row=2, column=0, columnspan=3)



# Funktion, die ermöglicht Excel-Dateien zu übergeben
def open_file():
    main_window.filename = filedialog.askopenfilename(initialdir="", title="Übergabe der Eingabedatei in Excel",
                                                      filetypes=(("Excel-Dateien", "*.xlsx"), ("alle Dateien", "*.*")))
    global excel_file_name
    excel_file_name = main_window.filename
    excel_name = Label(uebergabe_frame, text=excel_file_name)
    excel_name.grid(row=0, column=0)

# Funktion, die ermöglicht Excel-Dateien zu übergeben
def open_sensitivity_file():
   main_window.filename = filedialog.askopenfilename(initialdir="", title="Übergabe der Sensitivitäts-Eingabedatei in Excel",
                                                      filetypes=(("Excel-Dateien", "*.xlsx"), ("alle Dateien", "*.*")))
   global excel_sensitivity_file_name
   excel_sensitivity_file_name = main_window.filename



auswahl_Button = Button(uebergabe_frame, text="Datei auswählen", command=open_file)
auswahl_Button.grid(row=1, column=0)

hinweis_label = Label(uebergabe_frame, text="Bitte beachten: Das Optimierungsmodell funktioniert nur mit \n"
                                           "nach dem Anwenderleitfaden formatierten Excel-Dateien")
hinweis_label.grid(row=2, column=0)

#Auswahlboxen, hinsichtlich was optimiert werden soll

optimize_cost = IntVar()
optimize_cust = IntVar()
optimize_extra = IntVar()

checkbox_cost = Checkbutton(uebergabe_frame, text="Optimierung hinsichtlich Gesamtkosten", variable=optimize_cost)
checkbox_cost.grid(row=3, column=0)
checkbox_cost.deselect()
checkbox_cust = Checkbutton(uebergabe_frame, text="Optimierung hinsichtlich Kundennähe", variable=optimize_cust)
checkbox_cust.grid(row=4, column=0)
checkbox_cust.deselect()
checkbox_extra = Checkbutton(uebergabe_frame, text="Optimierung hinsichtlich Nacharbeit", variable=optimize_extra)
checkbox_extra.grid(row=5, column=0)
checkbox_extra.deselect()


# Neues Fenster und die verschiedenen Funktionen des neuen Fensters
# Power BI öffnen



#Funktion, die später als Schnittstelle zu POWER BI dient
def open_powerbi():
    webbrowser.open('https://powerbi.microsoft.com')


#Funktion, die bei Problemen Website öffnet
def problem():
    webbrowser.open('https://wbk.kit.edu')


#Funktion, die Excel öffnet
def open_excel(excel_file_result_name):
    os.system('start excel.exe ' + excel_file_result_name)


def func(value):
    print(value)

# Post-optimale Analyse starten

def poptana(model, excel_sensitivity_file_name):

    #Erstellen eines neuen Fensters zur Ergebnisvermittlung
    poptana_window = Toplevel()
    poptana_window.title("Ergebnisse der post-optimalen Analyse")
    poptana_window.iconbitmap("C:/GUI/PycharmProjects/ProdAlloPlan.net/images/BVL123_RZ_Logo_rgb.ico")
    poptana_window.resizable(False, False)

    image_result_labelbvl = Label(poptana_window, image=image_bvl)
    image_result_labelbvl.grid(row=4, column=0, columnspan=2)
    image_result_labelpap = Label(poptana_window, image=image_prodalloplan)
    image_result_labelpap.grid(row=0, column=0, columnspan=2)



    #Aufruf des Modells basierend auf Nutzerwahl

    if model == "Hour":
        poptana_result_file_name = PAPModel_Sensitivity_Hour.hour(excel_file_name, excel_sensitivity_file_name, optimize_cost, optimize_cust, optimize_extra)
    if model == "Relocation":
        poptana_result_file_name = PAPModel_Sensitivity_Relocation.relocation(excel_file_name, excel_sensitivity_file_name, optimize_cost, optimize_cust, optimize_extra)
    if model == "Demand":
        poptana_result_file_name = PAPModel_Sensitivity_Demand.demand(excel_file_name, excel_sensitivity_file_name, optimize_cost, optimize_cust, optimize_extra)
    if model == "Supplier":
        poptana_result_file_name = PAPModel_Sensitivity_Supplier.supplier(excel_file_name, excel_sensitivity_file_name, optimize_cost, optimize_cust, optimize_extra)
    if model == "Transport":
        poptana_result_file_name = PAPModel_Sensitivity_Transport.transport(excel_file_name, excel_sensitivity_file_name, optimize_cost, optimize_cust, optimize_extra)

     # Knopf zum Anzeigen in Excel
    excel_button = Button(poptana_window, text="\n         Ergebnis in Excel anzeigen          \n",
                          command=lambda: open_excel(poptana_result_file_name))
    excel_button.grid(row=2, column=0)

    # Knopf zum Anzeigen in POWER BI
    powerbi_button = Button(poptana_window, text="\n   Ergebnis in MS Power BI anzeigen   \n",
                            command=lambda: open_powerbi())
    powerbi_button.grid(row=3, column=0)


    # Knopf zum Melden eines Problems
    problem_button = Button(poptana_window, text="\n                 Problem melden                  \n",
                            command=lambda: problem())
    problem_button.grid(row=3, column=1)


def poptana_menu():

    #Erstellen eines neuen Fensters für die Post-optimale Analyse
    poptana_menu_window = Toplevel()
    poptana_menu_window.title("Post-optimale Analyse")
    poptana_menu_window.iconbitmap("C:/GUI/PycharmProjects/ProdAlloPlan.net/images/BVL123_RZ_Logo_rgb.ico")
    poptana_menu_window.resizable(False, False)

    image_result_labelpap = Label(poptana_menu_window, image=image_prodalloplan)
    image_result_labelpap.grid(row=0, column=0, columnspan=3)

    poptana_uebergabe_frame = LabelFrame(poptana_menu_window, text="Übergabe der Exceldatei mit den Sensitivitätsparametern",
                                 padx=50, pady=50)
    poptana_uebergabe_frame.grid(row=2, column=0, columnspan=3)

    poptana_auswahl_Button = Button(poptana_uebergabe_frame, text="Datei auswählen",
                                    command=lambda: open_sensitivity_file())
    poptana_auswahl_Button.grid(row=1, column=0)

    poptana_hinweis_label = Label(poptana_uebergabe_frame, text="Bitte beachten:\n"
                                                " Die Post-optimale Analyse funktioniert nur mit \n"
                                                "nach dem Anwenderleitfaden formatierten Excel-Dateien")
    poptana_hinweis_label.grid(row=2, column=0)

    #Erstellen eines Auswahlmenüs
    model = StringVar()
    options = [
        "Hour",
        "Relocation",
        "Demand",
        "Supplier",
        "Transport"
    ]
    model = StringVar()
    model.set(options[0])

    # Knopf zum Melden eines Problems
    problem_button = Button(poptana_menu_window, text="Problem melden", command=lambda: problem())
    problem_button.grid(row=3, column=0)

    #Text Label zur Abfrage
    poptana_label = Label(poptana_menu_window, text="Post-optimale Analyse hinsichtlich:")
    poptana_label.grid(row=3, column=1)

    #Auswahl der Optionen bei der Post-optimalen Analyse
    poptana_menu = OptionMenu(poptana_menu_window, model, *options, command=func(options))
    poptana_menu.grid(row=3, column=2)


    #Start der post-optimalen Analyse
    poptana_button = Button(poptana_menu_window, text="\nPost-optimale Analyse starten\n",
                            command=lambda: poptana(model.get(), excel_sensitivity_file_name))
    poptana_button.grid(row=4, column=0, columnspan=3)



#Start der Berechnung
def start():
    #Erstellen eines neuen Fensters zum Anzeigen von Ergebnissen
    result_window = Toplevel()
    result_window.title("Ergebnisse des Optimierungsmodells")
    result_window.iconbitmap("C:/GUI/PycharmProjects/ProdAlloPlan.net/images/BVL123_RZ_Logo_rgb.ico")
    result_window.resizable(False, False)

    image_result_labelbvl = Label(result_window, image=image_bvl)
    image_result_labelbvl.grid(row=4, column=0, columnspan=2)
    image_result_labelpap = Label(result_window, image=image_prodalloplan)
    image_result_labelpap.grid(row=0, column=0, columnspan=2)

    # Knopf zum Anzeigen in Excel
    excel_button = Button(result_window, text="\n         Ergebnis in Excel anzeigen          \n",
                          command=lambda: open_excel(excel_file_result_name))
    excel_button.grid(row=2, column=0)

    # Knopf zum Anzeigen in POWER BI
    powerbi_button = Button(result_window, text="\n   Ergebnis in MS Power BI anzeigen   \n",
                            command=lambda: open_powerbi())
    powerbi_button.grid(row=3, column=0)

    # Knopf zum Starten der Post-optimalen Analyse
    poptana_button = Button(result_window, text="\n     Post-Optimale Analyse starten      \n",
                            command=lambda: poptana_menu())
    poptana_button.grid(row=2, column=1)

    # Knopf zum Melden eines Problems
    problem_button = Button(result_window, text="\n                 Problem melden                  \n",
                            command=lambda: problem())
    problem_button.grid(row=3, column=1)

    #Start der Berechnung als Übergabe der Parameter an das Programm
    excel_file_result_name = ProdAlloPlanNet_Model.pathfinder(excel_file_name, optimize_cost, optimize_cust, optimize_extra)
    print(excel_file_result_name)


    #Speichern der verschiedenen Kosten für Pie-Chart

    #Materialkosten
    #material_costs = (pd.read_excel(excel_file_result_name, sheet_name='Kosten',
     #                         index_col=[0], header=0, skiprows=0, usecols='C'))
    #sum_costs_material = material_costs.sum(axis=1, numeric_only=True)
    #print(material_costs)

    #Ressourcenkosten
    #resource_costs = (pd.read_excel(excel_file_result_name, sheet_name='Kosten',
     #                      index_col=[0], header=0, skiprows=0, usecols='I'))
    #sum_costs_resource = resource_costs.sum(axis=1, numeric_only=True)
    #print(resource_costs)

    #Bearbeitungskosten
    #processing_costs = (pd.read_excel(excel_file_result_name, sheet_name='Kosten',
     #                      index_col=[0], header=0, skiprows=0, usecols='J'))
    #sum_costs_processing = processing_costs.sum(axis=1, numeric_only=True)

    #Erstellung der Pie-Chart
    #labels = 'Materialkosten', 'Ressourcenkosten', 'Bearbeitungskosten',
    #sizes = [45, 25, 30]
    #explode = (0, 0.1, 0)

    #fig1, ax1 = plt.subplots()
    #ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
    #        shadow=True, startangle=90)
    #ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

    #plt.show()

    #plot = sumcosts.plot.pie(subplots=True)
    #pie_frame = LabelFrame(result_window, text='Plot Area')
    #pie_frame.grid(row=1, column=0, columnspan=3, padx=3, pady=3)
    #canvas = FigureCanvasTkAgg(plot, master=pie_frame)
    #canvas.show()
    #canvas.get_tk_widget().grid(row=2, column=0)



#Knopf, der die Berechnung des Modells startet

start_button = Button(main_window, text="\n         Berechnung starten        \n", command=start)
start_button.grid(row=3, column=0, columnspan=3)

main_window.mainloop()
