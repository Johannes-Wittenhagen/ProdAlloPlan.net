# -*- coding: utf-8 -*-
"""
Optimierungsmodell zur integrierten Produktallokation und Netzwerkkonfiguration
in globalen Wertschöpfungsnetzwerken

@author: felix.kerndl
"""

# Import PuLP
import pulp

#import pulp as pl
# Import pandas
import pandas as pd

# Import xlsxWriter
import xlsxwriter

# Import math
import math

import numpy as np

def relocation(name, name2, opt_1, opt_2, opt_3):
    global opt1
    opt1 = opt_1
    global opt2
    opt2 = opt_2
    global opt3
    opt3 = opt_3

    #####Einlesen aller Eingabedaten
    #Excel sheet das eingelesen werden soll
    #Pfade auf dem Mac mit "/", auf Windows mit "\"
    excel_file_name=name

    #speichert die Anzahl Perioden, die im Modell betrachtet werden sollen
    periods = (pd.read_excel(excel_file_name, sheet_name='Perioden', header=6)
               ).loc[(0),'Perioden']
    #print(periods)

    #speichert alle Werke und deren Parameter
    factories = pd.read_excel(excel_file_name, sheet_name='Werke',
                              index_col=[0], header=6, skiprows=0, usecols='B:K')
    #print(factories.head())

    #for  factory in factories.index:   print(factory)

    #speichert alle Lieferanten-Produktzustand Kombinationen
    suppliers= pd.read_excel(excel_file_name, sheet_name='Lieferanten',
                             index_col=[0,1], header=6, usecols='B:F')
    #print(suppliers.head())

    #speichert die Nachfrage von Kundenregionen nach einem Produktzustand in
    #einer bestimmten Periode
    demand= pd.read_excel(excel_file_name, sheet_name='Nachfrage',
                          index_col=[0,1,2], header=6, usecols='B:E')
    demand2 = demand.copy()


    #speichert die Stücklistenauflösung
    bom=pd.read_excel(excel_file_name, sheet_name="BOM",
                      index_col=[0,1,3,4], header=6, usecols='B:H')
    #print(bom.head())
    #print(bom.drop(index = 'C', level=1))
    #print(bom.xs('C', level = 1).index.get_level_values(0).unique())
    #print(bom.xs('b', level = 1).index.get_level_values(0).unique())

    ressources=pd.read_excel(excel_file_name, sheet_name="Ressourcen",
                             index_col=[0,1,], header=6, usecols='B:P')

    segments=pd.read_excel(excel_file_name, sheet_name="Segmente",
                           index_col=[0,1], header=6, usecols='B:P')

    resProd =pd.read_excel(excel_file_name, sheet_name="RessourceProdukt",
                           index_col=[0,1,2,3], header=6, usecols='B:K')
    #print(resProd.head())

    transportCon = pd.read_excel(excel_file_name, sheet_name="Transportverbindungen",
                           index_col=[0,1,2], header=6, usecols='B:H')

    #speichert alle ExterneKapa-Produktzustand Kombinationen
    externs = pd.read_excel(excel_file_name, sheet_name='ExterneKapa',
                             index_col=[0,1], header=6, usecols='B:F')

    # speichert die Skala für die Kundennähe
    customerCloseness = pd.read_excel(excel_file_name, sheet_name='Kundennähe',
                             index_col=[0,1], header=6, usecols='B:D')


    ressources['r_Stunden'].replace('1241','40',inplace=True)
    #ressources['r_Stundensatz'] = ressources['r_Stundensatz'].replace('86.8','40')
    print(ressources)

    #####Einlesen der Sensitivitätsparameter

    #excel_file_name2='20201117_sensitivity.xlsx'
    excel_file_name2=name2





    nachfrage_sensitivity=pd.read_excel(excel_file_name2, sheet_name="Nachfrage",
                             index_col=[0,1,], header=6, usecols='B:D')



    #ressources.assign(r_Stundensatz = ressources_sensitivity['r_Stundensatz'])
    #ressources['r_Stundensatz'] = ressources_sensitivity['r_Stundensatz']
    #for ID_WERK, ID_Ressourcengruppe in ressources_sensitivity:
     #   for ID_WERK, ID_Ressourcengruppe in ressources:



    #### Entscheidungsvariablen
    ## Grundmodell
    # Anazhl Einheiten eines Produktes p, die von Werk u an Kundenregion k in
    # Periode t geliefert werden
    Delivery_U_K = pulp.LpVariable.dicts('Delivery_U_K',
        ((period, factory, customer, product)
            for period, customer, product in demand.index
            for  factory in factories.index),
            lowBound = 0, cat='Integer')

    # Anzahl Einheiten eines Produktes p, die von Werk u an Werk u' (u <> u')
    # in Periode t geliefert werden
    Delivery_U_U = pulp.LpVariable.dicts('Delivery_U_U',
        ((period, sendFactory, recFactory, product)
            for period in range(1, periods + 1)
            for sendFactory in factories.index
            for recFactory in factories.index.drop(sendFactory)
            for product in bom.xs('C', level = 1).index.get_level_values(0).unique()),
            lowBound = 0, cat='Integer')


    # Anzahl Einheiten eines Produktes p, die von Lieferant l an Werk u
    # in Periode t geliefert werden
    Delivery_L_U = pulp.LpVariable.dicts('Delivery_L_U',
        ((period, supplier, factory, product)
            for period in range(1, periods + 1)
            for supplier, product in suppliers.index
            for factory in factories.index),
            lowBound = 0, cat='Integer')

    # Anzahl Einheiten eines Prodktes p, die in Werk u auf Segment s von
    # Ressourcengruppe r in Periode t gefertigt werden
    AnzProdP = pulp.LpVariable.dicts('AnzProdP',
        ((period, factory, segment, ressource, product)
            for period in range(1, periods + 1)
            for factory, segment, ressource, product in resProd.index),
            lowBound = 0, cat='Integer')




    # Anzahl regulärer Ressourcen r in Werk u in Periode t
    AnzRU = pulp.LpVariable.dicts('AnzRU',
                                  ((period, factory, ressource)
                                   for period in range(1, periods + 1)
                                   for factory, ressource in ressources.index),
                                  lowBound = 0, cat='Integer')

    # Einsatzzeit regulärer Ressourcen r für die Montage von Produktzustand p an
    # Segment s in Werk u in Periode t
    OperatingTimeR = pulp.LpVariable.dicts('OperatingTimeR',
        ((period, factory, segment, ressource, product)
            for period in range(1, periods + 1)
            for factory, segment, ressource, product in resProd.index),
            lowBound = 0, cat='Continuous')


    # Binärvariablen
    # Modus des Werks in Periode t: 0 für Werk u geschlossen, 1 für Werk u geöffnet
    ActU =  pulp.LpVariable.dicts('ActU',
                                  ((period, factory)
                                   for period in range(1, periods + 1)
                                   for factory in factories.index),
                                  cat='Binary')

    # Erweiterung für das Öffnen und Schließen von Werken
    UOpen = pulp.LpVariable.dicts('UOpen',
                                  ((period, factory)
                                   for period in range(1, periods + 1)
                                   for factory in factories.index),
                                  cat='Binary')

    UClose = pulp.LpVariable.dicts('UClose',
                                  ((period, factory)
                                   for period in range(1, periods + 1)
                                   for factory in factories.index),
                                  cat='Binary')

    # Modus des Segments s in Periode t: 0 für Segment s geschlossen,
    # 1 für Segment s geöffnet
    ActS = pulp.LpVariable.dicts('ActS',
                                  ((period, factory, segment)
                                   for period in range(1, periods + 1)
                                   for factory, segment in segments.index),
                                  cat='Binary')

    # Erweiterung für das Öffnen und Schließen von Segmenten s
    SOpen = pulp.LpVariable.dicts('SOpen',
                                  ((period, factory, segment)
                                   for period in range(1, periods + 1)
                                   for factory, segment in segments.index),
                                  cat='Binary')

    SClose = pulp.LpVariable.dicts('SClose',
                                  ((period, factory, segment)
                                   for period in range(1, periods + 1)
                                   for factory, segment in segments.index),
                                  cat='Binary')


    ## Erweiterungen
    # genutzte Geleitzeit einer Ressource r in Periode t
    GleitzeitR = pulp.LpVariable.dicts('GleitzeitR',
                                       ((period, factory, ressource)
                                        for period in range(1, periods + 1)
                                        for factory, ressource in ressources.index),
                                       cat='Continuous')

    # Anzahl Einheiten eines Produktes p, die von Lieferant l an Werk u
    # in Periode t geliefert werden
    Delivery_E_U = pulp.LpVariable.dicts('Delivery_E_U',
        ((period, extern, factory, product)
            for period in range(1, periods + 1)
            for extern, product in externs.index
            for factory in factories.index),
            lowBound = 0, cat='Integer')

    # Anzahl Ressourcen r von Werk u die in Periode t eingestellt bzw. entlassen
    # werden
    AnzRUAdpPlus = pulp.LpVariable.dicts('AnzRUAdpPlus',
                                       ((period, factory, ressource)
                                        for period in range(1, periods + 1)
                                        for factory, ressource in ressources.index),
                                       lowBound = 0, cat='Integer')

    AnzRUAdpMinus = pulp.LpVariable.dicts('AnzRUAdpMinus',
                                       ((period, factory, ressource)
                                        for period in range(1, periods + 1)
                                        for factory, ressource in ressources.index),
                                       lowBound = 0, cat='Integer')

    # Anzahl Schichten an einem Segment s
    AnzShift = pulp.LpVariable.dicts('AnzShift',
                                  ((period, factory, segment)
                                   for period in range(1, periods + 1)
                                   for factory, segment in segments.index),
                                  lowBound = 0, cat='Integer')



    ### Funktion baut das Modell auf
    # modelName: Name des Modells
    # valueFunction: Zielfunktion des Modells
    # constraint: weitere Restriktion die dem Modell hinzugefügt werden soll
    #               notwendig für das Goal Programming
    def buildModel(valueFunction, constraintNew = [], modelName = "ProdAlloPlan.net"):


        # Die milp Variable enthält das Optimierungsproblem
        milp = pulp.LpProblem(modelName)

        # Zielfunktion dem Modell hinzufügen
        milp.setObjective(valueFunction)


        # Hinzufügen einer weiteren Restriktion
        for constr in constraintNew:
            milp += constr


        #### Initialisierung der Binärvariablen
        # geöffnete/geschlossene Werke u in Periode 1
        for factory in factories.index:
            milp += ActU[(1, factory)] == factories.loc[factory, 'ini_Werk']

        # geöffnete/geschlossene Segmente in Periode 1
        for factory, segment in segments.index:
            milp += ActS[(1, factory, segment)] == segments.loc[(factory, segment), 'ini_Segment']

        # Anzahl Ressourcen
        for factory, ressource in ressources.index:
            milp += AnzRU[(1, factory, ressource)] == ressources.loc[(factory, ressource), 'ini_Ressource']

        # Anzahl Schichten
        for factory, segment in segments.index:
            milp += AnzShift[(1, factory, segment)] == segments.loc[(factory, segment), 'ini_Schicht']


        ####Restriktionen

        # Materialfluss Restriktionen
        # Nachfrage jeder Kundenzone k nach Produkt p muss in jeder Periode t erfüllt werden
        for period, customer, product in demand.index:
            milp+=(pulp.lpSum(Delivery_U_K[(period, factory, customer, product)]
                    for factory in factories.index)

                == demand.loc[(period, customer, product), 'Nachfrage']
                    if demand.index.isin([(period, customer, product)]).any() else 0)

        # Versorgung der Produktion mit Zwischenprodukten
        for period in range(1, periods + 1):
            for factory in factories.index:
                for product in bom.xs('C', level = 1).index.get_level_values(0).unique():

                    # Anzahl Einheiten eines Produktes p, die in Werk u in Perode t gefertigt werden
                    helpAssembly = 0

                    if ((factory, product) in resProd.index.droplevel([1,2]).unique()) == True:
                        for segment in resProd.xs((factory, product), level=[0,3]).index.get_level_values(0).unique():
                            for ressource in resProd.xs((factory,segment,product),level=[0,1,3]).index.get_level_values(0).unique():
                                helpAssembly += AnzProdP[(period, factory, segment, ressource, product)]

                    # Anzahl Einheiten eines Produktes p, die von Werken u'<> u in Periode t geliefert werden
                    helpReceive = 0

                    helpReceive += pulp.lpSum(Delivery_U_U[(period, sendFactory, factory, product)]
                                              for sendFactory in factories.index.drop(factory))

                    # Anzahl Einheiten eines Produktes p, die von Werken u'<> u in Periode t empfangen werden
                    helpSend = 0
                    helpSend += pulp.lpSum(Delivery_U_U[(period, factory, recFactory, product)]
                                           for recFactory in factories.index.drop(factory))

                    # Anzahl Einheiten eines Produktes p, die von Werk u an die Kunden d in Periode t geliefert werden
                    helpDemand = 0

                    if ((period, product) in demand.index.droplevel(1).unique()) == True:
                        for customer in demand.xs((period, product), level=[0,2]).index:
                            helpDemand += Delivery_U_K[(period, factory, customer, product)]

                    # Anzahl Einheiten eines Produktes p, die für Folgeprodukte p', die in Werk u in Periode t produziert werden, notwendig sind
                    helpFollowProductionStep = 0

                    for sucProduct in bom.xs(product, level=0).index.get_level_values(1).unique():
                        if sucProduct == product:
                            continue
                        if ((factory, sucProduct) in resProd.index.droplevel([1,2]).unique()) == True:
                            for segment in resProd.xs((factory, sucProduct), level=[0,3]).index.get_level_values(0).unique():
                                for ressource in resProd.xs((factory, segment, sucProduct),level=[0,1,3]).index.get_level_values(0).unique():
                                    helpFollowProductionStep += (
                                        bom.xs((product, sucProduct), level=[0,2]).loc[:,'FaktorPsucP'][0]
                                        * AnzProdP[(period, factory, segment, ressource, sucProduct)]
                                        )

                    ### Erweiterung um externe Kapazitäten
                    # Anzahl Einheiten eines Produktes p, die von externer Einheit e
                    # an Werk u in Periode t geliefert werden
                    helpExternKapa = 0
                    if product in  externs.index.get_level_values(1).unique():
                        for extern in externs.xs(product, level = 1).index:
                            helpExternKapa += (
                                Delivery_E_U[(period, extern, factory, product)]
                                )

                    milp += (
                        helpAssembly + helpReceive + helpExternKapa
                        - helpSend - helpDemand - helpFollowProductionStep
                        >= 0
                        )


        # Bedarf an Rohmaterial
        for period in range(1, periods + 1):
            for factory in factories.index:
                for product in bom.xs('M', level = 1).index.get_level_values(0).unique():

                    # Rohmaterialbedarf für die Folgeprodukte, die am Standort u in Periode t gefertigt werden
                    helpRawMaterialDemand = 0

                    for sucProduct in bom.xs(product, level=0).index.get_level_values(1).unique():
                        if sucProduct == product:
                            continue
                        if ((factory, sucProduct) in resProd.index.droplevel([1,2]).unique()) == True:
                            for segment in resProd.xs((factory, sucProduct), level=[0,3]).index.get_level_values(0).unique():
                                for ressource in resProd.xs((factory, segment, sucProduct),level=[0,1,3]).index.get_level_values(0).unique():
                                    helpRawMaterialDemand += (
                                        bom.xs((product, sucProduct), level=[0,2]).loc[:,'FaktorPsucP'][0]
                                        * AnzProdP[(period, factory, segment, ressource, sucProduct)]
                                        )

                    # Liefermenge des Produktes p an Werk u in Periode t von allen Lieferanten
                    helpRawMaterialDelivery = 0

                    if (product in suppliers.index.droplevel(0).unique()) == True:
                        for supplier in suppliers.xs(product, level = 1).index:
                            helpRawMaterialDelivery += Delivery_L_U[(period, supplier, factory, product)]

                    milp += helpRawMaterialDemand - helpRawMaterialDelivery <= 0

        # Kapazitätsbeschränkungen
        # Segmentkapazitäten
        for period in range(1, periods + 1):
            for factory in factories.index:
                if (factory in segments.index.droplevel(1).unique()) == True:
                    for segment in segments.xs(factory, level = 0).index:
                        helpUsedSegmentCap = 0
                        if ((factory, segment) in resProd.index.droplevel([2,3]).unique()) == True:
                            for (ressource, product) in resProd.xs((factory, segment), level = [0, 1]).index:
                                helpUsedSegmentCap += (
                                    resProd.loc[(factory, segment, ressource, product), 'capUSRP']
                                    * AnzProdP[(period, factory, segment, ressource, product)]
                                    )

                        helpSegmentCap = 0
                        helpSegmentCap =  (
                            ActS[(period, factory, segment)]
                            * segments.loc[(factory, segment), 'cap_Segment']
                            )

                        milp += helpUsedSegmentCap <= helpSegmentCap

                        # Erweiterung um das Schichtmodell
                        helpSegmentCap = 0
                        helpAnzSchichtenEingabe = segments.loc[(factory, segment),
                                                               'maxSchichten_Segment']
                        if helpAnzSchichtenEingabe != 0:
                            helpSegmentCap = (
                                segments.loc[(factory, segment), 'cap_Segment']
                                * AnzShift[(period, factory, segment)]
                                * (1 / helpAnzSchichtenEingabe)
                                )
                        else:
                            print('Eingabe Anzahl Schichten = 0, daher unzulässig')
                            helpSegmentCap = 0
                            AnzShift[(period, factory, segment)] = 0

                        milp += helpUsedSegmentCap <= helpSegmentCap

        # Ressourcenkapazitäten
        for period in range(1, periods + 1):
            for factory in factories.index:
                if (factory in resProd.index.droplevel([1, 2, 3]).unique()) == True:
                    for ressource in resProd.xs(factory, level = 0).index.get_level_values(1).unique():
                        # Verfügbare Zeit einer Ressourcengruppe in einer Periode t
                        helpAvailableTimeR = 0
                        if ((factory, ressource) in ressources.index) == False:
                            print('Für Werk', factory, 'ist die Ressource',
                                  ressource, 'nicht definiert.')
                        else:
                            helpAvailableTimeR = (
                                AnzRU[(period, factory, ressource)]
                                * ressources.loc[(factory, ressource), 'r_Stunden']
                                )

                        # Operative Einsatzzeit einer Ressourcengruppe in Periode t
                        helpOperatingTimeR = 0

                        for (segment, product) in resProd.xs((factory, ressource), level = [0, 2]).index:
                            helpOperatingTimeR += OperatingTimeR[(period, factory, segment, ressource, product)]

                        # Erweiterung um Gleitzeit
                        helpGleitzeit = 0
                        if ((factory, ressource) in ressources.index) == True:
                            helpGleitzeit = GleitzeitR[(period, factory, ressource)]

                        milp += helpOperatingTimeR <= helpAvailableTimeR + helpGleitzeit


        # Die Einsatzzeit einer Ressource
        for period in range(1, periods + 1):
            for (factory, segment, ressource, product) in resProd.index:
                milp += (
                    OperatingTimeR[(period, factory, segment, ressource, product)]
                    == resProd.loc[(factory, segment, ressource, product), 'capUSRP']
                                    * AnzProdP[(period, factory, segment, ressource, product)]
                         )



        # Lieferantenkapazitäten
        for period in range(1, periods + 1):
            for product in bom.xs('M', level = 1).index.get_level_values(0).unique():
                if (product in suppliers.index.droplevel(0).unique()) == True:
                    for supplier in suppliers.xs(product, level= 1).index:
                        helpCapL = 0
                        helpCapL = suppliers.loc[(supplier, product), 'capL']
                        helpRawMaterialFromSupplier = 0
                        for factory in factories.index:
                            helpRawMaterialFromSupplier += (
                                Delivery_L_U[(period, supplier, factory, product)]
                                )

                        milp += helpRawMaterialFromSupplier <= helpCapL



                else:
                    print('Für Material ', product, ' existiert kein Lieferant!')

        # Werksflächen
        for period in range(1, periods + 1):
            for factory in factories.index:
                helpAvailableSpaceU = 0
                helpUsedSpaceByS = 0

                helpAvailableSpaceU = (
                    ActU[(period, factory)]
                    * factories.loc[(factory), 'ini_space_Werk']
                    )

                if (factory in segments.index.droplevel(1).unique()) == True:
                    for segment in segments.xs(factory, level = 0).index:
                        helpUsedSpaceByS += (
                            ActS[(period, factory, segment)]
                            * segments.loc[(factory, segment), 'space_Segment']
                            )

                milp += helpUsedSpaceByS <= helpAvailableSpaceU

        # Restriktion für Transportverbindungen
        # Transportverbindungen von oder zu Werken können nur genutzt werden,
        # wenn das Werk aktiv ist
        # Definition von zwei Listen: 'transportA' und 'transportB'
        # das Tupel (transportA, transportB) stellt die jeweiligen
        # Transportverbindungen dar
        transportA = []
        transportB = []

        transportA = (
            suppliers.index.droplevel(1).unique().values.tolist()
            + externs.index.droplevel(1).unique().values.tolist()
            + factories.index.values.tolist()
            )

        transportB = (
            factories.index.values.tolist()
            + demand.index.get_level_values(1).unique().values.tolist()
            )

        # alle Produkte in der BoM
        allProductsBom = []
        allProductsBom = (
            bom.index.get_level_values(0).unique().values.tolist()
            + bom.index.get_level_values(2).unique().values.tolist()
            )

        # Duplikate entfernen
        allProductsBom = list(dict.fromkeys(allProductsBom))

        for period in range(1, periods + 1):
            for i in transportA:
                newTransportB = transportB.copy()
                # Werke liefern nicht an sich selbst
                if i in transportB:
                    newTransportB.remove(i)
                for j in newTransportB:
                    for p in allProductsBom:
                        # Transportverbindung
                        helpTransportConnection = 0
                        helpTransportConMinCap = (
                            transportCon.loc[('general', 'general', 'general'),
                                             'cap_min_Transportverbindung']
                            )
                        helpTransportConMaxCap = (
                            transportCon.loc[('general', 'general', 'general'),
                                             'cap_max_Transportverbindung']
                            )
                        # check if index (period, i, j, p) existiert in den
                        #Entscheidungsvariablen
                        if (period, i, j, p) in Delivery_L_U:
                            helpTransportConnection = (
                                Delivery_L_U[(period, i, j, p)]
                                )

                        elif (period, i, j, p) in Delivery_U_U:
                            helpTransportConnection = (
                                Delivery_U_U[(period, i, j, p)]
                                )

                        elif (period, i, j, p) in Delivery_U_K:
                            helpTransportConnection = (
                                Delivery_U_K[(period, i, j, p)]
                                )

                        elif (period, i, j, p) in Delivery_E_U:
                            helpTransportConnection = (
                                Delivery_E_U[(period, i, j, p)]
                                )
                        else:
                            continue

                        # Existiert die Transportverbindung in der Nutzereingabe
                        if ((p, i, j) in transportCon.index) == True:
                            # check if Min Cap is not NULL
                            if pd.isnull(
                                    transportCon.loc[(p, i, j),
                                                     'cap_min_Transportverbindung'
                                                     ]) == False:
                                helpTransportConMinCap = (
                                transportCon.loc[(p, i, j), 'cap_min_Transportverbindung']
                                )
                            # check if Max Cap is not NULL
                            if pd.isnull(
                                    transportCon.loc[(p, i, j),
                                                     'cap_max_Transportverbindung'
                                                     ]) == False:
                                helpTransportConMaxCap = (
                                transportCon.loc[(p, i, j), 'cap_max_Transportverbindung']
                                )

                        # Tranpsortverbindung existiert in einer der Entscheidungsvariablen

                        statusFactory = 0
                        if i in factories.index:
                            statusFactory = ActU[(period, i)]
                        elif j in factories.index:
                            statusFactory = ActU[(period, j)]
                        else:
                            print('Transportverbindung nicht zulässig: (', i, ', ', j,')' )


                        milp += (
                            -(statusFactory * helpTransportConMaxCap)
                            + helpTransportConnection
                            <= 0
                                 )


                        milp += (
                            helpTransportConnection
                            >= statusFactory * helpTransportConMinCap
                            )


        ### Erweiterungen
        ## Flexibilität
        # Externe Kapazitäten: Kapaziätsrestriktion
        for period in range(1, periods + 1):
            for extern, product in externs.index:
                helpUsedExternCap = 0
                helpAvailableExternCap = 0

                for factory in factories.index:
                    helpUsedExternCap += (
                        Delivery_E_U[(period, extern, factory, product)]
                        )

                helpAvailableExternCap += (
                    externs.loc[(extern, product), 'capE']
                    )

                milp += helpUsedExternCap <= helpAvailableExternCap

        # Gleitzeit
        # Restriktion der Gleitzeit je Periode
        for period in range(1, periods + 1):
            for factory, ressource in ressources.index:
                helpUsedFlexTime = 0
                helpAvailableFlexTime = 0

                helpUsedFlexTime = (
                    GleitzeitR[(period, factory, ressource)]
                    )

                helpAvailableFlexTime = (
                    AnzRU[(period, factory, ressource)]
                    * ressources.loc[(factory, ressource), 'r_maxGZ']
                    )

                milp += helpUsedFlexTime <= helpAvailableFlexTime
                milp += helpUsedFlexTime >= -(helpAvailableFlexTime)

        ## Restriktion der Gleitzeit über einen Zyklus hinweg
        # Variablen und definierte Listen werden im weitern Programmablauf genutzt
        # (Kosten und Ausgabe)
        # Perioden in Gleitzeitzyklus
        periodsInCylce = int(ressources.loc[(factory, ressource),
                                                        'r_Perioden_in_Zyklus'])
        # maximale Zyklenanzahl
        maxCycle = math.floor(periods/periodsInCylce)
        # Zyklusliste erstellen
        listCycle = [periodsInCylce * i for i in range(0, maxCycle + 1)]

        for factory, ressource in ressources.index:
            # alle Zyklen durchlaufen
            for cyc in listCycle:
                helpFTEndOfCycle = 0
                helpAvailableFTEndOfCycle = 0
                for per in range(1, periodsInCylce + 1):
                    helpFTEndOfCycle += (
                        GleitzeitR[((cyc + per), factory, ressource)]
                        )

                    helpAvailableFTEndOfCycle += (
                        AnzRU[((cyc + per), factory, ressource)]
                        )

                    #Abbruchbedingung wenn maximale Periodenzahl des Modells erreicht wird
                    if ((cyc + per) == periods):
                        break

                helpAvailableFTEndOfCycle = (
                    helpAvailableFTEndOfCycle
                    * (1/periodsInCylce)
                    * ressources.loc[(factory, ressource), 'r_maxGZ_zyklus']
                    )

                milp += helpFTEndOfCycle <= helpAvailableFTEndOfCycle

                # für Negative Gleitzeit am Zyklusende entstehen dem fokalen Unternehmen
                # keine zusätzlichen Kosten, zusätzlich soll negative Gleitzeit zum
                # Zyklusende möglichst gar nicht vorkommen.
                # Hier wird im follgenden davon ausgegangen, dass es in Summe keine
                # negative Gleitzeit am Zyklusende geben darf, daher >= 0
                milp += helpFTEndOfCycle >= 0

        ## Rekonfigurierbarkeit
        # Anpassung von Ressourcen
        # gilt für perioden >= 2
        if periods > 1:
            for period in range(2, periods + 1):
                for factory, ressource in ressources.index:
                    milp += (
                        AnzRU[(period, factory, ressource)]
                        == AnzRU[((period - 1), factory, ressource)]
                        + AnzRUAdpPlus[(period, factory, ressource)]
                        - AnzRUAdpMinus[(period, factory, ressource)]
                        )

                    # Beschränkung für Einstellungen/Entlassungen
                    milp += (
                        AnzRUAdpPlus[(period, factory, ressource)]
                        <= ressources.loc[(factory, ressource), 'r_Einstellen']
                        )
                    milp += (
                        AnzRUAdpMinus[(period, factory, ressource)]
                        <= ressources.loc[(factory, ressource), 'r_Entlassen']
                        )
        # # für periode = 1
        # for factory, ressource in ressources.index:
        #         milp += (
        #             AnzRU[(1, factory, ressource)]
        #             == AnzRUAdpPlus[(1, factory, ressource)]
        #             - AnzRUAdpMinus[(1, factory, ressource)]
        #             )

        #         # Beschränkung für Einstellungen/Entlassungen
        #         milp += (
        #             AnzRUAdpPlus[(1, factory, ressource)]
        #             <= ressources.loc[(factory, ressource), 'r_Einstellen']
        #             )
        #         milp += (
        #             AnzRUAdpMinus[(1, factory, ressource)]
        #             <= ressources.loc[(factory, ressource), 'r_Entlassen']
        #             )


        # Anpassung des Status von Werken
        # gilt für perioden >= 2
        if periods > 1:
            for period in range(2, periods + 1):
                for factory in factories.index:
                    milp += (
                        ActU[(period, factory)]
                        == ActU[((period - 1), factory)]
                        + UOpen[(period, factory)]
                        - UClose[(period, factory)]
                        )
        # Werke werden initialisiert,
        # daher ist Periode 1 explizit durch die Benutzereingabe gegeben
        # for factory in factories.index:
        #     milp += (
        #         ActU[(1, factory)]
        #         == UOpen[(1, factory)]
        #         - UClose[(1, factory)]
        #         )



        # Zu Beginn einer Periode kann ein Werk entweder geöffnet oder geschlossen
        # werden
        for period in range(2, periods + 1):
            for factory in factories.index:
                milp += (
                    UOpen[(period, factory)]
                    + UClose[(period, factory)]
                    <= 1
                    )

        # Anpassung des Status von Segmenten
        # Der Status von Werken für Periode 1 ist durch die Benuztereingabe explizit
        # gegeben und muss daher nicht nochmals restringiert werden
        # perioden >= 2
        if periods > 1:
            for period in range(2, periods + 1):
                for factory, segment in segments.index:
                    milp += (
                        ActS[(period, factory, segment)]
                        == ActS[((period - 1), factory, segment)]
                        + SOpen[(period, factory, segment)]
                        - SClose[(period, factory, segment)]
                        )

        # Zu Beginn einer Periode kann ein Segment entweder geöffnet oder geschlossen
        # werden
        for period in range(2, periods + 1):
            for factory, segment in segments.index:
                milp += (
                    SOpen[(period, factory, segment)]
                    + SClose[(period, factory, segment)]
                    <= 1
                    )

        # Anzahl Schichten je Segment sind begrenzt
        for period in range(1, periods + 1):
            for factory, segment in segments.index:
                milp += (
                    AnzShift[(period, factory, segment)]
                    <= segments.loc[(factory, segment), 'maxSchichten_Segment']
                    )

        #### Das Optimierungsproblem wird in eine .lp Datei geschrieben
        milp.writeLP("OptimierungsproblemProdAlloPlanNet.lp")

        return milp

    ### Funktion baut die Kostenzielfunktion auf
    # hier als Funktion, um Erweiterungen zu ermöglichen
    def buildObjectiveTotalCost():

        #### Zielfunktionen
        objectiveTotalCost = 0

        ## Rohmaterialkosten
        helpRohmaterialkosten = 0
        for period in range(1, periods + 1):
            for supplier, product in suppliers.index:
                for factory in factories.index:
                    helpRohmaterialkosten += (
                        suppliers.loc[(supplier, product), 'cost_Material']
                        * math.pow(suppliers.loc[(supplier, product), 'factor_cost'], period - 1)
                        * Delivery_L_U[(period, supplier, factory, product)]
                        )

        ## Transportkosten
        helpTransportkosten = 0

        #Lieferant an Werke
        for period in range(1, periods + 1):
            for supplier, product in suppliers.index:
                for factory in factories.index:
                    if (('general', supplier, factory) in transportCon.index) == True:
                        helpTransportkosten += (suppliers.loc[(supplier, product),'cost_Material']
                            * transportCon.loc[('general', supplier, factory),
                                              'cost_Transport']
                            * math.pow(transportCon.loc[('general' ,supplier, factory),
                                                        'factor_cost'], period - 1)
                            * Delivery_L_U[(period, supplier, factory, product)]
                            )
                    elif (('general', 'general', 'general') in transportCon.index) == True:
                        helpTransportkosten += (
                            transportCon.loc[('general', 'general', 'general'),
                                              'cost_Transport']
                            * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                        'factor_cost'], period - 1)
                            * Delivery_L_U[(period, supplier, factory, product)]
                            )

        # Werk an Werk
        for period in range(1, periods + 1):
            for sendFactory in factories.index:
                for recFactory in factories.index.drop(sendFactory):
                    for product in bom.xs('C', level = 1).index.get_level_values(0).unique():
                        if ((product, sendFactory, recFactory) in transportCon.index) == True:
                            helpTransportkosten += (
                                transportCon.loc[(product, sendFactory, recFactory),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[(product, sendFactory, recFactory),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_U[(period, sendFactory, recFactory, product)]
                                )
                        elif (('general', 'general', 'general') in transportCon.index) == True:
                            helpTransportkosten += (
                                transportCon.loc[('general', 'general', 'general'),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_U[(period, sendFactory, recFactory, product)]
                                )

        # Werk an Kundenregion
        for period, customer, product in demand.index:
            for  factory in factories.index:
                if ((product, factory, customer) in transportCon.index) == True:
                    helpTransportkosten += (
                        transportCon.loc[(product, factory, customer),
                                         'cost_Transport']
                        * math.pow(transportCon.loc[(product, factory, customer),
                                                    'factor_cost'], period - 1)
                        * Delivery_U_K[(period, factory, customer, product)]
                        )
                elif (('general', 'general', 'general') in transportCon.index) == True:
                    helpTransportkosten += (
                        transportCon.loc[('general', 'general', 'general'),
                                         'cost_Transport']
                        * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                    'factor_cost'], period - 1)
                        * Delivery_U_K[(period, factory, customer, product)]
                        )

        ## Erweiterung
        # Externe Kapazitäten
        for period in range(1, periods + 1):
            for  factory in factories.index:
                for extern, product in externs.index:
                    if ((product, extern, factory) in transportCon.index) == True:
                        helpTransportkosten += (
                            Delivery_E_U[(period, extern, factory, product)]
                            * transportCon.loc[(product, extern, factory),
                                               'cost_Transport']
                            * math.pow(transportCon.loc[(product, extern, factory),
                                                        'factor_cost'], period - 1)
                            )
                    elif (('general', 'general', 'general') in transportCon.index) == True:
                        helpTransportkosten += (
                            transportCon.loc[('general', 'general', 'general'),
                                             'cost_Transport']
                            * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                        'factor_cost'], period - 1)
                            * Delivery_E_U[(period, extern, factory, product)]
                            )


        # Fixkosten für Werke
        helpFixCWerk = 0
        for period in range(1, periods + 1):
            for factory in factories.index:
                helpFixCWerk += (
                    ActU[(period, factory)]
                    * factories.loc[factory, 'c_Werk_fix']
                    * math.pow(factories.loc[factory, 'factor_fix'], period - 1)
                    )

        # Kosten für das Öffnen und Schließen von Werken
        helpWerkOpenClose = 0
        for period in range(1, periods + 1):
            for factory in factories.index:
                helpWerkOpenClose += (
                    UOpen[(period, factory)]
                    * factories.loc[factory, 'c_Werk_open']
                    * math.pow(factories.loc[factory, 'factor_open'], period - 1)
                    )
                helpWerkOpenClose += (
                    UClose[(period, factory)]
                    * factories.loc[factory, 'c_Werk_close']
                    * math.pow(factories.loc[factory, 'factor_close'], period - 1)
                    )

        # Fixkosten für Segmente
        helpFixCSegment = 0
        for period in range(1, periods + 1):
            for factory, segment in segments.index:
                helpFixCSegment += (
                    ActS[(period, factory, segment)]
                    * segments.loc[(factory, segment), 'c_Segment_fix']
                    * math.pow(segments.loc[(factory, segment),
                                            'factor_fix'], period - 1)
                    )

        # Erweiterung der Fixkosten für Segmente um die entstehenden Kosten je Schicht
        for period in range(1, periods + 1):
            for factory, segment in segments.index:
                helpFixCSegment += (
                    AnzShift[(period, factory, segment)]
                    * segments.loc[(factory, segment), 'cost_per_Shift']
                    * math.pow(segments.loc[(factory, segment),
                                            'factor_Shift'], period - 1)
                    )

        # Kosten für das Öffnen und Schließen von Segmenten
        helpSegmentOpenClose = 0
        for period in range(1, periods + 1):
            for factory, segment in segments.index:
                helpSegmentOpenClose += (
                    SOpen[(period, factory, segment)]
                    * segments.loc[(factory, segment), 'c_Segment_open']
                    * math.pow(segments.loc[(factory, segment),
                                            'factor_open'], period - 1)
                    )
                helpSegmentOpenClose += (
                    SClose[(period, factory, segment)]
                    * segments.loc[(factory, segment), 'c_Segment_close']
                    * math.pow(segments.loc[(factory, segment),
                                            'factor_open'], period - 1)
                    )

        ## Kosten für Ressourcen
        helpRessourcenKosten = 0
        for period in range(1, periods + 1):
            for factory, ressource in ressources.index:
                helpRessourcenKosten += (
                    AnzRU[(period, factory, ressource)]
                    * ressources.loc[(factory, ressource), 'r_Stunden']
                    * ressources.loc[(factory, ressource), 'r_Stundensatz']
                    * math.pow(ressources.loc[(factory, ressource),
                                              'r_Stundensatz_factor'], period - 1)
                    )


        # Perioden in Gleitzeitzyklus
        periodsInCylce = int(ressources.loc[(factory, ressource),
                                                        'r_Perioden_in_Zyklus'])
        # maximale Zyklenanzahl
        maxCycle = math.floor(periods/periodsInCylce)
        # Zyklusliste erstellen
        listCycle = [periodsInCylce * i for i in range(0, maxCycle + 1)]

        # Kosten für Gleitzeit
        helpRessourcenKostenGZ = 0
        for factory, ressource in ressources.index:
            # alle Zyklen durchlaufen
            for cyc in listCycle:
                helpGZEndOfCycle = 0
                cycPer = 1

                for per in range(1, periodsInCylce + 1):
                    cycPer = (cyc + per)
                    helpGZEndOfCycle += (
                        GleitzeitR[(cycPer, factory, ressource)]
                        )

                    #Abbruchbedingung wenn maximale Periodenzahl des Modells erreicht wird
                    if (cycPer == periods):
                        break

                helpRessourcenKostenGZ += (
                    helpGZEndOfCycle
                    * ressources.loc[(factory, ressource), 'r_Stundensatz']
                    * math.pow(ressources.loc[(factory, ressource),
                                              'r_Stundensatz_factor'], cycPer - 1)
                    * ressources.loc[(factory, ressource), 'r_GZ_percent']
                    )


        #print(helpRessourcenKostenGZ)

        helpRessourcenKosten =  helpRessourcenKosten + helpRessourcenKostenGZ

        # Kosten für Ressourcen Adaption
        # gilt für perioden >= 2
        helpRessourcenKostenAdaption = 0

        for period in range(1, periods + 1):
            for factory, ressource in ressources.index:
                helpRessourcenKostenAdaption += (
                    AnzRUAdpPlus[(period, factory, ressource)]
                    * ressources.loc[(factory, ressource), 'r_cost_Einstellen']
                    * math.pow(ressources.loc[(factory, ressource),
                                          'r_costFactor_Adp'], period - 1)
                    )
                helpRessourcenKostenAdaption += (
                    AnzRUAdpMinus[(period, factory, ressource)]
                    * ressources.loc[(factory, ressource), 'r_cost_Entlassen']
                    * math.pow(ressources.loc[(factory, ressource),
                                          'r_costFactor_Adp'], period - 1)
                    )

        helpRessourcenKosten += helpRessourcenKostenAdaption

        # Kosten Produktion: Bearbeitungskosten
        helpProduktionskosten = 0
        for period in range(1, periods + 1):
            for factory, segment, ressource, product in resProd.index:
                helpProduktionskosten += (
                    AnzProdP[(period, factory, segment, ressource, product)]
                    * resProd.loc[(factory, segment, ressource, product),
                                  'var_ProdCost']
                    *  math.pow(resProd.loc[(factory, segment, ressource, product),
                                              'factor_cost'], period - 1)
                    )

        ### Erweiterungen
        # Externe Einheiten
        helpExternKosten = 0
        for period in range(1, periods + 1):
            for extern, product in externs.index:
                for factory in factories.index:
                    helpExternKosten += (
                        Delivery_E_U[(period, extern, factory, product)]
                        * externs.loc[(extern, product), 'cost_Material']
                        * math.pow(externs.loc[(extern, product), 'factor_cost'],
                                   period - 1)
                        )

        # Zielfunktion Gesamtkosten
        objectiveTotalCost = (
            helpRohmaterialkosten
            + helpTransportkosten
            + helpFixCWerk
            + helpWerkOpenClose
            + helpFixCSegment
            + helpSegmentOpenClose
            + helpRessourcenKosten
            + helpProduktionskosten
            + helpExternKosten
            )

        return objectiveTotalCost



    ### Funktion gibt die Zielfunktion für die Kundennähe zurück
    def buildObjectiveCustomerCloseness():
        objectiveCustomerCloseness = 0

        for period, customer, product in demand.index:
            for  factory in factories.index:
                helpParCustomer = 0
                if ((factory, customer) in customerCloseness.index):
                    helpParCustomer = customerCloseness.loc[(factory, customer),
                                                            'Kundennähe']
                else:
                    helpParCustomer = customerCloseness.loc[('general', 'general'),
                                                            'Kundennähe']

                objectiveCustomerCloseness += (
                    helpParCustomer
                    * Delivery_U_K[(period, factory, customer, product)]
                    )

        return objectiveCustomerCloseness

    ### Funktion gibt die Zielfunktion für die Minimierung der Nacharbeitungszeit zurück
    def buildObjectiveExtraWork():
        objectiveExtraWork = 0

        for period in range(1, periods + 1):
            for factory, segment, ressource, product in resProd.index:
                objectiveExtraWork += (
                    AnzProdP[(period, factory, segment, ressource, product)]
                    * resProd.loc[(factory, segment, ressource, product),
                                  'cap_Error']
                    )
        return objectiveExtraWork

    ### Funktion gibt eine Liste mit den zu verwendenden Zielfunktionen zurück
    def get_multi_objectives():
        list_objectives = []

        list_objectives.append([buildObjectiveTotalCost(),
                                pulp.LpMinimize, 'R', 0.10])

        #list_objectives.append([buildObjectiveCustomerCloseness(),
         #                     pulp.LpMaximize, 'R', 0.2])

        #list_objectives.append([buildObjectiveExtraWork(),
                #                pulp.LpMinimize, 'R', 0.2])

        return list_objectives

    ### Dataframe mit genutzen Produktionstupeln und der jeweiligen Stückzahl, etc.
    def get_Produktion_Output():
        # Produktion
        productionOutput = []
        for period in range(1, periods + 1):
            for factory, segment, ressource, product in resProd.index:
                dict_output = {
                            'Periode': period,
                            'Werk': factory,
                            'Segment': segment,
                            'Ressource': ressource,
                            'Produktzustand': product,

                            'Stückzahl': AnzProdP[(period, factory, segment,
                                                   ressource, product)].varValue,
                            'Einsatzzeit': OperatingTimeR[(period, factory, segment,
                                                           ressource, product)].varValue,
                            'Anzahl Schichten': AnzShift[(period, factory, segment)].varValue
                            }

                # nur wenn Stückzahl bzw. Einsatzzeit > 0
                if dict_output['Stückzahl'] > 0 or dict_output['Einsatzzeit'] > 0:
                    productionOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        productionOutput_df = (pd.DataFrame.from_records(productionOutput)
                              .sort_values(['Periode', 'Werk', 'Segment',
                                            'Ressource', 'Produktzustand']))

        # Index definieren
        productionOutput_df.set_index(['Periode', 'Werk', 'Segment',
                                       'Ressource', 'Produktzustand'], inplace=True)

        return productionOutput_df

    def get_Produktion_Output2():
        # Produktion
        productionOutput = []
        for period in range(1, periods + 1):
            for factory, segment, ressource, product in resProd.index:
                dict_output = {
                            'Periode': period,
                            'Werk': factory,
                            'Segment': segment,
                            'Ressource': ressource,
                            'Produktzustand': product,
                            'Stückzahl':AnzProdP[(period, factory, segment,
                                                   ressource, product)].varValue,
                            'Einsatzzeit': OperatingTimeR[(period, factory, segment,
                                                           ressource, product)].varValue,
                            'Anzahl Schichten': AnzShift[(period, factory, segment)].varValue
                            }

                # nur wenn Stückzahl bzw. Einsatzzeit > 0
                if dict_output['Stückzahl'] > 0 or dict_output['Einsatzzeit'] > 0:
                    productionOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        productionOutput_df = (pd.DataFrame.from_records(productionOutput)
                              .sort_values(['Periode', 'Werk', 'Segment',
                                            'Ressource', 'Produktzustand']))

        # Index definieren
        productionOutput_df.set_index(['Periode', 'Werk', 'Segment',
                                       'Ressource', 'Produktzustand'], inplace=True)

        return productionOutput_df


    #### Lösen des linearen Programms


    #####Ausgabe der Werte für die Entscheidungsvariablen

    ##### Folgende Funktionen geben jeweils einen pandas Dataframe zurück
    ### Dataframe mit den gentzten Transportverbindungen und Liefermengen
    def get_Liefermengen_Output():
        # Liefermengen auf den einzelnen Transportverbindungen
        # Definition von zwei Listen: 'transportA' und 'transportB'
        # das Tupel (transportA, transportB) stellt die jeweiligen
        # Transportverbindungen dar
        transportA = []
        transportB = []

        transportA = (
            suppliers.index.droplevel(1).unique().values.tolist()
            + externs.index.droplevel(1).unique().values.tolist()
            + factories.index.values.tolist()
            )

        transportB = (
            factories.index.values.tolist()
            + demand.index.get_level_values(1).unique().values.tolist()
            )

        # alle Produkte in der BoM
        allProductsBom = []
        allProductsBom = (
            bom.index.get_level_values(0).unique().values.tolist()
            + bom.index.get_level_values(2).unique().values.tolist()
            )

        # Duplikate entfernen
        allProductsBom = list(dict.fromkeys(allProductsBom))

        transportOutput = []

        for period in range(1, periods + 1):
            for i in transportA:
                newTransportB = transportB.copy()
                # Werke liefern nicht an sich selbst
                if i in transportB:
                    newTransportB.remove(i)
                for j in newTransportB:
                    for p in allProductsBom:
                        dict_output = {
                            'Periode': period,
                            'Von': i,
                            'Nach': j,
                            'Produktzustand': p,
                            'Liefermenge': 0
                            }
                        # check if index (period, i, j, p) existiert in den
                        #Entscheidungsvariablen
                        if (period, i, j, p) in Delivery_L_U:
                            dict_output['Liefermenge'] = (
                                Delivery_L_U[(period, i, j, p)].varValue
                                )

                        elif (period, i, j, p) in Delivery_U_U:
                            dict_output['Liefermenge'] = (
                                Delivery_U_U[(period, i, j, p)].varValue
                                )

                        elif (period, i, j, p) in Delivery_U_K:
                            dict_output['Liefermenge'] = (
                                Delivery_U_K[(period, i, j, p)].varValue
                                )

                        elif (period, i, j, p) in Delivery_E_U:
                            dict_output['Liefermenge'] = (
                                Delivery_E_U[(period, i, j, p)].varValue
                                )

                        # nur wenn die Liefermnege > 0, wird dict zur transportOutput
                        # Liste hinzugefügt
                        if dict_output['Liefermenge'] > 0:
                            transportOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        transportOutput_df = (pd.DataFrame.from_records(transportOutput)
                              .sort_values(['Periode', 'Von', 'Nach', 'Produktzustand']))

        # Index definieren
        transportOutput_df.set_index(['Periode', 'Von', 'Nach', 'Produktzustand'], inplace=True)

        return transportOutput_df



    ### Dataframe für die Ressourcen
    def get_Ressourcen_Output():
        # Ressourcen
        ressourceOutput = []
        for period in range(1, periods + 1):
            for factory, ressource in ressources.index:
                dict_output = {
                    'Periode': period,
                    'Werk': factory,
                    'Ressource': ressource,
                    'Anzahl': AnzRU[(period, factory, ressource)].varValue,
                    'Gleitzeit [h]': GleitzeitR[(period, factory, ressource)].varValue,
                    'Einstellen': AnzRUAdpPlus[(period, factory, ressource)].varValue,
                    'Entlassen': AnzRUAdpMinus[(period, factory, ressource)].varValue
                    }


                # nur wenn Anzahl > 0
                if dict_output['Anzahl'] is None:
                    dict_output['Anzahl'] =0
                if dict_output['Einstellen'] is None:
                    dict_output['Einstellen'] = 0
                if dict_output['Entlassen'] is None:
                    dict_output['Entlassen'] = 0

                if (dict_output['Anzahl'] > 0 or dict_output['Einstellen'] > 0
                    or dict_output['Entlassen'] > 0):
                    ressourceOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        ressourceOutput_df = (pd.DataFrame.from_records(ressourceOutput)
                              .sort_values(['Periode', 'Werk', 'Ressource']))

        # Index definieren
        ressourceOutput_df.set_index(['Periode', 'Werk', 'Ressource'], inplace=True)

        return ressourceOutput_df



    ### Dataframe für den Status von Werken und Segmenten
    def get_Status_Output():
        # Status Werke und Segmente
        statusOutput = []
        for period in range(1, periods + 1):
            for factory in factories.index:
                if (factory in segments.index.droplevel(1).unique()) == True:
                    for segment in segments.xs(factory, level = 0).index:
                        dict_output = {
                            'Periode': period,
                            'Werk': factory,
                            'Status Werk': ActU[(period, factory)].varValue,
                            'Segment': segment,
                            'Status Segment': ActS[(period, factory, segment)].varValue
                        }

                        statusOutput.append(dict_output)

                else:
                    dict_output = {
                        'Periode': period,
                        'Werk': factory,
                        'Status Werk': ActU[(period, factory)].varValue
                        }
                    statusOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        statusOutput_df = (pd.DataFrame.from_records(statusOutput)
                              .sort_values(['Periode', 'Werk']))

        # Index definieren
        statusOutput_df.set_index(['Periode', 'Werk','Status Werk','Segment'], inplace=True)

        return statusOutput_df

    ### Dataframe mit den verschiedenen Kostenkategorien
    def get_Kosten_Output():
        # Hilfsvariable die die Summe der genutzten Gleitzeit eines Zyklus speichert
        # helpCostFTCycleOutput = 0
        costOutput = []
        for period in range(1, periods + 1):
            for factory in factories.index:
                dict_output = {
                    'Periode': period,
                    'Werk': factory,
                    'Rohmaterialkosten': 0,
                    'Transportkosten': 0,
                    'Werk Fixkosten': 0,
                    'Kosten O/C Werk': 0,
                    'Segment Fixkosten': 0,
                    'Kosten O/C Segment': 0,
                    'Ressourcenkosten': 0,
                    'Bearbeitungskosten': 0,
                    'Kosten Externe Kapazitäten': 0
                    }

                # Rohmaterialkosten
                for supplier, product in suppliers.index:
                    dict_output['Rohmaterialkosten'] += (
                        suppliers.loc[(supplier, product), 'cost_Material']
                        * math.pow(suppliers.loc[(supplier, product), 'factor_cost'], period - 1)
                        * Delivery_L_U[(period, supplier, factory, product)].varValue
                        )

                # Transportkosten: Lieferant -> Werk
                for supplier, product in suppliers.index:
                    if (('general', supplier, factory) in transportCon.index) == True:
                        dict_output['Transportkosten'] += (suppliers.loc[(supplier, product),'cost_Material']
                            *transportCon.loc[('general', supplier, factory),
                                             'cost_Transport']
                            * math.pow(transportCon.loc[('general', supplier, factory),
                                                        'factor_cost'], period - 1)
                            * Delivery_L_U[(period, supplier, factory, product)].varValue
                            )
                    elif (('general', 'general', 'general') in transportCon.index) == True:
                        dict_output['Transportkosten'] += (
                            transportCon.loc[('general', 'general', 'general'),
                                             'cost_Transport']
                            * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                        'factor_cost'], period - 1)
                            * Delivery_L_U[(period, supplier, factory, product)].varValue
                            )

                # Transportkosten: Werk -> Werk
                for recFactory in factories.index.drop(factory):
                    for product in bom.xs('C', level = 1).index.get_level_values(0).unique():
                        if ((product, factory, recFactory) in transportCon.index) == True:
                            dict_output['Transportkosten'] += (
                                transportCon.loc[(product, factory, recFactory),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[(product, factory, recFactory),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_U[(period, factory, recFactory, product)].varValue
                                )
                        elif (('general', 'general', 'general') in transportCon.index) == True:
                            dict_output['Transportkosten'] += (
                                transportCon.loc[('general', 'general', 'general'),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_U[(period, factory, recFactory, product)].varValue
                                )

                # Transportkosten: Werk -> Kundenregion
                if period in demand.index.get_level_values(0).unique():
                    for customer, product in demand.xs(period, level = 0).index:
                        if ((product, factory, customer) in transportCon.index) == True:
                            dict_output['Transportkosten'] += (
                                transportCon.loc[(product, factory, customer),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[(product, factory, customer),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_K[(period, factory, customer, product)].varValue
                                )
                        elif (('general', 'general', 'general') in transportCon.index) == True:
                            dict_output['Transportkosten'] += (
                                transportCon.loc[('general', 'general', 'general'),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                            'factor_cost'], period - 1)
                                * Delivery_U_K[(period, factory, customer, product)].varValue
                                )


                ### Erweiterung externe Kapazitäten
                # Transportkosten Externe Einheit -> Werk:
                for extern, product in externs.index:
                    if ((product, extern, factory) in transportCon.index) == True:
                        dict_output['Transportkosten'] += (
                            Delivery_E_U[(period, extern, factory, product)].varValue
                            * transportCon.loc[(product, extern, factory),
                                               'cost_Transport']
                            * math.pow(transportCon.loc[(product, extern, factory),
                                                        'factor_cost'], period - 1)
                            )
                    elif (('general', 'general', 'general') in transportCon.index) == True:
                            dict_output['Transportkosten'] += (
                                transportCon.loc[('general', 'general', 'general'),
                                                 'cost_Transport']
                                * math.pow(transportCon.loc[('general', 'general', 'general'),
                                                            'factor_cost'], period - 1)
                                * Delivery_E_U[(period, extern, factory, product)].varValue
                                )

                # Fixkkosten für das Werk
                dict_output['Werk Fixkosten'] += (
                    ActU[(period, factory)].varValue
                    * factories.loc[factory, 'c_Werk_fix']
                    * math.pow(factories.loc[factory, 'factor_fix'], period - 1)
                    )

                # Kosten für das Öffnen und Schließen von Werken
                dict_output['Kosten O/C Werk'] += (
                    UOpen[(period, factory)].varValue
                    * factories.loc[factory, 'c_Werk_open']
                    * math.pow(factories.loc[factory, 'factor_open'], period - 1)
                    )
                dict_output['Kosten O/C Werk'] += (
                    UClose[(period, factory)].varValue
                    * factories.loc[factory, 'c_Werk_close']
                    * math.pow(factories.loc[factory, 'factor_close'], period - 1)
                    )

                # Fixkosten für Segmente des Werkes
                if factory in segments.index.get_level_values(0).unique():
                    for segment in segments.xs(factory, level = 0).index:
                        dict_output['Segment Fixkosten'] += (
                            ActS[(period, factory, segment)].varValue
                            * segments.loc[(factory, segment), 'c_Segment_fix']
                            * math.pow(segments.loc[(factory, segment),
                                                    'factor_fix'], period - 1)
                            )
                # Erweiterung der Fixkosten der Segmente um die entstehenden Kosten
                # je Schicht
                if factory in segments.index.get_level_values(0).unique():
                    for segment in segments.xs(factory, level = 0).index:
                        dict_output['Segment Fixkosten'] += (
                            AnzShift[(period, factory, segment)].varValue
                            * segments.loc[(factory, segment), 'cost_per_Shift']
                            * math.pow(segments.loc[(factory, segment),
                                            'factor_Shift'], period - 1)
                            )

                # Kosten für das Öffnen und Schließen von Segmenten
                if factory in segments.index.get_level_values(0).unique():
                    for segment in segments.xs(factory, level = 0).index:
                        dict_output['Kosten O/C Segment'] += (
                            SOpen[(period, factory, segment)].varValue
                            * segments.loc[(factory, segment), 'c_Segment_open']
                            * math.pow(segments.loc[(factory, segment),
                                                    'factor_open'], period - 1)
                            )
                        dict_output['Kosten O/C Segment'] += (
                            SClose[(period, factory, segment)].varValue
                            * segments.loc[(factory, segment), 'c_Segment_close']
                            * math.pow(segments.loc[(factory, segment),
                                                    'factor_open'], period - 1)
                            )

                # Ressourcenkosten
                if factory in ressources.index.get_level_values(0).unique():
                    for ressource in ressources.xs(factory, level = 0).index:
                        dict_output['Ressourcenkosten'] += (
                            AnzRU[(period, factory, ressource)].varValue
                            * ressources.loc[(factory, ressource), 'r_Stunden']
                            * ressources.loc[(factory, ressource), 'r_Stundensatz']
                            * math.pow(ressources.loc[(factory, ressource),
                                                      'r_Stundensatz_factor'], period - 1)
                            )


                # Kosten für Ressourcen Adaption
                if factory in ressources.index.get_level_values(0).unique(): # and period > 1:
                    for ressource in ressources.xs(factory, level = 0).index:
                        dict_output['Ressourcenkosten'] += (
                            AnzRUAdpPlus[(period, factory, ressource)].varValue
                            * ressources.loc[(factory, ressource), 'r_cost_Einstellen']
                            * math.pow(ressources.loc[(factory, ressource),
                                                      'r_costFactor_Adp'], period - 1)
                            )

                        dict_output['Ressourcenkosten'] += (
                            AnzRUAdpMinus[(period, factory, ressource)].varValue
                            * ressources.loc[(factory, ressource), 'r_cost_Entlassen']
                            * math.pow(ressources.loc[(factory, ressource),
                                                      'r_costFactor_Adp'], period - 1)
                            )


                # variable Produktionskosten
                if factory in resProd.index.get_level_values(0).unique():
                    for segment, ressource, product in resProd.xs(factory, level = 0).index:
                        dict_output['Bearbeitungskosten'] += (
                            AnzProdP[(period, factory, segment, ressource, product)].varValue
                            * resProd.loc[(factory, segment, ressource, product),
                                          'var_ProdCost']
                            *  math.pow(resProd.loc[(factory, segment, ressource, product),
                                                      'factor_cost'], period - 1)
                            )

                # Kosten durch die Nutzung externer Kapazitäten
                for extern, product in externs.index:
                    dict_output['Kosten Externe Kapazitäten'] += (
                        Delivery_E_U[(period, extern, factory, product)].varValue
                        * externs.loc[(extern, product), 'cost_Material']
                        * math.pow(externs.loc[(extern, product), 'factor_cost'],
                                   period - 1)
                        )

                costOutput.append(dict_output)

        # Werte einem Dataframe hinzufügen und sortieren
        costOutput_df = (pd.DataFrame.from_records(costOutput)
                              .sort_values(['Periode', 'Werk']))

        # Index definieren
        costOutput_df.set_index(['Periode', 'Werk'], inplace=True)

        # Kosten für Gleitzeit werden seperat ermittelt, da die Zyklen berücksichtig
        # werden müssen
        # die entstehenden Kosten für die Gleitzeit werden im Dataframe costOutput_df
        # den Ressourcenkosten hinzuaddiert
        # Perioden in Gleitzeitzyklus
        periodsInCylce = int(ressources.loc[(factory, ressource),
                                                        'r_Perioden_in_Zyklus'])
        # maximale Zyklenanzahl
        maxCycle = math.floor(periods/periodsInCylce)
        # Zyklusliste erstellen
        listCycle = [periodsInCylce * i for i in range(0, maxCycle + 1)]
        for factory, ressource in ressources.index:
            # alle Zyklen durchlaufen
            for cyc in listCycle:
                helpGZEndOfCycle = 0
                cycPer = 1

                for per in range(1, periodsInCylce + 1):
                    cycPer = (cyc + per)
                    helpGZEndOfCycle += (
                        GleitzeitR[(cycPer, factory, ressource)].varValue
                        )

                    #Abbruchbedingung wenn maximale Periodenzahl des Modells erreicht wird
                    if (cycPer == periods):
                        break

                costOutput_df.loc[((cyc + per), factory), 'Ressourcenkosten'] += (
                    helpGZEndOfCycle
                    * ressources.loc[(factory, ressource), 'r_Stundensatz']
                    * math.pow(ressources.loc[(factory, ressource),
                                              'r_Stundensatz_factor'], cycPer - 1)
                    * ressources.loc[(factory, ressource), 'r_GZ_percent']
                    )

        return costOutput_df






    def writeSolutionToExcel():

        path = 'Ergebnisse/'
        workbookName = (pd.Timestamp.now().strftime(format = "%Y-%m-%d_%H_%M_%S")
        + '_Ergebnis.xlsx'
        )
        ### write results to Excel
        # Name der Ausgabedatei mit Zeitstempel
        # Pfade auf dem Mac mit "/", auf Windows mit "\"
        excel_file_result = path + workbookName

        # Aufrufen des pandas ExcelWriter
        writer = pd.ExcelWriter(excel_file_result, engine='xlsxwriter')

        df1.to_excel(writer, index=True, sheet_name='Zusammenfassung')


        Restriktionen_df.to_excel(writer, index=True, sheet_name="Restriktionen")


        # für jede Outputkategorie wird ein eigenes Excelsheet erstellt
        # Excelsheet für die Transportmengen
        get_Liefermengen_Output().to_excel(writer, index=True, sheet_name='Transportmengen')

        # Excelsheet für die Produktionsmengen
        get_Produktion_Output().to_excel(writer, index=True, sheet_name='Produktionsmengen')

        # Excelsheet für die Ressourcen
        get_Ressourcen_Output().to_excel(writer, index=True, sheet_name='Ressourcen')

        # Excelsheet für den Status von Werken und Segmenten
        get_Status_Output().to_excel(writer, index=True, sheet_name='Status')

        # Excelsheet für die Kostenaufschlüsselung
        get_Kosten_Output().to_excel(writer, index=True, sheet_name='Kosten')



        # passt die Spaltenbreite im Reiter Kosten ans
        worksheet = writer.sheets['Kosten']
        my_format = writer.book.add_format({'num_format': '#,###\€'} )
        worksheet.set_column('C:K', 18, my_format)
        # speichern der Excel-Datei (Pfad und Name siehe oben)
        writer.save()

        # workbook = xlsxwriter.Workbook('output.xlsx')

        # worksheet = writer.bookworksheet = writer.sheets['report']

        # header_fmt = workbook.add_format({'bold': True})
        # worksheet.set_row(0, None, header_fmt)

        # writer.save()


    d = {}
    d2 = {}
    Prod = {}
    productionOutput = []






    def suche(Nachfrage2):

        Nachfrage = Nachfrage2

        Steigerung = 30

        demand.loc[(1,'USA','K6067000'),'Nachfrage'] = 30
        demand.loc[(2,'USA','K6067000'),'Nachfrage'] = round(Nachfrage*1.06)
        demand.loc[(3,'USA','K6067000'),'Nachfrage'] = round(Nachfrage*1.06*1.06)
        demand.loc[(4,'USA','K6067000'),'Nachfrage'] = round(Nachfrage*1.06*1.06*1.06)




        targets = get_multi_objectives()
        restriktionAdd = []
        for target in targets:
                prblToGo = buildModel(target[0], restriktionAdd)
                prblToGo.sense = target[1]

                # Eingabe der Zielgenauigkeit und der maximalen Optimierungszeit
                #prblToGo.solve(pulp.apis.PULP_CBC_CMD( timeLimit = 60))
                prblToGo.solve()
                #prblToGo.solve(pulp.apis.GLPK_CMD(keepFiles=True, timeLimit=300))
                #prblToGo.solve(pulp.apis.GLPK_CMD(options=['--ranges test.sen']))

                #schreibe Status

                print(pulp.LpStatus[prblToGo.status])

                # schreibe den Zielfunktionswert
                print(pulp.value(prblToGo.objective))
                df1 = pd.Series([pulp.LpStatus[prblToGo.status],pulp.value(prblToGo.objective)])

                helpDeviation = 0
                # Relative Abweichung
                if target[2] == 'R':
                    helpDeviation = pulp.value(prblToGo.objective)
                elif target[2] == 'A':
                    helpDeviation = target[3]
                else:
                    print('Eingabe der Abweichung fehlerhaft!')
                    helpDeviation = 0







                Restriktionen= []

                for v in prblToGo.variables():
                    dict_output = {
                        'Restriktion':v.name,
                        'Wert': v.varValue
                        }
                    Restriktionen.append(dict_output)

                # Werte einem Dataframe hinzufügen und sortieren
                Restriktionen_df = (pd.DataFrame.from_records(Restriktionen).sort_values(['Restriktion', 'Wert']))



                # LpMaximize = -1, lpMinimize = 1
                if target[1] == pulp.LpMaximize:
                    restriktionAdd.append(pulp.LpConstraint(
                        e = prblToGo.objective,
                        sense = pulp.LpConstraintGE,
                        rhs = pulp.value(prblToGo.objective) - helpDeviation
                        ))
                else:
                    restriktionAdd.append(pulp.LpConstraint(
                        e = prblToGo.objective,
                        sense = pulp.LpConstraintLE,
                        rhs = pulp.value(prblToGo.objective) + helpDeviation
                        ))

                del prblToGo


        Status = get_Status_Output()
        Test =  Status.loc[(2,'IMUS',1,'finalAsWave'),'Status Segment']
        print(Status)
        print(Test)
        return(Test)






    links = {}
    links[0] = 40
    rechts = {}
    rechts[0] = 200
    t = 0
    maximum = 5
    genauigkeit = 10
    ergebnis1 =[]
    ergebnis0 =[]



    while (abs(rechts[t]-links[t]) > genauigkeit):
        print(links[t],rechts[t])

        if links[t] in ergebnis1:
            ergebnislinks = 1
        elif links[t] in ergebnis0:
            ergebnislinks = 0
        else:
            ergebnislinks = suche(links[t])


        if rechts[t] in ergebnis1:
            ergebnisrechts = 1
        elif links[t] in ergebnis0:
            ergebnisrechts = 0
        else:
            ergebnisrechts = suche(rechts[t])




        if ergebnislinks == 1:
            ergebnis1.append(links[t])
        elif ergebnislinks == 0:
            ergebnis0.append(links[t])

        if ergebnisrechts == 1:
            ergebnis1.append(rechts[t])
        elif ergebnisrechts == 0:
            ergebnis0.append(rechts[t])



        if ergebnislinks == 0 and ergebnisrechts ==1:
            links[t+1] = round((links[t]+rechts[t])/2)
            rechts[t+1] = rechts[t]
        elif ergebnislinks ==1 and ergebnisrechts ==0:
            rechts[t+1] = round((links[t]+rechts[t])/2)
            links[t+1] = links[t]
        elif ergebnislinks ==1 and ergebnisrechts ==1:
            rechts[t+1] = links[t]
            links[t+1] = links[t-1]
        else: print("fehler")
        print(links,rechts)
        t+=1
        print(t)



    productionOutput_df = (pd.DataFrame.from_records(productionOutput)).drop_duplicates().sort_values(['Durchgang', 'Periode', 'Werk'])

    path = 'Ergebnisse/'
    workbookName2 = (pd.Timestamp.now().strftime(format = "%Y-%m-%d_%H_%M_%S")
    + '_Zusammenfassung.xlsx'
    )
    ### write results to Excel
    # Name der Ausgabedatei mit Zeitstempel
    # Pfade auf dem Mac mit "/", auf Windows mit "\"
    excel_file_result = path + workbookName2

    # Aufrufen des pandas ExcelWriter
    writer = pd.ExcelWriter(excel_file_result, engine='xlsxwriter')

    # Excelsheet für die Produktionsmengen
    productionOutput_df.to_excel(writer, index=True, sheet_name='Produktionsmengen')

    table = pd.pivot_table(productionOutput_df,index=['Durchgang','Periode'], columns=('Werk') , values=['Summe'],aggfunc=np.sum)

    final_Table = table.droplevel(0,axis=1).reset_index().fillna(0).set_index('Durchgang')

    max_Columns= final_Table.shape[0]

    final_Table.to_excel(writer, index=True, sheet_name='Zusammenfassung')
    #table.plot.bar(stacked=True)

    workbook = writer.book
    worksheet = writer.sheets['Zusammenfassung']

    chart2 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    print(len(final_Table.columns))
    # Configure the first series.
    for col_num in range(2, len(final_Table.columns)+1 ):
        chart2.add_series({
            'name':       ['Zusammenfassung', 0, col_num],
            'categories': ['Zusammenfassung', 1, 1, max_Columns,0],
            'values':     ['Zusammenfassung', 1, col_num, max_Columns, col_num],
            'gap':        40,
        })

    worksheet.insert_chart('I2', chart2, {'x_offset': 25, 'y_offset': 10})

    workbook.close()
    writer.save()

    print(d[1])
