import xlwings as xlw
import pandas as pd
import datetime as dt
import tkinter as tk
import re


EXCEL = r"C:\Users\PC\Proy-Codigo\filtrar\app_filtrar\Base Suministros SIDAE FEB desktop.xlsx"
PESTAÑA = "M-GI-06"

COLEGIO_COLS = ['Id Sitio Entrega', 'Nombre Institución Educativa', 'Sitio de Entrega', 'Codigo_Dane_Sede']
FILTRO_COLS = ['Id Sitio Entrega', 'FECHA']


class ConexionExcel():
    def __init__(self, excel, pestaña):
        self.ws = xlw.Book(excel).sheets[pestaña]
        try:
            self.ws.api.ShowAllData()
            print("Rango desfiltrado.")
        except:
            print("Rango previamente desfiltrado.")
        self.df = self.ws.range("A1").expand().options(pd.DataFrame, header=1, index=False, numbers=int).value
        self.df = self.df[COLEGIO_COLS]
        self.df = self.df.drop_duplicates()
        self.n_filas = self.ws.range("D1").end("down").row
        print("número de filas", self.n_filas) 
        # print(self.df)

    def filtrar_colegio_e_ids(self, colegio, ids):
        self.ws.range("A1").api.AutoFilter(Field=18, Criteria1=colegio, Operator=7)     # filtra el colegio
        self.ws.range("A1").api.AutoFilter(Field=4, Criteria1=ids, Operator=7)          # filtra los ids
        
    def filtrar_fechas(self, fechas):
        self.ws.range("A1").api.AutoFilter(Field=47, Criteria2=fechas, Operator=7)      # filtra las fechas

        	
    
    def filtrar_avanzado(self, filas, var_dias):
        rows = len(filas)
        # print(self.n_filas+3)
        # print(self.ws.range(f"A{self.n_filas+3}").value )               
        if self.ws.range(f"A{self.n_filas+3}").value == None:           #para capturar si la celdas despues de la base se han ocupado, toca reiniciar el programa, puede ocurrir cuando se crean nuevas planillas
            self.ws.range(f"A{self.n_filas+3}").value = FILTRO_COLS     #encabezados tabla de referencia
            self.ws.range(f"B{self.n_filas+4}:B{self.n_filas+4+rows-1}").api.NumberFormat ="dd/mm/yyyy"      #formato para las fechas
            self.ws.range(f"A{self.n_filas+4}").options(date=dt.datetime).value = filas           #valores para filtrar con tabla de referencia
            if var_dias == "apartir":
                self.ws.range(f"B{self.n_filas+4}:B{self.n_filas+4+rows}").value = ""
            rango_ref = self.ws.api.Range(f"A{self.n_filas+3}:B{self.n_filas+3+rows}")
            self.ws.api.Range(f"A1:BC{self.n_filas}").AdvancedFilter(Action=1, CriteriaRange=rango_ref)
            self.ws.range(f"A{self.n_filas+3}:B{self.n_filas+3+rows}").value = ""       #borra el rango de referencia
        else:
            print("\nEspacio de referencia ocupado, revisar.")
            self.n_filas = self.ws.range("D1").end("down").row

    def desfiltrar(self):
        self.ws.api.ShowAllData()



class App():
    def __init__(self, excel):

        self.mi_excel = excel
        self.app = tk.Tk()
        self.app.title("( :")

        self.var_filtrar = tk.IntVar(self.app)
        self.var_dias = tk.StringVar(self.app, "deldia")
        self.var_exacto = tk.IntVar(self.app)


        self.textEntry = tk.Text(self.app, height=6, width =46)

        self.filterButton = tk.Checkbutton(self.app,
                                    text="filtrar",
                                    variable= self.var_filtrar,
                                    onvalue=  self.var_filtrar.set(value=1),
                                    offvalue= self.var_filtrar.set(value=0),
                                    indicator=0,
                                    width= 7,
                                    command=self.check_filtrar)
        
        self.favanzadoButton = tk.Checkbutton(self.app,
                                           text="avanzado",
                                           variable=self.var_exacto,
                                           onvalue=self.var_exacto.set(value=1),
                                           offvalue=self.var_exacto.set(value=0),
                                           font=("Arial", 6),
                                           bd=2)


        self.radioButton_deldia = tk.Radiobutton(self.app,
                                                text="del día",
                                                variable= self.var_dias,
                                                value=  "deldia",
                                                font=("Arial", 8))
        self.radioButton_apartir =tk.Radiobutton(self.app,
                                                text="a partir de",
                                                variable= self.var_dias,
                                                value= "apartir",                             
                                                font=("Arial", 8))


        self.textEntry.pack()
        self.filterButton.pack(side = "left", padx = 30, pady = 1, fill="x")
        self.favanzadoButton.pack(side = "left", padx = 0, pady = 1)
        self.radioButton_apartir.pack(side = "right", padx = 5, pady = 1)
        self.radioButton_deldia.pack(side = "right", padx = 5, pady = 1)



    
    def check_filtrar(self):

        texto = self.textEntry.get("1.0",'end-1c')
        ids = self.extraer_ids(texto)
        colegio = self.extraer_colegio(texto, ids)
        fechas = self.extraer_fechas(texto)
        filas = self.extraer_filas(texto)
        
        if self.var_filtrar.get() == 1:
            
            print("")
            print("             filt  exc     dias")
            print("comandos   | ",self.var_filtrar.get()," | ",f"{self.var_exacto.get()}"," | ", self.var_dias.get(), " |")
            print(colegio)
            print("ids:     ", ids)
            print("fechas:  ", fechas)
            print("filas:    ")
            for fila in filas:
                print("         ",fila)  

            if self.var_exacto.get() == 1:              #filtro nuevo, "exacto"
                if self.var_dias.get() == "deldia":
                    self.mi_excel.filtrar_avanzado(filas, self.var_dias.get())
                else:
                    self.mi_excel.filtrar_avanzado(filas, self.var_dias.get())

            elif self.var_exacto.get() == 0:            #filtro antiguo
                if self.var_dias.get() == "deldia":
                    self.mi_excel.filtrar_colegio_e_ids(colegio, ids)
                    self.mi_excel.filtrar_fechas(fechas)
                    
                else:
                    self.mi_excel.filtrar_colegio_e_ids(colegio, ids)
        
        elif self.var_filtrar.get() == 0:
            self.mi_excel.desfiltrar()



    def extraer_ids(self, texto):
        patron_ids = r"\d{5}"
        ids = []
        ids = re.findall(patron_ids, texto)
        ids = list(set(ids)) #quita duplicados de la lista
        # print(ids)
        return ids

    def extraer_colegio(self, texto, ids):
        ids2 = [int(n) for n in ids]
        # print(ids2)
        colegio = self.mi_excel.df[self.mi_excel.df["Id Sitio Entrega"].isin(ids2)]["Nombre Institución Educativa"].iloc[0]
        return colegio

    def extraer_fechas(self, texto):
        patron_fechas = r"\d{1,2}/\d{2}/\d{4}"
        fechas = []
        #para el "metodo normal" que filtra por todos los ids y fechas
        fechas = re.findall(patron_fechas, texto)
        fechas = list(set(fechas))
        fechas = self.ajustar_fechas(fechas) #ajusta las fechas a una lista de tuples (2, fecha) para excel
        # print(fechas)
        return fechas

    def ajustar_fechas(self, fechas):
        """Convierte las fechas a tuplas de fechas para excel (2, fecha1, 2, fecha2, 2, fecha3, ...)"""
        fechas2 = []
        for fecha in fechas:
            fechas2.append(2)
            fechas2.append(fecha)
        return fechas2

    def extraer_filas(self, texto):
        patron_ids = r"\d{5}"
        patron_fechas = r"\d{1,2}/\d{2}/\d{4}"
        filas = []
        for fila in texto.split("\n"):
            # print(fila)
            id = re.search(patron_ids, fila)
            id = id.group()
            fecha = re.search(patron_fechas, fila)
            fecha = "'"+fecha.group()               #Se les pone una comilla simple para que al escribirlas en excel queden como texto
            # print(id, fecha)
            filas.append([id, fecha])
        # print(filas)

        return filas





   



mi_excel = ConexionExcel(EXCEL, PESTAÑA)

app = App(mi_excel)


tk.mainloop()