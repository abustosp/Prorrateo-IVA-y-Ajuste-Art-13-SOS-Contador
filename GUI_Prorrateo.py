#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
import LIB.Ajuste_Art_13_xls as Ajuste_Art_13_xls , LIB.Ajuste_Art_13_csv as Ajuste_Art_13_csv , LIB.Prorrateo_mensual_xls as Prorrateo_mensual_xls , LIB.Prorrateo_mensual_csv as Prorrateo_mensual_csv

class App_Prorrateo_Art13:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=250,
            width=325)
        Toplevel_1.iconbitmap("LIB/BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.minsize(325, 250)
        Toplevel_1.overrideredirect("False")
        Toplevel_1.title("Prorrateo y Ajuste de Art 13 ")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="LIB/BIN/ABP blanco en fondo negro111.png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            takefocus=False,
            text='Calcular Prorrateos de Crédito Fiscal de IVA en base a los Archivos de BackUp de SOS-Contador.\n',
            wraplength=325)
        Label_1.pack(expand="true", side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand="true", side="top")
        self.Mensual_XLS = ttk.Button(Toplevel_1)
        self.Mensual_XLS.configure(text='Prorrateo Mensual XLS' , command=Prorrateo_mensual_xls.Prorrateo_Mensual_XLS)
        self.Mensual_XLS.pack(expand="true", pady=4, side="top")
        self.Mensual_CSV = ttk.Button(Toplevel_1)
        self.Mensual_CSV.configure(text='Prorrateo Mensual CSV' , command=Prorrateo_mensual_csv.Prorrateo_Mensual_CSV)
        self.Mensual_CSV.pack(expand="true", padx=0, pady=4, side="top")
        self.Anual_XLS = ttk.Button(Toplevel_1)
        self.Anual_XLS.configure(text='Prorrateo Anual + Ajuste Art 13 XLS' , command=Ajuste_Art_13_xls.Ajuste_Art13_XLS)
        self.Anual_XLS.pack(expand="true", pady=4, side="top")
        self.Anual_CSV = ttk.Button(Toplevel_1)
        self.Anual_CSV.configure(text='Prorrateo Anual + Ajuste Art 13 CSV' , command=Ajuste_Art_13_csv.Ajuste_Art13_CSV)
        self.Anual_CSV.pack(expand="true", pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = App_Prorrateo_Art13()
    app.run()
