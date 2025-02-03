import tkinter as tk
from tkinter import messagebox
from habeasData import habeasData  # Asegúrate de que el nombre del archivo y la clase coincidan

class classApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Aplicación de Cobranzas")
        self.geometry("500x400")
        self.config(bg="#2C3E50")

        label = tk.Label(self, text="Seleccione una opción", font=("Helvetica", 18), fg="white", bg="#2C3E50")
        label.pack(pady=30)

        self.boton1 = self.create_button("Habeas Data", self.abrir_habeas_data)
        self.boton2 = self.create_button("Prejuridica", self.opcion_2)
        self.boton3 = self.create_button("Ultimo Aviso", self.opcion_3)

        self.boton1.pack(pady=10)
        self.boton2.pack(pady=10)
        self.boton3.pack(pady=10)

    def create_button(self, text, command):
        button = tk.Button(self, text=text, width=30, height=2, font=("Helvetica", 14),
                        bg="#16A085", fg="white", relief="solid", bd=2, command=command)
        button.config(activebackground="#1ABC9C", activeforeground="white")
        return button

    def opcion_1(self):
        messagebox.showinfo("Opción 1", "Has seleccionado la opción de Ingresos")

    def opcion_2(self):
        messagebox.showinfo("Opción 2", "Has seleccionado la opción de Gastos")

    def opcion_3(self):
        messagebox.showinfo("Opción 3", "Has seleccionado la opción de Reportes")

    def abrir_habeas_data(self):
        ventana_habeas = habeasData(self)
        ventana_habeas.grab_set()
        ventana_habeas.mainloop()
