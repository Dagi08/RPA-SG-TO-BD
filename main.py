import tkinter
import customtkinter
import json
import os
import sys
import pygetwindow as gw
from PIL import Image
from CTkMessagebox import CTkMessagebox # type: ignore
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD

from script.account_details import extra



class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        self.iconbitmap("iconobn.ico")

        self.title("SEGURO DESGRAVAMEN (SSSF)")
        self.geometry("990x615")
        
        #self.geometry("1450x950")
        
        self.resizable(False, False)

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images
        self.load_images()

        # Establecer el modo claro como predeterminado
        customtkinter.set_appearance_mode("light")

        #self.directorio_in = ""

        # Variable para almacenar la ventana seleccionada
        self.ventana_seleccionada = None

        # Variable de instancia compartida
        self.file_entry_down = None  

        # Crear el marco de navegación
        self.navigation_frame = NavigationFrame(self)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")

        # Crear los marcos de contenido
        self.script_eicd=ScriptEICD(self)
        self.home_frame = HomeFrame(self)
        self.second_frame = SecondFrame(self)

        # Seleccionar el marco predeterminado
        self.select_frame_by_name("eicd")


        #Guarda las rutas de en un json
        if os.path.exists("paths.json"):
            with open("paths.json", "r") as f:
                data = json.load(f)
            self.directorio_in = data.get("input_path")
    

        # Establecer las rutas guardadas como valores predeterminados en los campos de entrada
        if self.directorio_in:
            self.script_eicd.File_Entry_down.configure(state='normal')
            self.script_eicd.File_Entry_down.insert(0, self.directorio_in)
            self.script_eicd.File_Entry_down.configure(state='readonly')

    
    #Manejo de la ubicación de los archivos cuando el script está empaquetado (capeta img)
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def load_images(self):
        image_path = self.resource_path("img")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "logo.png")), size=(150, 50))
        self.files = customtkinter.CTkImage(Image.open(os.path.join(image_path, "calendar-on.png")), size=(40, 40))
        self.cargar = customtkinter.CTkImage(Image.open(os.path.join(image_path, "detail-on.png")), size=(40, 40))
        self.seleccionar = customtkinter.CTkImage(Image.open(os.path.join(image_path, "logo.png")), size=(40, 40))

    def select_frame_by_name(self, name):
        # Actualizar el color del botón seleccionado
        self.navigation_frame.update_buttons(name)

        # Mostrar el marco seleccionado
        if name == "eicd":
            self.script_eicd.grid(row=0, column=1, sticky="nsew")
        else:
            self.script_eicd.grid_forget()
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()
          

class NavigationFrame(customtkinter.CTkFrame):
    def __init__(self, master):
        super().__init__(master, corner_radius=0)
        self.master = master

        self.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self, text="", image=self.master.logo_image,
                                                             compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20, sticky="s")

        self.file_button = customtkinter.CTkButton(self, corner_radius=0, height=40, border_spacing=10, text="CALENDARIO - MOV.",font=customtkinter.CTkFont(size=12, weight="bold"),
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.master.files, anchor="w", command=lambda: self.master.select_frame_by_name("eicd"))
        self.file_button.grid(row=1, column=0, sticky="ew")

        self.home_button = customtkinter.CTkButton(self, corner_radius=0, height=40, border_spacing=10, text="DETALLES CUENT.",font=customtkinter.CTkFont(size=12, weight="bold"),
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.master.cargar, anchor="w", command=lambda: self.master.select_frame_by_name("home"))
        self.home_button.grid(row=2, column=0, sticky="ew")

        self.frame_2_button = customtkinter.CTkButton(self, corner_radius=0, height=40, border_spacing=10, text="PASO 03",font=customtkinter.CTkFont(size=12, weight="bold"),
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.master.seleccionar, anchor="w", command=lambda: self.master.select_frame_by_name("frame_2"))
        self.frame_2_button.grid(row=3, column=0, sticky="ew")

    def update_buttons(self, selected_frame):
        self.file_button.configure(fg_color="#FFFFFF" if selected_frame == "eicd" else "transparent")
        self.home_button.configure(fg_color="#FFFFFF" if selected_frame == "home" else "transparent")
        self.frame_2_button.configure(fg_color="#FFFFFF" if selected_frame == "frame_2" else "transparent")

class BaseFrame(customtkinter.CTkFrame):
    def __init__(self, master):
        super().__init__(master, corner_radius=0, fg_color="#FFFFFF")
        self.grid_columnconfigure(0, weight=1)

class ScriptEICD(BaseFrame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        self.file_title_label= customtkinter.CTkLabel(self, text="Extraer Información Cuotas",font=customtkinter.CTkFont(size=20, weight="bold"), text_color="#BF0615")
        self.file_title_label.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")

        #Frame de seleccion de name ventan emulacion

        self.name_ventana_frame = customtkinter.CTkFrame(self, border_color="#B8B2B2", border_width=1, fg_color="transparent")
        self.name_ventana_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.Title_Ventana_ComboBox = customtkinter.CTkLabel(master=self.name_ventana_frame, text="Seleccionar ventana de produccion: ", font=customtkinter.CTkFont(size=15,weight="bold"))
        self.Title_Ventana_ComboBox.grid(row=0, column=0, padx=10, pady=10, sticky="")

        self.Ventana_ComboBox = customtkinter.CTkComboBox(master=self.name_ventana_frame, values=[""], width=200,
                                                              fg_color="#FFFFFF", button_color="#B8B2B2", border_color="#BF0615")
        self.Ventana_ComboBox.configure(state='readonly')
        self.Ventana_ComboBox.grid(row=1, column=0, columnspan=4, padx=40, pady=20, sticky="n")

        self.Ventana_Button = customtkinter.CTkButton(master= self.name_ventana_frame, text="Seleccionar Ventana", font=customtkinter.CTkFont(size=15,weight="bold"), 
                                                  border_width=1, border_color="#BF0615", fg_color="#FFFFFF", text_color="#BF0615", hover_color="#F2CDD0", command=self.seleccionar_ventana)
        self.Ventana_Button.grid(row=1, column=5, padx=270, pady=20, sticky="n")


        #Frame de Ruta de resultados de script (.txt)
        self.file_download_frame=customtkinter.CTkFrame(self, border_color="#B8B2B2", border_width=1, fg_color="transparent")
        self.file_download_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")

        self.Title_Entry_down= customtkinter.CTkLabel(master=self.file_download_frame, text="Seleccionar carpeta para RESULTADO archivos: ", font=customtkinter.CTkFont(size=15,weight="bold"))
        self.Title_Entry_down.grid(row=0, column=0, padx=10, pady=10, sticky="")

        self.File_Entry_down = customtkinter.CTkEntry(master= self.file_download_frame, placeholder_text="Definir la ruta para el resultado de los archivos (.txt)", width=450)
        self.File_Entry_down.configure(state='readonly')
        self.File_Entry_down.grid(row=1, column=0, columnspan=6, padx=40, pady=20, sticky="n")

        # Asignar el valor a la variable de instancia compartida
        self.master.file_entry_down = self.File_Entry_down

        self.File_Button_down= customtkinter.CTkButton(master=self.file_download_frame, text="Seleccionar Carpeta", font=customtkinter.CTkFont(size=15,weight="bold"), 
                                                  border_width=1, border_color="#BF0615", fg_color="#FFFFFF", text_color="#BF0615", hover_color="#F2CDD0", command=self.select_folder)
        self.File_Button_down.grid(row=1, column=8, padx=37, pady=20, sticky="n")

        #Frame drag and drop resltado de macros (.xlsx)
        self.file_upload_frame=customtkinter.CTkFrame(self, border_color="#B8B2B2", border_width=1, fg_color="transparent")
        self.file_upload_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")

        self.Title_Entry_up=customtkinter.CTkLabel(master=self.file_upload_frame, text="Cargar resultado de la Macros: ", font=customtkinter.CTkFont(size=15,weight="bold"))
        self.Title_Entry_up.grid(row=0, column=0, padx=10, pady=10, sticky="")

        self.drag_and_drop_entry = customtkinter.CTkEntry(master= self.file_upload_frame, placeholder_text="Arrastrar y soltar archivo para leer (.xlsx)", width=680, height=150)
        self.drag_and_drop_entry.configure(state='readonly', justify="center")
        self.drag_and_drop_entry.grid(row=1, column=0, columnspan=5, padx=40, pady=10, sticky="n")
        
        # Agregar soporte para arrastrar y soltar
        self.drag_and_drop_entry.drop_target_register(DND_FILES)
        self.drag_and_drop_entry.dnd_bind('<<Drop>>', self.on_drop)

        self.EjecutarScript1_button= customtkinter.CTkButton(master=self.file_upload_frame, text="Ejecutar", font=customtkinter.CTkFont(size=15,weight="bold"), 
                                                  border_width=1, border_color="#BF0615", fg_color="#FFFFFF", text_color="#BF0615", hover_color="#F2CDD0", command=self.ejecutar_script)
        self.EjecutarScript1_button.grid(row=2, column=4, padx=10, pady=10, sticky="n")

        # Obtener las ventanas y actualizar el ComboBox
        self.update_combo_box()

        # Variable para almacenar la ventana seleccionada
        #self.ventana_seleccionada = None

    def update_combo_box(self):
        ventanas = [window.title for window in gw.getWindowsWithTitle('desa') + gw.getWindowsWithTitle('prod')]
        self.Ventana_ComboBox.configure(values=ventanas)

    def seleccionar_ventana(self):
        self.master.ventana_seleccionada = self.Ventana_ComboBox.get()
        if self.master.ventana_seleccionada:
            CTkMessagebox(
                        title ="Nombre Ventana", message= f"Ventana seleccionada: {self.master.ventana_seleccionada}", icon="check", option_1="Ok",
                        bg_color="#FFFFFF", fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                        )
        else:
            CTkMessagebox(
                            title="Advertencia", message="No se ha seleccionado ninguna ventana.", icon="cancel", option_1="Ok",
                            bg_color="#FFFFFF", fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615",button_text_color="#FFFFFF"
                        )
    
    def select_folder(self):
        file_path = filedialog.askdirectory()
        if file_path:
            self.File_Entry_down.configure(state='normal')
            self.File_Entry_down.delete(0, tkinter.END)
            self.File_Entry_down.insert(0, file_path)
            self.master.directorio_in = file_path  # Save the selected path
            self.File_Entry_down.configure(state='readonly')
            self.save_folder_path("input_path", file_path)

    def save_folder_path(self,key, path):
        data = {}
        if os.path.exists("paths.json"):
            with open("paths.json", "r") as f:
                data = json.load(f)
        data[key] = path
        with open("paths.json", "w") as f:
            json.dump(data, f)

    def on_drop(self, event):
        
        filepath = event.data.strip('{}')
        print(filepath)
        #filepath[1:-1]
        
        if filepath.endswith('.xlsx'):
            filename =os.path.basename(filepath)
            print (filename)
            if filename.startswith('PRTMO. MULTIRED - SEGURO FALL_'):
                
                #filename = os.path.basename(filepath)
                self.drag_and_drop_entry.configure(state='normal')
                self.drag_and_drop_entry.delete(0, tkinter.END)
                self.drag_and_drop_entry.insert(0, filepath)
                self.drag_and_drop_entry.configure(state='readonly')
                CTkMessagebox(
                                title="Archivo Seleccionado", message=f"Archivo seleccionado: {filepath}", icon="check", option_1="Ok",
                                bg_color="#FFFFFF", fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                            )
            else:
                CTkMessagebox(
                    title="Archivo Incorrecto", 
                    message="El archivo seleccionado no es el esperado. Por favor seleccione un archivo que comience con 'MAQUETA-PRTMO. MULTIRED - SEGURO FALL'", 
                    icon="cancel", option_1="Ok", bg_color="#FFFFFF", fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615",
                    button_text_color="#FFFFFF"
                )
        else:
            CTkMessagebox(
                            title="Advertencia", message="Por favor seleccione un archivo .xlsx", icon="cancel", option_1="Ok",
                            bg_color="#FFFFFF", fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615",button_text_color="#FFFFFF"
                        )

    def ejecutar_script(self):

        try:
            filepath = self.drag_and_drop_entry.get()
            ventana_titulo = self.master.ventana_seleccionada
            ruta = self.File_Entry_down.get()
            if filepath and ventana_titulo and ruta:
                #S01_extraer_info_fallecidos.S01_extraer_info_fallecidos(filepath,ventana_titulo, ruta)
                CTkMessagebox(
                            title="Ejecucion Exitosa", message="Datos extraidos exitosamente.", icon="check", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                        )
            else:
                CTkMessagebox(
                            title="Error", message="Por favor, completar los campos.", icon="cancel", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                        )
        except Exception as e:
            CTkMessagebox(
                    title="Error", message=f"Ocurrió un error: {str(e)}", icon="cancel", option_1="Ok", bg_color="#FFFFFF", fg_color="#FFFFFF",
                    button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                )

class HomeFrame(BaseFrame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master 

        self.scrip02_title_label= customtkinter.CTkLabel(self, text="Abono Fallecidos",font=customtkinter.CTkFont(size=20, weight="bold"), text_color="#BF0615")
        self.scrip02_title_label.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")

        #Frame de ejecucion de S02_extraer_info_cuotas
        self.Script2_frame = customtkinter.CTkFrame(self, border_color="#B8B2B2", border_width=1, fg_color="transparent")
        self.Script2_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.label_Script2 = customtkinter.CTkLabel(master=self.Script2_frame, text="Ingresar a la pantalla PRAH6302 abono regulación - fallecidos ", font=customtkinter.CTkFont(size=15,weight="bold"))
        self.label_Script2.grid(row=0, column=0, padx=10, pady=10, sticky="")

        #self.image_scrip2 = customtkinter.CTkImage(Image.open("img/PRAH6302.png"), size=(450, 400))
        #self.image_scrip2_label = customtkinter.CTkLabel(master=self.Script2_frame, compound="center", image=self.image_scrip2, text="")
        #self.image_scrip2_label.configure(justify="center")
        #self.image_scrip2_label.grid(row=1, column=0, columnspan=4, padx=100, pady=10, sticky="")

        self.button_script2 = customtkinter.CTkButton(master=self.Script2_frame , text="Ejecutar", font=customtkinter.CTkFont(size=15,weight="bold"), 
                                                  border_width=1, border_color="#BF0615", fg_color="#FFFFFF", text_color="#BF0615", hover_color="#F2CDD0", command =self.ejecutar_script2)
        self.button_script2.grid(row=2, column=3, padx=50, pady=10, sticky="n")

    def ejecutar_script2(self):
        if os.path.exists("paths.json"):
            with open("paths.json", "r") as f:
                data = json.load(f)
            ruta = data.get("input_path")

        try:
            ventana_titulo = self.master.ventana_seleccionada
            resultado = 1 #S02_abono_fallecidos.S02_abono_fallecidos(ruta, ventana_titulo)

            if resultado == False:
                CTkMessagebox(
                            title="Error", message="Por favor, verificar si hay campos vacios antes de continuar.", icon="cancel", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                            )
            else:
                CTkMessagebox(
                            title="Ejecucion Exitosa", message="Cuentas abondas exitosamente.", icon="check", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                            )
        except Exception as e:
                CTkMessagebox(
                        title="Error", message=f"Ocurrió un error: {str(e)}", icon="cancel", option_1="Ok", bg_color="#FFFFFF", fg_color="#FFFFFF",
                        button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                        )


class SecondFrame(BaseFrame):
    def __init__(self, master):
        super().__init__(master)

        self.scrip03_title_label= customtkinter.CTkLabel(self, text="Modificación de Situación a FALLECIDO",font=customtkinter.CTkFont(size=20, weight="bold"), text_color="#BF0615")
        self.scrip03_title_label.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")

        #Frame de ejecucion de S02_extraer_info_cuotas
        self.Script3_frame = customtkinter.CTkFrame(self, border_color="#B8B2B2", border_width=1, fg_color="transparent")
        self.Script3_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.label_Script3 = customtkinter.CTkLabel(master=self.Script3_frame, text="Ingresar a la pantalla PRAH6307 administración de situaciones", font=customtkinter.CTkFont(size=15,weight="bold"))
        self.label_Script3.grid(row=0, column=0, padx=10, pady=10, sticky="")

        #self.image_scrip3 = customtkinter.CTkImage(Image.open("img/PRAH6307.png"), size=(450, 400))
        #self.image_scrip3_label = customtkinter.CTkLabel(master=self.Script3_frame, compound="center", image=self.image_scrip3, text="")
        #self.image_scrip2_label.configure(justify="center")
        #self.image_scrip3_label.grid(row=1, column=0, columnspan=4, padx=100, pady=10, sticky="")

        self.button_script3 = customtkinter.CTkButton(master=self.Script3_frame , text="Ejecutar", font=customtkinter.CTkFont(size=15,weight="bold"), 
                                                  border_width=1, border_color="#BF0615", fg_color="#FFFFFF", text_color="#BF0615", hover_color="#F2CDD0", command =self.ejecutar_script3)
        self.button_script3.grid(row=2, column=3, padx=50, pady=10, sticky="n")

    def ejecutar_script3(self):
        if os.path.exists("paths.json"):
            with open("paths.json", "r") as f:
                data = json.load(f)
            ruta = data.get("input_path")

        try:
            ventana_titulo = self.master.ventana_seleccionada
            if ruta:
                #S03_modificar_estado_fallecido.S03_modificar_estado_fallecido (ruta, ventana_titulo)
                CTkMessagebox(
                            title="Ejecucion Exitosa", message="Situacion modifacada exitosamente.", icon="check", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                            )
            else:
                CTkMessagebox(
                            title="Error", message="Por favor, verificar ruta del archivo.", icon="cancel", option_1="Ok", bg_color="#FFFFFF",
                            fg_color="#FFFFFF", button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                            )
        except Exception as e:
                CTkMessagebox(
                        title="Error", message=f"Ocurrió un error: {str(e)}", icon="cancel", option_1="Ok", bg_color="#FFFFFF", fg_color="#FFFFFF",
                        button_color="#4D4D4D", button_hover_color="#BF0615", button_text_color="#FFFFFF"
                        )

if __name__ == "__main__":
    app = App()
    app.mainloop()