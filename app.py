import tkinter as tk 
from tkinter import END, Toplevel, ttk
from tkcalendar import Calendar
import openpyxl

app = tk.Tk()
app.title("Mi aplicación Tkinter")
app.geometry("900x600")

# Lista para almacenar los reportes
reportes = []

# FUNCIONES
def seleccionar_opcion(opcion):
    print("Opción seleccionada:", opcion)

    # Botón de enviar

def obtener_fecha(event):
    global cal, date_window
    date_window = Toplevel()
    date_window.grab_set()
    date_window.title("selecciona la fecha")
    date_window.geometry("250x220+590+370")
    cal = Calendar(date_window, selectmode="day", date_pattern="dd/mm/y")
    cal.place(x=0,y=0)

    submit_btn = tk.Button(date_window, text ="guardar", command = seleccionar_fecha)
    submit_btn.place(x=80,y=190)

def seleccionar_fecha():
    entry_fecha.delete(0, END)
    entry_fecha.insert(0, cal.get_date())
    date_window.destroy()

def añadir_reporte():
    reporte = texto_nuevo_reporte.get("1.0", "end-1c")
    # Si el texto no está vacío, agregarlo al cuadro de texto principal
    if reporte.strip():
        texto_reporte.insert("end", f"\n\nReporte {len(reportes) + 1}:\n{texto_nuevo_reporte.get('1.0', 'end-1c')}")
        reportes.append(texto_reporte)
        texto_nuevo_reporte.delete("1.0", "end")  # Limpiar el cuadro de texto de entrada

def guardar_reportes():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reportes"
    
    for idx, reporte in enumerate(reportes, start=1):
        ws[f'A{idx}'] = reporte.get("1.0", "end-1c")
    
    wb.save("reportes.xlsx")

def on_select(event):
    widget = event.widget
    index = widget.curselection()[0]
    value = widget.get(index)
    print(value)

# FUNCIONES FIN
TIPOS_REPORTES = ["mantenimiento", "avionicos/radios", "electricos", "otros"]
# Datos para el menú desplegable
DATOS = [
    "00 Introduction	Introducción",
    "05 Time Limits/ Maintenance Check	Límites de tiempo",
    "06 Dimension and Areas	Dimensiones y Áreas",
    "07 Lifting and Shoring	Levantamiento y Anclaje",
    "08 Levelling and Weighing	Nivelación y Peso",
    "09 Towing and Taxiing	Remolque y Rodaje",
    "10 Parking and Mooring	Estacionamiento y Anclaje",
    "11 Placards and Markings	Letreros y Señalamientos",
    "12 Servicing - Routing Maintenance	Servicio - Mantenimiento de Rutina",
    "20 Standard Practices - Airframe	Prácticas Estándar - Airframe",
    "21 Air Conditioning	Aire Acondicionado",
    "22 Auto Flight	Piloto Automático",
    "23 Communications	Comunicaciones",
    "24 Electrical Power	Sistema Eléctrico",
    "25 Equipment/Furnishing	Equipo y Accesorios",
    "26 Fire Protection	Protección Contra Fuego",
    "27 Flight Control	Controles de Vuelo",
    "28 Fuel	Combustible",
    "29 Hydraulic Power	Sistema Hidráulico",
    "30 Ice and Rain Protection	Protección contra hielo y lluvia",
    "31 Indicating/Recording System	Sistemas de Indicación y Grabación",
    "32 Landing Gear	Tren de Aterrizaje",
    "33 Lights	Luces",
    "34 Navigation	Navegación",
    "35 Oxygen	Oxígeno",
    "36 Pneumatic	Sistema Neumático",
    "37 Vacuum	Presión y Vacío",
    "38 Water/Waste	Aguas y Desechos",
    "39 Electrical/Electronic Panel	Panel Eléctrico/Electrónico",
    "41 Water ballast	Balance de Agua",
    "45 Central Maintenance System (CMS)	Sistema de Mantenimiento Central",
    "46 Information Systems	Sistemas de Información",
    "47 Nitrogen Generation System	Sistema de Generación de Nitrógeno",
    "49 Airborne Auxiliary Power	Unidad de Potencia Auxiliar (APU)",
    "50 Cargo and Accessory Compartments	Compartimientos de Carga y Accesorios",
    "51 Standard Practices and Structures - General	Prácticas Estándar y Estructuras - General",
    "52 Doors	Puertas",
    "53 Fuselage	Fuselaje",
    "54 Nacelles/Pylons	Nacelles/Pylons",
    "55 Stabilizers	Estabilizadores",
    "56 Windows	Ventanas",
    "57 Wings	Alas",
    "60 Standar Practices - Propeller/Rotor	Prácticas Estándar - Propelas/Rotores",
    "61 Propellers/Propulsors	Hélices y Propulsores",
    "62 Main Rotor	Rotor Pincipal",
    "63 Main Rotor Drives	Impulsor del Rotor",
    "64 Tail Rotor	Rotor de Cola",
    "65 Tail Rotor Drive	Impulsor del Rotor de Cola",
    "66 Rotor Blade and Tail Pylon Folding	Palas Plegables y Pilones",
    "67 Rotors Flight Control	Controles de Vuelo del Rotor",
    "70 Standard Practices - Engines	Prácticas Estándar - Motores",
    "71 Powerplant	Planta Propulsora",
    "72 Engine	Motor",
    "73 Engine Fuel and Control	Sistema de Combustible del motor",
    "74 Ignition	Ignición",
    "75 Air	Aire",
    "76 Engine Controls	Controles del Motor",
    "77 Enging Indicating	Indicadores del Motor",
    "78 Exhaust	Escape",
    "79 Oil	Aceite",
    "80 Starting	Arranque",
    "81 Turbines (Reciprocating Engines)	Turbinas (Motores Recíprocos)",
    "82 Water Injection	Inyección de Agua",
    "83 Accessory Gear Boxes (Engine Driven)	Cajas de Engranajes de Accesorios",
    "84 Propulsion Augmentation	Incremento de la Propulsión",
    "91 Charts	Gráficos y Diagramas",
    "95 Special Equipment	Equipamiento Especial"
]

# Crear el Listbox
listbox = tk.Listbox(app)
listbox.grid(row=4, column=0, sticky="nsew")

# Configurar el scrollbar
scrollbar = ttk.Scrollbar(app, orient="vertical", command=listbox.yview)
scrollbar.grid(row=4, column=1, sticky="ns")


for dato in DATOS:
    listbox.insert("end", dato)

listbox.bind("<<ListboxSelect>>", on_select)
# Configurar el grid para que se expanda con la ventana
# app.grid_columnconfigure(0, weight=1)
# app.grid_rowconfigure(0, weight=1)

opcion_seleccionada = tk.StringVar(app)
opcion_seleccionada.set("Tipo Reporte") 

menu_desplegable = tk.OptionMenu(app, opcion_seleccionada, *TIPOS_REPORTES, command=seleccionar_opcion)
menu_desplegable.grid(row=3, column=0, padx=10, pady=10)


label_matricula = tk.Label(app, text="matricula aeronave HK:")
label_matricula.grid(row=0, column=0)
entry_matricula = tk.Entry(app)
entry_matricula.grid(row=0, column=1)


label_fecha = tk.Label(app, text="fecha:")
label_fecha.grid(row=1, column=0)

entry_fecha = tk.Entry(app)
entry_fecha.grid(row=2, column=0)
entry_fecha.bind("<1>", obtener_fecha)

# Cuadro de texto para ingresar nuevos reportes
texto_nuevo_reporte = tk.Text(app, height=4, width=40)
texto_nuevo_reporte.grid(row=5, column=0, padx=10, pady=5)

# Crear la casilla de texto
texto_reporte = tk.Text(app, height=10, width=40)
texto_reporte.grid(row=6, column=0, padx=10, pady=10)

# Botón para añadir reporte
boton_añadir_reporte = tk.Button(app, text="Añadir Reporte", command=añadir_reporte)
boton_añadir_reporte.grid(row=7, column=0, pady=5)

# Botón para guardar reportes
boton_guardar_reportes = tk.Button(app, text="Guardar Reportes en Excel", command=guardar_reportes)
boton_guardar_reportes.grid(row=8, column=0, pady=5)

app.mainloop()

