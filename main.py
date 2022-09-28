import pandas as pd
import PySimpleGUI as sg


EXCEL = "mi_archivo_excel.xlsx" # Aqui el archivo de excel en el mismo directorio (carpeta) que este archivo main.py
df = pd.read_excel(EXCEL)

# Tema.
sg.theme("reddit") # Nombre del tema PySimpleGUI

# Diseño del (Layout) de la Pestaña 1: Registro.
pestana1 = [
        [sg.Text()],
        [sg.Text('Nombre', size=(15, 1), font="bold"), sg.In(key='Nombre'), ],
        [sg.Text('Apellido', size=(15, 1), font="bold"), sg.In(key='Apellido')],
        [sg.Text('DNI', size=(15, 1), font="bold"), sg.In(key='DNI')],
        [sg.Text('Obra Social', size=(15, 1), font="bold"), sg.Combo(['OSDE', 'IOMA', 'Sancor Salud'],
                                                                     key='Obra Social', size=(10, 1)),
         sg.Text('Nro. Afiliado', size=(10, 1), font="bold"), sg.In(key='Nro Afiliado', size=(16, 1))],
        [sg.Text('Consultorio', size=(15, 1), font="bold"), sg.Combo(['Capital Federal', "Virtual"],
                                                                     key='Consultorio', size=(10, 1)),
         sg.Text('Edad', size=(10, 1), font="bold"), sg.Combo(['Niño', 'Adolescente', "Adulto", "Adulto Mayor"],
                                                              key='Edad', size=(14, 1))],
        [sg.Text("")],
        [sg.Text("", size=(23, 1)), sg.Submit("Guardar"), sg.Button('Limpiar campos'), sg.Exit("Salir")]
        # Botones eventos
]

# Diseño (Layout) de la Pestaña 2: Búsqueda.
pestana2 = [
        [sg.Text("Ingrese Nombre", font="bold"), sg.In(size=(55, 1), key="-NOM-")],
        [sg.Text("Ingrese Apellido", font="bold"), sg.In(size=(55, 1), key="-APE-")],

        # Botones
        [sg.Text("", size=(18, 1)),
         sg.Submit("Buscar"),
         sg.Cancel("Nueva Búsqueda"),
         sg.Button("Eliminar"),
         sg.Exit("Salir")],

        # Frame de los resultados obtenidos
        [sg.Frame('RESULTADOS',
                  [[sg.Text("Nro Afiliado", size=(14, 1), justification="right"),
                    sg.Multiline(key="-OUTAFL-", size=(18, 1), background_color="white",
                                 font="bold", sbar_relief="flat"),  # Salida Nro afiliado.

                    sg.Text("Nro DNI", size=(7, 1), justification="right"),
                    sg.Multiline(key="-OUTDNI-", size=(20, 1), background_color="white",
                                 font="bold", sbar_relief="flat")],  # Salida Nro. DNI

                   [sg.Text("Consultorio", size=(14, 1), justification="right"),
                    sg.Multiline(key="-OUTCON-", size=(18, 1), background_color="white",
                                 font="bold", sbar_relief="flat"),  # Salida consultorio

                    sg.Text("O. Social", size=(7, 1), justification="right"),
                    sg.Multiline(key="-OUTOBS-", size=(20, 1), background_color="white",
                                 font="bold", sbar_relief="flat")],
                   [sg.Text("Edad", size=(14, 1), justification="right"),
                    sg.Multiline(key="-OUTEDA-", size=(18, 1), background_color="white",
                                 font="bold", sbar_relief="flat")]],  # Salida Obra Soc.

                  title_color="lightblue4",
                  title_location="n",
                  font="bold",
                  relief="ridge",
                  size=(590, 150))]
]

pestana3 = [
        [sg.Text("Aquí algunas instrucciones sobre el uso de ésta app",
                 font="bold", background_color="white", text_color="lightblue3")],
        [sg.Text("El registro del paciente se crea segun lo rellenado en los campos de la pestaña \n'Registro',"
                 " por eso debes tener cuidado con los errores de tipeo, incluidos los \nespacios extras,"
                 " las mayúsculas, las minúsculas, etc.", background_color="white", text_color="grey50")],
        [sg.Text(" Para buscar, simplemente rellena los campos con los datos exactamente\n"
                 + "igual a cuando fueron ingresados. En caso de que no exista el registro, o se \n"
                 + "ingrese mal, no se mostrará nada en pantalla. También está la opción de Eliminar\n"
                 + "un registro. No se pueden borrar registros que no existen. Por favor evitar enviar\n"
                 + "el formulario de registro totalmente en blanco." + "\n\n¡GRACIAS!", background_color="white",
                 text_color="grey50")]
]

# Agrupación de pestañas
tab_group_layout = [[sg.Tab('Registro', pestana1, key="-PES1-"),
                     sg.Tab('Búsqueda', pestana2, key="-PES2-"),
                     sg.Tab('Cómo usar', pestana3, background_color="white")]]

# Diseño del grupo de pestañas
layout = [[sg.TabGroup(tab_group_layout, size=(600, 250),
                       background_color="grey89",
                       selected_title_color="white",
                       tab_background_color="grey90",
                       selected_background_color="lightblue",
                       tab_border_width=2,
                       border_width=0,
                       key="-TG-")]]

# Especificaciones de la ventana.
window = sg.Window('REGISTRO DE PACIENTES', layout,
                   use_custom_titlebar=True,
                   titlebar_background_color="lightblue",
                   #titlebar_icon="tbar.png",       ## Opcional agregar icono en la barra del título
                   size=(600, 350),
                   resizable=False)


# Para borrar los campos.
def limpiar_camps():
    for key in values:
        window[key]('')
        window["-OUTAFL-"]("")
        window["-OUTDNI-"]("")
        window["-OUTCON-"]("")
        window["-OUTOBS-"]("")
        window["-OUTEDA-"]("")
    return None


# Para buscar los datos del paciente.
def buscar():
    try:
        fnom = df[df['Nombre'] == values["-NOM-"]]  # Retorna la fila del valor buscado
        indx = fnom.index[0]
        if df.at[indx, "Nombre"] == values["-NOM-"] and df.at[indx, "Apellido"] == values["-APE-"]:
            dni = df.at[indx, "DNI"]
            osoc = df.at[indx, "Obra Social"]
            nroa = df.at[indx, "Nro Afiliado"]
            cons = df.at[indx, "Consultorio"]
            edad = df.at[indx, "Edad"]
            return dni, osoc, nroa, cons, edad
        else:
            return "s/reg", "s/reg", "s/reg", "s/reg", "s/reg"
    except:
        sg.popup("Verifique que los datos sean correctos," +
                 "inclusive mayúsculas y minúsculas en el/los nombre(s) y apellido(s).")
        return "s/reg", "s/reg", "s/reg", "s/reg", "s/reg"


def guardar():
    df.to_excel(EXCEL, index=False)


# Funcion para borrar un registro
def borrar():
    fnom = df[df['Nombre'] == values["-NOM-"]]  # Fila del valor a borrar
    indx = fnom.index[0]  # Indice de la fila del registro
    if df.at[indx, "Nombre"] == values["-NOM-"] and df.at[indx, "Apellido"] == values["-APE-"]:  # Condicional.
        df.drop(indx, axis=0, inplace=True)
    else:
        pass


# Mantener ventana abierta y realizar acciones
while True:
    event, values = window.read()  # Abre la ventana

    if event == sg.WIN_CLOSED or event == 'Salir' or event == 'Salir0':
        break  # Sale del bucle y cierra la ventana.

    if event == 'Limpiar campos' or event == "Nueva Búsqueda":
        limpiar_camps()

    if values["-TG-"] == '-PES1-':  # Válido la pestaña 1
        if event == 'Guardar':
            grabar_datos = pd.DataFrame(values, index=[0])
            df = pd.concat([df, grabar_datos], ignore_index=True)
            df.to_excel(EXCEL, index=False)
            sg.popup('¡Guardado con éxito!')
            limpiar_camps()

    elif values["-TG-"] == '-PES2-':  # Válido solo para la pestaña 2
        if event == 'Buscar':
            dni, osoc, nroa, cons, edad = buscar()
            window["-OUTDNI-"](dni)
            window["-OUTOBS-"](osoc)
            window["-OUTAFL-"](nroa)
            window["-OUTCON-"](cons)
            window["-OUTEDA-"](edad)
        # Condiciones para eliminar el registro.
        if event == "Eliminar":
            try:
                click = sg.popup_ok_cancel("¿Realmente quieres eliminar este paciente?",
                                           button_color="red",
                                           text_color="white",
                                           font="bold",
                                           no_titlebar=True)
                # Eventos del popup
                if click == 'OK':
                    borrar()
                    guardar()
                if click == 'Cancel':
                    pass
                if click is None:
                    pass
            except:
                pass

window.close()  # Cierra la ventana.
