'''GEST 2020: Gestor de compras y materiales
 para pequeñas empresas basado en software GEST;
diseñado para la empresa LIE SRL. Por Agustin Arnaiz'''

import sqlite3  # BDD
import pandas as pd
import tkinter  # GUI
from tkinter import *
from tkinter import filedialog  # para abrir archivos
from tkinter import messagebox
from tkinter import ttk

#-----------------------------CLASES--------------------------------
class convert_sheet:    #selecciona un XLS y convierte a SQLite *.db con el mismo nombre
    def __init__(self, window):
        self.wind = window
        self.open_file()

    #ABRE XLS y su PATH
    def open_file(self):
        self.file_path = filedialog.askopenfilename(title='Abrir archivo XLS', filetypes=(
                                                    ('planilla de cálculo', '*.xls'), ('todos los archivos', '*.*')))
        try:
            self.xls_file = pd.read_excel(self.file_path)
            self.create_db()
        except:
            pass

    # checkea que no exista el db y lo crea
    def create_db(self):
        self.db_filename = self.change_extension(self.file_path)
        try:
            self.database = sqlite3.connect(self.db_filename)
            self.copy_xls_db()
            messagebox.showinfo('XLS to SQLite',f'Database {self.db_filename}\ncreado con éxito')
        except:
            messagebox.showwarning('XLS to SQLite','El archivo ya existe')

    # replica el PATH con la extension *.db
    def change_extension(self, path):
        n=-1
        while path[n] != '.':
            n -= 1
            if path[n] == '.':
                return path[0:(n)]+'.db'

    #Copia Archivo XLS a Database
    def copy_xls_db(self):
        self.xls_file.to_sql(name='hoja1', con=self.database)    #convierte la hoja 0 a db
        self.database.commit()
        self.database.close()

class db_listas:
    db_name = 'LISTAS.db'

    def __init__(self, window):
        self.wind = window
        self.sheet_name = 'hoja1'

        #define nombres de columnas de la database
        self.sheet_columns = []
        self.read_columns()

        # define FRAMEWORK
        frame = LabelFrame(self.wind, text='Nuevo Elemento', labelanchor=N)
        frame.grid(row=1, column=0, columnspan=3, pady=20)  # el frame esta en la fila 0

        # Mensaje de salida en la ventana
        self.message = Label(text='', fg='red')
        self.message.grid(row=3, column=0, columnspan=2, sticky=W + E)

        # creacion de tabla
        self.tree = ttk.Treeview(height=20, columns=4)
        self.tree.grid(row=4, column=0, columnspan=2)
        self.tree["columns"] = ("Elemento", "Cantidad", "Unidad")
        self.tree.column("#0", width=150, minwidth=50)
        self.tree.column("#1", width=150, minwidth=50)
        self.tree.column("#2", width=50, minwidth=50)
        self.tree.column("#3", width=50, minwidth=50)
        self.tree.heading('#0', text='Lista', anchor=CENTER)
        self.tree.heading('#1', text='Elemento', anchor=CENTER)
        self.tree.heading('#2', text='Cant', anchor=CENTER)
        self.tree.heading('#3', text='UN', anchor=CENTER)
        scrollVertical = Scrollbar(command=self.tree.yview)  # barra vertical de desplazamiento
        scrollVertical.grid(row=4, column=3, sticky='nsew')
        self.tree.config(yscrollcommand=scrollVertical.set)


        # acciones sobre la lista
        ttk.Button(text='Agregar Elemento', command=self.add_element).grid(row=0, column=0, sticky=W + E)
        #ttk.Button(text='Buscar Lista', command=self.search_lista).grid(row=0, column=1, sticky=W + E)
        ttk.Button(text='Editar Elemento', command=self.edit_lista).grid(row=1, column=0, sticky=W + E)
        ttk.Button(text='Copiar Lista', command=self.copy_list).grid(row=1, column=1, sticky=W + E)
        ttk.Button(text='Eliminar Elemento', command=self.delete_elemento).grid(row=2, column=0, sticky=W + E)
        ttk.Button(text='Eliminar Lista', command=self.delete_list).grid(row=2, column=1, sticky=W + E)

        # lee la base de datos completa
        self.get_products()

    def read_columns(self):
        query = f'SELECT * FROM "{self.sheet_name}"'
        columns = self.run_query(query)
        for columna in columns.description:
            self.sheet_columns.append(columna[0]) #crea lista de nombres de columnas el 0 es index

    def run_query(self, query, parameters={}):  # conexion a base de datos
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    def get_products(self):
        records = self.tree.get_children()  # obtiene todos los datos de la tabla tree
        for element in records:
            self.tree.delete(element)  # limpia todos los datos de tree
        # consulta de datos
        query = f'SELECT * from {self.sheet_name} ORDER BY "{self.sheet_columns[1]}" DESC '  # {nombre tabla} {nombre de columna}
        db_rows = self.run_query(query)  # lee los datos de la database
        # rellenando los datos de database en el tree
        for row in db_rows:
            self.tree.insert('', 0, text=row[1],
                             values=(row[2], row[3], row[4]))  # el 1ro no va nada, xq es el indice de fila

    def validation(self):  # aprueba la escritura en DB
        return len(self.cod_lista.get()) !=0 and len(self.elemento.get()) !=0# retorna true si los valores no estan vacios

    def add_element(self):
        self.add_window = Toplevel()
        self.add_window.title = 'Agregar Elemento'
        frame = LabelFrame(self.add_window, text='Agregar Elemento')
        frame.grid(row=0, column=0, columnspan=2, pady=20)

        # Lista input
        Label(frame, text='Lista: ').grid(row=1, column=0)
        self.cod_lista = Entry(frame)  # se guarda en lista para despues ser manipulada
        self.cod_lista.grid(row=1, column=1)
        self.cod_lista.focus()  # posiciona el cursor sobre este input en cada inicio

        # Elemento input
        Label(frame, text='Código: ').grid(row=2, column=0)
        self.elemento = Entry(frame)
        self.elemento.grid(row=2, column=1)

        # cantidad input
        Label(frame, text='Cantidad: ').grid(row=3, column=0)
        self.cantidad = Entry(frame)
        self.cantidad.grid(row=3, column=1)

        # unidad input
        Label(frame, text='Unidad: ').grid(row=4, column=0)
        self.unidad = Entry(frame)
        self.unidad.grid(row=4, column=1)

        # BOTON guardar elemento
        Button(self.add_window, text='Guardar Elemento', command=self.save_element).grid(row=5, columnspan=2, sticky=W + E)

    def save_element(self):
        if self.validation():
            query = f'INSERT INTO {self.sheet_name} VALUES(NULL, ?, ?, ?, ?)'  # products es la hoja de DB, null es el indice, y los 4 ? son los valores a escribir
            parameters = (self.cod_lista.get(), self.elemento.get(), self.cantidad.get(), self.unidad.get())
            self.run_query(query, parameters)
            self.message['fg'] = 'green'
            self.message['text'] = 'El Elemento ha sido guardado con éxito'
            self.cod_lista.delete(0, END)  # despues de guardar el elemento borra las casillas
            self.elemento.delete(0, END)
            self.cantidad.delete(0, END)
            self.unidad.delete(0, END)
        else:
            self.message['fg'] = 'red'
            self.message['text'] = 'Lista y Elemento son requeridos'
        self.get_products()

    def delete_elemento(self):
        self.message['text'] = ''  # limpiar el texto del mensaje
        try:
            self.tree.item(self.tree.selection())['text'][0]  # obtiene el nombre del elemento seleccionado
        except IndexError as e:
            self.message['fg'] = 'red'
            self.message['text'] = 'Debe seleccionar una fila para borrar'
            return
        lista_sel = self.tree.item(self.tree.selection())['text']
        elemento_sel= self.tree.item(self.tree.selection())['values'][0] # pasa la seleccion de lista mas elemento
        query = f'DELETE FROM "{self.sheet_name}" WHERE "{self.sheet_columns[1]}" = ? AND "{self.sheet_columns[2]}" = ?'
        self.run_query(query, (lista_sel, elemento_sel))  # pone la coma para que se entienda que es una tupla
        self.message['fg'] = 'green'
        self.message['text']= f'El elemento {elemento_sel} ha sido eliminado'
        self.get_products()  # actualiza la tabla

    def delete_list(self):
        self.message['text'] = ''  # limpiar el texto del mensaje
        try:
            self.tree.item(self.tree.selection())['text'][0]  # obtiene el nombre del elemento seleccionado
        except IndexError as e:
            self.message['fg'] = 'red'
            self.message['text'] = 'Debe seleccionar una fila para borrar'
            return
        lista_sel = self.tree.item(self.tree.selection())['text']
        query = f'DELETE FROM "{self.sheet_name}" WHERE "{self.sheet_columns[1]}" = ?'
        self.run_query(query, (lista_sel, ))
        self.message['fg'] = 'green'
        self.message['text']= f'La lista {lista_sel} ha sido eliminado'
        self.get_products()

    def edit_lista(self):
        self.message['text'] = ''
        try:  # comprobacion de seleccion de elemento a editar
            self.tree.item(self.tree.selection())['text'][0]
        except IndexError as e:
            self.message['fg'] = 'red'
            self.message['text'] = 'Debe seleccionar una fila para editar'
            return
        lista_cod = self.tree.item(self.tree.selection())['text']
        elemento = self.tree.item(self.tree.selection())['values'][0]
        cantidad = self.tree.item(self.tree.selection())['values'][1]
        unidad = self.tree.item(self.tree.selection())['values'][2]
        self.edit_window = Toplevel()

        # valor anterior
        Label(self.edit_window, text='Previo: ').grid(row=0, column=2)
        Entry(self.edit_window, textvariable=StringVar(self.edit_window, value=lista_cod), state='readonly').grid(row=1,
                                                                                                              column=2)
        Entry(self.edit_window, textvariable=StringVar(self.edit_window, value=elemento), state='readonly').grid(row=2,
                                                                                                                 column=2)
        Entry(self.edit_window, textvariable=StringVar(self.edit_window, value=cantidad), state='readonly').grid(row=3,
                                                                                                                 column=2)
        Entry(self.edit_window, textvariable=StringVar(self.edit_window, value=unidad), state='readonly').grid(row=4,
                                                                                                               column=2)

        # valor nuevo
        Label(self.edit_window, text='Actual:').grid(row=0, column=3)
        nueva_lista_cod = Entry(self.edit_window)
        nueva_lista_cod.insert(END, lista_cod)
        nueva_lista_cod.grid(row=1, column=3)
        nueva_lista_cod.focus()
        nuevo_elemento = Entry(self.edit_window)
        nuevo_elemento.insert(END, elemento)
        nuevo_elemento.grid(row=2, column=3)
        nueva_cantidad = Entry(self.edit_window)
        nueva_cantidad.insert(END, cantidad)
        nueva_cantidad.grid(row=3, column=3)
        nueva_unidad = Entry(self.edit_window)
        nueva_unidad.insert(END, unidad)
        nueva_unidad.grid(row=4, column=3)

        Button(self.edit_window, text='Modificar elemento',
               command=lambda: self.edit_record(nueva_lista_cod.get(), lista_cod,
                                                nuevo_elemento.get(),elemento,
                                                nueva_cantidad.get(), cantidad,
                                                nueva_unidad.get(), unidad)).grid(row=5, column=2, columnspan=2, sticky=W+E)

    def edit_record(self, nueva_lista_cod, lista_cod, nuevo_elemento, elemento, nueva_cantidad, cantidad, nueva_unidad, unidad):
        query = f'UPDATE {self.sheet_name} SET "{self.sheet_columns[1]}"= ?, "{self.sheet_columns[2]}" = ?, "{self.sheet_columns[3]}" = ?, "{self.sheet_columns[4]}" = ? WHERE "{self.sheet_columns[1]}" = ? AND "{self.sheet_columns[2]}" = ? AND "{self.sheet_columns[3]}" = ? AND "{self.sheet_columns[4]}" = ? '
        parameters = (nueva_lista_cod, nuevo_elemento, nueva_cantidad, nueva_unidad, lista_cod, elemento, cantidad, unidad)
        self.run_query(query, parameters)
        self.edit_window.destroy()
        self.message['fg'] = 'green'
        self.message['text'] = f'La lista {lista_cod} ha sido actualizada'
        self.get_products()

# ------------FUNCIONES---------------------------------------------
def exit_app():  # funcion para cerrar app
    raiz.destroy()

def convert_xls():
    convert_sheet(raiz)

def call_listas():
    db_listas(raiz)

def SinDeterminar():
    messagebox.showinfo('ADVERTENCIA', 'Acción no programada')

def infoAyuda():
    messagebox.showinfo('Ayuda GEST2020', 'Solicitar manual a su proveedor')

def infoAdicional():  # funcion para vent emergente que muestra info con icono de info
    messagebox.showinfo('Gestor de empresa', 'GEST2020 Versión: v1.0')  # primero titulo ventana, luego texto a mostrar

def infoLicencia():  # funcion para ventana emergente que muestra un warning con icono warning
    messagebox.showinfo('GEST2020', 'licencia válida')

#---------------MAIN, Window, FRAMEWORK-----------------------------
if __name__ == '__main__':
    # ------------VENTANA PRINCIPAL---------------------------
    raiz = tkinter.Tk()
    raiz.title('GEST2020 | Gestor y administrador de empresa')  # titulo de la ventana
    raiz.resizable(True, True)  # redimencionar ancho? y alto?
    raiz.iconbitmap('LOGOLIE x3.ico')  # icono e la ventana
    raiz.config(bg="grey")  # config de la ventana bg= back ground color

    # -------------FRAMEWORK---------------------------------
    miFrame = Frame(raiz)  # pertenece a Tk
    miFrame.config(bg="grey")  # color de fondo del frame
    miFrame.config(bd=5)  # ancho borde de frame
    miFrame.config(relief='groove')  # modela tipo de bordes del frame
    miFrame.config(cursor='arrow')  # tipo cursor


    # ------------BARRA DE MENU--------------------------------------
    barraMenu = tkinter.Menu(raiz)
    raiz.config(menu=barraMenu, width=800, height=600)

    archivoMenu = tkinter.Menu(barraMenu, tearoff=0)
    archivoMenu.add_command(label="Convertir Planilla a base de datos", command=convert_xls)
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Maestro de articulos")
    archivoMenu.add_command(label="Abrir Listas", command=call_listas)
    archivoMenu.add_command(label="Ordenes de trabajo")
    archivoMenu.add_separator()
    archivoMenu.add_command(label="Cerrar programa", command=exit_app)

    ayudaMenu = tkinter.Menu(barraMenu, tearoff=0)
    ayudaMenu.add_command(label="Ayuda", command=infoAyuda)
    ayudaMenu.add_command(label="Licencia", command=infoLicencia)
    ayudaMenu.add_separator()
    ayudaMenu.add_command(label="Acerca de", command=infoAdicional)

    barraMenu.add_cascade(label="Archivo", menu=archivoMenu)
    barraMenu.add_cascade(label="Ayuda", menu=ayudaMenu)
    # ------------BARRA DE MENU--------------------------------------

    raiz.mainloop()
