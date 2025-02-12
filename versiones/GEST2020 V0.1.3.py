'''GEST 2020: Gestor de compras y materiales para pequeñas empresas basado en software GEST;
diseñado para la empresa LIE SRL. Por Agustin Arnaiz
* Salvo la clase UpperEntry'''

import sqlite3
import pandas as pd
import tkinter
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from datetime import date

#-----------------------------CLASES--------------------------------
#Pasa automaticamente a MAYUSCULAS textos de los entrys
class UpperEntry(tk.Entry):
    def __init__(self, parent, *args, **kwargs):
        self._var = kwargs.get("textvariable") or tk.StringVar(parent)
        super().__init__(parent, *args, **kwargs)
        self.configure(textvariable=self._var)
        self._to_upper()

    def config(self, cnf=None, **kwargs):
        self.configure(cnf, **kwargs)

    def configure(self, cnf=None, **kwargs):
        var = kwargs.get("textvariable")
        if var is not None:
            var.trace_add('write', self._to_upper)
            self._var = var
        super().config(cnf, **kwargs)

    def __setitem__(self, key, item):
        if key == "textvariable":
            item.trace_add('write', self._to_upper)
            self._var = item
        super.__setitem__(key, item)

    def _to_upper(self, *args):
        self._var.set(self._var.get().upper())

#convierte una planilla XLS a database SQLite3
class Convert_sheet:    #selecciona un XLS y convierte a SQLite *.db con el mismo nombre
    def __init__(self):
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
            self.db_sql = sqlite3.connect(self.db_filename)
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
        self.xls_file.to_sql(name='hoja1', con=self.db_sql)    #convierte la hoja 0 a db
        self.db_sql.commit()
        self.db_sql.close()

#Clase principal de base de datos
class Database:
    def __init__(self, window, db_name, sheet_name='hoja1'):
        self.window = window
        self.db_name = db_name
        self.sheet_name = sheet_name
        self.entrys_array = []

        #define nombres de COLUMNAS de la database
        self.sheet_columns = []
        self.read_columns()

        # define FRAMEWORK controles y mensajes
        self.frame_msg = LabelFrame(self.window, text='')
        self.frame_msg.grid(row=10, column=0, columnspan=len(self.sheet_columns), pady=10, padx=10)
        # define FRAMEWORK tabla de datos
        self.frame_tree = LabelFrame(self.window, text='')
        self.frame_tree.grid(row=0, column=0, columnspan=len(self.sheet_columns), pady=10, padx=10, sticky=E)

        # MENSAJE de salida en la ventana
        self.message = Label(self.window, text='', fg='blue')
        self.message.grid(row=10, column=0, columnspan=20, sticky=W + E)

        # Moldea ventana y lee base de datos
        self.build_main_view()
        self.database_to_tree()

        #EVENTOS de la ventana y comandos rápidos
        self.window.bind('<Return>', self.database_to_tree)   #busca según código en 'self.entrys_array[0]'
        self.window.bind('<Control-Return>', self.clean_entrys) #borra elementos de busqueda
        self.window.bind('<Control-a>', self.add_record)
        self.window.bind('<Control-e>', self.edit_record)
        self.window.bind('<Control-d>', self.delete_record)
        self.tree.bind('<space>', self.auto_scroll) #baja de a una hoja la vista de tree
        self.tree.bind('<<TreeviewSelect>>', self.load_edit_item)   #carga datos seleccionados
        self.tree.bind('<Double-Button-1>', call_listas)    #doble click abre la lista

        #Scroll vertical del TREE
        self.scroll_tree = Scrollbar(self.frame_tree, command=self.tree.yview)
        self.scroll_tree.grid(row=10, column=len(self.sheet_columns), sticky='nsew')
        self.tree.config(yscrollcommand=self.scroll_tree.set)

    #desplaza de a una hoja en la vista de tree con la barra 'space'
    def auto_scroll(self, *args):
        self.tree.yview_scroll(1,what='page')

    #Armado de ventana, TREEVIEW adaptable segun Database
    def build_main_view(self):
        # creacion de tabla para visualizar
        self.tree = ttk.Treeview(self.frame_tree, height=25, columns=len(self.sheet_columns))
        self.tree.grid(row=10, column=0, columnspan=len(self.sheet_columns), pady=10)

        # NOMBRES COLUMNAS
        nombres_columnas = []
        for each_column in self.sheet_columns:
            nombres_columnas.append(each_column)
        self.tree["columns"] = nombres_columnas[2:] #desde 2 xq el 1ro es index y el 2do esta en text (no en values)

        #Esta query es para dar TAMAÑO DINAMICO al ancho de columnas del tree
        query = f'SELECT * from "{self.sheet_name}" ORDER BY "{self.sheet_columns[1]}" DESC '
        dataframe = self.run_query(query)
        all_table = dataframe.fetchall()

        #se cargan titulos y tamaños de las columnas del tree, asi como Entrys para la edicion de records
        num_column = 0
        item_table = ''
        for each_column in self.sheet_columns[1:]:
            largest = ''
            for row_table in all_table[:25]: #lee las 1ras 25 lineas de cada columna, para estimar el ancho de la misma
                item_table = str(row_table[num_column+1])
                if len(largest) < len(item_table):
                    largest = item_table
            self.tree.column("#" + str(num_column), width=(50+len(largest)*5), minwidth=30, stretch=False)
            self.tree.heading('#' + str(num_column), text=f'{self.sheet_columns[num_column+1]}', anchor=CENTER)

            # crea tantos ENTRYS como columnas, para editar registro
            self.entrys_array.append(self.entrys(frame=self.frame_tree,
                                                 name=each_column, row=7, column=num_column, width=len(largest) + 4))

            num_column += 1

    #Crea Entrys con label superior
    def entrys(self, frame, name='entry', row=0, column=0, width=50, textvariable=''):
        # define label y entrys segun llamada
        Label(frame, text=name).grid(row=row, column=column)
        entry = UpperEntry(frame, width=width, textvariable=textvariable)
        entry.config(fg="blue")
        entry.grid(row=row+2, column=column, columnspan=1)
        return entry

    #Borra los entrys de registro
    def clean_entrys(self, *args):
        for entry_element in self.entrys_array:
            entry_element.delete(0, 'end')
        self.database_to_tree()

    #Lee los nombres de las columnas de la database y su cantidad
    def read_columns(self):
        query = f'SELECT * FROM "{self.sheet_name}"'
        columns = self.run_query(query)
        for columna in columns.description:
            self.sheet_columns.append(columna[0]) #crea lista de nombres de columnas el 0 es index

    #Ejecuta una QUERY SQLite3 con cursor usando parametros
    def run_query(self, query, parameters={}):  # conexion a base de datos
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    #Copia datos masivos pasados por parametro (la usa listas y OT)
    def run_query_many(self, query, parameters={}):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.executemany(query, parameters)
            conn.commit()
        return result

    #Limpia el tree, y lo re-hace segun codido exacto o simil (like) y ordena segun column
    def database_to_tree(self, *args, like=True, col_search=1, col_order=1, order='DESC'):   #la columna 0 suele ser index, el código es la #1
        self.delete_tree()
        valores=[]
        if like:
            query = f'SELECT * from "{self.sheet_name}" WHERE "{self.sheet_columns[col_search]}" ' \
                f'LIKE "{self.entrys_array[col_search-1].get()}%" ORDER BY "{self.sheet_columns[col_order]}" {order}'
        else:
            query = f'SELECT * from "{self.sheet_name}" WHERE "{self.sheet_columns[col_search]}" ' \
                    f'LIKE "{self.entrys_array[col_search].get()}" ORDER BY "{self.sheet_columns[col_order]}" {order}'
        dataframe = self.search_database(*args, like=like, col_search=col_search, col_order=col_order)
        for row in dataframe:
            self.tree.insert('', 0, text=row[1], values=row[2:])#el 1ro no va nada xq es el indice de fila,
        # Hace foco en buscar registro por código
        self.entrys_array[0].focus()

    #Ejecuta query de busqueda en DB, retorna dataframe
    def search_database(self, *args, like=True, col_search=1, col_order=1, order='DESC'):
        if like:                #hay un entry menos que columnas (por index) de ahi que se resta uno
            query = f'SELECT * from "{self.sheet_name}" WHERE "{self.sheet_columns[col_search]}" ' \
                f'LIKE "{self.entrys_array[col_search - 1].get()}%" ORDER BY "{self.sheet_columns[col_order]}" {order}'
        else:
            query = f'SELECT * from "{self.sheet_name}" WHERE "{self.sheet_columns[col_search]}" ' \
                    f'LIKE "{self.entrys_array[col_search - 1].get()}" ORDER BY "{self.sheet_columns[col_order]}" {order}'
        return self.run_query(query)

    #Valida operacion si hay seleccion de registro en TREE
    def valid_selection(self):  # aprueba la escritura en DB
        seleccion = self.tree.item(self.tree.selection())['text']
        return seleccion != ''

    #Valida operación para agregar registro si el mismo no es vacío o repetido
    def valid_add(self):
        if self.entrys_array[0].get() == '':
            return False
        query = f'SELECT * from {self.sheet_name}'
        db_rows = self.run_query(query)
        for row in db_rows:
            if self.entrys_array[0].get() == row[1]:
                    return False
        return True

    #Borra el tree, para nueva visualizacion
    def delete_tree(self):
        records = self.tree.get_children()  # obtiene todos los datos de la tabla tree
        for element in records:
            self.tree.delete(element)  # limpia todos los datos de tree

    #agrega un registro en la base de datos
    def add_record(self, *args):
        if self.valid_add():
            arg_query = ''
            parameters = []
            for index in range(len(self.entrys_array)):
                arg_query += ', ?'
                parameters.insert(index, self.entrys_array[index].get())
            query = f'INSERT INTO {self.sheet_name} VALUES(NULL{arg_query})'#null es el indice
            self.run_query(query, parameters)
            self.message['text'] = 'El Registro ha sido guardado con éxito'
            fail_add = False
        else:
            messagebox.showwarning('Advertencia', 'El Registro ya existe o se encuentra vacío')
            fail_add = True #evita borrar el tree cuando se agrega algo vacio
        self.database_to_tree(like=fail_add)    #like = True: busca la db con "codigo%"
        return not fail_add #devuelve True si agrego, false si no agrego registro

    #Edita un registro en la base de datos
    def edit_record(self, *args):
        if self.valid_selection() and self.entrys_array[0].get() != '':
            query_text_column = ''
            query_text_item = ''
            parameters = []
            param_anterior = []

            #CREA la query y parametros segun cantidad de entrys de columnas haya
            for index in range(len(self.entrys_array)):
                if index == 0:
                    parameters.insert(index, self.entrys_array[index].get())
                    query_text_column += f'"{self.sheet_columns[index + 1]}" = ?'
                    query_text_item += f'"{self.sheet_columns[index + 1]}" = ?'
                    param_anterior.insert(index, self.tree.item(self.tree.selection())['text'])
                else:
                    param_anterior.insert(index, self.tree.item(self.tree.selection())['values'][index-1])
                    parameters.insert(index, self.entrys_array[index].get())
                    query_text_column += f', "{self.sheet_columns[index+1]}" = ?'
                    query_text_item += f' AND "{self.sheet_columns[index+1]}" = ?'

            query = f'UPDATE {self.sheet_name} SET {query_text_column} WHERE {query_text_item}'
            parameters.extend(param_anterior) #primero estan los datos actuales, y despues los anteriores
            self.run_query(query, parameters)
            self.message['text'] = f'El elemento {self.entrys_array[0].get()} ha sido actualizado'
            self.database_to_tree()
        else:
            messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
            self.database_to_tree(like=True)

    # Borra un registro en la base de datos
    def delete_record(self, *args):
        try:
            self.tree.item(self.tree.selection())['text'][0]  # intenta con los 1ros dos elementos que no esten vacios
        except IndexError as e:
            messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para eliminar')
            return

        #Elimina el registro con mismo código y descripción o misma Lista y elemento
        cod_sel = self.tree.item(self.tree.selection())['text']
        descr_sel = self.tree.item(self.tree.selection())['values'][0]  # pasa la seleccion de lista mas elemento
        query = f'DELETE FROM "{self.sheet_name}" WHERE "{self.sheet_columns[1]}" = ? AND "{self.sheet_columns[2]}" = ?'
        self.run_query(query, (cod_sel, descr_sel))  # pone la coma para que se entienda que es una tupla
        self.message['text'] = f'El registro {cod_sel} {descr_sel} ha sido eliminado'

        #Limpia y lee los datos
        #self.clean_entrys()
        self.database_to_tree()  # actualiza la tabla

    # Carga el registro seleccionado de TREE en los entrys de edicion
    def load_edit_item(self, event):
        try:  # comprobacion de seleccion de elemento a editar
            self.tree.item(self.tree.selection())['values'][0]
        except IndexError as e:
            messagebox.showwarning('Advertencia', 'Debe seleccionar una fila para editar')
            return
        # Carga en el array de entrys los valores de row seleccionados
        for index in range(len(self.sheet_columns)-1):
            if index == 0:
                self.entrys_array[0].delete(0, 50)
                self.entrys_array[0].insert(END, self.tree.item(self.tree.selection())['text'])
            else:
                self.entrys_array[index].delete(0,100)
                self.entrys_array[index].insert(END, self.tree.item(self.tree.selection())['values'][index-1])

# Base de datos especifica para listas
class Listas(Database):
    def __init__(self, db_name='LISTAS.db'):
        #define una ventana nueva para ver las listas
        listas_window = Toplevel()
        listas_window.iconbitmap('LOGOLIE x3.ico')
        listas_window.resizable(False, False)

        #con super inicializa el init del padre como propio
        super().__init__(listas_window, db_name)

        #comandos rápidos
        self.tree.bind('<Double-Button-1>', self.database_to_tree)
        listas_window.bind('<Control-p>', self.print_list)

        #carga los entrys, segun seleccion de maestro si != vacio
        self.load_lista()

    # funcion override de database para listas, solo impide cuando lista y elemento son identicos a algun registro
    def valid_add(self):
        if self.entrys_array[0].get() == '' or self.entrys_array[1].get() == '':
            return False
        query = f'SELECT * from {self.sheet_name}'
        db_rows = self.run_query(query)
        for row in db_rows:
            if self.entrys_array[0].get() == row[1] and self.entrys_array[1].get() == row[2]:
                return False
        return True

    # Genera un XLS con la lista a imprimir
    def print_list(self, *args):

        #genera un path con el nombre del archivo XLS
        path = str(date.today()) + '-' + self.entrys_array[0].get() + '.xls'
        file_print = pd.ExcelWriter(path)

        #lee dataframe desde listas
        dataframe = self.search_database(like=False, col_order=2, order='ASC')
        all_table = dataframe.fetchall()

        #genera el encabezado con la data de maestro de articulos
        all_table.insert(0, ('', ''))    #deja un renglon vacío por estética
        all_table.insert(0, ('', '', f'Fecha de alta: {maestro.entrys_array[5].get()}'))
        all_table.insert(0, ('', '', f'Nombre: {maestro.entrys_array[1].get()}'))
        all_table.insert(0, ('', '', f'Código: {self.entrys_array[0].get()}'))

        # crea el DATAFRAME a imprimir
        dframe = pd.DataFrame(all_table)
        dframe.drop(dframe.columns[[0, 1]], axis=1, inplace=True)    #Elimina columna 0 con los indices
        dframe.to_excel(file_print, 'Hoja1', index=False, header=False) #index quita columna de indice, header nom columnas
        try:
            file_print.save()
            messagebox.showinfo('Mensaje', 'Archivo XLS guardado con éxito')
        except:
            messagebox.showerror('Advertencia', 'No se pudo guardar el archivo')

    # carga una lista segun el codigo en maestro con doble click
    def load_lista(self, *args):
        self.entrys_array[0].delete(0, 50)
        self.entrys_array[0].insert(END, maestro.tree.item(maestro.tree.selection())['text'])
        bool_like = maestro.tree.item(maestro.tree.selection())['text'] == ''#si se abre sin seleccion, carga db completa
        self.database_to_tree(like=bool_like) #LIKE cambia forma de busqueda, a exacto o parecido a...

    #Copia una lista con diferente código
    def copy_list(self, *args):

        # crea el registro en maestro con nuevo código
        if maestro.add_record():
            dataframe = self.search_database(like=False)    #lee la lista de listas con un código exacto
            all_table = dataframe.fetchall()

            #arma argumento de query, segun cantidad de columnas, menos la indice
            arg_query = ''
            for column in self.sheet_columns[1:]:
                arg_query += ', ?'

            # arma nuevo dataframe para ingresar en database listas (reemplaza la columna del codigo)
            parameters = []
            index=0
            for row in all_table:
                tupla = (maestro.entrys_array[0].get(), row[2], row[3], row[4])
                parameters.insert(index, tupla)
                index += 1

            #Ingresa masivos datos en database de listas
            query = f'INSERT INTO {self.sheet_name} VALUES(NULL{arg_query})'  # null es el indice
            self.run_query_many(query, parameters)
            self.entrys_array[0].delete(0,50)

            self.entrys_array[0].insert(END, maestro.entrys_array[0].get())
            self.database_to_tree()
            maestro.message['text'] = "Nueva lista creada con éxito"
        else:
            maestro.message['text'] = "No se copió ningún registro"

#Base de datos especifica para Ordenes de trabajo
class Ordenes_trabajo(Database):
    def __init__(self, db_name):
        ot_window = Toplevel()
        ot_window.iconbitmap('LOGOLIE x3.ico')
        self.db_name = db_name
        ot_window.resizable(False, False)
        Database(ot_window, db_name)

# ------------FUNCIONES---------------------------------------------
#Instancia base de datos de LISTAS
def call_listas(*args):
    global lista
    lista = Listas("LISTAS.db") #instancia una lista

    # si encuentra 'keysym=c' --> ejecuta copiar lista
    if str(args).find('keysym=c', 0, -1) != -1:
        lista.copy_list()

#Instancia base de datos de ordenes de trabajo
def call_ot(*args):
    global ot
    ot = Ordenes_trabajo("OT.db")

#Info de ayuda
def help_info():
    messagebox.showinfo('Ayuda GEST2020', 'Solicitar manual a su proveedor')

#Detalle de los HOTKEYS del programa
def hotkeys():
    messagebox.showinfo('Accesos rápidos', 'Enter: filtrar búsqueda por código\n'
                                           'Ctrl+Enter: limpiar búsqueda\n'
                                           'SpaceBar: Desplazar una hoja\n\n'
                                           'Ctrl+C: Copiar Lista\n'
                                           'Ctrl+A: Agrega registro\n'
                                           'Ctrl+E: Editar registro\n'
                                           'Ctrl+D: Eliminar registro\n\n'
                                           'Ctrl+L: Abrir Listas\n'
                                           'Ctrl+O: Abrir Órdenes de trabajo\n\n'
                                           'Ctrl+P: Imprimir\n'
                                           'ALT+F4: Cerrar programa\n')

#Info de licencia
def license():  # funcion para ventana emergente que muestra un warning con icono warning
    messagebox.showinfo('GEST2020', 'licencia válida')

#Info del programa y version
def help_about():  # funcion para vent emergente que muestra info con icono de info
    messagebox.showinfo('Gestor de empresa', 'GEST2020 Versión: V2\nProgramado por Agustin Arnaiz')

#------------MAIN-BARRAMENU-LOOP-INSTANCIAS  DATABASES-----------------
if __name__ == '__main__':

    #ROOT, ventana principal
    root = tkinter.Tk()
    root.title('GEST2020 | Gestor y administrador de empresa')  # titulo de la ventana
    root.resizable(True, True)  # redimencionar ancho? y alto?
    root.iconbitmap('LOGOLIE x3.ico')  # icono e la ventana
    root.config(bg="grey")  # config de la ventana bg= back ground color

    #Define la BARRA de menu
    barra_menu = tkinter.Menu(root)
    root.config(menu=barra_menu, width=400, height=200)
    root.resizable(False, False)

    #menu de ARCHIVO
    file_menu = tkinter.Menu(barra_menu, tearoff=0)
    file_menu.add_command(label="Convertir Planilla a base de datos", command= Convert_sheet)
    file_menu.add_separator()
    file_menu.add_command(label="Abrir Maestro de artículos", command=lambda: Database(root, 'MAESTRO.db'))
    file_menu.add_command(label="Abrir Listas (ctrl+l)", command=call_listas)
    file_menu.add_command(label="Órdenes de trabajo (ctrl+o)", command=call_ot)
    file_menu.add_separator()
    file_menu.add_command(label="Cerrar programa (alt+F4)", command=lambda: root.destroy())

    #MENU edicion
    edit_menu = tkinter.Menu(barra_menu, tearoff=0)
    edit_menu.add_command(label="Copiar Lista (ctrl+c)", command=lambda: call_listas('keysym=c'))
    edit_menu.add_separator()
    edit_menu.add_command(label="Agregar Registro (ctrl+a)", command=lambda: maestro.add_record())
    edit_menu.add_command(label="Editar Registro (ctrl+e)", command=lambda: maestro.edit_record())
    edit_menu.add_command(label="Eliminar Registro (ctrl+d)", command=lambda: maestro.delete_record())

    #Menu de ayuda
    help_menu = tkinter.Menu(barra_menu, tearoff=0)
    help_menu.add_command(label="Ayuda", command=help_info)
    help_menu.add_command(label="Licencia", command=license)
    help_menu.add_command(label="Hotkeys", command=hotkeys)
    help_menu.add_separator()
    help_menu.add_command(label="Acerca de", command=help_about)

    #items de la barra de menu
    barra_menu.add_cascade(label="Archivo", menu=file_menu)
    barra_menu.add_cascade(label="Edición", menu=edit_menu)
    barra_menu.add_cascade(label="Ayuda", menu=help_menu)
    # ------------BARRA DE MENUS-------------------------------

    #Instancia a Maestro de articulos (database)
    try:
        maestro = Database(root, 'Maestro.db')
    except:
        Convert_sheet(root) #si no existe el MAESTRO.db, pide convertir un xls del mismo nombre

    #Instancia Listas con CTRL+L
    root.bind('<Control-l>', call_listas)
    root.bind('<Control-c>', call_listas)

    # Instancia OT con CTRL+O
    root.bind('<Control-o>', call_ot)

    root.mainloop()
