''' GEST 2020: Gestor de compras y materiales para pequeñas empresas basado en software GEST;
Consta de una base de datos con 4 tablas: MAESTRO, LISTAS, OT, RUBROS. Su fin es conseguir una orden de compra de materiales
unificada como "orden de trabajo", ordenada por PLAN, RUBRO, LISTA
diseñado para la empresa LIE SRL. Por Agustin Arnaiz
* Clase UpperEntry tomada y modificada de stackoverflow (usuario FJSevilla): https://es.stackoverflow.com/questions/356082/como-lograr-poner-mayúscula-en-los-campos-de-tipo-entry-y-formatear-números-en-p'''

import sqlite3                  # SQL, manejo base de datos
import pandas as pd             # Dataframes y XLS files
import numpy as np				# se usa para listas de datos, acceso no consecutivo a sus items
import os                       # usa la impresión de shell "print"
import inspect
import tkinter as tk
from tkinter import *
from tkinter import messagebox  # mensajes de salida
from tkinter import ttk         # tkinter mas facha
from ttkthemes import ThemedTk	# themes :D
from datetime import date       # fecha
from datetime import datetime	# fecha con hora, now
from win32api import GetSystemMetrics	# lee  resolucion de la pantalla


#-----------------------------CLASES--------------------------------

# Configuraciones, usuarios, etc
class ProgramManager:

	def __init__(self, *args):
		# Configuraciones
		self.configs = {'db_name': 'db_gest2020.db', 
						'theme_name': 'clearlooks',
						'path_config': 'config.cfg'}
		self.theme_names = {1: 'clearlooks', 
							2: 'elegance', 
							3: 'plastik', 
							4: 'radiance', 
							5: 'breeze', 
							6: 'equilux', 
							7: 'black', 
							8: 'blue'}

	#Carga la configuración desde archivo, o la crea por default
	def load_config(self):
	
		# Lee configuraciones de archivo
		lista = []
		if os.path.isfile(self.configs['path_config']):
			config_file = open(self.configs['path_config'], 'r')
			for line in config_file:
				equal = line.find('=', 0, -1)
				lista.append(line[equal+1:-1])
			index = 0
			for key in self.configs:
				self.configs[key] = lista[index]
				index += 1
			config_file.close()
		# si no existe el archivo, lo crea con valores por defecto
		else:
			self.save_config()

	# Guarda la configuración en archivo
	def save_config(self):
		for line in self.configs:
			config_file = open(self.configs['path_config'], 'w')
			for key in self.configs:
				config_file.write(key+'=')
				config_file.write(f'{self.configs.get(key)}\n')
		config_file.close()
		

# Formato ventana por defecto
class WindowConfig:
	def __init__(self, parent, *args, **kwargs):
		self.window = parent
		self.window.config(bg="grey", width=800, height=200)  # config de la ventana bg= back ground color
		self.window.iconbitmap('_LOGOLIE x3.ico')  # icono e la ventana
		self.window.rowconfigure(0, weight=1)
		self.window.columnconfigure(0, weight=1)	
		self.window.resizable(True, True)
		self.window.geometry('+50+5')

		self.bar_menu = tk.Menu(self.window)
		self.window.config(menu=self.bar_menu)
		self.file_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.edit_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.tools_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.help_menu = tk.Menu(self.bar_menu, tearoff=0)
		
		self.name_window(**kwargs)
		self.create_menu()
		self.bindings()

	def name_window(self, *args, **kwargs):
		title = kwargs.get('title')
		if title != None:
			self.window.title(title)

	def create_menu(self):
		# MENU FILE
		self.file_menu.add_command(label="Abrir Listas (ctrl+L)", command=Listas)
		self.file_menu.add_command(label="Abrir Órdenes de trabajo (ctrl+O)", command=Ordenes_trabajo)
		self.file_menu.add_command(label="Ver Rubros (ctrl+R)", command=Rubros)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Imprimir (ctrl+p)", command=ToPrinter)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Cerrar programa (alt+F4)", command=root.destroy)

		# MENU edicion
		self.edit_menu.add_command(label="Agregar Registro (ctrl+A)", command= lambda: maestro.add_record())
		self.edit_menu.add_command(label="Editar Registro (ctrl+E)", command=lambda: maestro.edit_record())
		self.edit_menu.add_command(label="Eliminar Registro (ctrl+D)", command=lambda: maestro.delete_record())

		# MENU Herramientas
		self.tools_menu.add_command(label="Cambiar apariencia (ctrl+T)", command=lambda: WindowTheme(win_to_mod=root).select_theme())
		self.tools_menu.add_separator()
		self.tools_menu.add_command(label="Crear base de datos GEST2020", command=CreateDatabase)

		# Menu ayuda
		self.help_menu.add_command(label="Ayuda", command=help_info)
		self.help_menu.add_command(label="Licencia", command=license)
		self.help_menu.add_command(label="Comandos rápidos", command=hotkeys)
		self.help_menu.add_separator()
		self.help_menu.add_command(label="Acerca de GEST2020", command=help_about)

		# items de la barra de menu
		self.bar_menu.add_cascade(label="Archivo", menu=self.file_menu)
		self.bar_menu.add_cascade(label="Edición", menu=self.edit_menu)
		self.bar_menu.add_cascade(label="Herramientas", menu=self.tools_menu)
		self.bar_menu.add_cascade(label="Ayuda", menu=self.help_menu)

	# Define los bindings y comandos rápidos para la ventana
	def bindings(self):
		# EVENTOS de la ventana y comandos rápidos
		self.window.bind('<Control-t>', lambda e: WindowTheme(win_to_mod=self.window).select_theme())
		self.window.bind('<Control-T>', lambda e: WindowTheme(win_to_mod=self.window).select_theme())
		self.window.bind('<Control-l>', Listas)
		self.window.bind('<Control-L>', Listas)
		self.window.bind('<Control-r>', Rubros)
		self.window.bind('<Control-R>', Rubros)
		self.window.bind('<Control-o>', Ordenes_trabajo)
		self.window.bind('<Control-O>', Ordenes_trabajo)
		self.window.bind('<Control-b>', CreateDatabase)
		self.window.bind('<Control-B>', CreateDatabase)


#Seleccion de THEME para ventana principal
class WindowTheme:
	def __init__(self, *args, win_to_mod):
		self.window = Toplevel()
		self.window.iconbitmap('_LOGOLIE x3.ico')
		self.window.resizable(False, False)
		self.win_to_mod = win_to_mod
		self.window.config(bg='black')

	def select_theme(self):	
		self.theme_sel = IntVar()
		# Crea la seleccion de THEMES
		for index in manager.theme_names:	 
			button = tk.Radiobutton(self.window, 
				text=manager.theme_names[index],
				variable=self.theme_sel, 
				indicator=0, 
				value=index,
				cursor='hand2', 
				background = "light blue",
				command=self.change_theme)
			button.grid(sticky=W+E, columnspan=1, pady=5, padx=20)
			
			if manager.theme_names[index] == manager.configs['theme_name']:
				button.select()

	def change_theme(self):
		# Carga el theme a la ventana
		name_theme = manager.theme_names[self.theme_sel.get()]
		self.win_to_mod.config(theme=name_theme)

		# Carga el theme al manager
		manager.configs['theme_name'] = manager.theme_names[self.theme_sel.get()]
		manager.save_config()


#Pasa automaticamente a MAYUSCULAS textos de los entrys
class UpperEntry(tk.Entry):
	def __init__(self, frame, *args, **kwargs):
		self.text_to_upper = tk.StringVar(frame)
		super().__init__(frame, *args, **kwargs)
		
		self.configure(textvariable=self.text_to_upper)
		self._to_upper()

	def configure(self, cnf=None, **kwargs):
		text_var = kwargs.get("textvariable")
		if text_var is not None:
			# De haber cambio en el texto, llama a _to_upper para pasarlo a mayuscula
			text_var.trace_add('write', self._to_upper)
			self.text_to_upper = text_var

		# Crea el entry original
		super().config(cnf, **kwargs)
		
	# Pasa el texto a mayúscula
	def _to_upper(self, *args):
		self.text_to_upper.set(self.text_to_upper.get().upper())


# Recibe DF de cualquier ventana e imprime a XLS y via shell envia a la default printer
class ToPrinter:
	def __init__(self, *args, to_print, from_obj):
		self.to_print = pd.DataFrame(to_print)
		# genera un path con el nombre del archivo XLS
		self.path = str(date.today()) + '-' + from_obj + '.xlsx'
		self.file_print = pd.ExcelWriter(self.path, engine='xlsxwriter')  # cambia el motor de pd.ExcelWriter para modificar las col
		self.final_print()

	# Envía a printer y file xls
	def final_print(self):
		self.to_print.to_excel(self.file_print, 'Hoja1',
							 index=False, header=False)  # index quita columna de indice, header nom columnas
		# calcula un ancho dinamico para cada columna
		worksheet = self.file_print.sheets['Hoja1']  # genera objeto worksheet para hacer uso de set_column

		for column in list(self.to_print):
			max_len = 4  # tamaño mínimo
			for row in range(8, len(self.to_print[:100])):
				item = self.to_print.loc[row, column]
				max_len = max(max_len, len(str(item)))
				worksheet.set_column(column, column, max_len+1)
		try:
			self.file_print.save()
			#os.startfile(path, "print")     #envia el archivo a impresora default
			messagebox.showinfo('Mensaje', 'Archivo XLS guardado con éxito')
		except Exception as err:
			messagebox.showerror('Advertencia', f'No se pudo guardar el archivo:\n{err}')


#Crea la base de datos SQLite3 desde los 3 XLS del GEST original
class CreateDatabase:
	def __init__(self, *args):
		self.windows = Toplevel
		self.conn = sqlite3.connect(manager.configs.get('db_name'))
		try:
			self.maestro = self.format_table('MAESTRO')
			self.listas = self.format_table('LISTAS')
			self.ot = self.format_table('OT')
			self.rubros = self.format_table('RUBROS')
			self.conn.commit()
			self.conn.close()
			messagebox.showinfo('Info', 'Base de datos creada con éxito\nDebe reiniciar el programa.')

		except Exception as err:
			messagebox.showerror('Error', f'No se pudo crear la base de datos: \n{err}')
		self.windows.destroy

	#Abre xls y genera la query para la tabla de db
	def format_table(self, name):
		try:
			xls_file = pd.read_excel(name+'.xls')
		except Exception as err:
			messagebox.showwarning('Advertencia', f'No se encuentra el archivo {name}.xls\n{err}' )
			return

		query = ''
		drop_columns = []
		nom_columns = []
		# CREA LA TABLA, ELIMINA COLUMNAS NO DESEADAS Y CAMBIA EL NOMBRE DE LAS EXISTENTES
		if name == 'MAESTRO':
			# Primero elimina decimales no deseados de columna d Precios
			xls_file = self.delete_decimal(xls_file, 'PRECIO1,N,10,3')

			# Lista para eliminar ciertas columnas no usadas por gest2020
			drop_columns = [6, 7, 9]  # cableado, stkmin, comprasug

			# creacion de tabla
			nom_columns = ["CÓDIGO", "DESCRIPCIÓN", "UN.", "PRECIO", "FECHA_PRECIO", "FECHA_ALTA", "RUBRO"]
			query = 'CREATE TABLE "MAESTRO" ' \
					'("CÓDIGO" TEXT NOT NULL, ' \
					'"DESCRIPCIÓN" TEXT, ' \
					'"UN."	TEXT DEFALUT "UN", ' \
					'"PRECIO"	REAL, ' \
					'"FECHA_PRECIO"	TIMESTAMP DEFAULT CURRENT_TIMESTAMP, ' \
					'"FECHA_ALTA"	TIMESTAMP DEFAULT CURRENT_TIMESTAMP, ' \
					'"RUBRO"	TEXT DEFAULT "VS.", ' \
					'FOREIGN KEY("RUBRO") REFERENCES "RUBROS"("RUBRO") ON UPDATE CASCADE ON DELETE SET NULL,' \
					'PRIMARY KEY("CÓDIGO"))'
					

		if name == "LISTAS":
			# Primero elimina decimales no deseados de columna cantidad
			xls_file = self.delete_decimal(xls_file, 'CANT,N,10,3')

			#lista para renombrar las columnas
			nom_columns = ["LISTA", "CÓDIGO", "CANT.", "UN."]

			#creacion de la tabla
			query = f'CREATE TABLE "LISTAS" ' \
					f'("{nom_columns[0]}"	TEXT NOT NULL,' \
					f'"{nom_columns[1]}"	TEXT,' \
					f'"{nom_columns[2]}"	REAL,' \
					f'"{nom_columns[3]}"	TEXT,' \
					f'FOREIGN KEY("LISTA") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE,' \
					f'FOREIGN KEY("CÓDIGO") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE)'

		if name == "OT":
			# Primero elimina decimales no deseados de columna precio unitario y monto
			xls_file = self.delete_decimal(xls_file, 'PRUNIT,N,10,3')
			xls_file = self.delete_decimal(xls_file, 'MONTO,N,10,3')

			drop_columns = [3, 4, 5, 6, 7, 8, 11, 12] #fecha_venta, fecha_cum, can_cum, nombre_ fecha_ing, nom_ant, plan, despi
			nom_columns = ["PLAN", "CÓDIGO", "CANT.", "PRECIO", "MONTO"]
			query = 'CREATE TABLE "OT" ' \
					'("PLAN"	INTEGER, ' \
					'"CÓDIGO"	TEXT NOT NULL, ' \
					'"CANT."	INTEGER, ' \
					'"PRECIO"	REAL, ' \
					'"MONTO"	REAL, ' \
					'FOREIGN KEY ("CÓDIGO") REFERENCES "MAESTRO"("CÓDIGO"))'

		if name == "RUBROS":
			nom_columns = ["RUBRO"]
			query = 'CREATE TABLE "RUBROS" ' \
					'("RUBRO"	TEXT NOT NULL,' \
					'PRIMARY KEY("RUBRO"))'

		# ELIMINA COLUMNAS No deseadas
		xls_file.drop(xls_file.columns[drop_columns], axis=1, inplace=True)
		# CAMBIA NOMBRE DE LAS COLUMNAS
		xls_file.columns = nom_columns

		self.copy_xls_db(xls_file, name, query)

	# Elimina decimales no significativos
	def delete_decimal(self, xls_file, column):
		for item in range(len(xls_file[column])):
			temp = round(xls_file[column][item], 3)
			xls_file[column][item] = temp

		return xls_file

	# crea la tabla dentro de la base de datos
	def copy_xls_db(self, xls_file, name, query):
		cursor = self.conn.cursor()
		print('copiando ', name, '...')
		try:
			cursor.execute(query)
			#copia los datos dentro de esa misma tabla
			xls_file.to_sql(name=name, con=self.conn, if_exists='append', index=False) #if_exist='replace' resulta en falla PK FK etc
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'Ocurrió un error al querer agregar la tabla {name} a la base de datos.\n{err}')


#Clase principal de base de datos (auto arma panel, manejo CRUD)
class Database:

	def __init__(self, window, table_name):
		# Definicion de atributos
		self.window = window
		#WindowConfig(self.window)
		self.table_name = table_name
		self.entry_array = []  # array de entrys editores de registro de database
		self.no_edit_entry = []  # Define los entry bloqueados para edicion
		self.focus_decoder = {}  # carga una lista de widgets entry para realizar busquedas

		# define FRAMEWORK tabla de datos
		self.frame_tree = LabelFrame(self.window, text=f'Editor de {self.table_name}')
		self.frame_tree.grid(row=0, column=0, columnspan=20, pady=10, padx=10, sticky=W+E+S+N)
		self.frame_tree.config(cursor='hand2')  #indica seleccion de los elementos del tree
		# define FRAMEWORK mensajes
		self.frame_msg = LabelFrame(self.window, text='')
		self.frame_msg.grid(row=10, column=0, columnspan=20, pady=10, padx=10, sticky=W+E+S+N)
		# MENSAJE de salida en la ventana
		self.message = Label(self.window, text='', fg='blue')
		self.message.grid(row=10, column=0, columnspan=20, sticky=W + E)
		
		# LEE nombres de las columnas de la tabla
		self.sheet_columns = np.array([]) #usa numpy para citar items no consecutivos
		self.read_columns()
		# TREE principal
		self.tree = ttk.Treeview(self.frame_tree, height=25)

		# permite expandir los widgets internos a la ventana en x e y
		self.frame_tree.rowconfigure(10, weight=1, minsize=50)
		for index in range(len(self.sheet_columns)):
			self.frame_tree.columnconfigure(index, weight=1, minsize=50)

		# Moldea ventana y lee base de datos
		self.build_main_view()
		self.show_data()

		# Hace foco en buscar registro por código al inicio
		self.entry_array[0].focus()

		# Bindings de DATABASE
		self.window.bind('<Return>', self.show_data)  # busca según código en 'self.entrys_array[0]'
		self.window.bind('<Control-Return>', self.clean_entrys)  # borra elementos de busqueda
		self.window.bind('<Control-a>', self.add_record)
		self.window.bind('<Control-A>', self.add_record)
		self.window.bind('<Control-e>', self.edit_record)
		self.window.bind('<Control-E>', self.edit_record)
		self.window.bind('<Control-d>', self.delete_record)
		self.window.bind('<Control-D>', self.delete_record)
		self.window.bind('<Control-p>', self.prepare_to_print)
		self.window.bind('<Control-P>', self.prepare_to_print)
		self.tree.bind('<space>', self.auto_scroll)  # baja de a una hoja la vista de tree
		self.tree.bind('<<TreeviewSelect>>', self.load_edit_item)  # carga datos seleccionados

	#Armado de ventana, TREEVIEW adaptable segun Database
	def build_main_view(self):

		self.tree.config(columns=len(self.sheet_columns))
		self.tree.grid(row=10, column=0, columnspan=20, rowspan=1, pady=10, sticky=N+S+W+E)

		# Scroll vertical del TREE
		self.scroll_tree_v = Scrollbar(self.frame_tree, command=self.tree.yview)
		self.scroll_tree_v.grid(row=10, column=len(self.sheet_columns), sticky=NS)
		self.tree.config(yscrollcommand=self.scroll_tree_v.set)

		# Scroll horizontal del TREE
		self.scroll_tree_h = Scrollbar(self.frame_tree, orient='horizontal', command=self.tree.xview)
		self.scroll_tree_h.grid(row=12, column=0, columnspan=20, sticky=W+E)
		self.tree.config(xscrollcommand=self.scroll_tree_h.set)

		# NOMBRES COLUMNAS
		nombres_columnas = []
		for each_column in self.sheet_columns:
			nombres_columnas.append(each_column)
		self.tree["columns"] = nombres_columnas[1:] #desde 2 xq el 1ro es index y el 2do esta en text (no en values)
		print(self.tree["columns"])


		
		#Esta query es para dar TAMAÑO DINAMICO al ancho de columnas del tree con 400 datos por columna
		if self.table_name == 'LISTAS':
			query = f'SELECT {self.table_name}.*, ' \
					f'{maestro.table_name}.{maestro.sheet_columns[1]},{maestro.table_name}.{maestro.sheet_columns[3]} ' \
					f'FROM {self.table_name} ' \
					f'INNER JOIN {maestro.table_name} LIMIT 1000'
		else:
			query = f'SELECT * from "{self.table_name}" LIMIT 1000'
		cursor = self.run_query(query)

		#se cargan titulos y tamaños de las columnas del tree, asi como Entrys para la edicion de records
		num_column = 0
		item_table = ''
		all_table = cursor.fetchall()

		for each_column in self.sheet_columns:
			largest = 6	# tamaño minimo
			for row_table in all_table:
				item_table = str(row_table[num_column])
				if largest < len(item_table):
					largest = len(item_table)
			if largest > 12:
				largest = 6 + int(largest / 2)
			
			self.tree.column("#" + str(num_column), width=largest*11, stretch=True)
			self.tree.heading('#' + str(num_column), text=f'{nombres_columnas[num_column]}', anchor=CENTER)

			# crea tantos ENTRYS como columnas, para editar registro
			self.entry_array.append(self.entrys(frame=self.frame_tree,
												name=each_column, row=7, column=num_column, width=largest*2))
			num_column += 1

	#Crea Entrys con label superior
	def entrys(self, frame, name='entry', row=0, column=0, width=50):

		# define label y entrys segun llamada
		Label(frame, text=name).grid(row=row, column=column)

		#permite ingresar con minusculas la descripcion
		if name == 'DESCRIPCIÓN':
			entry = Entry(frame, width=width)
			entry.config(fg="blue")

		#menu fijo para Rubro ComboBox de Maestro solamente
		elif name == 'RUBRO' and self.table_name == 'Maestro':
			#lee de la tabla rubros todos los items y los agrega como una lista
			query = f'select * from "RUBROS" ORDER BY "RUBRO" ASC'
			rubros = self.run_query(query)
			lista_rubros = []
			for item in rubros.fetchall():
				lista_rubros.append(item)
			entry = ttk.Combobox(frame, width=width, state='readonly', values=lista_rubros)

		#todos los demas son en mayúscula
		else:
			entry = UpperEntry(frame, width=width)
			entry.config(fg="blue")

		entry.grid(row=row + 2, column=column, columnspan=1, sticky=W+E, padx='5')

		#define decodificacion de focus, para hacer busqueda por columnas segun el focus de entry
		self.focus_decoder[str(entry)] = column

		return entry

	#Borra los entrys de registro
	def clean_entrys(self, *args):
		# Habiita todos los entrys despues de borrarlos, permite buscar por cada campo
		self.hab_entry(self.no_edit_entry)
		for entry_element in self.entry_array:
			entry_element.delete(0, 'end')
		
		# Hace foco en buscar registro por código
		self.entry_array[0].focus()

	#Lee los nombres de las columnas de la database
	def read_columns(self):
		query = f'SELECT * FROM "{self.table_name}" LIMIT 1'
		columns = self.run_query(query)
		for columna in columns.description:
			self.sheet_columns = np.append(self.sheet_columns, columna[0])

	#Ejecuta una QUERY SQLite3 con cursor usando parametros
	def run_query(self, query, parameters={}):
		with sqlite3.connect(manager.configs.get('db_name')) as conn:
			conn.execute('PRAGMA foreign_keys = True')  # habilita las constraints de las FK, por defecto = False
			cursor = conn.cursor()
			try:
				result = cursor.execute(query, parameters)
				conn.commit()
				return result
			except Exception as err:
				messagebox.showerror('ERROR', f'Ocurrió un error en la base de datos: \n{err}')

	#Copia datos masivos pasados por parametro (la usa listas y OT)
	def run_query_many(self, query, parameters={}):
		with sqlite3.connect(manager.configs.get('db_name')) as conn:
			#conn.execute('PRAGMA foreign_keys = True')  # habilita las constraints de las FK, por defecto = False
			cursor = conn.cursor()
			result = cursor.executemany(query, parameters)
			conn.commit()
		return result

	#Limpia el tree, y lo re-hace segun codido exacto o simil (like) y ordena segun column
	def show_data(self, *args, like='%', col_search=0, col_order=0, limit='', open=0):

		#Si se llama a la funcion con ENTER, hace busqueda de columna segun foco de entry
		if str(args).find('keysym=Return', 0, -1) != -1:     #(texto a buscar, inicio, final)
			#se fija donde esta el foco de entry, para hacer la busqueda segun esa columna
			focus = self.window.focus_get()
			
			if focus is not None:
				if open == 0:
					# Para OT: abre o cierra el tree en la busqueda
					if str(focus) == '.!toplevel.!labelframe.!upperentry2':
						open=False
					else:
						#open true hace que se expanda la lista de OT cuando viene de una busqueda puntual
						#open false para cuando esa busqueda es por código (no por ORDEN)
						open = bool(self.entry_array[0].get()) or bool(self.entry_array[1].get())

				try:
					col_order = col_search = self.focus_decoder[str(focus)]
				except:
					col_order = col_search = 0
		
		# Habilita los no edit entry para cargar los datos y limpia el tree previamente
		self.hab_entry(self.no_edit_entry)
		self.delete_tree()

		#BUSCA EN LA DB y copia en el TREE
		cursor = self.query_search(*args, 
									search=self.entry_array[col_search].get(), 
									like=like, 
									col_search=col_search, 
									col_order=col_order, 
									limit=limit)
		self.data_into_tree(cursor, open)

		# Deshabilita los no edit entrys despues de cargar los datos
		self.deshab_entry(self.no_edit_entry)

	# Ingresa los datos de busqueda en el TREE
	def data_into_tree(self, cursor, open=0):
		for row in cursor:
			line = self.tree.insert('', 0, text=row[0], values=row[1:], open=open)
		
	#Ejecuta query de busqueda en DB, retorna dataframe, se programó separado de "database_to_tree" para ser llamada por separado (por print_list y copy_list)
	def query_search(self, *args, search='%', like='%', col_search=0, col_order=0, order='DESC', limit=''):
		query = f'SELECT * from "{self.table_name}" ' \
				f'WHERE "{self.sheet_columns[col_search]}" ' \
				f'LIKE "{search}{like}" ' \
				f'ORDER BY "{self.sheet_columns[col_order]}" {order} {limit}'

		return self.run_query(query)

	#Valida operacion si hay seleccion de registro en TREE
	def valid_selection(self):  # aprueba la escritura en DB
		seleccion = self.tree.item(self.tree.selection())['text']
		return seleccion != ''

	#Valida operación para agregar registro si el mismo no es vacío o repetido
	def valid_add(self):
		if self.entry_array[0].get() == '':
			return False
		query = f'SELECT "{self.sheet_columns[0]}" from "{self.table_name}"'
		db_rows = self.run_query(query)
		for row in db_rows:
			if self.entry_array[0].get() == row[0]:
				return False
		return True

	#Borra el tree, para nueva visualizacion
	def delete_tree(self):
		records = self.tree.get_children()  # obtiene todos los datos de la tabla tree
		for element in records:
			self.tree.delete(element)  # limpia todos los datos de tree

	# forma la query para agregar item a tabla
	def add_query(self):
		nom_column, arg_query, parameters = [[],[],[]]

		# Solo agrega si no se encuentra en no_edit_entry o si el entry no está vacío
		for index, column in enumerate(self.sheet_columns):
			print(type(self.entry_array[index]))
			if index not in self.no_edit_entry:
				if self.entry_array[index].get() != '':
					nom_column.insert(index, '"'+column+'"')
					arg_query.insert(index, '?')
					parameters.insert(index, self.entry_array[index].get())
			else:
				print(column)
		return nom_column, arg_query, parameters

	# agrega un registro en la base de datos
	def add_record(self, *args):
		if self.valid_add():
			nom_column, arg_query, parameters = self.add_query()
			
			print(nom_column, arg_query, parameters)
			
			query = f'INSERT INTO {self.table_name} ({", ".join(nom_column)}) ' \
				f'VALUES({" ,".join(arg_query)})' #une la lista con join

			print(query)

			self.run_query(query, parameters)
			self.message['text'] = f'{maestro.entry_array[0].get()} ha sido guardado con éxito'
			fail_add = ''
		else:
			messagebox.showwarning('Advertencia', 'El Registro ya existe o se encuentra vacío')
			fail_add = '%'  # evita borrar el tree cuando se agrega algo vacio al buscar con LIKE

		self.show_data(like=fail_add, open=True)  # like = %: busca la db con "codigo%"
		return not fail_add  # devuelve True si agrego, false si no agrego registro

	#Edita un registro en la base de datos
	def edit_record(self, *args):

		# se habilitan los entrys, ya que se modifican a veces antes de editar un registro
		self.hab_entry(self.no_edit_entry)
		
		if self.valid_selection() and self.entry_array[0].get() != '':

			query_text_column, query_text_item, parameters, param_anterior = [[],[],[],[]]

			# Edita los las columnas que no pertenecen a no_edit_entry
			for index, column in enumerate(self.sheet_columns):
				if index not in self.no_edit_entry:
					#para editar un item, lo comprueba el "cod y descr", xq algunos precios daban error
					if index == 0:
						param_anterior.insert(index, self.tree.item(self.tree.selection())['text'])
					else:
						param_anterior.insert(index, self.tree.item(self.tree.selection())['values'][index - 1])

					query_text_column.insert(index, f'"{column}" = ? ')
					query_text_item.insert(index, f'"{column}" = ? ')
					parameters.insert(index, self.entry_array[index].get())
				'''
				# para el caso de maestro, en la columna de precio, compara si el mismo cambió
				if self.sheet_columns[index] == 'PRECIO' \
						and self.entry_array[index].get() \
						!= self.tree.item(self.tree.selection())['values'][index - 1]:
					#si cambió, actualiza la fecha de precio que esta en index mas 1 (le sigue a precio)
					self.entry_array[index+1].delete(0, 50)
					self.entry_array[index+1].insert(END, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))'''

			
			query = f'UPDATE {self.table_name} SET {", ".join(query_text_column)} WHERE {"AND ".join(query_text_item)}'

			parameters.extend(param_anterior) #primero estan los datos actuales, y despues los anteriores
			self.run_query(query, parameters)
			self.message['text'] = f'El elemento {self.entry_array[0].get()} ha sido actualizado'
			self.show_data(like='', open=True)
		else:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
			self.show_data(like='%')

	# Borra un registro en la base de datos
	def delete_record(self, *args):
		selection = self.tree.selection()
		for id in selection:
			item = self.tree.item(id)
			
			# Elimina el registro con mismo código y descripción o misma Lista y elemento
			cod_sel = self.tree.item(id)['text']
			try:
				descr_sel = self.tree.item(id)['values'][0]  # pasa la seleccion de lista mas elemento
				query = f'DELETE FROM "{self.table_name}" ' \
						f'WHERE "{self.sheet_columns[0]}" = ? ' \
						f'AND "{self.sheet_columns[1]}" = ?'

			#este artificio funciona para rubros que solo tiene una columna ['text'] y no la columna ['values']
			except:
				descr_sel = 'None'

			#para los escasos registros donde la descr es NULL, esto permite borrar esa entrada
			if descr_sel == 'None':
				descr_sel = ''
				query = f'DELETE FROM "{self.table_name}" ' \
						f'WHERE "{self.sheet_columns[0]}" = ? '
				self.run_query(query, (cod_sel,))  # pone la coma para que se entienda que es una tupla

			else:
				self.run_query(query, (cod_sel, descr_sel))
			self.message['text'] = f'El registro {cod_sel} {descr_sel} ha sido eliminado'
		self.show_data(open=True)  # actualiza la tabla

	# Carga el registro seleccionado de TREE en los entrys de edicion
	def load_edit_item(self, *args):
		self.hab_entry(self.no_edit_entry)
		# Carga en el array de entrys los valores de row seleccionados
		
		for index in range(len(self.entry_array)):
			if index == 0:
				self.entry_array[0].delete(0, 50)
				self.entry_array[0].insert(END, self.tree.item(self.tree.selection())['text'])

			#maneja por separado el combobox de rubro, se setea diferente
			elif self.sheet_columns[index] == 'RUBRO':
				self.entry_array[index].set(self.tree.item(self.tree.selection())['values'][index - 1])

			else:
				self.entry_array[index].delete(0, 100)
				self.entry_array[index].insert(END, self.tree.item(self.tree.selection())['values'][index - 1])
	
		self.deshab_entry(self.no_edit_entry)

	# contruye dataframe a partir de la vista del tree
	def prepare_to_print(self, *args):
		dframe = self.build_print()

		#manda a imprimir
		ToPrinter(to_print=dframe, from_obj=self.entry_array[0].get() or self.table_name)

	# genera el encabezado genérico de impresión
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			item = self.tree.item(self.tree.selection())['text']
			descr_item = self.tree.item(self.tree.selection())['values']
			descr_item.insert(0, item)
			list_row.append(descr_item)

		# incorpora el nombre de las columnas
		list_row.insert(0, (self.sheet_columns))  # deja de lado la columna de código de lista
		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Impreso desde {self.table_name}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row

	#deshabilita varios entrys para evitar su edicion
	def deshab_entry(self, deshab_entry=[]):
		for item in deshab_entry:
			self.entry_array[item].config(state='readonly')

	#habilita los entry para el ingreso de datos
	def hab_entry(self, hab_entry=[]):
		for item in hab_entry:
			self.entry_array[item].config(state='normal')

	#desplaza de a una hoja en la vista de tree con la barra 'space'
	def auto_scroll(self, *args):
		self.tree.yview_scroll(1, what='page')


#Base de datos especifica para Maestro
class Maestro(Database):

	def __init__(self, window, table_name='MAESTRO'):
		# Suma acciones al menú
		main_window.edit_menu.add_separator()
		main_window.edit_menu.add_command(label="Copiar Lista (ctrl+C)", command=lambda: Listas('keysym=c'))
		main_window.edit_menu.add_command(label="Eliminar Lista (ctrl+F)", command=lambda: Listas('keysym=f'))
		main_window.edit_menu.add_separator()
		main_window.edit_menu.add_command(label="Buscar Item (ENTER)", command=lambda: self.show_data())

		super().__init__(window, table_name)

		# Orden de entrys no disponibles para editar "readonly"
		self.no_edit_entry = [4, 5]

		# solo desde MAESTRO abre una lista con doble click
		self.tree.bind('<Double-Button-1>', Listas) 
		window.bind('<Control-c>', Listas)  # copia lista con nuevo codigo
		window.bind('<Control-C>', Listas)  # copia lista con nuevo codigo

	# override para eliminar una columna de fecha de alta
	def build_print(self):
		list = super().build_print()
		df = pd.DataFrame(list)
		df.drop(df.columns[5], axis=1, inplace=True)
		return df


#Base de datos especifica para listas (asincorpora crud listas, print to file and printer)
class Listas(Database):

	def __init__(self, *args, table_name='LISTAS'):

		#define una ventana nueva para ver las listas
		listas_window = Toplevel()
		listas_win = WindowConfig(listas_window)
		listas_win.edit_menu.add_separator()
		listas_win.edit_menu.add_command(label="Buscar Item (ENTER)", command=lambda: self.show_data())
		# con super inicializa el init del padre como propio para la nueva ventana
		super().__init__(listas_window, table_name)

		# Deshabilita entrys foraneos
		self.no_edit_entry = [4, 5]
		self.deshab_entry(self.no_edit_entry)

		#carga los entrys, segun seleccion de maestro y deshabilita los entrys
		self.load_lista()

		# Con doble click, muestra la lista seleccionada
		self.tree.bind('<Double-Button-1>', self.show_data) 

		# si encuentra 'keysym=c' --> ejecuta copiar lista
		if str(args).find('keysym=c', 0, -1) != -1:     #(texto a buscar, inicio, final)
			self.copy_list()

		# si encuentra 'keysym=double click' --> ejecuta cargar lista
		if str(args).find('ButtonPress', 0, -1) != -1:
			self.load_lista()

	def read_columns(self):
		query = f'SELECT {self.table_name}.*, ' \
				f'{maestro.table_name}.{maestro.sheet_columns[1]},{maestro.table_name}.{maestro.sheet_columns[3]} ' \
				f'FROM {self.table_name} ' \
				f'INNER JOIN {maestro.table_name} LIMIT 1'
		columns = self.run_query(query)
		for columna in columns.description:
			self.sheet_columns = np.append(self.sheet_columns, columna[0])

	#query search de listas inner join maestro
	def query_search(self, *args, search='', like='%', col_search=0, col_order=1, order='DESC', limit=''):
		query = f'SELECT {self.table_name}.*' \
				f', {maestro.table_name}.{maestro.sheet_columns[1]}, {maestro.table_name}.{maestro.sheet_columns[3]} ' \
				f'FROM "{self.table_name}" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON {self.table_name}.{self.sheet_columns[1]} = {maestro.table_name}.{maestro.sheet_columns[0]} ' \
				f'WHERE {self.table_name}.{self.sheet_columns[col_search]} ' \
				f'LIKE "{search}{like}" ' \
				f'ORDER BY "{self.sheet_columns[0]}" {order}, "{self.sheet_columns[1]}" {order} {limit}'
		return self.run_query(query)

	#override para listas, agrega item desde maestro
	def add_query(self, *args):
		parameters = []
		nom_column = []
		arg_query = ''
		columns = super().query_search(limit='LIMIT 1')
		for column in columns.description:
			nom_column.append('"'+column[0]+'"')
			arg_query += '?'

		parameters.append(self.entry_array[0].get()) #lo agrega bajo el cod de lista
		parameters.append(maestro.entry_array[0].get()) # elemento
		parameters.append(maestro.entry_array[3].get()) # cant (pone el precio en realidad)
		parameters.append(maestro.entry_array[2].get()) # unidad
		return nom_column, arg_query, parameters

	# edita un registro en la base de datos de listas (que tiene 2 columnas menos que las visibles)
	def edit_record(self, *args):
		columns_list = self.sheet_columns
		self.sheet_columns = self.sheet_columns[0:4]

		if self.valid_selection() and self.entry_array[0].get() != '':
			query_text_column = ''
			query_text_item = ''
			parameters = []
			param_anterior = []

			# CREA la query y parametros segun cantidad de entrys de columnas haya
			for index in range(len(self.sheet_columns)):
				# para editar un item, solo comprueba el "cod y descr"
				if index == 0:
					query_text_column += f'"{self.sheet_columns[index]}" = ?'
					query_text_item += f'"{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())
					param_anterior.insert(index, self.tree.item(self.tree.selection())['text'])

				elif index == 1:
					query_text_column += f', "{self.sheet_columns[index]}" = ?'
					query_text_item += f' AND "{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())
					param_anterior.insert(index, self.tree.item(self.tree.selection())['values'][index - 1])

				else:
					query_text_column += f', "{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())

					# para el caso de maestro, en la columna de precio, compara si el mismo cambió
					if self.sheet_columns[index] == 'PRECIO' \
							and self.entry_array[index].get() != self.tree.item(self.tree.selection())['values'][
						index - 1]:
						# si cambió, actualiza la fecha de precio que esta en index mas 1 (le sigue a precio)
						self.entry_array[index + 1].delete(0, 50)
						self.entry_array[index + 1].insert(END, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

			query = f'UPDATE {self.table_name} ' \
					f'SET {query_text_column} ' \
					f'WHERE {query_text_item}'

			parameters.extend(param_anterior)  # primero estan los datos actuales, y despues los anteriores
			self.run_query(query, parameters)
			self.message['text'] = f'El elemento {self.entry_array[1].get()} ha sido actualizado'
			self.show_data()
		else:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
			self.show_data(like='%')

		self.sheet_columns = columns_list

	# funcion override de database para listas, solo impide cuando lista y elemento son identicos a algun registro
	def valid_add(self):
		if maestro.entry_array[0].get() == '' or maestro.entry_array[1].get() == '':
			return False
		query = f'SELECT "{self.sheet_columns[0]}", "{self.sheet_columns[1]}" from {self.table_name}'
		db_rows = self.run_query(query)
		for row in db_rows:
			if self.entry_array[0].get() == row[0] and maestro.entry_array[0].get() == row[1]:
				return False
		return True

	# carga una lista segun el codigo en maestro con doble click
	def load_lista(self, *args):
		self.entry_array[0].delete(0, 50)
		self.entry_array[0].insert(END, maestro.tree.item(maestro.tree.selection())['text'])

		#carga el código exacto sin like %
		self.show_data(like='')

	#Copia una lista con diferente código
	def copy_list(self, *args):
		# crea el registro en maestro con nuevo código
		if maestro.add_record():
			# Busca en listas, la lista a copiar
			cursor = super().query_search(search=self.entry_array[0].get(), like='')

			#arma argumento de query, segun cantidad de columnas
			arg_query = ''
			for column in range(len(self.sheet_columns)-2): #quita las 2 columnas de maestro
				arg_query += '?'

			# arma nuevo dataframe para ingresar en database listas (reemplaza con codigo nuevo la 1er columna)
			parameters = []
			index = 0
			for row in cursor.fetchall():
				tupla = (maestro.entry_array[0].get(), row[1], row[2], row[3])
				parameters.insert(index, tupla)
				index += 1

			# Ingresa masivos datos en database de listas
			query = f'INSERT INTO {self.table_name} ' \
					f'VALUES({", ".join(arg_query)})'

			self.run_query_many(query, parameters)

			# Actualiza en nombre de la lista a mostrar
			self.entry_array[0].delete(0, 50)
			self.entry_array[0].insert(END, maestro.entry_array[0].get())
			self.show_data(like='')			
			maestro.message['text'] = "Nueva lista creada con éxito"
		else:
			maestro.message['text'] = "No se copió ningún registro"

	# encabezado para listas, incluye precio total
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		precio_lista = 0
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			descr_item = self.tree.item(self.tree.selection())['values']

			#item = self.tree.item(self.tree.selection())['text']
			#descr_item.insert(0, item)

			precio = round(float(self.tree.item(self.tree.selection())['values'][1])
						   * float(self.tree.item(self.tree.selection())['values'][4]), 3)
			precio_lista += precio
			descr_item.append(precio)
			list_row.append(descr_item)

		# incorpora el nombre de las columnas y agrega la de precio total de cada item
		self.sheet_columns = np.append(self.sheet_columns, 'MONTO')

		list_row.insert(0, (self.sheet_columns[1:])) #deja de lado la columna de código de lista

		#intercambia orden columnas de unidad por descripción
		for row in list_row:
			temp = row[1]
			row[1] = row[3]
			row[3] = temp

		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Costo Total:  u$s {round(precio_lista, 3)}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código Lista: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row

	# En el caso de LISTAS, no se desea habilitar columnas no habilitadas para busqueda
	def clean_entrys(self, *args):
		super().clean_entrys()
		self.deshab_entry(self.no_edit_entry)


#Base de datos especifica para Ordenes de trabajo
class Ordenes_trabajo(Database):

	def __init__(self, *args, table_name='OT'):
		ot_window = Toplevel()
		ot_win = WindowConfig(ot_window)
		ot_win.edit_menu.add_separator()
		ot_win.edit_menu.add_command(label="Buscar Item (ENTER)", command=lambda: self.show_data(open=True))

		super().__init__(ot_window, table_name)

		# Deshabilita entrys de precios (son automatizados)
		self.no_edit_entry = [3, 4]

	# Ingresa los datos de busqueda en el TREE y los agrupa por ORDEN de TRABAJO
	def data_into_tree(self, cursor, open=0):
		temp=''
		for row in cursor:
			# registra cada codido nuevo y crea la linea para esa OT
			if row[0] != temp:
				line = self.tree.insert('', 0, text=row[0], values=row[1:], open=open)
			else:
				self.tree.insert(line, tk.END, text=row[0], values=row[1:])
			temp = row[0]

	#override para OT, agrega item desde maestro y completa columnas
	def add_query(self, *args):
		parameters = []
		arg_query = ''
		nom_column = []
		
		for column in self.sheet_columns:
			nom_column.append('"'+column+'"')
			arg_query += '?'

		parameters.append(self.entry_array[0].get()) #lo agrega bajo el cod de OT
		parameters.append(maestro.entry_array[0].get()) # codigo
		parameters.append(0) # cant por defecto = 0
		parameters.append(maestro.entry_array[3].get()) # precio
		parameters.append(0) # monto
		return arg_query, parameters

	# edita un registro en la base de datos OT
	def edit_record(self, *args):
		
		if self.valid_selection() and self.entry_array[0].get() != '':
			self.hab_entry(self.no_edit_entry)
			query_text_column = ''
			query_text_item = ''
			parameters = []
			param_anterior = []

			# CREA la query y parametros segun cantidad de entrys de columnas haya
			for index in range(len(self.sheet_columns)):
				# para editar un item, solo comprueba el "cod y descr" o "ot y codigo"
				if index == 0:
					query_text_column += f'"{self.sheet_columns[index]}" = ?'
					query_text_item += f'"{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())
					param_anterior.insert(index, self.tree.item(self.tree.selection())['text'])

				elif index == 1:
					query_text_column += f', "{self.sheet_columns[index]}" = ?'
					query_text_item += f' AND "{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())
					param_anterior.insert(index, self.tree.item(self.tree.selection())['values'][index - 1])

				else:
					if self.sheet_columns[index] == 'MONTO':
						self.entry_array[index].delete(0, 50)
						monto = float(self.entry_array[index-1].get()) * int(self.entry_array[index-2].get())
						self.entry_array[index].insert(END, round(monto, 3))

					query_text_column += f', "{self.sheet_columns[index]}" = ?'
					parameters.insert(index, self.entry_array[index].get())

			query = f'UPDATE {self.table_name} ' \
					f'SET {query_text_column} ' \
					f'WHERE {query_text_item}'

			parameters.extend(param_anterior)  # primero estan los datos actuales, y despues los anteriores
			self.run_query(query, parameters)
			self.deshab_entry(self.no_edit_entry)
			self.message['text'] = f'El elemento {self.entry_array[1].get()} ha sido actualizado'
			self.show_data(open=True)
		else:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
			self.show_data(like='%')

	# encabezado para OT, varias impresiones
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		precio_lista = 0
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			descr_item = self.tree.item(self.tree.selection())['values']

			item = self.tree.item(self.tree.selection())['text']
			descr_item.insert(0, item)

			precio = round(float(self.tree.item(self.tree.selection())['values'][1])
						   * float(self.tree.item(self.tree.selection())['values'][4]), 3)
			precio_lista += precio
			descr_item.append(precio)
			list_row.append(descr_item)

		# incorpora el nombre de las columnas y agrega la de precio total de cada item
		self.sheet_columns = np.append(self.sheet_columns, 'MONTO')

		list_row.insert(0, (self.sheet_columns[1:])) #deja de lado la columna de código de lista

		#intercambia orden columnas de unidad por descripción
		for row in list_row:
			temp = row[1]
			row[1] = row[3]
			row[3] = temp

		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Costo Total:  u$s {round(precio_lista, 3)}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código Lista: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row


#Clase para manejar los rubros
class Rubros(Database):
	def __init__(self, *args, table_name='RUBROS'):
		rubros_window = Toplevel()
		WindowConfig(rubros_window)
		super().__init__(rubros_window, table_name)

	# genera el encabezado genérico de impresión
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			item = self.tree.item(self.tree.selection())['text']
			list_row.append([item,''])

		# incorpora el nombre de las columnas
		list_row.insert(0, (self.sheet_columns))  # deja de lado la columna de código de lista
		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Impreso desde {self.table_name}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row


# ------------FUNCIONES DE MENU------------------------------------------

#Info de ayuda
def help_info():
	# Ventana y titulo
	help_window = Toplevel()
	WindowConfig(help_window)

	# FRAME
	help_frame = LabelFrame(help_window, text='Menu de ayuda')
	help_frame.grid(column=0, row=0, sticky='nsew')
	help_frame.rowconfigure(0, weight=1)
	help_frame.columnconfigure(0, weight=1)

	# TREE
	help_tree = ttk.Treeview(help_frame)
	help_tree.grid(column=0, row=0, sticky='nsew')
	help_tree.column('#0', width=750, minwidth=50, stretch=True)

	# SCROLL BAR
	scroll_bar = Scrollbar(help_window, command=help_tree.yview)
	scroll_bar.grid(column=1, row=0, sticky='nsew')
	help_tree.config(yscrollcommand=scroll_bar.set)

	# ITEMS DE AYUDA
	help_one = help_tree.insert('', tk.END, text='Generar las bases de datos')
	help_tree.insert(help_one, tk.END, text='Primero debe importar las bases de datos ORIGINALES con Excel: '
											'"MAESTRO.dbf", "LISTAS.dbf", "OT.dbf" y "RUBROS.dbf"')
	help_tree.insert(help_one, tk.END, text='y guardar en la misma carpeta del programa los archivos "*.xls"')
	help_tree.insert(help_one, tk.END, text='Luego debe ir a menú "Herramientas/Crear base de datos".')
	help_tree.insert(help_one, tk.END, text='Y reiniciar el programa." ')

	help_two = help_tree.insert('', tk.END, text='Agregar elementos a una "lista" u "órden de trabajo"')
	help_tree.insert(help_two, tk.END, text='Abra la lista u órden de trabajo a editar')
	help_tree.insert(help_two, tk.END, text='Seleccione el item a agregar en Maestro de articulos')
	help_tree.insert(help_two, tk.END, text='y desde la ventana de la lista destino, ir a menú "Edición/Agregar" o pulsar "CTRL-A"')

	help_three = help_tree.insert('', tk.END, text='Copiar o Eliminar Listas completas')
	help_tree.insert(help_three, tk.END, text='Para eliminar una lista basta con eliminar el item de lista desde maestro')
	help_tree.insert(help_three, tk.END, text='Para eliminar una lista basta con eliminar el item de lista desde maestro')

	
	help_last = help_tree.insert('', tk.END, text='Para mas acciones vea el apartado "comandos rápidos" desde el menú ayuda')


#Detalle de los HOTKEYS del programa
def hotkeys():
	messagebox.showinfo('Accesos rápidos', 'Enter: Búsqueda por campo\n'
										   'Ctrl+Enter: limpiar búsqueda\n'
										   'SpaceBar: Desplazar una hoja\n\n'
										   'Ctrl+A: Agrega registro\n'
										   'Ctrl+E: Editar registro\n'
										   'Ctrl+D: Eliminar registro\n\n'
										   'Ctrl+O: Abrir Órdenes de trabajo\n'
										   'Ctrl+L: Abrir Listas (doble click)\n'
										   'Ctrl+R: Abrir Rubros\n\n'
										   'Ctrl+C: Copiar Lista\n'
										   'Ctrl+F: Eliminar Lista\n\n'
										   'Ctrl+B: Crear base de datos\n'
										   'Ctrl+P: Imprimir\n'
										   'Ctrl+T: Cambiar apariencia\n'
										   'ALT+F4: Cerrar\n')


#Info de licencia
def license():  # funcion para ventana emergente que muestra un warning con icono warning
	messagebox.showinfo('GEST2020', 'Licencia válida para uso exclusivo de L.I.E. S.R.L.')


#Info del programa y version
def help_about():  # funcion para vent emergente que muestra info con icono de info
	messagebox.showinfo('Gestor de artículos', 'GEST2020 Versión: V2\n\nProgramado por Agustin Arnaiz'
											 '\n\nEn memoria a Rodolfo Alfredo Taparello.. el Rody')


#------------MAIN-BARRAMENU-LOOP-INSTANCIAS  DATABASES-----------------
if __name__ == '__main__':

	# Primero carga la configuración desde archivo
	manager = ProgramManager()
	manager.load_config()

	# ROOT, ventana principal (carga el maestro de articulos)
	root = ThemedTk(theme=manager.configs.get('theme_name'))
	# Setea ventana al 90% de la pantalla
	monitor_x = str(int(GetSystemMetrics(0)*0.9))
	monitor_y = str(int(GetSystemMetrics(1)*0.9))
	root.geometry(f'{monitor_x}x{monitor_y}+20+5')  # "width_x_height_+off_x_+off_y"

	main_window = WindowConfig(root, title='GEST2020 | Sistema de Administración de Producción L.I.E. S.R.L.')

	# si existe la base de datos la instancia
	if os.path.isfile(manager.configs['db_name']):
		maestro = Maestro(window=root, table_name='Maestro')
	else:
		messagebox.showwarning('Advertencia', 'No se encuentra la base de datos\n\n'
					'vaya a menú "Herramientas/crear base de datos" para generarla.\n'
					'de no existir los archivos XLS, deberá crearlos previamente con "MS Excel" o "CALC OpenOffice".')

	#LOOP CIERRE
	root.mainloop()
print('END.')