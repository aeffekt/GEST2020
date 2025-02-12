'''
GEST2020: Gestor de compras y materiales para pequeñas empresas basado en software GEST;
Consta de una base de datos con 4 tablas: MAESTRO, LISTAS, OT, RUBROS. 
Su fin es conseguir una orden de compra de materiales unificada como "orden de trabajo", 
ordenada por PLAN, y por RUBROS. Diseñado para la empresa L.I.E. S.R.L. por Agustin Arnaiz.
'''

import sqlite3                  # SQL, manejo base de datos
import pandas as pd             # Dataframes y XLS files
import numpy as np				# se usa para listas de datos, acceso no consecutivo a sus items
import os                       # usa la impresión de shell "print"
import inspect
import tkinter as tk
from tkinter import *
from tkinter import filedialog	# abrir archivos
from tkinter import messagebox  # mensajes de salida
from tkinter import ttk         # tkinter mas facha
from ttkthemes import ThemedTk	# themes :D
from datetime import date       # fecha
from datetime import datetime	# fecha con hora, now


#-----------------------------CLASES--------------------------------

# Configuraciones, usuarios (a futuro), etc
class ProgramManager:

	def __init__(self, *args):
		# Configuraciones default
		self.configs = {'db_name': 'db_gest2020.db', 
						'theme_name': 'clearlooks',
						'geometry': '800x600+10+10',
						'fullscreen': 'zoomed',	# o 'normal'
						'path_config': 'config.cfg'}

		self.frame_labels = {'Maestro':'MAESTRO DE ARTÍCULOS',  \
						'LISTAS':'EDITOR DE LISTAS', \
						'OT':'PLANES DE PRODUCCIÓN', \
						'RUBROS':'RUBROS'}

		self.theme_names = {}
		themes_list = (	"clearlooks",			
						"black",
						"blue",
						"equilux",
						"itft1",
						"keramik",
						"kroc",
						"plastik",
						"radiance",
						"smog",
						"winxpblue",
						"xpnative")

		for index, theme in enumerate(themes_list, start=1):
			self.theme_names[index] = theme

	#Carga la configuración desde archivo, o la crea por default
	def load_config(self):
		# Lee configuracioknes de archivo
		lista = []
		if os.path.isfile(self.configs['path_config']):
			config_file = open(self.configs['path_config'], 'r')
			for line in config_file:
				equal = line.find('=', 0, -1)
				lista.append(line[equal+1:-1])
			
			for index, key in enumerate(self.configs):
				self.configs[key] = lista[index]
				
			config_file.close()
		# si no existe el archivo, lo crea con valores por defecto
		else:
			self.save_config()

	# Guarda la configuración en archivo
	def save_config(self, *args):
		for line in self.configs:
			config_file = open(self.configs['path_config'], 'w')
			for key in self.configs:
				config_file.write(key+'=')
				config_file.write(f'{self.configs.get(key)}\n')
		config_file.close()
		
	# acciones antes del cierre de programa
	def exit_handler(*args):
		manager.configs['geometry'] = maestro.window.geometry()
		manager.configs['fullscreen'] = root.state()
		manager.save_config()
		maestro.window.destroy()
		print('END.')

# Formato ventana por defecto: titulo, menu, bindings
class WindowConfig:
	def __init__(self, parent, *args, **kwargs):
		self.window = parent
		self.window.config(bg="grey")  # config de la ventana bg= back ground color
		self.window.iconbitmap('_LOGOLIE x3.ico')  # icono e la ventana
		self.window.rowconfigure(0, weight=1)
		self.window.columnconfigure(0, weight=1)	
		self.window.resizable(True, True)

		self.bar_menu = tk.Menu(self.window)

		self.window.config(menu=self.bar_menu)
		self.file_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.edit_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.tools_menu = tk.Menu(self.bar_menu, tearoff=0)
		#self.windows_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.help_menu = tk.Menu(self.bar_menu, tearoff=0)
		
		self.name_window(**kwargs)
		self.create_menu()
		self.bindings()

	def name_window(self, *args, **kwargs):
		title = kwargs.get('title')
		if title != None:
			self.window.title(title)

	def create_menu(self):
		# items de la barra de menu
		self.bar_menu.add_cascade(label="Archivo", menu=self.file_menu)
		self.bar_menu.add_cascade(label="Edición", menu=self.edit_menu)
		self.bar_menu.add_cascade(label="Herramientas", menu=self.tools_menu)
		#self.bar_menu.add_cascade(label="Ventanas", menu=self.windows_menu)
		self.bar_menu.add_cascade(label="Ayuda", menu=self.help_menu)

		# MENU FILE
		self.file_menu.add_command(label="Abrir Base de datos (CTRL+F)", command=OpenDatabase)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Abrir Listas (CTRL+L)", command=Listas)
		self.file_menu.add_command(label="Abrir Órdenes de trabajo (CTRL+O)", command=OrdenTrabajo)
		self.file_menu.add_command(label="Ver Rubros (CTRL+R)", command=Rubros)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Imprimir (CTRL+P)", command=ToPrinter)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Cerrar ventana (alt+F4)", command=manager.exit_handler)

		# MENU Herramientas
		self.tools_menu.add_command(label="Cambiar apariencia (CTRL+T)", command=lambda: WindowTheme(main_window.window).select_theme())
		self.tools_menu.add_separator()
		self.tools_menu.add_command(label="Crear base de datos (CTRL-B)", command=CreateDatabase)
		
		# Menu ayuda
		self.help_menu.add_command(label="Ayuda", command=help_info)
		self.help_menu.add_command(label="Licencia", command=license)
		self.help_menu.add_command(label="Comandos rápidos", command=hotkeys)
		self.help_menu.add_separator()
		self.help_menu.add_command(label="Acerca de GEST2020", command=help_about)

	# Define los bindings y comandos rápidos para la ventana
	def bindings(self):
		# EVENTOS de la ventana y comandos rápidos
		self.window.bind('<Control-t>', lambda e: WindowTheme(main_window.window).select_theme())
		self.window.bind('<Control-T>', lambda e: WindowTheme(main_window.window).select_theme())
		self.window.bind('<Control-l>', Listas)
		self.window.bind('<Control-L>', Listas)
		self.window.bind('<Control-r>', Rubros)
		self.window.bind('<Control-R>', Rubros)
		self.window.bind('<Control-o>', OrdenTrabajo)
		self.window.bind('<Control-O>', OrdenTrabajo)
		self.window.bind('<Control-b>', CreateDatabase)
		self.window.bind('<Control-B>', CreateDatabase)
		self.window.bind('<Control-f>', OpenDatabase)
		self.window.bind('<Control-F>', OpenDatabase)


#Seleccion de THEME para ventana principal
class WindowTheme:
	def __init__(self, win_to_mod, *args):
		self.window = Toplevel()
		self.window.iconbitmap('_LOGOLIE x3.ico')
		self.window.resizable(False, False)
		self.win_to_mod = win_to_mod
		self.window.config(bg='dark grey')

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


# Crea la base de datos SQLite3 desde los 3 XLS del GEST original
class CreateDatabase:

	def __init__(self, *args):
		if messagebox.askquestion('Pregunta', 
			'¿Desea crear la base de datos ahora?\nEsta acción puede demorarse un momento') == 'yes':
			try:
				self.conn = sqlite3.connect(manager.configs.get('db_name'))
			except:
				OpenDatabase()
				self.conn = sqlite3.connect(manager.configs.get('db_name'))
			try:
				self.maestro = self.format_table('MAESTRO')
				self.listas = self.format_table('LISTAS')
				self.ot = self.format_table('OT')
				self.rubros = self.format_table('RUBROS')

				self.create_triggers()

				self.conn.commit()
				self.conn.close()
				messagebox.showinfo('Info', 'Base de datos creada con éxito\nDebe reiniciar el programa.')

			except Exception as err:
				messagebox.showerror('Error', f'No se pudo crear la base de datos:\n{err}')
			

	# Crea los TRIGGERS
	def create_triggers(self):

		# Actualiza fecha de precio al editar el precio
		query = '''CREATE TRIGGER maestro_fecha_precio
             AFTER UPDATE ON "MAESTRO"
             WHEN old.PRECIO <> new.PRECIO
             BEGIN
                 UPDATE "MAESTRO" SET "FECHA_PRECIO" = CURRENT_TIMESTAMP WHERE "CÓDIGO" = new.CÓDIGO;
             END
             ;
             '''
		self.conn.execute(query)

		# Trigger actualiza Monto de ordenes de trabajo al editar precio o cantidad
		query = '''CREATE TRIGGER ot_monto_update
             AFTER UPDATE ON "OT"
             WHEN old.CANT <> new.CANT
             	 OR old.PRECIO <> new.PRECIO
             BEGIN
                 UPDATE "OT" SET "MONTO" = new.PRECIO * new.CANT WHERE "PLAN" = new.PLAN AND "CÓDIGO" = new.CÓDIGO;
             END
             ;
             '''
		self.conn.execute(query)

		# Trigger actualiza monto de ordenes de trabajo al insertar un registro nuevo
		query = '''CREATE TRIGGER ot_monto_insert
             AFTER INSERT ON "OT"
             BEGIN
                 UPDATE "OT" SET "MONTO" = new.PRECIO * new.CANT WHERE "PLAN" = new.PLAN AND "CÓDIGO" = new.CÓDIGO;
             END
             ;
             '''
		self.conn.execute(query)

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
			nom_columns = ["CÓDIGO", "DESCRIPCIÓN", "UN", "PRECIO", "FECHA_PRECIO", "FECHA_ALTA", "RUBRO"]
			query = 'CREATE TABLE "MAESTRO" ' \
					'("CÓDIGO" TEXT NOT NULL, ' \
					'"DESCRIPCIÓN" TEXT, ' \
					'"UN"	TEXT DEFAULT "UN", ' \
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
			nom_columns = ["LISTA", "CÓDIGO", "CANT", "UN"]

			#creacion de la tabla
			query = 'CREATE TABLE "LISTAS" ' \
					'("LISTA"	TEXT NOT NULL,' \
					'"CÓDIGO"	TEXT,' \
					'"CANT"	REAL,' \
					'"UN"	TEXT,' \
					'FOREIGN KEY("LISTA") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE,' \
					'FOREIGN KEY("CÓDIGO") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE)'	

		if name == "OT":
			# Primero elimina decimales no deseados de columna precio unitario y monto
			xls_file = self.delete_decimal(xls_file, 'PRUNIT,N,10,3')
			xls_file = self.delete_decimal(xls_file, 'MONTO,N,10,3')

			drop_columns = [3, 4, 5, 6, 7, 8, 11, 12] #fecha_venta, fecha_cum, can_cum, nombre_ fecha_ing, nom_ant, plan, despi
			nom_columns = ["PLAN", "CÓDIGO", "CANT", "PRECIO", "MONTO"]
			query = 'CREATE TABLE "OT" ' \
					'("PLAN"	INTEGER, ' \
					'"CÓDIGO"	TEXT NOT NULL, ' \
					'"CANT"	INTEGER DEFAULT 0, ' \
					'"PRECIO"	REAL DEFAULT 0, ' \
					'"MONTO"	REAL DEFAULT 0, ' \
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


# Clase principal de base de datos, maneja una tabla 
class ManageTable:
	
	def __init__(self, window, table_name):
		self.no_edit_entry = []  # Define los entry bloqueados para edicion
		self.window = window
		self.win = WindowConfig(self.window)

		# Atributos principales de cada tabla/ventana
		self.table_name = table_name
		self.entry_array = []  # array de entrys editores de registro de database
		self.table_columns = np.array([]) #usa numpy para citar items no consecutivos
		self.focus_decoder = {}  # carga una lista de widgets entry para realizar busquedas		

		# define FRAMEWORK tabla de datos
		self.frame_tree = ttk.LabelFrame(self.window, text=manager.frame_labels[self.table_name], labelanchor=N)
		self.frame_tree.grid(row=0, column=0, columnspan=20, pady=2, padx=2, sticky=W+E+S+N)
		self.frame_tree.config(cursor='hand2')  #indica seleccion de los elementos del tree

		self.tree = ttk.Treeview(self.frame_tree, height=25)

		# define FRAMEWORK mensajes
		self.frame_msg = ttk.LabelFrame(self.window, text='')
		self.frame_msg.grid(row=10, column=0, columnspan=20, pady=2, padx=2, sticky=W+E+S+N)
		# MENSAJE de salida en la ventana
		self.message = ttk.Label(self.window, text='')
		self.message.grid(row=10, column=0, columnspan=20, sticky=W + E)
		
		# Lee nombres de columna de la tabla, Moldea ventana y lee base de datos
		self.build_main_view()

		# permite expandir los widgets internos a la ventana en x e y
		self.frame_tree.rowconfigure(10, weight=1, minsize=50)
		for index in range(len(self.table_columns)):
			self.frame_tree.columnconfigure(index, weight=1, minsize=50)
		
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

		# MENU edicion
		self.win.edit_menu.add_command(label="Agregar Registro (CTRL+A)", command=self.add_record)
		self.win.edit_menu.add_command(label="Editar Registro (CTRL+E)", command=self.edit_record)
		self.win.edit_menu.add_command(label="Eliminar Registro (CTRL+D)", command=self.delete_record)
		self.win.edit_menu.add_separator()
		self.win.edit_menu.add_command(label="Buscar.. (ENTER)", command=self.show_data)

	#Armado de ventana, TREEVIEW adaptable segun Database
	def build_main_view(self):
		cursor = self.query_search(limit='limit 1000')

		# Nombres de columnas y config de barras de desplaz del tree
		for index, column in enumerate(cursor.description):
			self.table_columns = np.append(self.table_columns, column[0])
		self.tree['columns'] = list(self.table_columns[1:])
		self.tree.grid(row=10, column=0, columnspan=20, rowspan=1, pady=10, sticky=N+S+W+E)
		# Scroll vertical del TREE
		self.scroll_tree_v = ttk.Scrollbar(self.frame_tree, command=self.tree.yview)
		self.scroll_tree_v.grid(row=10, column=len(self.table_columns), sticky=NS)
		self.tree.config(yscrollcommand=self.scroll_tree_v.set)
		# Scroll horizontal del TREE
		self.scroll_tree_h = ttk.Scrollbar(self.frame_tree, orient='horizontal', command=self.tree.xview)
		self.scroll_tree_h.grid(row=12, column=0, columnspan=20, sticky=W+E)
		self.tree.config(xscrollcommand=self.scroll_tree_h.set)

		#se cargan titulos y tamaños de las columnas del tree, asi como Entrys para la edicion de records
		all_table = cursor.fetchall()
		for index, column in enumerate(self.table_columns):
			largest = 6	# tamaño minimo
			for row_table in all_table:
				item_table = str(row_table[index])
				if largest < len(item_table):
					largest = len(item_table)
			if largest > 12:
				largest = 6 + int(largest / 2)
			self.tree.column('#'+str(index), width=largest*11, stretch=True)
			self.tree.heading('#'+str(index), text=f'{self.table_columns[index]}', anchor=CENTER)

			# crea tantos ENTRYS como columnas, para editar registro
			
			self.entry_array.append(self.entrys(frame=self.frame_tree,
												name=column, 
												row=7, 
												column=index, 
												width=largest*2))		

	#Crea Entrys con label superior
	def entrys(self, frame, name='entry', row=0, column=0, width=50):

		# define label y entrys segun llamada
		ttk.Label(frame, text=name).grid(row=row, column=column, pady=5)
		
		#permite ingresar con minusculas la descripción
		if name == 'DESCRIPCIÓN':
			entry = ttk.Entry(frame, width=width)

		#menu fijo para Rubro = ComboBox
		elif self.table_name == 'Maestro' and name == 'RUBRO':
			#lee de la tabla rubros todos los items y los agrega como una lista			
			temp = self.table_name
			self.table_name = 'RUBROS'
			rubros = self.query_search(like='%', col_order=0, order='ASC')
			self.table_name = temp
			lista_rubros = []
			for item in rubros.fetchall():
				lista_rubros.append(item)

			entry = ttk.Combobox(frame, values=lista_rubros, width=width, state='readonly')

		#todos los demas son en mayúscula
		else:
			entry = UpperEntry(frame, width=width)

		entry.grid(row=row + 2, column=column, columnspan=1, sticky=W+E, padx='5', pady=2)
		
		#define decodificacion de focus, para hacer busqueda por columnas segun el focus de entry
		self.focus_decoder[str(entry)] = column
		
		return entry

	#Ejecuta una QUERY SQLite3 con cursor usando parametros
	def run_query(self, query, parameters={}, many=False):
		with sqlite3.connect(manager.configs.get('db_name')) as conn:
			#try:
			conn.execute('PRAGMA foreign_keys = True')
			cursor = conn.cursor()
			if many:
				result = cursor.executemany(query, parameters)
			else:
				result = cursor.execute(query, parameters)
			conn.commit()
			return result
			#except Exception as err:
			#	messagebox.showerror('ERROR', f'Ocurrió un error en la base de datos: \n{err}')

	#Limpia el tree, y lo re-hace segun codido exacto o simil (like) y ordena segun column
	def show_data(self, *args, like='%', col_search=0, col_order=0, limit='', open=True):	
		#Si se llama a la funcion con ENTER, hace busqueda de columna segun foco de entry
		if str(args).find('keysym=Return', 0, -1) != -1:     #(texto a buscar, inicio, final)
			focus = self.window.focus_get()
			if focus is not None:
				
				# Si la busqueda se hace desde el tree, el resultado es exacto y abre el indexed
				if self.table_name != 'Maestro' and self.table_name != 'RUBROS':
					open = str(focus).find('treeview') != -1 or self.focus_decoder[str(focus)] != 0
					if open:
						like=''
				# Si el foco no se decodifica, por defecto busca en 0
				try:
					col_order = col_search = self.focus_decoder[str(focus)]
				except:
					col_order = col_search = 0
		#BUSCA EN LA DB y copia en el TREE
		cursor = self.query_search(*args,
									search=self.entry_array[col_search].get(), 
									like=like, 
									col_search=col_search, 
									col_order=col_order, 
									limit=limit)
		# Habilita los no edit entry para cargar los datos y limpia el tree previamente
		self.hab_entry(self.no_edit_entry)
		self.delete_tree()
		self.data_into_tree(cursor, open)
		self.deshab_entry(self.no_edit_entry)

	# Ingresa los datos de busqueda en el TREE y los agrupa por listas de haberlas
	def data_into_tree(self, cursor, open):
		list_tittle=''
		for row in cursor:
			# si va abierto, no va indexado
			if open:
				self.tree.insert('', 0, text=row[0],  values=row[1:], open=True)
			else:
				# registra cada codigo nuevo y crea la linea para esa lista
				if row[0] != list_tittle:
					line = self.tree.insert('', 0, text=row[0], open=False)
					self.tree.insert(line, tk.END, text=row[0], values=row[1:])
				else:
					self.tree.insert(line, tk.END, text=row[0], values=row[1:])
			last = row
			list_tittle = row[0]
		
	#Ejecuta query de busqueda en DB, retorna dataframe, se programó separado de "database_to_tree" para ser llamada por separado (por print_list y copy_list)
	def query_search(self, *args, search='', like='%', col_search=0, col_order=0, order='DESC', limit=''):

		# si existen las columnas, las usa para definir la busqueda
		if self.table_columns.size > 0:
			# Si hay segunda columna, ordena ASC
			try:
				col_search_two = f', "{self.table_columns[col_order+1]}" ASC '
			except:
		 		col_search_two = ''
			where_line = f' WHERE "{self.table_columns[col_search]}"'
			like_line = f' LIKE "{search}{like}"'
			order_line = f' ORDER BY "{self.table_columns[col_order]}" {order}{col_search_two}'
		else:
			where_line = order_line = like_line = ''

		query = f'SELECT * from "{self.table_name}"{where_line}{like_line}{order_line} {limit}'
				
		return self.run_query(query)

	# valida el agregar un item
	def valid_add(self):
		if maestro.entry_array[0].get() == '':
			return False
		db_rows = self.query_search(search=self.entry_array[0].get(), like='')
		try:
			for row in db_rows:
				if str(row).find(maestro.entry_array[0].get()) != -1:
					return False
			return True
		except:
			return True

	#Borra el tree, para nueva visualizacion
	def delete_tree(self):
		records = self.tree.get_children()  # obtiene todos los datos de la tabla tree
		for element in records:
			self.tree.delete(element)  # limpia todos los datos de tree

	# forma la query para agregar item a tabla
	def add_query(self):
		nom_column, arg_query, parameters = [[],[],[]]
		for index, column in enumerate(self.table_columns):
			# Solo agrega si no se encuentra en self.no_edit_entry o si el entry no está vacío	
			if index not in self.no_edit_entry:
				if self.entry_array[index].get() != '' or self.entry_array[index].get() != None:
					nom_column.insert(index, '"'+column+'"')
					arg_query.insert(index, '?')
					parameters.insert(index, self.entry_array[index].get())
		
		return nom_column, arg_query, parameters

	# agrega un registro en la base de datos
	def add_record(self, *args):

		if self.valid_add():
			nom_column, arg_query, parameters= self.add_query()			

			query = f'INSERT INTO {self.table_name} ({", ".join(nom_column)}) ' \
				f'VALUES({" ,".join(arg_query)})' #une la lista con join

			self.run_query(query, parameters)

			self.message['text'] = f'{maestro.entry_array[0].get()} ha sido guardado con éxito'
			fail_add = ''
		else:
			messagebox.showwarning('Advertencia', 'No se puede agregar ese registro.\nYa existe, está vacío o no corresponde')
			fail_add = '%'  # evita borrar el tree cuando se agrega algo vacio al buscar con LIKE

		self.show_data(like=fail_add, open=True)  # like = %: busca la db con "codigo%"
		return not fail_add  # devuelve True si agrego, false si no agrego registro

	#Edita un registro en la base de datos
	def edit_record(self, *args):

		# se habilitan los entrys, ya que se modifican a veces antes de editar un registro
		self.hab_entry(self.no_edit_entry)
		if self.tree.item(self.tree.selection())['text'] != '' and self.entry_array[0].get() != '':
			try:
				query_text_column, query_text_item, parameters, param_anterior = [[],[],[],[]]

				# Edita los las columnas que no pertenecen a self.no_edit_entry
				for index, column in enumerate(self.table_columns):

					if index not in self.no_edit_entry:
						#para editar un item, lo comprueba el "cod y descr" o "lista y cod"
						if index == 0:
							param_anterior.insert(index, self.tree.item(self.tree.selection())['text'])
							query_text_item.insert(index, f'"{column}" = ? ')
						elif index == 1:
							param_anterior.insert(index, self.tree.item(self.tree.selection())['values'][index - 1])
							query_text_item.insert(index, f'"{column}" = ? ')

						query_text_column.insert(index, f'"{column}" = ? ')
						parameters.insert(index, self.entry_array[index].get())
						
					
				query = f'UPDATE {self.table_name} ' \
						f'SET {", ".join(query_text_column)} ' \
						f'WHERE {"AND ".join(query_text_item)}'

				# primero estan los datos actuales, y despues los anteriores
				parameters.extend(param_anterior) 

				self.run_query(query, parameters)
				try:
					self.message['text'] = f'El elemento {self.entry_array[0].get()} {self.entry_array[1].get()} ha sido actualizado'
				except:
					self.message['text'] = f'El elemento {self.entry_array[0].get()} ha sido actualizado'

				self.show_data(like='', open=True)

			except Exception as err:
				messagebox.showerror('Error', f'No se pudo editar:\n{err}')
			
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
						f'WHERE "{self.table_columns[0]}" = ? ' \
						f'AND "{self.table_columns[1]}" = ?'

			#este artificio funciona para rubros que solo tiene una columna ['text'] y no la columna ['values']
			except:
				descr_sel = 'None'

			#para los escasos registros donde la descr es NULL, esto permite borrar esa entrada
			if descr_sel == 'None':
				descr_sel = ''
				query = f'DELETE FROM "{self.table_name}" ' \
						f'WHERE "{self.table_columns[0]}" = ? '
				self.run_query(query, (cod_sel,))  # pone la coma para que se entienda que es una tupla

			else:
				self.run_query(query, (cod_sel, descr_sel))
			self.message['text'] = f'El registro {cod_sel} {descr_sel} ha sido eliminado'
		self.show_data(like='', open=True)  # actualiza la tabla

	#Borra los entrys de registro
	def clean_entrys(self, *args):
		# Habiita todos los entrys despues de borrarlos, permite buscar por cada campo
		self.hab_entry(self.no_edit_entry)
		for index, entry_element in enumerate(self.entry_array):
			
			if str(entry_element) == ".!labelframe.!combobox":
				self.entry_array[index].set("VS.")
			else:
				entry_element.delete(0, 'end')	
		self.deshab_entry(self.no_edit_entry)
		# Hace foco en buscar registro por código
		self.entry_array[0].focus()

	# Carga el registro seleccionado de TREE en los entrys de edicion
	def load_edit_item(self, *args):
		self.hab_entry(self.no_edit_entry)
		
		# Carga solo cuando hay un item seleccionado y tiene algun value
		if len(self.tree.selection()) == 1:

			# Carga en el array de entrys los valores de row seleccionados
			for index, entry in enumerate(self.entry_array):
				
				if index == 0:
					entry.delete(0, 50)
					entry.insert(END, self.tree.item(self.tree.selection())['text'])

				else:
					if self.tree.item(self.tree.selection())['values'] != '':

						# maneja por separado el combobox de rubro, se setea diferente
						if str(entry).find('combobox') != -1:
							
							entry.set(self.tree.item(self.tree.selection())['values'][index - 1])

						# Carga todos los values del tree
						else:
							entry.delete(0, 100)
							entry.insert(END, self.tree.item(self.tree.selection())['values'][index - 1])
					
		self.deshab_entry(self.no_edit_entry)

	# contruye dataframe a partir de la vista del tree
	def prepare_to_print(self, *args):
		dframe = self.build_print()

		#manda a imprimir
		ToPrinter(to_print=dframe, from_obj=self.table_name+self.entry_array[0].get())

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
		list_row.insert(0, (self.table_columns))  # deja de lado la columna de código de lista
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


# Base de datos especifica para tabla Maestro
class Maestro(ManageTable):
	
	def __init__(self, window, table_name='MAESTRO'):
		self.window = window
		super().__init__(self.window, table_name)
		self.no_edit_entry = [4, 5]
		self.show_data()

		# Suma acciones al menú
		self.win.edit_menu.add_separator()
		self.win.edit_menu.add_command(label="Copiar Lista (CTRL+C)", command=lambda: Listas('keysym=c'))
		self.win.edit_menu.add_command(label="Eliminar Lista (CTRL+F)", command=lambda: Listas('keysym=f'))

		# solo desde MAESTRO abre una lista con doble click
		self.tree.bind('<Double-Button-1>', Listas) 
		self.window.bind('<Control-c>', Listas)  # copia lista con nuevo codigo
		self.window.bind('<Control-C>', Listas)  # copia lista con nuevo codigo

	# override para eliminar una columna de fecha de alta
	def build_print(self):
		list = super().build_print()
		df = pd.DataFrame(list)
		df.drop(df.columns[5], axis=1, inplace=True)
		return df


# Base de datos especifica para tabla listas
class Listas(ManageTable):

	def __init__(self, *args, table_name='LISTAS'):

		#define una ventana nueva para ver las listas
		listas_window = Toplevel()
		
		# con super inicializa el init del padre como propio para la nueva ventana
		super().__init__(listas_window, table_name)
		self.no_edit_entry = [4, 5]
		self.deshab_entry(self.no_edit_entry)

		# si encuentra 'keysym=c' --> ejecuta copiar lista
		if str(args).find('keysym=c', 0, -1) != -1:     #(texto a buscar, inicio, final)
			self.copy_list()

		# si encuentra 'keysym=double click' --> ejecuta cargar lista
		if str(args).find('ButtonPress', 0, -1) != -1:
			self.load_lista()
		else:
			self.show_data(open=False)

	# query search de listas inner join maestro
	def query_search(self, *args, search='', like='%', col_search=0, col_order=1, order='DESC', limit=''):
		if self.table_columns.size > 0:
			where_line = f' WHERE {self.table_name}.{self.table_columns[col_search]}'
			like_line = f' LIKE "{search}{like}"'
			order_line = f' ORDER BY "{self.table_columns[0]}" {order}, "{self.table_columns[1]}" {order}'
		else:
			where_line = like_line = order_line = ''

		query = f'SELECT {self.table_name}.*' \
				f', {maestro.table_name}.{maestro.table_columns[1]}, {maestro.table_name}.{maestro.table_columns[3]} ' \
				f'FROM "{self.table_name}" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON {self.table_name}.CÓDIGO = {maestro.table_name}.CÓDIGO ' \
				f'{where_line}{like_line}{order_line} {limit}'
		
		return self.run_query(query)

	# override para listas, agrega los registros desde maestro
	def add_query(self, *args):
		# Carga los entrys los datos desde maestro
		self.entry_array[1].delete(0, 100)
		self.entry_array[2].delete(0, 50)
		self.entry_array[3].delete(0, 50)
		self.entry_array[1].insert(END, maestro.tree.item(maestro.tree.selection())['text']) # elemento
		self.entry_array[2].insert(END, maestro.entry_array[3].get()) # cant (pone el precio en realidad)
		self.entry_array[3].insert(END, maestro.tree.item(maestro.tree.selection())['values'][1]) # unidad

		return super().add_query()

	# carga una lista segun el codigo en maestro con doble click
	def load_lista(self, *args):
		self.entry_array[0].delete(0, 50)
		self.entry_array[0].insert(END, maestro.tree.item(maestro.tree.selection())['text'])

		#carga el código exacto sin like %
		self.show_data(like='', open=False)

	#Copia una lista con diferente código
	def copy_list(self, *args):
		# crea el registro en maestro con nuevo código
		if maestro.add_record():
			# Busca en listas, la lista a copiar
			cursor = super().query_search(search=self.entry_array[0].get(), like='')

			#arma argumento de query, segun cantidad de columnas
			arg_query = ''
			for column in range(len(self.table_columns)-2): #quita las 2 columnas de maestro
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

			self.run_query(query, parameters, many=True)

			# Actualiza en nombre de la lista a mostrar
			self.entry_array[0].delete(0, 50)
			self.entry_array[0].insert(END, maestro.entry_array[0].get())
			self.show_data(like='', open=True)			
			maestro.message['text'] = "Nueva lista creada con éxito"
		else:
			maestro.message['text'] = "No se copió ningún registro"

	# encabezado para listas, incluye precio total
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		precio_lista = precio = 0
		for row in self.tree.get_children():
			self.tree.selection_set(row)

			#item = self.tree.item(self.tree.selection())['text']
			#descr_item.insert(0, item)
			try:
				descr_item = self.tree.item(self.tree.selection())['values']
				print(descr_item)
				precio = round(float(self.tree.item(self.tree.selection())['values'][1])
							   * float(self.tree.item(self.tree.selection())['values'][4]), 3)
				precio_lista += precio
			except:
				precio = precio
			descr_item.append(precio)
			list_row.append(descr_item)

		# incorpora el nombre de las columnas y agrega la de precio total de cada item
		self.table_columns = np.append(self.table_columns, 'MONTO')

		list_row.insert(0, (self.table_columns[1:])) #deja de lado la columna de código de lista

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


# Base de datos especifica para planes de trabajo
class OrdenTrabajo(ManageTable):

	def __init__(self, *args, table_name='OT'):
		self.window = Toplevel()
		super().__init__(self.window, table_name)
		self.no_edit_entry = [4, 5, 6]
		
		self.show_data(open=False)

		self.win.tools_menu.add_separator()
		self.win.tools_menu.add_command(label="Generar compra (CTRL-X)", command=lambda: OrdenCompra(buy_order=self.entry_array[0].get()))
		self.window.bind('<Control-X>', lambda b: OrdenCompra(buy_order=self.entry_array[0].get()))
		self.window.bind('<Control-x>', lambda b: OrdenCompra(buy_order=self.entry_array[0].get()))

	# query search de OT compra inner join maestro
	def query_search(self, *args, search='', like='%', col_search=0, col_order=6, order='DESC', limit=''):
		if self.table_columns.size > 0:
			where_line = f' WHERE {self.table_name}.{self.table_columns[col_search]}'
			like_line = f' LIKE "{search}{like}"'
			order_line = f' ORDER BY {self.table_columns[0]} {order}, "{self.table_columns[1]}" {order}'
		else:
			where_line = like_line = order_line = ''

		query = f'SELECT {self.table_name}.*' \
				f', {maestro.table_name}.{maestro.table_columns[1]}, {maestro.table_name}.{maestro.table_columns[6]} ' \
				f'FROM "{self.table_name}" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON {self.table_name}.CÓDIGO = {maestro.table_name}.CÓDIGO ' \
				f'{where_line}{like_line}{order_line} {limit}'
		
		return self.run_query(query)

	#override para OT, agrega los registros desde maestro
	def add_query(self, *args):
		# Carga los entrys los datos desde maestro
		self.entry_array[1].delete(0, 100)
		self.entry_array[2].delete(0, 50)
		self.entry_array[3].delete(0, 50)
		self.entry_array[4].delete(0, 50)

		self.entry_array[1].insert(END, maestro.tree.item(maestro.tree.selection())['text']) # codigo
		self.entry_array[2].insert(END, maestro.entry_array[3].get()) # cantidad, usa el entry de precio
		self.entry_array[3].insert(END, maestro.tree.item(maestro.tree.selection())['values'][2]) # precio

		return super().add_query()

	# encabezado para OT, varias impresiones
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		costo_plan = 0
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			descr_item = self.tree.item(self.tree.selection())['values']

			item = self.tree.item(self.tree.selection())['text']
			descr_item.insert(0, item)

			costo_plan += float(self.tree.item(self.tree.selection())['values'][3])
			
			list_row.append(descr_item)

		list_row.insert(0, (self.table_columns[1:])) #deja de lado la columna de código de PLAN

		#intercambia orden columnas de unidad por descripción
		for row in list_row:
			temp = row[1]
			row[1] = row[3]
			row[3] = temp

		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Costo Total:  u$s {round(costo_plan, 3)}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código Lista: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row


# Orden de compra
class OrdenCompra(OrdenTrabajo):
	
	def __init__(self, *args, table_name='OT', buy_order):
		self.no_edit_entry = []		
		super().__init__(table_name)
		self.entry_array[0].insert(END, buy_order)
		self.show_data(like='')


# Clase para manejar los rubros
class Rubros(ManageTable):
	def __init__(self, *args, table_name='RUBROS'):
		rubros_window = Toplevel()
		WindowConfig(rubros_window)
		super().__init__(rubros_window, table_name)
		self.show_data()

	# genera el encabezado genérico de impresión
	def build_print(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			item = self.tree.item(self.tree.selection())['text']
			list_row.append([item,''])

		# incorpora el nombre de las columnas
		list_row.insert(0, (self.table_columns))  # deja de lado la columna de código de lista
		list_row.insert(0, ('', ''))
		list_row.insert(0, (f'Impreso desde {self.table_name}', ''))
		list_row.insert(0, (f'Fecha de Impresión: {date.today()}', ''))
		list_row.insert(0, (f'Fecha de alta: {maestro.entry_array[5].get()}', ''))
		list_row.insert(0, (f'Descripción: {maestro.entry_array[1].get()}', ''))
		list_row.insert(0, (f'Código: {self.entry_array[0].get()}', ''))
		list_row.insert(0, ('', ''))
		list_row.insert(0, ('L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020', ''))
		return list_row


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


#Pasa automaticamente a MAYUSCULAS textos de los entrys
class UpperEntry(tk.ttk.Entry):
	def __init__(self, frame, *args, **kwargs):
		self.text_to_upper = tk.StringVar(frame)
		super().__init__(frame, *args, **kwargs)
		self.configure(textvariable=self.text_to_upper)
		self.text_to_upper.trace_add('write', self._to_upper)

	def configure(self, cnf=None, **kwargs):		
		# Crea el entry original
		super().config(cnf, **kwargs)

	# Pasa el texto a mayúscula
	def _to_upper(self, *args):
		self.text_to_upper.set(self.text_to_upper.get().upper())


# ------------FUNCIONES DE MENU Y BASICOS-----------------------------------------

# Abre un archivo Database y lo setea como default (path)
def OpenDatabase(*args):
		path = filedialog.askopenfilename(title='Abrir Base de datos', filetypes=(
                                                    ('SQLite3', '*.db'), ('todos los archivos', '*.*')))
		if path:
			manager.configs["db_name"] = path
			manager.save_config()
			messagebox.showwarning('Advertencia', 'Debe reiniciar el programa para leer la nueva base de datos')


# Info de ayuda para el usuario
def help_info():
	# Ventana y titulo
	help_window = Toplevel()
	WindowConfig(help_window)

	# FRAME
	help_frame = ttk.LabelFrame(help_window, text='Menu de ayuda')
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
	help_tree.insert(help_two, tk.END, text='y desde la ventana de la lista destino, ir a menú "Edición/Agregar" o pulsar "CTRL+A"')
	help_tree.insert(help_two, tk.END, text='Puede ingresar en el campo de edición de PRECIO (en MAESTRO), la CANTIDAD que desea agregar para la lista')

	help_three = help_tree.insert('', tk.END, text='Copiar o Eliminar Listas completas')
	help_tree.insert(help_three, tk.END, text='Para copiar una lista, desde maestro seleccionar la lista a copiar')
	help_tree.insert(help_three, tk.END, text='Se carga la lista en los campos de edición arriba, cambiar el nombre')
	help_tree.insert(help_three, tk.END, text=' y pulsar CTRL-C y desde menú "edición/copiar lista"')
	help_tree.insert(help_three, tk.END, text='Para eliminar una lista basta con eliminar el item de lista desde maestro')
	
	help_last = help_tree.insert('', tk.END, text='Para mas acciones vea el apartado "comandos rápidos" desde el menú ayuda')


# Detalle de los HOTKEYS del programa
def hotkeys():
	messagebox.showinfo('Accesos rápidos',  'CTRL+S: Abrir Base de datos\n'
											'CTRL+B: Crear base de datos\n\n'
											'Enter: Búsqueda por campo\n'
											'CTRL+Enter: limpiar búsqueda\n'
											'SpaceBar: Desplazar una hoja\n\n'
											'CTRL+A: Agregar registro\n'
											'CTRL+E: Editar registro\n'
										    'CTRL+D: Eliminar registro\n'
										    'CTRL+C: Copiar Lista\n\n'
										    'CTRL+L: Abrir Listas (doble click)\n'
										    'CTRL+O: Abrir Órdenes de trabajo\n'
										    'CTRL+X: Emitir Órden de Compra\n'
										    'CTRL+R: Abrir Rubros\n\n'
										    'CTRL+P: Imprimir\n'
										    'CTRL+T: Cambiar apariencia\n'
										    'ALT+F4: Cerrar ventana')


# Info de licencia
def license():  # funcion para ventana emergente que muestra un warning con icono warning
	messagebox.showinfo('GEST2020', 'Licencia válida para uso exclusivo de L.I.E. S.R.L.')


# Info del programa y versión
def help_about():  # funcion para vent emergente que muestra info con icono de info
	messagebox.showinfo('Gestor de artículos', 'GEST2020 Versión: V2\n\nProgramado por Agustin Arnaiz'
											 '\n\nEn memoria a Rodolfo Alfredo Taparello, "el Rody".')


#------------MAIN-BARRAMENU-LOOP-INSTANCIAS  DATABASES-----------------
if __name__ == '__main__':
	print('Start:')
	# Primero carga la configuración desde archivo
	manager = ProgramManager()
	manager.load_config()

	# ROOT, ventana principal (carga el maestro de articulos)
	root = ThemedTk(theme=manager.configs.get('theme_name'))
	root.state(manager.configs.get('fullscreen'))
	root.geometry(manager.configs.get('geometry'))
	main_window = WindowConfig(root, title='GEST2020 | Sistema de Administración de Producción L.I.E. S.R.L.')
	

	# si existe la base de datos la instancia
	if os.path.isfile(manager.configs['db_name']):
		#try:
		maestro = Maestro(window=root, table_name='Maestro')
		#except Exception as err:
		#	messagebox.showwarning('Advertencia', f'Hubo un problema al leer base de datos:\n{err}')
	else:
		messagebox.showwarning('Advertencia', 'No se encuentra la base de datos\n\n'
					'vaya a menú "Inicio/Abrir base de datos".\n'
					'o vaya "Herramientas/crear base de datos" para generarla.')

	root.protocol("WM_DELETE_WINDOW", manager.exit_handler)
	#LOOP CIERRE
	root.mainloop()
##------------------------ END OF CODE --------------------------------##