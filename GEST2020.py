'''GEST2020: Gestor de compras y materiales para pequeñas empresas (basado en software DOS GEST)
Consta de una base de datos con 4 tablas: MAESTRO, LISTAS, OT, RUBROS. 
Su fin es conseguir una orden de compra de materiales unificada como "orden de trabajo", 
ordenada por ORDEN, y por RUBROS y poder imprimirla via S.O. Windows.
Diseñado para la empresa L.I.E. S.R.L. por Agustin Arnaiz. '''

version = "1.0.10"

import os, sys							# check path + impresión de shell "print" *(not in use)
import win32print, win32ui, win32con 	# uso de impresora
from tabulate import tabulate			# justify left dataframe (for printing)
import functools						# se usa para sort column del tree
import logging 							# log to file
import sqlite3							# SQL, manejo base de datos integrada
import pandas as pd						# Dataframes y XLS files
import numpy as np						# uso en listas de datos, acceso no consecutivo a sus items
import time 							# uso de timer (decorador, mide el tiempo de ejecucion)
import tkinter as tk 					# GUI
from tkinter import *					# sys.executable, etc
from tkinter import filedialog, messagebox, ttk			# abrir archivos, mensajes de salida, looks
from ttkthemes import ThemedTk			# themes :D
from datetime import date				# fecha
from datetime import datetime			# fecha con hora, now
from UpperEntry import UpperEntry		# Pasa entrys a mayúsculas (código tomado de internet)


#-----------------------------CLASES--------------------------------#
# Manejo de aspectos generales del programa: Logger, Configuraciones, @timer, backup, cierre final
class ProgManager:

	def __init__(self, *args):
		# instancia de logger maestro, asi permite autonomía de configuracion
		self.logger = logging.getLogger(__name__) 
		self.logger.setLevel(logging.WARNING)	# nivel de logger [INFO:DEBUG:WARNING:ERROR:CRITICAL]
		formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(funcName)s:%(message)s:%(lineno)d') # formateador del texto
		file_handler = logging.FileHandler('_debug.log')	# un manejador del archivo de log
		file_handler.setFormatter(formatter)
		stream_handler = logging.StreamHandler()
		self.logger.addHandler(file_handler)	# asocia para escribir a archivo
		self.logger.addHandler(stream_handler)	# asocia para mostrar en consola
		# Configuraciones default
		self.configs = {'db_name': 'db_gest2020.db', 		# DATABASE
						'theme_name': 'clearlooks',			# THEME
						'geometry': '800x600+10+10', 		# MAIN WINDOW
						'fullscreen': 'zoomed',				
						'geometryL': '800x600+10+10', 		# LISTAS WINDOW
						'fullscreenL': 'normal',
						'geometryO': '800x600+10+10', 		# OT WINDOW
						'fullscreenO': 'normal',
						'path_config': 'config.cfg',		# ARCHIVO DE CONFIGURACION
						'backup_max': '1'}					# Maxima cantidad de archivos backup
		self.frame_labels = {'Maestro':'MAESTRO DE ARTÍCULOS',  \
						'LISTAS':'EDITOR DE LISTAS', \
						'OT':'ORDENES DE TRABAJO', \
						'RUBROS':'RUBROS'}
		self.theme_names = {}
		themes_list = (	"clearlooks",			
						"black",
						"blue",
						"equilux",
						#"itft1",
						#"keramik",
						"kroc",
						#"plastik",
						"radiance",
						#"smog",
						# "winxpblue",
						"xpnative")
		
		for index, theme in enumerate(themes_list, start=1):
			self.theme_names[index] = theme

	# WRAPPER FUNCTION imprime en consola el tiempo de ejecucion de una funcion
	@staticmethod
	def timer(func):
		def wrapper(*args, **kwargs):
			start = time.time()
			rv = func(*args, **kwargs)	
			total_time = time.time() - start
			func_name = str(func).split( )
			print(f'Time by: {func_name[1]}: ', round(total_time, 3), 's')
			return rv
		return wrapper

	# genera un archivo igual con nombre backup en el mismo directorio
	def file_backup(self, file_path):
		self.backup_max = int(self.configs['backup_max'])
		if self.backup_max > 0:
			with open(file_path, 'rb') as file:
				path_backup = self.backup_name(file_path)
				file_backup = open(path_backup, 'wb')
				for line in file:
					file_backup.write(line)
				file_backup.close()

	# recursion para nombre del archivo backup, hasta 10 archivos permite, reescribe el mas antiguo despues
	def backup_name(self, file_path, num=0, fechas_mod={}):
		path, ext = file_path.split('.')	
		path_backup = path+'_backup_'+str(num)+'.'+ext
		if os.path.isfile(path_backup) and num < self.backup_max:
			fechas_mod[os.path.getmtime(path_backup)] = num
			return self.backup_name(file_path, num+1, fechas_mod)
		# busca el archivo mas viejo y lo pasa para reescribir
		elif num == self.backup_max:
			older_num = fechas_mod[min(fechas_mod)]
			return path+'_backup_'+str(older_num)+'.'+ext
		else:	
			return path_backup

	# Carga la configuración desde archivo, o la crea por default
	def load_config(self):
		lista = []
		# revisa si existe el archivo de configuraciones y que el mismo no este vacío
		if os.path.isfile(self.configs['path_config']) and os.path.getsize(self.configs['path_config']) != 0:
			with open(self.configs['path_config'], 'r') as config_file:
				for line in config_file:
					equal = line.find('=', 0, -1)
					lista.append(line[equal+1:-1])
				try:
					for index, key in enumerate(self.configs):
						if index<len(lista):	# evita errores por cantidad de items en el archivo vs configs
							self.configs[key] = lista[index]
				except Exception as err:
					messagebox.showwarning('LEER', f'Ocurrió un error al cargar las configuraciones iniciales.{err}', parent=root)
		# si no existe el archivo, lo crea con valores por defecto
		else:
			self.save_config()

	# Guarda la configuración en archivos
	def save_config(self, *args):
		for line in self.configs:
			with open(self.configs['path_config'], 'w') as config_file:
				for key in self.configs:
					config_file.write(key+'=')
					config_file.write(f'{self.configs.get(key)}\n')	

	# acciones antes del cierre de programa
	def exit_handler(self):
		try:
			self.configs['geometry'] = root.geometry()
			self.configs['fullscreen'] = root.state()
			manager.save_config()
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'No se pudo guardar la configuración al salir: {err}')
		finally:
			manager.logger.debug('END.')
			root.destroy()
			sys.exit()


# Formato ventana por defecto: titulo, menu, bindings, ttk style
class WindowConfig:

	# crea el menú por defecto de cada ventana y los bindings e icono
	def __init__(self, parent, *args, **kwargs):
		self.window = parent
		self.window.config(bg="grey")  # config de la ventana bg= back ground color
		self.window.iconbitmap(sys.executable)  # icono e la ventana
		self.window.rowconfigure(0, weight=1)
		self.window.columnconfigure(0, weight=1)	
		self.window.resizable(True, True)
		self.bar_menu = tk.Menu(self.window, tearoff=0)
		self.window.config(menu=self.bar_menu)
		self.file_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.edit_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.tools_menu = tk.Menu(self.bar_menu, tearoff=0)
		self.help_menu = tk.Menu(self.bar_menu, tearoff=0)
		# configura letra tipo BOLD
		self.bold = ttk.Style()
		self.bold.configure("Bold.TButton", font = ('Sans','10','bold'))
		self.name_window(**kwargs)
		self.create_menu()
		self.bindings()

	# cambia el titulo de una ventana
	def name_window(self, *args, **kwargs):
		title = kwargs.get('title')
		if title != None:
			self.window.title(title)

	# carga el menu en la barra de menu en la window
	def create_menu(self):
		# items de la barra de menu
		self.bar_menu.add_cascade(label="Archivo", menu=self.file_menu)
		self.bar_menu.add_cascade(label="Edición", menu=self.edit_menu)
		self.bar_menu.add_cascade(label="Herramientas", menu=self.tools_menu)
		self.bar_menu.add_cascade(label="Ayuda", menu=self.help_menu)
		# MENU FILE
		self.file_menu.add_command(label="Abrir Base de datos (CTRL+F)", command=CreateDatabase.OpenDatabase)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Abrir Listas (CTRL+L)", command=Listas)
		self.file_menu.add_command(label="Abrir Órdenes de trabajo (CTRL+O)", command=OrdenTrabajo)
		self.file_menu.add_command(label="Ver Rubros (CTRL+R)", command=Rubros)
		self.file_menu.add_separator()
		self.file_menu.add_command(label="Cerrar ventana (alt+F4)", command=manager.exit_handler)
		# MENU Herramientas
		self.tools_menu.add_command(label="Cambiar apariencia (CTRL+T)", command=lambda: WindowTheme(main_window.window))
		self.tools_menu.add_separator()
		self.tools_menu.add_command(label="Crear base de datos (CTRL-B)", command=CreateDatabase)
		# Menu ayuda
		self.help_menu.add_command(label="Ayuda (CTRL+H)", command=self.help_info)
		self.help_menu.add_command(label="Licencia", command=self.license)
		self.help_menu.add_command(label="Comandos rápidos", command=self.hotkeys)
		self.help_menu.add_separator()
		self.help_menu.add_command(label="Acerca de GEST2020", command=self.help_about)

	# Define los bindings y comandos rápidos para la ventana
	def bindings(self):
		# EVENTOS de la ventana y comandos rápidos
		self.window.bind('<Control-t>', lambda e: WindowTheme(main_window.window))
		self.window.bind('<Control-T>', lambda e: WindowTheme(main_window.window))
		self.window.bind('<Control-l>', Listas)
		self.window.bind('<Control-L>', Listas)
		self.window.bind('<Control-r>', Rubros)
		self.window.bind('<Control-R>', Rubros)
		self.window.bind('<Control-o>', OrdenTrabajo)
		self.window.bind('<Control-O>', OrdenTrabajo)
		self.window.bind('<Control-b>', CreateDatabase)
		self.window.bind('<Control-B>', CreateDatabase)
		self.window.bind('<Control-f>', CreateDatabase.OpenDatabase)
		self.window.bind('<Control-F>', CreateDatabase.OpenDatabase)
		self.window.bind('<Control-h>', WindowConfig.help_info)
		self.window.bind('<Control-H>', WindowConfig.help_info)

	# Info de ayuda para el usuario
	def help_info(*args):
		# Ventana y titulo
		help_window = Toplevel()
		WindowConfig(help_window)
		# FRAME
		help_frame = ttk.LabelFrame(help_window, text='Menu de ayuda', labelanchor=N)
		help_frame.grid(column=0, row=0, sticky='nsew')
		help_frame.rowconfigure(0, weight=1)
		help_frame.columnconfigure(0, weight=1)
		# TREE
		help_tree = ttk.Treeview(help_frame)
		help_tree.grid(column=0, row=0, sticky='nsew')
		help_tree.column('#0', width=750, minwidth=50, stretch=True)
		# SCROLL BAR Y
		scroll_bary = Scrollbar(help_window, command=help_tree.yview)
		scroll_bary.grid(column=1, row=0, sticky='ns')
		help_tree.config(yscrollcommand=scroll_bary.set)
		# ITEMS DE AYUDA - Generar las bases de datos
		help_one = help_tree.insert('', tk.END, text='Generar las bases de datos:')
		help_tree.insert(help_one, tk.END, text='Primero debe importar las bases de datos ORIGINALES con Excel:')
		help_tree.insert(help_one, tk.END, text='"MAESTRO.dbf", "LISTAS.dbf", "OT.dbf" y "RUBROS.dbf"')
		help_tree.insert(help_one, tk.END, text='y guardar en la misma carpeta del programa los archivos "*.xls"')
		help_tree.insert(help_one, tk.END, text='Luego debe ir a menú "Herramientas/Crear base de datos".')
		# 2 a - Como crear una "lista
		help_two_a = help_tree.insert('', tk.END, text='Como crear una "Lista": ')
		help_tree.insert(help_two_a, tk.END, text='Primero debe crear la lista en el maestro de artículos: La misma debe tener un código único (CTRL+A)')
		help_tree.insert(help_two_a, tk.END, text='Abrir la lista haciendo doble click sobre la misma, ')
		help_tree.insert(help_two_a, tk.END, text='y puede ir agregando los elementos con su código, o desde maestro de artículos, con el comando de agregar (CTRL+A).')
		# 2 b - Agregar elementos a una "lista" u "órden de trabajo
		help_two = help_tree.insert('', tk.END, text='Agregar elementos a una "Lista" u "Órden de trabajo":')
		help_tree.insert(help_two, tk.END, text='Abra una Lista u Órden a editar y complete los campos de edición para agregar un ítem (CTRL+A)')
		help_tree.insert(help_two, tk.END, text='Si no conoce el código exacto, deje el campo CÓDIGO vacío, y vaya a la ventana de Maestro de articulos.')
		help_tree.insert(help_two, tk.END, text='desde allí selecciona el item deseado, vuelve a la ventana de Lista u Órden de destino. ')
		help_tree.insert(help_two, tk.END, text='y desde menú "Edición/Agregar" o el comando "CTRL+A" lo agrega.')
		# 3 - Copiar o Eliminar Listas completas
		help_three = help_tree.insert('', tk.END, text='Copiar o Eliminar Listas completas:')
		help_tree.insert(help_three, tk.END, text='Para copiar una lista, desde maestro seleccionar la lista a copiar')
		help_tree.insert(help_three, tk.END, text='Se carga la lista en los campos de edición arriba, cambiar el nombre')
		help_tree.insert(help_three, tk.END, text=' y pulsar CTRL-M o desde menú "edición/copiar lista"')
		help_tree.insert(help_three, tk.END, text='Para eliminar una lista basta con eliminar el item de lista desde maestro')
		# n - Ver hotkeys
		help_last = help_tree.insert('', tk.END, text='Para mas acciones vea el apartado "comandos rápidos" desde el menú ayuda')

	# Detalle de los HOTKEYS del programa
	def hotkeys(self):
		messagebox.showinfo('Accesos rápidos',  'CTRL+H: Ayuda\n'
												'CTRL+S: Abrir Base de datos\n'
												'CTRL+B: Crear base de datos\n\n'
												'Enter: Búsqueda por campo\n'
												'CTRL+Enter: limpiar búsqueda\n'
												'SpaceBar: Desplazar una hoja\n\n'
												'CTRL+A: Agregar registro\n'
												'CTRL+E: Editar registro\n'
												'CTRL+D: Eliminar registro\n'
												'CTRL+M: Copiar Lista\n\n'
												'CTRL+L: Abrir Listas (doble click)\n'
												'CTRL+O: Abrir Órdenes de trabajo\n'
												'CTRL+X: Emitir Órden de Compra\n'
												'CTRL+R: Abrir Rubros\n\n'
												'CTRL+P: Imprimir\n'
												'CTRL+T: Cambiar apariencia\n'
												'ALT+F4: Cerrar ventana', parent=self.window)

	# Info de licencia
	def license(self):  
		# funcion para ventana emergente que muestra un warning con icono warning
		messagebox.showinfo('GEST2020', 'Licencia válida para uso exclusivo de L.I.E. S.R.L.', parent=self.window)

	# Info del programa y versión
	def help_about(self):  # funcion para vent. emergente que muestra info con icono de info
		messagebox.showinfo('Gestor de artículos',
			f'GEST2020 Versión: {version}\n\nDesarrollado por Agustin Arnaiz', parent=self.window)


# Seleccion de THEME para ventana principal
class WindowTheme:
	window = None

	def __init__(self, win_to_mod, *args):
		if WindowTheme.window == None:
			WindowTheme.window = Toplevel()
			WindowTheme.window.attributes('-topmost', 'true')
			WindowTheme.window.iconbitmap(sys.argv[0])
			WindowTheme.window.resizable(False, False)
			WindowTheme.window.config(bg='grey')
			WindowTheme.window.protocol('WM_DELETE_WINDOW', self.exit_handler)
			self.win_to_mod = win_to_mod
			self.select_theme()

	def select_theme(self):	
		self.theme_sel = IntVar()
		# Crea la seleccion de THEMES
		for index in manager.theme_names:	 
			button = tk.Radiobutton(WindowTheme.window, 
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

	# permite una sola ventana a la vez
	@classmethod
	def exit_handler(cls):
		cls.window.destroy()
		cls.window = None


# Crea la base de datos SQLite3 desde los 3 XLS del GEST original
class CreateDatabase:
	conn = None

	@ProgManager.timer
	def __init__(self, *args):
		if messagebox.askquestion('Pregunta', 
			'¿Desea crear la base de datos ahora?\nEsta acción puede demorarse un momento') == 'yes':
			try:
				CreateDatabase.conn = sqlite3.connect(manager.configs.get('db_name'))
				self.maestro = CreateDatabase.format_table('MAESTRO')
				self.listas = CreateDatabase.format_table('LISTAS')
				self.ot = CreateDatabase.format_table('OT')
				self.rubros = CreateDatabase.format_table('RUBROS')
				CreateDatabase.create_triggers()
				CreateDatabase.conn.commit()
				CreateDatabase.conn.close()
				messagebox.showinfo('Info', 'Base de datos creada con éxito\n')
				inst_master()
			except Exception as err:
				messagebox.showerror('ERROR', f'Hubo un error: {err}!')
				manager.logger.warning(f'ERROR: {err}')

	# Abre un archivo Database y lo setea como default (path)
	@staticmethod
	def OpenDatabase(*args):
		path = filedialog.askopenfilename(title='Abrir Base de datos', filetypes=(
								('SQLite3', '*.db'), ('todos los archivos', '*.*')))
		if path:
			manager.configs["db_name"] = path
			manager.save_config()
			inst_master()
			
	# Crea los TRIGGERS
	@staticmethod
	def create_triggers():
		# Actualiza fecha de precio al editar el precio
		query = '''CREATE TRIGGER maestro_fecha_precio
			AFTER UPDATE ON "MAESTRO"
			WHEN old.PRECIO <> new.PRECIO
			BEGIN
				UPDATE "MAESTRO" SET "FECHA_PRECIO" = CURRENT_TIMESTAMP 
				WHERE "CÓDIGO" = new.CÓDIGO;
			END;'''
		CreateDatabase.conn.execute(query)
		# Trigger actualiza Monto de ordenes de trabajo al editar precio o cantidad
		query = '''CREATE TRIGGER ot_monto_update
			AFTER UPDATE ON "OT"
			WHEN old.CANT <> new.CANT
				OR old.PRECIO <> new.PRECIO
			BEGIN
				UPDATE "OT" SET "MONTO" = ROUND(new.PRECIO * new.CANT, 3) 
				WHERE "ORDEN" = new.ORDEN AND "CÓDIGO" = new.CÓDIGO;
			END;'''
		CreateDatabase.conn.execute(query)
		# Trigger actualiza monto de ordenes de trabajo al insertar un registro nuevo
		query = '''CREATE TRIGGER ot_monto_insert
			AFTER INSERT ON "OT"
			BEGIN
				UPDATE "OT" SET "MONTO" = ROUND(new.PRECIO * new.CANT, 3) 
				WHERE "ORDEN" = new.ORDEN AND "CÓDIGO" = new.CÓDIGO;
			END;'''
		CreateDatabase.conn.execute(query)

	# Abre xls de la tabla y genera la query para agregar a la db
	@staticmethod
	def format_table(name):
		try:
			df = pd.read_excel(name+'.xls')
		except Exception as err:
			messagebox.showwarning('Advertencia', f'No se encuentra el archivo {name}.xls\n{err}' )
			return
		query = ''
		drop_columns = []
		nom_columns = []
		# CREA LA TABLA, ELIMINA COLUMNAS NO DESEADAS Y CAMBIA EL NOMBRE DE LAS EXISTENTES
		if name == 'MAESTRO':
			# Lista para eliminar ciertas columnas no usadas por gest2020
			drop_columns = [6, 7, 9]  # cableado, stkmin, comprasug
			# Primero elimina decimales no deseados de columna d Precios
			df = CreateDatabase.delete_decimal(df, 'PRECIO1,N,10,3')
			df = CreateDatabase.delete_null(df, 'RUBRO,C,10')
			# creacion de tabla
			nom_columns = ["CÓDIGO", "DESCRIPCIÓN", "UN", "PRECIO", "FECHA_PRECIO", "FECHA_ALTA", "RUBRO"]
			query = 'CREATE TABLE "MAESTRO" ' \
					'("CÓDIGO" TEXT NOT NULL, ' \
					'"DESCRIPCIÓN" TEXT, ' \
					'"UN"	TEXT DEFAULT "UN", ' \
					'"PRECIO"	REAL NOT NULL DEFAULT "0", ' \
					'"FECHA_PRECIO"	TIMESTAMP DEFAULT CURRENT_TIMESTAMP, ' \
					'"FECHA_ALTA"	TIMESTAMP DEFAULT CURRENT_TIMESTAMP, ' \
					'"RUBRO"	TEXT, ' \
					'FOREIGN KEY("RUBRO") REFERENCES "RUBROS"("RUBRO") ON UPDATE CASCADE ON DELETE SET NULL,' \
					'PRIMARY KEY("CÓDIGO"))'				
		elif name == "LISTAS":
			# Primero elimina decimales no deseados de columna cantidad
			df = CreateDatabase.delete_decimal(df, 'CANT,N,10,3')
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
		elif name == "OT":
			# Primero elimina decimales no deseados de columna precio unitario y monto
			df = CreateDatabase.delete_decimal(df, 'PRUNIT,N,10,3')
			df = CreateDatabase.delete_decimal(df, 'MONTO,N,10,3')
			drop_columns = [3, 4, 5, 6, 7, 8, 11, 12] #fecha_venta, fecha_cum, can_cum, nombre_ fecha_ing, nom_ant, plan, despi
			nom_columns = ["ORDEN", "CÓDIGO", "CANT", "PRECIO", "MONTO"]
			query = 'CREATE TABLE "OT" ' \
					'("ORDEN"	INTEGER, ' \
					'"CÓDIGO"	TEXT, ' \
					'"CANT"	INTEGER DEFAULT 0, ' \
					'"PRECIO"	REAL DEFAULT 0, ' \
					'"MONTO"	REAL DEFAULT 0, ' \
					'FOREIGN KEY ("CÓDIGO") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE)'
		elif name == "RUBROS":
			nom_columns = ["RUBRO"]
			#agrega como ultimo elemento "SIN_RUBRO"
			df.loc[len(df.index)]="SIN_RUBRO" 
			query = 'CREATE TABLE "RUBROS" ' \
					'("RUBRO"	TEXT NOT NULL,' \
					'PRIMARY KEY("RUBRO"))'
		# ELIMINA COLUMNAS No deseadas
		df.drop(df.columns[drop_columns], axis=1, inplace=True)
		# CAMBIA NOMBRE DE LAS COLUMNAS
		df.columns = nom_columns
		CreateDatabase.copy_xls_db(df, name, query)

	# Elimina valores NULL de la columna
	@staticmethod
	def delete_null(df, col):
		df[col] = df[col].fillna('SIN_RUBRO')
		return df

	# Elimina decimales no significativos
	@staticmethod
	def delete_decimal(df, col):
		df[col] = df[col].astype(float)
		df[col] = df[col].round(3)
		return df

	# crea la tabla dentro de la base de datos
	@staticmethod
	def copy_xls_db(df, name, query):
		label = Label(root, text=f'copiando {name} ...')
		label.grid(row=0, column=0)
		root.update()
		cursor = CreateDatabase.conn.cursor()
		cursor.execute(query)
		try:	
			#copia los datos dentro de esa misma tabla
			df.to_sql(name=name, con=CreateDatabase.conn, if_exists='append', index=False) #if_exist='replace' resulta en falla PK FK etc
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'Ocurrió un error al querer agregar datos de la tabla {name} a la base de datos.\n{err}')
		finally:
			label.destroy()


# Clase principal de base de datos, maneja una tabla 
class ManageTable:
	# atributo de clase que impide la edicion de algunos campos si hay mas de un item seleccionado
	multiple_no_edit = [0,1]
	# atributo de clase que impide la edicion de algunos campos en todo momento
	no_edit_entry = []
	
	def __init__(self, window, table_name):
		self.window = window
		# Atributos principales de cada tabla/ventana
		self.table_name = table_name
		self.entry_array = np.array([])  # array de entrys editores de registro de database
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
		self.message = ttk.Label(self.frame_msg, text='')
		self.message.grid(row=0, column=10, columnspan=1, sticky=W)	
		# barra de progreso
		self.progress_bar = ttk.Progressbar(self.frame_msg, 
											orient=tk.HORIZONTAL, 
											mode="determinate",
											length=200,
											maximum=100,
											value=0)
		self.progress_bar.grid(row=0, column=0, columnspan=10, sticky=E)
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
		self.window.bind('<Button-3>', self.menu_popup)
		self.win = WindowConfig(self.window)
		# MENU archivo
		self.win.file_menu.add_separator()
		self.win.file_menu.add_command(label="Imprimir (CTRL+P)", command=self.prepare_to_print)
		# MENU edicion
		self.win.edit_menu.add_command(label="Agregar Registro (CTRL+A)", command=self.add_record)
		self.win.edit_menu.add_command(label="Editar Registro (CTRL+E)", command=self.edit_record)
		self.win.edit_menu.add_command(label="Eliminar Registro (CTRL+D)", command=self.delete_record)
		self.win.edit_menu.add_separator()
		self.win.edit_menu.add_command(label="Buscar.. (ENTER)", command=self.show_data)
		self.win.edit_menu.add_command(label="Limpiar Busqueda (CTRL+ENTER)", command=self.clean_entrys)

	# Identificacion para instancias de la clase
	def __repr__(self):
		return f'{self.table_name}("{self.window}", "{self.table_name}")'

	# Menu del boton derecho mouse
	def menu_popup(self, event):
		iid = self.tree.identify_row(event.y)
		if iid:
			# mouse pointer over item
			self.tree.selection_set(iid)
			self.win.edit_menu.tk_popup(event.x_root, event.y_root)
			self.win.edit_menu.grab_release()
		else:
			self.win.bar_menu.tk_popup(event.x_root, event.y_root)
			self.win.bar_menu.grab_release()

	# Armado de ventana, TREEVIEW adaptable segun Database
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
			sort_f = functools.partial(self.treeview_sort_column,'#'+str(index), False)
			self.tree.heading('#'+str(index), text=f'{self.table_columns[index]}', anchor=CENTER, command=sort_f)
			# crea tantos ENTRYS como columnas, para editar registro
			self.entry_array = np.append(self.entry_array, self.entrys(frame=self.frame_tree,
																		name=column,
																		row=7,
																		column=index,
																		width=largest*2))

	# Crea Entrys con label superior
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
			rubros = self.query_search(search='%', order='ASC')
			self.table_name = temp
			lista_rubros = [item for item in rubros]
			lista_rubros.sort()
			entry = ttk.Combobox(frame, values=lista_rubros, width=width, state='readonly')
		#todos los demas son en mayúscula
		else:
			entry = UpperEntry(frame, width=width)
		entry.grid(row=row + 2, column=column, columnspan=1, sticky=W+E, padx='5', pady=2)
		#define decodificacion de focus, para hacer busqueda por columnas segun el focus de entry
		self.focus_decoder[str(entry)] = column	
		return entry

	# Ejecuta una QUERY SQLite3 con cursor usando parametros
	def run_query(self, query, parameters={}, many=False):
		with sqlite3.connect(manager.configs.get('db_name')) as conn:
			try:
				conn.execute('PRAGMA foreign_keys = True')
				cursor = conn.cursor()
				if many:
					result = cursor.executemany(query, parameters)
				else:
					result = cursor.execute(query, parameters)
			except Exception as err:
				messagebox.showerror('ERROR', f'Ocurrió un error en la base de datos: \n{err}\n{query}\n{parameters}', parent=self.window)
				manager.logger.exception(f'error con:{err}-{query}-{parameters}')
			else:
				conn.commit()
				return result

	# Limpia el tree, y lo re-hace segun codido exacto o simil (like) y ordena segun column
	def show_data(self, *args, like='%', col_search=0, col_order=0, limit='', open=True):	
		# Al llamar la funcion con ENTER, hace busqueda de columna segun foco de entry
		if str(args).find('keysym=Return', 0, -1) != -1:     #(texto a buscar, inicio, final)
			focus = self.window.focus_get()
			if focus is not None:
				# Si el foco no se decodifica, por defecto busca en 0
				try:
					col_order = col_search = self.focus_decoder[str(focus)]
				except:
					col_order = col_search = 0
				# Si la busqueda se hace desde el tree y desde listas o OT, busca exacto y abre el display
				if self.table_name != 'Maestro' and self.table_name != 'RUBROS':
					# prohibe busquedas fuera de lista y codigo
					if col_search > 1:
						messagebox.showinfo('Advertencia','Solo se admite búsqueda por Lista y Código.', parent=self.window)
						return
					open = str(focus).find('treeview') != -1 or self.focus_decoder[str(focus)] != 0
					# actualiza el dato de busqueda, sea de lista o inversa
					self.message['text'] = f'{self.text_frame(True, col_search)}'
					if open:
						like=''
		# Habilita los no edit entry para cargar los datos y limpia el tree previamente
		self.hab_entry(self.no_edit_entry)
		self.hab_entry(self.multiple_no_edit)
		self.delete_tree()
		#BUSCA EN LA DB y copia en el TREE
		cursor = self.query_search(*args, search=self.entry_array[col_search].get(), like=like, col_search=col_search, col_order=col_order, limit=limit)
		self.data_into_tree(cursor, open)
		self.deshab_entry(self.no_edit_entry)

	# titulo a mostrar despues de cada busqueda, define datos de la ultima busqueda
	def text_frame(self, open=False, col_search=0):
		if open:
			return self.entry_array[col_search].get()

	# Ingresa los datos de busqueda en el TREE y los agrupa por listas de haberlas
	def data_into_tree(self, cursor, open):
		row_text='~_#~_#_' # un valor inexistente arbitrario
		for row in cursor:
			# si va abierto, no va indexado
			if open:
				self.tree.insert('', 0, text=row[0],  values=row[1:], open=True)
			else:
				# registra cada codigo nuevo y crea la linea para esa lista
				if row[0] != row_text:
					line = self.tree.insert('', 0, text=row[0], open=False)
					self.tree.insert(line, 0, text=row[0], values=row[1:])
					row_text = row[0]
				else:
					self.tree.insert(line, 0, text=row[0], values=row[1:])
	
	# Ejecuta query de busqueda en DB, retorna dataframe, se programó separado de "database_to_tree" para ser llamada por separado (por print_list y copy_list)
	def query_search(self, *args, search='', like='%', col_search=0, col_order=0, order='DESC', limit=''):
		# si existen las columnas, las usa para definir la busqueda
		if self.table_columns.size > 0:
			# Si hay segunda columna, ordena ASC
			try:
				col_search_two = f', "{self.table_columns[col_order+1]}" ASC '
			except:
				col_search_two = ''
			where_line = f' WHERE "{self.table_columns[col_search]}"'
			like_line = f' LIKE "{like}{search}{like}"'
			order_line = f' ORDER BY "{self.table_columns[col_order]}" {order}{col_search_two}'
		else:
			where_line = order_line = like_line = ''
		query = f'SELECT * from "{self.table_name}"{where_line}{like_line}{order_line} {limit}'
		return self.run_query(query)

	# valida el agregar un item
	def valid_add(self):
		# ni el item a agregar o el codigo de "lista" u "orden" debe estar vacio
		try:
			if self.entry_array[0].get() == '':
				raise Exception('El campo de código se encuentra vacío!')
			# checkea que el item no exista previamente en la database
			db_rows = self.query_search(search=self.entry_array[0].get(), like='')
			found = db_rows.fetchone()
			if found:
				raise Exception('El código que desea ingresar ya se encuentra en la base de datos')
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'No se pudo agregar registro:\n{err}', parent=self.window)
			return False
		else:
			return True

	# Borra el tree, para nueva visualizacion
	def delete_tree(self):
		records = self.tree.get_children()  # obtiene todos los datos de la tabla tree
		for element in records:
			self.tree.delete(element)  # limpia todos los datos de tree

	# forma la query para agregar item a tabla
	def add_query(self):
		nom_column, arg_query, parameters = [[],[],[]]
		for index, column in enumerate(self.table_columns):
			# Solo agrega si no se encuentra en "self.no_edit_entry" o si el entry no está vacío	
			if index not in self.no_edit_entry and self.entry_array[index].get() != '':
				nom_column.insert(index, '"'+column+'"')
				arg_query.insert(index, '?')
				parameters.insert(index, self.entry_array[index].get())	
		return nom_column, arg_query, parameters

	# agrega un registro en la base de datos
	def add_record(self, *args):
		if self.valid_add():
			nom_column, arg_query, parameters= self.add_query()			
			query = f'INSERT INTO "{self.table_name}" ({", ".join(nom_column)}) ' \
					f'VALUES({" ,".join(arg_query)})' # une la lista con join
			self.run_query(query, parameters)
			self.message['text'] = f'{maestro.entry_array[0].get()} ha sido guardado con éxito'
			fail_add = ''
		else:
			fail_add = '%'  # evita borrar el tree cuando se agrega algo vacio al buscar con LIKE
		self.show_data(like=fail_add, open=True)  # like = %: busca la db con "codigo%"
		return not fail_add  # devuelve True si agregó, False si no agregó registro

	# valida edicion de registro
	def valid_edit(self):
		self.edit_column = []
		self.selection = self.tree.selection()
		# condiciones básicas para editar:
		try:
			# debe haber algo seleccionado
			if len(self.selection) == 0:
				#Si hay un solo item en pantalla, lo selecciona por defecto
				if len(self.tree.get_children()) == 1:
					self.selection = self.tree.get_children()
				else:
					raise Exception('Debe seleccionar un item para editarlo')
			# Si se edita un solo item, debe tener código
			if self.entry_array[0].get() == '' and len(self.selection) == 1:
				raise Exception('El campo de código no puede estar vacío')
			elif len(self.selection) > 1:
				if messagebox.askquestion('Advertencia!', 
					'No se aconseja editar varios campos en simultáneo' \
					'\nNi cantidades muy grandes de items seleccionados \n' 
					'¿Desea continuar con la edición?') == 'yes':
					# para multiple edicion, checkea la(s) columna(s) que se modificó
					for index, column in enumerate(self.table_columns[1:], start=1):
						if self.entry_array[index].get() != self.tree.item(self.selection[-1])['values'][index-1]:
							self.edit_column.append(column)
				else:
					return False
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'{err}!', parent=self.window)
			return False
		else:
			return True

	# Edita registros en la base de datos
	def edit_record(self, *args):
		column_to_edit, column_check, parameters, parameters_prev = [[],[],[],[]]
		if self.valid_edit():
			# se habilitan los entrys para permitir su edición y se recorre item a item de la seleccion
			self.hab_entry(self.no_edit_entry)
			step = 100/len(self.selection)
			for index, iid in enumerate(self.selection, start=1):
				self.progress_bar['value'] = step * index
				column_to_edit, column_check, parameters, parameters_prev = [[],[],[],[]]
				# Edita los las columnas que no pertenecen a self.no_edit_entry
				item = self.tree.item(iid)['text']
				values = self.tree.item(iid)['values']
				for index, column in enumerate(self.table_columns):
					# No edita los campos no_edit o vacios
					if index not in self.no_edit_entry and self.entry_array[index].get() != '' and str(self.entry_array[index].get()) != 'None':
						# para editar un item luego, guarda el "COD y descr" o "lista y COD" o "ORDEN y COD"
						if index == 0:
							column_check.insert(index, f'"{column}" = ? ')
							parameters_prev.insert(index, item)
						elif index == 1 and values[index-1] != '' and values[index-1] != 'None':
							column_check.insert(index, f'"{column}" = ? ')
							parameters_prev.insert(index, values[index-1])
						column_to_edit.insert(index, f'"{column}" = ? ')
						# Si hay una sola seleccion, permite editar el codigo
						if len(self.selection) == 1:
							parameters.insert(index, self.entry_array[index].get())
						# si hay mas de una, para el codigo toma el original
						else:
							if index == 0:
								parameters.insert(index, item)
							elif index == 1:
								parameters.insert(index, values[index-1])
							# para los demas items, toma los entrys (con posible edicion)
							else:
								if str(self.edit_column).find(column) != -1:
									parameters.insert(index, self.entry_array[index].get())
								else:
									parameters.insert(index, values[index-1])
				# al final de los parametros a editar, se agregan los de check column
				query = f'UPDATE {self.table_name} ' \
						f'SET {", ".join(column_to_edit)} ' \
						f'WHERE {"AND ".join(column_check)}'
				parameters.extend(parameters_prev)
				self.run_query(query, parameters)
				try:
					self.message['text'] = f'Elemento editado: {item} {values[0]}'
				except:
					self.message['text'] = f'Elemento editado: {item}'
				self.window.update()
			self.progress_bar['value'] = 0
			self.show_data(like='', open=True)

	# Borra un registro en la base de datos
	def delete_record(self, *args):
		try:
			selection = self.tree.selection()
			if len(selection) == 0:
				raise Exception('Debe seleccionar un ítem para poder borrarlo.')
		except Exception as err:
			messagebox.showwarning('Atención','{err}', parent=self.window)
		else:
			step_bar = 100 / len(selection)
			for index, iid in enumerate(selection, start=1):
				self.progress_bar['value'] = step_bar * index
				item = self.tree.item(iid)
				# Elimina el registro con mismo código y descripción o misma Lista y elemento
				cod_selected = item['text']
				try:
					descr_selected = item['values'][0]
				#este artificio funciona para rubros que solo tiene una columna ['text'] y no la columna ['values']
				except:
					descr_selected = 'None'
				#para los escasos registros donde la descr es NULL, esto permite borrar esa entrada
				if descr_selected == 'None' or descr_selected == None:
					query = f'DELETE FROM "{self.table_name}" ' \
							f'WHERE "{self.table_columns[0]}" = ? '
					parameters = [cod_selected]
				else:
					query = f'DELETE FROM "{self.table_name}" ' \
							f'WHERE "{self.table_columns[0]}" = ? ' \
							f'AND "{self.table_columns[1]}" = ?'
					parameters = [cod_selected, descr_selected]
				self.run_query(query, parameters)
				self.message['text'] = f'Eliminado: {cod_selected} {descr_selected}'
				self.window.update()
			if len(selection) > 1:
				self.message['text'] = f'{len(selection)} registros han sido eliminados'
			self.progress_bar['value'] = 0
			self.show_data(like='', open=True)  # actualiza la tabla

	# Borra los entrys de registro
	def clean_entrys(self, *args):
		# Habiita todos los entrys despues de borrarlos, permite buscar por cada campo
		self.hab_entry(self.no_edit_entry)
		for index, entry_element in enumerate(self.entry_array):
			if str(entry_element) == ".!labelframe.!combobox":
				self.entry_array[index].set("")
			else:
				entry_element.delete(0, 'end')
		self.deshab_entry(self.no_edit_entry)
		# Hace foco en buscar registro por código
		self.entry_array[0].focus()

	# Carga el registro seleccionado de TREE en los entrys de edicion
	def load_edit_item(self, *args):
		# habilita los entrys para cargar los datos
		self.hab_entry(self.no_edit_entry+self.multiple_no_edit)
		# Carga solo cuando hay un item seleccionado y tiene algun value
		selection = self.tree.selection()
		# Carga en el array de entrys los valores de campos seleccionados
		for index, entry in enumerate(self.entry_array):
			# por el uso de TREE el indice cero se maneja por separado ['text']
			if index != 0:
				# En algunas vistas (rubro) no hay ['values'], por eso el try
				try:
					# maneja por separado el combobox de rubro, se setea diferente
					if str(entry).find('combobox') != -1:	
						entry.set(self.tree.item(selection[-1])['values'][index - 1])
					# Carga todos los values del tree
					else:
						entry.delete(0, 100)
						entry.insert(END, self.tree.item(selection[-1])['values'][index - 1])
				except:
					pass
			else:
				entry.delete(0, 50)
				entry.insert(END, self.tree.item(selection[-1])['text'])
		if len(selection) > 1:
			self.deshab_entry(self.multiple_no_edit)
			self.message['text'] = f'{len(selection)} items seleccionados'
		else:
			self.hab_entry(self.multiple_no_edit)
			self.message['text'] = f'{self.text_frame(True)}'
		self.deshab_entry(self.no_edit_entry)

	# contruye dataframe a partir de la vista del tree
	def prepare_to_print(self, *args):
		heading_list, list_data_tree = [], []
		heading_list.insert(0, '')
		heading_list.insert(0, f'Impreso desde {self.table_name}')
		heading_list.insert(0, f'Fecha de Impresión: {date.today()}')
		heading_list.insert(0, '')
		heading_list.insert(0, 'L.I.E. S.R.L. - Sistema de Administración de Producción GEST2020')
		for iid in self.tree.get_children():
			values = self.tree.item(iid)['values']
			if values == '':
				values = []
			values.insert(0, self.tree.item(iid)['text'])
			list_data_tree.append(values)
		try:
			df = pd.DataFrame(list_data_tree, columns=self.table_columns)
		except:
			messagebox.showwarning('ADVERTENCIA', 'No se puede imprimir. Filtre la búsqueda a una lista u órden de compra.', parent=self.window)
		else:
			df, from_obj, heading_list = self.build_print(df, heading_list)
			heading_df = pd.DataFrame(heading_list)
			#manda a imprimir
			ToPrinter(to_print=df, from_obj=from_obj, heading=heading_df)

	# edita el dataframe y heading segun cada tabla
	def build_print(self, df, heading_list):
		from_obj = self.table_name+'-'+self.entry_array[0].get()
		return df, from_obj, heading_list

	# deshabilita varios entrys para evitar su edicion
	def deshab_entry(self, deshab_entry=[]):
		for item in deshab_entry:
			self.entry_array[item].config(state='readonly')

	# habilita los entry para el ingreso de datos
	def hab_entry(self, hab_entry=[]):
		for item in hab_entry:
			self.entry_array[item].config(state='normal')

	# scroll del tree por página
	def auto_scroll(self, *args):
		#desplaza de a una hoja en la vista de tree con la barra 'space'
		self.tree.yview_scroll(1, what='page')

	# ordena el tree al hacer click en columnas, asc y desc
	def treeview_sort_column(self, col, reverse=False):
		# la columna 0 se programa diferente, sino no funciona
		if col == '#0':
			tree_col_id = [(self.tree.item(iid)["text"], 0,iid) for iid in self.tree.get_children()] #el 0 se usa solo para obtener 3 valores como el caso de otras columnas
		else:
			# col a ordenar + col de codigo 2do orden + iid del tree
			tree_col_id = [(self.tree.set(iid, col), self.tree.set(iid, '#1'), iid) for iid in self.tree.get_children('')]
		tree_col_id.sort(reverse=reverse)
		for index, (_,_, iid) in enumerate(tree_col_id):
			self.tree.move(iid, '', index)
		# invierte el orden para la siguente vez
		self.tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))


# Base de datos especifica para tabla Maestro
class Maestro(ManageTable):
	no_edit_entry = [4, 5]
	multiple_no_edit = [0]

	# agrega items al menu, bindings
	def __init__(self, window, table_name='MAESTRO'):
		self.window = window
		super().__init__(self.window, table_name)
		self.show_data()
		# Suma acciones al menú
		self.win.edit_menu.add_separator()
		self.win.edit_menu.add_command(label="Copiar Lista (CTRL+M)", command=lambda: Listas('keysym=m'))
		self.win.edit_menu.add_command(label="Eliminar Lista (CTRL+F)", command=lambda: Listas('keysym=f'))
		# solo desde MAESTRO abre una lista con doble click
		self.tree.bind('<Double-Button-1>', lambda l: Listas('ButtonPress', maestro.entry_array[0].get())) 
		self.window.bind('<Control-m>', Listas)  # copia lista con nuevo codigo
		self.window.bind('<Control-M>', lambda l: Listas('keysym=m'))  # en caso 'M', le pasa forzado 'm'

	# override para eliminar una columna de fecha de alta
	def build_print(self, df, heading_list):
		df.drop(df.columns[5], axis=1, inplace=True)
		heading_list.insert(2, f'Código: {self.entry_array[0].get()}')
		heading_list.insert(3, f'Descripción: {self.entry_array[1].get()}')
		heading_list.insert(4, f'Fecha de alta: {self.entry_array[5].get()}')
		from_obj = self.table_name+'-'+self.entry_array[0].get()
		return df, from_obj, heading_list

# Ventana simple para ver listas valorizadas sin imprimir
class VerListaValorizada(tk.Tk):
	def __init__(self, df, heading_list):
		self.df = df
		self.heading_list = heading_list
		tk.Tk.__init__(self)		
		max_length = max(len(str(row)) for row in df.values) + 10
		num_rows = df.shape[0] + 10 # 10 renglones del encabezado
		self.text_area = tk.Text(self, width=max_length, height=num_rows)  # Ajusta la altura según tus necesidades
		self.text_area.pack()
		self.write_lista()
		sys.stdout = self
		# Resto de la configuración de tu aplicación Tkinter

	def write_lista(self):
		for item in self.heading_list:		
			self.text_area.insert(tk.END, item+'\n')		
		self.text_area.insert(tk.END, self.df.to_string())
		

# Base de datos especifica para tabla listas
class Listas(ManageTable):
	window = None
	no_edit_entry = [4, 5, 6]

	def __init__(self, *args, table_name='LISTAS'):
		#define una ventana nueva para ver las listas
		Listas.window = Toplevel()
		super().__init__(Listas.window, table_name)
		# update carga bien tamaño de ventana, sino no toma el extra en Y del menu dado por super
		Listas.window.update_idletasks()
		Listas.window.state(manager.configs.get('fullscreenL'))
		Listas.window.geometry(manager.configs.get('geometryL'))
		Listas.window.protocol('WM_DELETE_WINDOW', self.exit_handler)
		self.deshab_entry(self.no_edit_entry)
		# si encuentra 'keysym=m' --> ejecuta copiar lista
		if str(args).find('keysym=m', 0, -1) != -1:     #(texto a buscar, inicio, final)
			self.copy_list()
		# si encuentra 'keysym=double click' --> ejecuta cargar lista
		elif str(args).find('ButtonPress', 0, -1) != -1:
			self.load_lista(*args)
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
		if col_search == 0:
			on_fk = 'CÓDIGO'
		else:
			on_fk = 'LISTA'
		query = f'SELECT {self.table_name}.*' \
				f', {maestro.table_name}.DESCRIPCIÓN, {maestro.table_name}.PRECIO, {maestro.table_name}.RUBRO ' \
				f'FROM "{self.table_name}" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON {self.table_name}.{on_fk} = {maestro.table_name}.CÓDIGO' \
				f'{where_line}{like_line}{order_line} {limit}'
		return self.run_query(query)

	# titulo a mostrar despues de cada busqueda, define datos de la ultima busqueda
	def text_frame(self, open=False, col_search=0):
		if open:
			result = maestro.query_search(search=self.entry_array[col_search].get(), limit='limit 1')
			self.lista_maestro_data = np.array(result.fetchone())
			return str(self.lista_maestro_data[[0,1,6]])

	# valida el agregar un item
	def valid_add(self):
		try:
			if self.entry_array[0].get() == '':
				raise Exception('El campo de Lista se encuentra vacío!')
			if self.entry_array[1].get() == '':
				if maestro.entry_array[0].get() == '':
					raise Exception('El campo de código (para agregar a la lista) se encuentra vacío, ' \
									'en Listas y en el Maestro de artículos')
				codigo = maestro.entry_array[0].get()
			else:
				codigo = self.entry_array[1].get()
			# checkea que el item exista previamente en la tabla maestro
			cursor_maestro = maestro.query_search(search=codigo, like='')
			found = cursor_maestro.fetchone()
			if not found:
				raise Exception('El código que desea agregar no se encuentra en el maestro de artículos')
			else:
				# checkea que el item no exista previamente en la tabla Listas
				query = 'SELECT LISTAS.LISTA, LISTAS.CÓDIGO ' \
						f'FROM LISTAS WHERE LISTA = "{self.entry_array[0].get()}" AND CÓDIGO = "{codigo}"'
				cursor_listas = self.run_query(query)
				found = cursor_listas.fetchone()
				if found:
					raise Exception('El código que desea ingresar ya se encuentra en la misma Lista. ' \
									'Puede editar la cantidad (CANT) del mismo dentro de la lista')
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'No se pudo agregar registro:\n{err}', parent=self.window)
			return False
		else:
			return True

	# override para listas, agrega los registros desde maestro
	def add_query(self, *args):
		# Borra la unidad del array para cargarla desde maestro de articulos
		self.entry_array[3].delete(0, 50)	
		# Carga los entrys con los datos desde maestro si no hay codigo en listas
		if self.entry_array[1].get() == '':
			self.entry_array[1].delete(0, 100)
			self.entry_array[1].insert(END, maestro.tree.item(maestro.tree.selection())['text']) # elemento
			self.entry_array[3].insert(END, maestro.tree.item(maestro.tree.selection())['values'][1]) # unidad
		else:
			cursor = maestro.query_search(search=self.entry_array[1].get(), like='')
			item_found = cursor.fetchone()
			self.entry_array[3].insert(END, item_found[2]) # unidad
		return super().add_query()

	# override: elimina NONE de CANT, ejecuta super() y borra el codigo de entry, para agregar otro luego (especial si se agrega desde maestro)
	def add_record(self, *args):
		# Si no se especifica una cantidad, le pone 0 (sino hay error al imprimir con valor None)
		if self.entry_array[2].get() == '' or  self.entry_array[2].get()== None:
			self.entry_array[2].insert(0, 0)
		super().add_record()
		self.entry_array[1].delete(0, 50)	

	# carga una lista segun el codigo en maestro con doble click
	def load_lista(self, *args):
		arg, codigo = args
		self.entry_array[0].delete(0, 50)
		self.entry_array[0].insert(END, codigo)
		#carga el código exacto sin like %
		self.show_data(like='', open=True)
		self.message['text'] = f'{self.text_frame(True)}'

	# Copia una lista con diferente código
	def copy_list(self, *args):
		old_code = maestro.tree.item(maestro.tree.selection())['text']
		new_code = maestro.entry_array[0].get()
		# crea el registro en maestro con nuevo código
		if maestro.add_record():
			# Busca en listas, la lista a copiar
			cursor = self.query_search(search=old_code, like='')
			arg_query = ['?' for column in range(len(self.table_columns)-len(self.no_edit_entry))]		
			parameters = [(new_code, row[1], row[2], row[3]) for row in cursor]
			query = f'INSERT INTO {self.table_name} ' \
					f'VALUES({", ".join(arg_query)})'
			# Ingresa masivos datos en database de listas
			self.run_query(query, parameters, many=True)
			self.load_lista(0, new_code)	# arg + codigo a cargar
			maestro.message['text'] = "Nueva lista creada con éxito"
		else:
			maestro.message['text'] = "No se pudo copiar la lista. Problemas al ingresar el código en Maestro."

	# override de listas, agrega valorizada, cantidad a imprimir y reordena
	def build_print(self, df, heading_list):
		window_print = Toplevel()
		window_print.attributes('-topmost', 'true')
		window_print.resizable(False, False)
		window_print.iconbitmap(sys.argv[0])
		window_print.config(bg='grey')
		frame_print = ttk.LabelFrame(window_print, text="Configuración de impresión")
		frame_print.grid(row=0, column=0, columnspan=5, pady=2, padx=2, sticky=W+E+S+N)
		self.multiplicador = 1
		self.valorizada = StringVar()
		Label(frame_print, text = 'Cantidad:').grid(row = 1, column = 0, sticky=W)
		entry_cant = ttk.Entry(frame_print, width=8)
		entry_cant.grid(row=1,column=1,sticky=W+E)
		entry_cant.insert(0, self.multiplicador)
		entry_cant.config(state='readonly')
		Label(frame_print, text = 'Valorizar: ').grid(row = 2, column = 0, sticky=W)
		sel_valorizada = ttk.Combobox(frame_print, 
									textvariable=self.valorizada,
									values=["Valorizada","No valorizada"], 
									width=15,
									state='readonly')
		sel_valorizada.grid(row=2,column=1, columnspan=3,sticky=W+E)
		sel_valorizada.set("Valorizada")
		boton_mas = ttk.Button(frame_print, text="+", command=lambda: cant_print('+'))
		boton_mas.grid(row=1,column=2)
		boton_menos = ttk.Button(frame_print, text="-", command=lambda: cant_print('-'))
		boton_menos.grid(row=1,column=3)
		boton_aceptar = ttk.Button(frame_print, text="Imprimir", command=lambda: window_print.destroy())
		boton_aceptar.grid(row=5,column=0,columnspan=5,sticky=S)
		self.ver_valorizada = False
		
		def ver_lista_valorizada():
			self.ver_valorizada = True			
			window_print.destroy()

		boton_ver = ttk.Button(frame_print, text="Ver", command=lambda: ver_lista_valorizada())
		boton_ver.grid(row=5,column=3,columnspan=5,sticky=S)	

		def cant_print(mas_menos):
			self.multiplicador = int(entry_cant.get())
			if mas_menos=='+':
				self.multiplicador+=1
			else:
				self.multiplicador-=1		
			if self.multiplicador < 1:
				self.multiplicador=1
			entry_cant.config(state='normal')
			entry_cant.delete(0,10)
			entry_cant.insert(0,self.multiplicador)
			entry_cant.config(state='readonly')
		
		# espera a que se termine la configuracion de impresion
		window_print.wait_window(window_print)
		cost_lista = 0
		#intercambia orden columnas: UN[2], DESCR[4], RUBRO[6] --> DESCR, RUBRO, UN
		re_order_cols = [column for column in self.table_columns[[0,1,4,3,6,5,2]]]		
		df = df.reindex(columns=re_order_cols)
		# Si hay mas de un item y son todos iguales, elimina columna LISTA por redundante
		cant_lista = len(pd.unique(df['LISTA']))
		cant_codigo = len(pd.unique(df['CÓDIGO']))
		# lista inversa
		if cant_lista > 1 and cant_codigo == 1:
			heading_list.insert(2, f'Listado inverso: {self.entry_array[1].get()}')
			from_obj = self.table_name+'-'+self.entry_array[1].get()
			df.drop(df.columns[1], axis=1, inplace=True)
		# lista común
		elif cant_lista == 1 and cant_codigo >= 1:
			heading_list.insert(2, f'Lista: {self.entry_array[0].get()} x {self.multiplicador}')
			from_obj = self.table_name+'-'+self.entry_array[0].get()
			df.drop(df.columns[0], axis=1, inplace=True)
		heading_list.insert(3, f'Descripción: {self.lista_maestro_data[1]}')
		heading_list.insert(4, f'Fecha de alta: {self.lista_maestro_data[5]}')
		# multiplica la cantidad de cada elemento
		if self.multiplicador!=1:
			df['CANT'] = df['CANT'].astype(float)
			df['CANT'] = df['CANT'] * self.multiplicador
		# si se imprime valorizada se crea la columna con el monto
		if self.valorizada.get()=="Valorizada":	
			df['PRECIO'] = df['PRECIO'].astype(float)
			df['CANT'] = df['CANT'].astype(float)
			df['MONTO'] = df['PRECIO'] * df['CANT']
			suma = df.sum(numeric_only=True)
			cost_lista = suma['MONTO']
			columns = np.append(re_order_cols,'MONTO')	
			heading_list.insert(6, f'Costo Total:  u$s {round(cost_lista, 3)}')
		else:
			# Elimina columna de PRECIO
			df.drop(df.columns[4], axis=1, inplace=True)

		if self.ver_valorizada:			
			VerListaValorizada(df, heading_list)

		return df, from_obj, heading_list

	# manejo manual del cierre de ventana (impide que se abra mas de una ventana)
	def exit_handler(self):
		manager.configs['geometryL'] = self.window.geometry()
		manager.configs['fullscreenL'] = self.window.state()
		self.window.destroy()


# Base de datos especifica para ORDENes de trabajo
class OrdenTrabajo(ManageTable):
	window = None
	no_edit_entry = [4, 5, 6]

	def __init__(self, *args, table_name='OT'):
		self.table_name = table_name
		if OrdenTrabajo.window == None:
			OrdenTrabajo.window = Toplevel()
			super().__init__(OrdenTrabajo.window, table_name)
			# con update carga bien tamaño de ventana, sino no toma el extra en 'Y' del 'menu' dado por super()
			OrdenTrabajo.window.update()
			OrdenTrabajo.window.state(manager.configs.get('fullscreenO'))
			OrdenTrabajo.window.geometry(manager.configs.get('geometryO'))
			OrdenTrabajo.window.protocol('WM_DELETE_WINDOW', self.exit_handler)
			OrdenTrabajo.window.bind('<Control-X>', self.orden_compra)
			OrdenTrabajo.window.bind('<Control-x>', self.orden_compra)
			self.win.edit_menu.add_separator()
			# simula doble click, para abrir la lista determinada
			self.win.edit_menu.add_command(label="Abrir Lista (CTRL-L)", command=lambda: Listas('ButtonPress',self.entry_array[1].get()))
			self.win.tools_menu.add_separator()
			self.win.tools_menu.add_command(label="Generar lista de compra (CTRL-X)", command=self.orden_compra)
			# re define binding ctrl-l para abrir lista determinada desde OT
			self.window.bind('<Control-l>', lambda l: Listas('ButtonPress', self.entry_array[1].get()))
			self.window.bind('<Control-L>', lambda l: Listas('ButtonPress', self.entry_array[1].get()))
			self.show_data(open=False)

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

	# Prepara la órden de compra
	def orden_compra(self, *args):
		try:
			if self.entry_array[0].get() == '':
				raise Exception('Debe seleccionar una "órden de trabajo", para generar la "órden de compra"')
			else:
				self.show_data(like='')
		except Exception as err:
			messagebox.showwarning('AVISO', f'{err}', parent=self.window)
		else:
			step = 100 / len(self.tree.get_children())
			for index, iid in enumerate(self.tree.get_children(), start=1):
				self.progress_bar['value'] = step * index
				self.window.update()
				cod_lista = self.tree.item(iid)['values'][0]
				cursor = self.search_lista(cod_lista)
				lista = cursor.fetchall()
				for item in lista:
					item = list(item) # formato lista
					# multiplica la cantidad del elemento por la cantidad de la lista
					item[1] = round(float(item[1])*float(self.tree.item(iid)['values'][1]), 3)
					# calcula el monto en base a la cantidad y precio
					item.insert(3, round(item[1]*item[2], 3))
					# si el código ya existe en la compra lo identifica en el tree, y solo suma a la cantidad y su monto
					found = False
					for id2 in self.tree.get_children():
						if item[0] == self.tree.item(id2)['values'][0]:
							found = True
							id3 = id2
					if found:
						tree_item = self.tree.item(id3)
						tree_item['values'][1] = round(float(item[1])+float(self.tree.item(id3)["values"][1]), 3)
						tree_item['values'][3] = round(float(item[3])+float(self.tree.item(id3)["values"][3]), 3)
						self.tree.item(id3, values=tree_item['values'])
					else:
						self.tree.insert('', 0, text=self.entry_array[0].get(), values=item)
			self.progress_bar['value'] = 0
			self.frame_tree['text'] = f'Órden de compra: {self.entry_array[0].get()}'
			self.treeview_sort_column('RUBRO', False)
			messagebox.showinfo('AVISO', 'Órden de compra finalizada.', parent=self.window)
				
	# query search de listas inner join maestro para COMPRAS
	def search_lista(self, search=''):
		query = f'SELECT LISTAS.CÓDIGO, LISTAS.CANT' \
				f', {maestro.table_name}.PRECIO, {maestro.table_name}.DESCRIPCIÓN, {maestro.table_name}.RUBRO ' \
				f'FROM "LISTAS" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON "LISTAS".CÓDIGO = {maestro.table_name}.CÓDIGO' \
				f' WHERE "LISTAS"."LISTA" = "{search}"' \
				f' ORDER BY "LISTA" DESC, LISTAS.CÓDIGO DESC'
		return self.run_query(query)

	# valida el agregar un item
	def valid_add(self):
		try:
			if self.entry_array[0].get() == '':
				raise Exception('El campo de ORDEN se encuentra vacío!')
			if self.entry_array[1].get() == '':
				if maestro.entry_array[0].get() == '':
					raise Exception('El campo de código se encuentra vacío, ' \
									'(en órdenes de trabajo y en el Maestro de artículos)')
				codigo = maestro.entry_array[0].get()
			else:
				codigo = self.entry_array[1].get()
			# checkea que el item exista previamente en la tabla maestro
			cursor_maestro = maestro.query_search(search=codigo, like='')
			found = cursor_maestro.fetchone()
			if not found:
				raise Exception('El código que desea agregar no se encuentra en el maestro de artículos')
			else:
				# checkea que el item no exista previamente en la tabla Listas
				query = f'SELECT OT.{self.table_columns[0]}, OT.{self.table_columns[1]} ' \
						f'FROM "OT" WHERE "ORDEN" = "{self.entry_array[0].get()}" AND CÓDIGO = "{codigo}"'
				cursor_listas = self.run_query(query)
				found = cursor_listas.fetchone()
				if found:
					raise Exception('El código que desea ingresar ya se encuentra en la misma Lista. ' \
									'Puede editar la cantidad (CANT) del mismo dentro de la lista')
		except Exception as err:
			messagebox.showwarning('ADVERTENCIA', f'No se pudo agregar registro:\n{err}', parent=self.window)
			return False
		else:
			return True

	# override para OT, agrega los registros desde maestro si COD = ''
	def add_query(self, *args):
		# si no se agrega un codigo manualmente, busca el item seleccionado en maestro
		if self.entry_array[1].get() == '':
			# Carga los entrys los datos desde maestro
			self.entry_array[1].delete(0, 100)
			self.entry_array[3].delete(0, 50)
			self.entry_array[4].delete(0, 50)
			self.entry_array[1].insert(END, maestro.tree.item(maestro.tree.selection())['text']) # codigo
			self.entry_array[3].insert(END, maestro.tree.item(maestro.tree.selection())['values'][2]) # precio
		else:
			cursor = maestro.query_search(search=self.entry_array[1].get(), like='')
			item_found = cursor.fetchone()
			self.entry_array[3].delete(0, 50)
			self.entry_array[3].insert(END, item_found[3]) # precio
		return super().add_query()

	# encabezado para OT, varias impresiones
	def build_print(self, df, heading_list):
		window = Toplevel()
		window.resizable(False, False)
		window.iconbitmap(sys.executable)
		frame = ttk.LabelFrame(window, text='Rubros a imprimir:', labelanchor=N)
		frame.grid(row=0, column=0)
		# carga lista de rubros en orden de compra
		cant_rubros = df['RUBRO'].unique()
		rubro_var = []
		for index, rubro in enumerate(cant_rubros):
			rubro_var.insert(index, StringVar())
			check = ttk.Checkbutton(frame, text=rubro, onvalue=0, offvalue=rubro , variable=rubro_var[index])
			check.grid(row=index, column=0, sticky=W)
			rubro_var[index].set(0)
		sel_all = ttk.Button(frame, text='Seleccionar todos', command=lambda: seleccion()).grid(row=index+1, column=0, sticky=N, pady=5)
		desel_all = ttk.Button(frame, text='Deseleccionar todos', command=lambda: deseleccion()).grid(row=index+2, column=0, sticky=N, pady=5)
		accept = ttk.Button(frame, text='Aceptar', command=window.destroy, style = "Bold.TButton")
		accept.grid(row=index+3, column=0, sticky=N, pady=5)
		accept.focus()
		def seleccion():
			for index, rubro in enumerate(cant_rubros):
				rubro_var[index].set(0)
		def deseleccion():
			for index, rubro in enumerate(cant_rubros):
				rubro_var[index].set(rubro)
		# espera a que se haga la seleccion de rubros a imprimir y se pulse aceptar
		window.wait_window(window)
		# Elimina los rubros no seleccionados
		for index, item in enumerate(rubro_var):
			if item.get() != 0:
				df = df.drop(df[df['RUBRO'] == item.get()].index)
		cost_order = 0
		df['PRECIO'] = df['PRECIO'].astype(float)
		df['CANT'] = df['CANT'].astype(float)
		df['MONTO'] = df['MONTO'].astype(float)
		# solo computa la suma de columnas numericas
		suma = df.sum(numeric_only=True)
		cost_order = suma['MONTO']
		heading_list.insert(2, f'Orden de trabajo: {self.entry_array[0].get()}')
		heading_list.insert(5, f'Costo Total:  u$s {round(cost_order, 3)}')
		#intercambia orden columnas
		df = df.reindex(columns=df.columns[[0,1,5,6,2,3,4]])
		# elimina columna de orden de trabajo
		df.drop(df.columns[0], axis=1, inplace=True)
		from_obj = self.table_name+'-'+self.entry_array[0].get()
		return df, from_obj, heading_list

	# manejo manual del cierre de ventana (impide que se abra mas e una ventana)
	@classmethod
	def exit_handler(cls):
		# resta 20 pixels a Y al cerrar la ventana
		manager.configs['geometryO'] = cls.window.geometry()
		manager.configs['fullscreenO'] = cls.window.state()
		cls.window.destroy()
		cls.window = None


# Clase para manejar los rubros
class Rubros(ManageTable):
	window = None
	multiple_no_edit = [0]

	def __init__(self, *args, table_name='RUBROS'):
		self.table_name = table_name
		if Rubros.window == None:
			Rubros.window = Toplevel()
			super().__init__(Rubros.window, table_name)
			self.show_data()
			Rubros.window.protocol('WM_DELETE_WINDOW', self.exit_handler)

	# solo permite una ventana de Rubros por vez, al salir actualiza la listas de rubros al maestro
	def exit_handler(self):
		rubros = self.query_search(search='%', order='ASC')
		lista = [item for item in rubros]
		maestro.entry_array[6]['values'] = lista
		Rubros.window.destroy()
		Rubros.window = None


# Manejo de Impresora
class ToPrinter:

	def __init__(self, to_print, from_obj='GEST2020-print', heading=''):
		self.window = Toplevel()
		self.window.iconbitmap(sys.executable)
		self.window.resizable(False, False)
		self.df = to_print
		self.heading = heading
		self.from_obj = from_obj
		self.default_printer = win32print.GetDefaultPrinter()
		self.printer = None
		self.level = 7
		self.select_printer()

	# lee las impresoras del sistema y permite elegir a cual imprimir
	def select_printer(self):
		printers = win32print.EnumPrinters(self.level)
		lista_printers = [p[2] for p in printers]
		printers_menu = ttk.Combobox(self.window, values=lista_printers, width=max(len(l) for l in lista_printers), state='readonly')
		printers_menu.set(self.default_printer)
		printers_menu.pack(side=TOP)
		print_button = ttk.Button(self.window, text='Imprimir', command=lambda: self.send_print(printers_menu.get()), style = "Bold.TButton")
		print_button.pack(side=BOTTOM)
		print_button.focus()
		self.window.bind('<Return>', lambda p: self.send_print(printers_menu.get()))  # usa ENTER para aceptar el envío a la impresora

	# moldea la data y envia impresion a printer usando printer device context de win32ui
	def send_print(self, printer, *args):
		self.window.destroy()
		self.printer = printer
		# Handler Device Context
		hDC = win32ui.CreateDC()
		# Handler Printer Device Context
		hDC.CreatePrinterDC(self.printer)
		hDC.StartDoc(self.from_obj)
		# seteo del modo MAP para escalar la salida, de este modo la escala Y baja de hoja con numero negativo
		hDC.SetMapMode(win32con.MM_TWIPS) # 1440 per inch
		scale_factor = 20 		# i.e. 20 twips to the point
		margin = 40
		# font Lucida Console 10 point.
		font = win32ui.CreateFont({"name": "Lucida Console", "height": int(scale_factor * 8), "weight": 400,})
		hDC.SelectObject(font)
		# Con tabulate justify a la derecha, otros métodos no funcionaron con pandas
		self.heading = tabulate(self.heading, showindex=False, headers=self.heading.columns)
		self.df = tabulate(self.df, showindex=False, headers=self.df.columns)
		multi_line_heading = self.heading.split('\n')
		multi_line_df = self.df.split('\n')
		hDC.StartPage()
		renglon = 0
		for index, row in enumerate(multi_line_df[2:]):
			if renglon == 0:
				# imprime encabezado en cada pagina
				for ind, line in enumerate(multi_line_heading[1:]):
					hDC.TextOut(margin * scale_factor, -1 * scale_factor * margin -(ind*scale_factor*10) , line)
				# agrega nombres de columnas a cada página y linea de puntos
				hDC.TextOut(margin * scale_factor, -1 * scale_factor * margin -((ind+1)*scale_factor*10) , multi_line_df[0])
				hDC.TextOut(margin * scale_factor, -1 * scale_factor * margin -((ind+2)*scale_factor*10) , multi_line_df[1])
			hDC.TextOut(margin * scale_factor, -1 * scale_factor * margin -((renglon+ind+3)*scale_factor*10), row)
			renglon += 1
			# Cambio de página
			if (index+1)%(scale_factor*3) == 0:
				hDC.EndPage()
				hDC.StartPage()
				renglon = 0
		hDC.EndPage()
		hDC.EndDoc()


##-------------------MAIN---INSTANCIA MAESTRO-------------------------##
if __name__ == '__main__':
	# Inicia el asistente de programa (logging, y variables de config)
	manager = ProgManager()
	manager.logger.debug('Start:')
	# Carga la configuración desde archivo
	manager.load_config()
	root = ThemedTk(theme=manager.configs['theme_name'])
	root.state(manager.configs['fullscreen'])
	root.geometry(manager.configs['geometry'])
	root.protocol("WM_DELETE_WINDOW", manager.exit_handler)
	main_window = WindowConfig(root, title='GEST2020 | Sistema de Administración de Producción L.I.E. S.R.L.')
	def inst_master():
		# si existe la base de datos la instancia
		if os.path.isfile(manager.configs['db_name']):
			# genera backups de la base de datosbackup_max
			manager.file_backup(manager.configs['db_name'])
			global maestro #variable global maestro de articulos database
			maestro = Maestro(window=root, table_name='Maestro')
			
		else:
			messagebox.showwarning('Advertencia', 'No se encuentra la base de datos:\n\n'
						'Puede abrir una desde menú: "Archivo/Abrir base de datos".\n'
						'o puede crear una nueva desde "Herramientas/crear base de datos" ' \
						'(requiere que existan previamenta los archivos *.xls de gest: MAESTRO, LISTAS, OT y RUBROS).')
	inst_master()
	root.mainloop()
##------------------------ END OF CODE -------------------------------##