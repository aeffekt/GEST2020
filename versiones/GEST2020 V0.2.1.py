''' GEST 2020: Gestor de compras y materiales para pequeñas empresas basado en software GEST;
diseñado para la empresa LIE SRL. Por Agustin Arnaiz
* Salvo la clase UpperEntry del usuario FJSevilla de stackoverflow: https://es.stackoverflow.com/questions/356082/como-lograr-poner-mayúscula-en-los-campos-de-tipo-entry-y-formatear-números-en-p'''

import sqlite3                  # SQL, manejo base de datos
import pandas as pd             # Dataframes y XLS files
import numpy as np				# se usa para listas de datos, acceso no consecutivo a sus items
import os                       # usa la impresión de shell "print"
import tkinter as tk
from tkinter import *
from tkinter import messagebox  # mensajes de salida
from tkinter import ttk         # tkinter mas facha
from ttkthemes import ThemedTk	# themes :D
from datetime import date       # fecha
from datetime import datetime	# fecha con hora, now

#-----------------------------CLASES--------------------------------

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
			for row in range(8, len(self.to_print[:20])):
				item = self.to_print.loc[row, column]
				max_len = max(max_len, len(str(item)))
				worksheet.set_column(column, column, max_len)
		try:
			self.file_print.save()
			#os.startfile(path, "print")     #envia el archivo a impresora default
			messagebox.showinfo('Mensaje', 'Archivo XLS guardado con éxito')
		except:
			messagebox.showerror('Advertencia', 'No se pudo guardar el archivo')


#Seleccion de THEME para ventana principal
class WindowTheme:
	def __init__(self, *args, win_to_mod):
		self.window = Toplevel()
		self.win_to_mod = win_to_mod
		self.theme = IntVar()
		# clearlooks, elegance, plastik, radiance, black, blue, breeze, equilux, yaru
		Radiobutton(self.window, text='Vista clara (default)', variable=self.theme, value=1, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Fondo Gris', variable=self.theme, value=2, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Fondo Blanco', variable=self.theme, value=3, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Tamaño Letra Grande', variable=self.theme, value=4, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Tamaño letra Mediana', variable=self.theme, value=5, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Fondo negro Bajo contraste', variable=self.theme, value=6, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Fondo Negro Alto contraste', variable=self.theme, value=7, command=self.window_theme).grid(sticky=W)
		Radiobutton(self.window, text='Fondo Azul', variable=self.theme, value=8, command=self.window_theme).grid(sticky=W)

	def window_theme(self):
		theme = {1: 'clearlooks', 2: 'elegance', 3: 'plastik', 4: 'radiance', 5: 'breeze', 6: 'equilux', 7: 'black', 8: 'blue'}
		self.win_to_mod.config(theme=theme[self.theme.get()])
		self.window.destroy()


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


#Crea la base de datos SQLite3 desde los 3 XLS del GEST original
class CreateDatabase:
	def __init__(self, *args):
		self.windows = Toplevel
		self.db_name = sqlite3.connect('db_gest2020.db')
		try:
			self.maestro = self.format_table('MAESTRO')
			self.listas = self.format_table('LISTAS')
			self.ot = self.format_table('OT')
			self.rubros = self.format_table('RUBROS')
			self.db_name.commit()
			self.db_name.close()
			messagebox.showinfo('Info', 'Base de datos creada con éxito')
			messagebox.showwarning('ADVERTENCIA',
							   f'Debe reiniciar el programa para leer la base de datos')
		except:
			messagebox.showerror('Error', 'No se pudo crear la base de datos')
		self.windows.destroy

	#Abre xls y genera la query para la tabla de db
	def format_table(self, name):
		try:
			xls_file = pd.read_excel(name+'.xls')
		except:
			messagebox.showwarning('Advertencia', f'No se encuentra el archivo {name}.xls' )
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
					'"UN."	TEXT, ' \
					'"PRECIO"	REAL, ' \
					'"FECHA_PRECIO"	TIMESTAMP, ' \
					'"FECHA_ALTA"	TIMESTAMP, ' \
					'"RUBRO"	TEXT, ' \
					'FOREIGN KEY("RUBRO") REFERENCES "RUBROS"("RUBRO") ON UPDATE CASCADE ON DELETE SET NULL,' \
					'PRIMARY KEY("CÓDIGO"))'

		if name == "LISTAS":
			# Primero elimina decimales no deseados de columna cantidad
			xls_file = self.delete_decimal(xls_file, 'CANT,N,10,3')

			#lista para renombrar las columnas
			nom_columns = ["CÓDIGO", "COMPONENTE", "CANT.", "UN."]

			#creacion de la tabla
			query = 'CREATE TABLE "LISTAS" ' \
					'("CÓDIGO"	TEXT NOT NULL,' \
					'"COMPONENTE"	TEXT,' \
					'"CANT."	REAL,' \
					'"UN."	TEXT,' \
					'FOREIGN KEY("CÓDIGO") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE,' \
					'FOREIGN KEY("COMPONENTE") REFERENCES "MAESTRO"("CÓDIGO") ON UPDATE CASCADE ON DELETE CASCADE)'

		if name == "OT":
			# Primero elimina decimales no deseados de columna precio unitario y monto
			xls_file = self.delete_decimal(xls_file, 'PRUNIT,N,10,3')
			xls_file = self.delete_decimal(xls_file, 'MONTO,N,10,3')

			drop_columns = [3, 4, 5, 6, 7, 8, 11, 12] #fecha_venta, fecha_cum, can_cum, nombre_ fecha_ing, nom_ant, plan, despi
			nom_columns = ["ORDEN", "CÓDIGO", "CANT.", "PRECIO", "MONTO"]
			query = 'CREATE TABLE "OT" ' \
					'("ORDEN"	INTEGER,' \
					'"CÓDIGO"	TEXT NOT NULL,' \
					'"CANT."	INTEGER,' \
					'"PRECIO"	REAL,' \
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
		cursor = self.db_name.cursor()
		print('copiando ', name, '...')
		try:
			cursor.execute(query)
			#copia los datos dentro de esa misma tabla
			xls_file.to_sql(name=name, con=self.db_name, if_exists='append', index=False) #if_exist='replace' resulta en falla PK FK etc
		except:
			messagebox.showwarning('ADVERTENCIA', f'Ocurrió un error al querer agregar la tabla {name} a la base de datos')


#Clase principal de base de datos (auto arma panel, manejo CRUD)
class Database:
	def __init__(self, window, table_name):
		self.window = window
		self.window.geometry("+100+5")  # posicion inicial de ventana ("+x +y")
		self.db_name = "db_gest2020.db"
		self.table_name = table_name
		self.entry_array = []

		#define nombres de COLUMNAS de la database
		self.sheet_columns = np.array([]) #usa numpy para citar items no consecutivos (al final no esta en uso)
		self.focus_deco = {}
		self.read_columns()

		# define FRAMEWORK tabla de datos
		self.frame_tree = LabelFrame(self.window, text=f'Editor de {self.table_name}')
		self.frame_tree.grid(row=0, column=0, columnspan=20, pady=10, padx=10, sticky=W+E+S+N)
		self.frame_tree.config(cursor='hand2')  #indica seleccion de los elementos del tree

		#permite expandir los widgets internos a la ventana
		for index in range(len(self.sheet_columns)):
			self.frame_tree.columnconfigure(index, weight=1, minsize=50)

		# define FRAMEWORK mensajes
		self.frame_msg = LabelFrame(self.window, text='')
		self.frame_msg.grid(row=10, column=0, columnspan=20, pady=10, padx=10, sticky=W+E+S+N)

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
		self.window.bind('<Control-A>', self.add_record)
		self.window.bind('<Control-e>', self.edit_record)
		self.window.bind('<Control-E>', self.edit_record)
		self.window.bind('<Control-d>', self.delete_record)
		self.window.bind('<Control-D>', self.delete_record)
		self.window.bind('<Control-p>', self.prepare_to_print)
		self.window.bind('<Control-P>', self.prepare_to_print)
		self.tree.bind('<space>', self.auto_scroll)     #baja de a una hoja la vista de tree
		self.tree.bind('<<TreeviewSelect>>', self.load_edit_item)   #carga datos seleccionados

		# Instancia RUBROS con CTRL+R
		self.window.bind('<Control-r>', Rubros)
		self.window.bind('<Control-R>', Rubros)

		# Instancia OT con CTRL+O
		self.window.bind('<Control-o>', Ordenes_trabajo)
		self.window.bind('<Control-O>', Ordenes_trabajo)

	#desplaza de a una hoja en la vista de tree con la barra 'space'
	def auto_scroll(self, *args):
		self.tree.yview_scroll(1, what='page')

	#Armado de ventana, TREEVIEW adaptable segun Database
	def build_main_view(self):
		# creacion de tabla para visualizar
		self.tree = ttk.Treeview(self.frame_tree, height=25, columns=len(self.sheet_columns))
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

		#Esta query es para dar TAMAÑO DINAMICO al ancho de columnas del tree con 400 datos por columna
		query = f'SELECT * from "{self.table_name}" LIMIT 1000'
		cursor = self.run_query(query)

		#se cargan titulos y tamaños de las columnas del tree, asi como Entrys para la edicion de records
		num_column = 0
		item_table = ''
		all_table = cursor.fetchall()

		for each_column in self.sheet_columns:
			largest = 2
			for row_table in all_table:
				item_table = str(row_table[num_column])
				if largest < len(item_table):
					largest = len(item_table)
			self.tree.column("#" + str(num_column), width=largest*9, stretch=True)
			self.tree.heading('#' + str(num_column), text=f'{nombres_columnas[num_column]}', anchor=CENTER)

			# crea tantos ENTRYS como columnas, para editar registro
			self.entry_array.append(self.entrys(frame=self.frame_tree,
												name=each_column, row=7, column=num_column, width=largest+4))
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

		entry.grid(row=row + 2, column=column, columnspan=1, sticky=W+E, padx='10')

		#define decodificacion de focus, para hacer busqueda por columnas segun el focus de entry
		self.focus_deco[str(entry)] = column

		return entry

	#Borra los entrys de registro
	def clean_entrys(self, *args):
		for entry_element in self.entry_array:
			entry_element.delete(0, 'end')
		self.database_to_tree()

	#Lee los nombres de las columnas de la database
	def read_columns(self):
		query = f'SELECT * FROM "{self.table_name}" LIMIT 1'
		columns = self.run_query(query)
		for columna in columns.description:
			self.sheet_columns = np.append(self.sheet_columns, columna[0])

	#Ejecuta una QUERY SQLite3 con cursor usando parametros
	def run_query(self, query, parameters={}):
		with sqlite3.connect(self.db_name) as conn:
			conn.execute('PRAGMA foreign_keys = True')  # habilita las constraints de las FK, por defecto = False
			cursor = conn.cursor()
			#try:
			result = cursor.execute(query, parameters)
			conn.commit()
			return result
			#except:
			#	messagebox.showerror('ERROR', 'Ocurrió un error en la base de datos')

	#Copia datos masivos pasados por parametro (la usa listas y OT)
	def run_query_many(self, query, parameters={}):
		with sqlite3.connect(self.db_name) as conn:
			cursor = conn.cursor()
			result = cursor.executemany(query, parameters)
			conn.commit()
		return result

	#Limpia el tree, y lo re-hace segun codido exacto o simil (like) y ordena segun column
	def database_to_tree(self, *args, like='%', col_search=0, col_order=0):

		#Si se llama a la funcion con ENTER, hace busqueda de columna segun foco de entry
		if str(args).find('keysym=Return', 0, -1) != -1:     #(texto a buscar, inicio, final)
			#se fija donde esta el foco de entry, para hacer la busqueda segun esa columna
			focus = self.window.focus_get()
			if focus is not None:
				try:
					col_order = col_search = self.focus_deco[str(focus)]
				except:
					col_order = col_search = 0

		self.delete_tree()

		#BUSCA EN LA DB y copia en el TREE
		cursor = self.query_search(*args, like=like, col_search=col_search, col_order=col_order)
		for row in cursor:
			self.tree.insert('', 0, text=row[0], values=row[1:])

		# Hace foco en buscar registro por código
		self.entry_array[0].focus()

	#Ejecuta query de busqueda en DB, retorna dataframe, se programó separado de "database_to_tree" para ser llamada por separado (por print_list y copy_list)
	def query_search(self, *args, like='%', col_search=0, col_order=0, order='DESC'):
		query = f'SELECT * from "{self.table_name}" ' \
				f'WHERE "{self.sheet_columns[col_search]}" ' \
				f'LIKE "{self.entry_array[col_search].get()}{like}" ' \
				f'ORDER BY "{self.sheet_columns[col_order]}" {order}'
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

	# agrega un registro en la base de datos
	def add_record(self, *args):
		if self.valid_add():
			arg_query = ''
			parameters = []
			item = 0	#usa item, solo para los entrys (pertenenecen a columnas visibles)
			for index in range(len(self.sheet_columns)):
				arg_query += '?'
				parameters.append(self.entry_array[index].get())
			query = f'INSERT INTO {self.table_name} ' \
					f'VALUES({" ,".join(arg_query)})' #une la lista con join

			self.run_query(query, parameters)
			self.message['text'] = 'El Registro ha sido guardado con éxito'
			fail_add = ''
		else:
			messagebox.showwarning('Advertencia', 'El Registro ya existe o se encuentra vacío')
			fail_add = '%'  # evita borrar el tree cuando se agrega algo vacio al buscar con LIKE

		self.database_to_tree(like=fail_add)  # like = %: busca la db con "codigo%"
		return not fail_add  # devuelve True si agrego, false si no agrego registro

	#Edita un registro en la base de datos
	def edit_record(self, *args):
		if self.valid_selection() and self.entry_array[0].get() != '':
			query_text_column = ''
			query_text_item = ''
			parameters = []
			param_anterior = []

			#CREA la query y parametros segun cantidad de entrys de columnas haya
			for index in range(len(self.sheet_columns)):
				#para editar un item, solo comprueba el "cod y descr", xq algunos precios daban error (ej: 0,20800000001)
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
						and self.entry_array[index].get() != self.tree.item(self.tree.selection())['values'][index - 1]:

						#si cambió, actualiza la fecha de precio que esta en index mas 1 (le sigue a precio)
						self.entry_array[index+1].delete(0, 50)
						self.entry_array[index+1].insert(END, datetime.now().strftime("%Y/%m/%d %H:%M:%S"))

			query = f'UPDATE {self.table_name} ' \
					f'SET {query_text_column} ' \
					f'WHERE {query_text_item}'

			parameters.extend(param_anterior) #primero estan los datos actuales, y despues los anteriores
			self.run_query(query, parameters)
			self.message['text'] = f'El elemento {self.entry_array[0].get()} ha sido actualizado'
			self.database_to_tree()
		else:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
			self.database_to_tree(like='%')

	# Borra un registro en la base de datos
	def delete_record(self, *args):
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para eliminar')
			return

		# Elimina el registro con mismo código y descripción o misma Lista y elemento
		cod_sel = self.tree.item(self.tree.selection())['text']
		try:
			descr_sel = self.tree.item(self.tree.selection())['values'][0]  # pasa la seleccion de lista mas elemento
			query = f'DELETE FROM "{self.table_name}" ' \
					f'WHERE "{self.sheet_columns[0]}" = ? ' \
					f'AND "{self.sheet_columns[1]}" = ?'

		#este artificio funciona para rubros que solo tiene una columna ['text'] y no la columna ['values']
		except:
			descr_sel = 'None'

		#para los escasos registros donde la descr es NULL, esto permite borrar esa entrada
		#REVISAR SI NO BORRA UNA LISTA COMPLETA!!!!!!!!!!!!!!!!!
		if descr_sel == 'None':
			descr_sel = ''
			query = f'DELETE FROM "{self.table_name}" ' \
					f'WHERE "{self.sheet_columns[0]}" = ? '
			self.run_query(query, (cod_sel,))  # pone la coma para que se entienda que es una tupla

		else:
			self.run_query(query, (cod_sel, descr_sel))
		self.message['text'] = f'El registro {cod_sel} {descr_sel} ha sido eliminado'

		self.database_to_tree()  # actualiza la tabla

	# Carga el registro seleccionado de TREE en los entrys de edicion
	def load_edit_item(self, *args):

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

	# contruye dataframe a partir de la vista del tree
	def prepare_to_print(self, *args):
		dframe = self.build_header()

		#manda a imprimir
		ToPrinter(to_print=dframe, from_obj=self.entry_array[0].get() or self.table_name)

	# genera el encabezado genérico de impresión
	def build_header(self):
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

#Base de datos especifica para Maestro
class Maestro(Database):
	def __init__(self, window, table_name='MAESTRO'):

		super().__init__(window, table_name)

		# solo desde MAESTRO abre una lista con doble click
		self.tree.bind('<Double-Button-1>', Listas)  # doble click abre la lista

	# agrega un registro en la base de datos de maestro y agrega fecha de alta
	def add_record(self, *args):
		if self.valid_add():
			arg_query = ''
			parameters = []
			item = 0	#usa item, solo para los entrys (pertenenecen a columnas visibles)
			for index in range(len(self.sheet_columns)):
				arg_query += '?'

				#carga la fecha de alta del registro
				if self.sheet_columns[index] == 'FECHA_ALTA':
					parameters.append(datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
				else:
					parameters.append(self.entry_array[index].get())

			query = f'INSERT INTO {self.table_name} ' \
					f'VALUES({" ,".join(arg_query)})' #une la lista con join

			self.run_query(query, parameters)
			self.message['text'] = 'El Registro ha sido guardado con éxito'
			fail_add = ''
		else:
			messagebox.showwarning('Advertencia', 'El Registro ya existe o se encuentra vacío')
			fail_add = '%'  # evita borrar el tree cuando se agrega algo vacio al buscar con LIKE

		self.database_to_tree(like=fail_add)  # like = %: busca la db con "codigo%"
		return not fail_add  # devuelve True si agrego, false si no agrego registro

	# override para eliminar una columna
	def build_header(self):
		list = super().build_header()
		df = pd.DataFrame(list)
		df.drop(df.columns[5], axis=1, inplace=True)
		return dframe


#Base de datos especifica para listas (incorpora crud listas, print to file and printer)
class Listas(Database):

	def __init__(self, *args, table_name='LISTAS'):

		#define una ventana nueva para ver las listas
		listas_window = Toplevel()
		listas_window.iconbitmap('_LOGOLIE x3.ico')
		listas_window.rowconfigure(0, weight=1)
		listas_window.columnconfigure(0, weight=1)
		listas_window.resizable(True, False)

		# Define la BARRA de menu de MAESTRO de articulos
		bar_menu = tk.Menu(listas_window)
		listas_window.config(menu=bar_menu, width=400, height=200)

		# menu de ARCHIVO
		file_menu = tk.Menu(bar_menu, tearoff=0)
		file_menu.add_command(label="Órdenes de trabajo (ctrl+o)", command=Ordenes_trabajo)
		file_menu.add_command(label="Imprimir Lista (ctrl+p)", command=self.prepare_to_print)
		file_menu.add_separator()
		file_menu.add_command(label="Cerrar Ventana (alt+F4)", command=listas_window.destroy)

		# MENU edicion
		edit_menu = tk.Menu(bar_menu, tearoff=0)
		edit_menu.add_command(label="Agregar desde maestro (ctrl+a)", command=lambda: self.add_record())
		edit_menu.add_command(label="Editar Registro (ctrl+e)", command=lambda: self.edit_record())
		edit_menu.add_command(label="Eliminar Registro (ctrl+d)", command=lambda: self.delete_record())

		# Menu ayuda
		help_menu = tk.Menu(bar_menu, tearoff=0)
		help_menu.add_command(label="Ayuda", command=help_info)
		help_menu.add_command(label="Licencia", command=license)
		help_menu.add_command(label="Comandos rápidos", command=hotkeys)
		help_menu.add_separator()
		help_menu.add_command(label="Acerca de GEST2020", command=help_about)

		# items de la barra de menu
		bar_menu.add_cascade(label="Archivo", menu=file_menu)
		bar_menu.add_cascade(label="Edición", menu=edit_menu)
		bar_menu.add_cascade(label="Ayuda", menu=help_menu)

		# Define entry foraneos, los culales los deshabilita para edicion
		self.foreign_entry = [4 ,5]

		# con super inicializa el init del padre como propio para la nueva ventana
		super().__init__(window=listas_window, table_name=table_name)

		# si encuentra 'keysym=c' --> ejecuta copiar lista
		if str(args).find('keysym=c', 0, -1) != -1:     #(texto a buscar, inicio, final)
			self.copy_list()

		# si encuentra 'keysym=f' --> ejecuta eliminar lista
		if str(args).find('keysym=f', 0, -1) != -1:
			self.delete_list()

		# si encuentra 'keysym=double click' --> ejecuta cargar lista
		if str(args).find('ButtonPress', 0, -1) != -1:
			self.load_lista()

	def read_columns(self):
		query = f'SELECT {self.table_name}.*, ' \
				f'{maestro.table_name}.{maestro.sheet_columns[1]},{maestro.table_name}.{maestro.sheet_columns[3]} ' \
				f'FROM {self.table_name} ' \
				f'INNER JOIN {maestro.table_name} LIMIT 1'
		# ON "LISTAS.COMPONENTE" = "MAESTRO.CÓDIGO"
		columns = self.run_query(query)
		for columna in columns.description:
			self.sheet_columns = np.append(self.sheet_columns, columna[0])

	# Armado de ventana, TREEVIEW adaptable segun LISTAS
	def  build_main_view(self):
		# creacion de tabla para visualizar
		self.tree = ttk.Treeview(self.frame_tree, height=25, columns=len(self.sheet_columns))
		self.tree.grid(row=10, column=0, columnspan=20, rowspan=1, pady=10, sticky=N + S + W + E)

		# Scroll vertical del TREE
		self.scroll_tree_v = Scrollbar(self.frame_tree, command=self.tree.yview)
		self.scroll_tree_v.grid(row=10, column=len(self.sheet_columns), sticky=NS)
		self.tree.config(yscrollcommand=self.scroll_tree_v.set)

		# Scroll horizontal del TREE
		self.scroll_tree_h = Scrollbar(self.frame_tree, orient='horizontal', command=self.tree.xview)
		self.scroll_tree_h.grid(row=12, column=0, columnspan=20, sticky=W + E)
		self.tree.config(xscrollcommand=self.scroll_tree_h.set)

		# NOMBRES COLUMNAS
		nombres_columnas = []
		for each_column in self.sheet_columns:
			nombres_columnas.append(each_column)
		self.tree["columns"] = nombres_columnas[
							   1:]  # desde 2 xq el 1ro es index y el 2do esta en text (no en values)

		# Esta query es para dar TAMAÑO DINAMICO al ancho de columnas del tree
		query = f'SELECT {self.table_name}.*, ' \
				f'{maestro.table_name}.{maestro.sheet_columns[1]},{maestro.table_name}.{maestro.sheet_columns[3]} ' \
				f'FROM {self.table_name} ' \
				f'INNER JOIN {maestro.table_name} LIMIT 25'
		cursor = self.run_query(query)

		# se cargan titulos y tamaños de las columnas del tree, asi como Entrys para la edicion de records
		num_column = 0
		item_table = ''
		all_table = cursor.fetchmany(25)
		for each_column in self.sheet_columns:
			largest = ''
			for row_table in all_table[:25]:#lee las 1ras 25 lineas de cada columna, para estimar el ancho de la misma
				item_table = str(row_table[num_column])
				if len(largest) < len(item_table):
					largest = item_table

			self.tree.column("#" + str(num_column), width=20 + (len(largest) * 6), minwidth=30, stretch=True)
			self.tree.heading('#' + str(num_column), text=f'{nombres_columnas[num_column]}', anchor=CENTER)

			# crea tantos ENTRYS como columnas, para editar registro
			self.entry_array.append(self.entrys(frame=self.frame_tree,
												name=each_column, row=7, column=num_column, width=len(largest) + 4))

			num_column += 1

	#query search de listas inner join maestro
	def query_search(self, *args, like='%', col_search=0, col_order=1, order='DESC'):
		query = f'SELECT {self.table_name}.*' \
				f', {maestro.table_name}.{maestro.sheet_columns[1]}, {maestro.table_name}.{maestro.sheet_columns[3]} ' \
				f'FROM "{self.table_name}" ' \
				f'INNER JOIN "{maestro.table_name}" ' \
				f'ON {self.table_name}.{self.sheet_columns[1]} = {maestro.table_name}.{maestro.sheet_columns[0]} ' \
				f'WHERE {self.table_name}.{self.sheet_columns[col_search]} ' \
				f'LIKE "{self.entry_array[col_search].get()}{like}" ' \
				f'ORDER BY "{self.sheet_columns[0]}" {order}, "{self.sheet_columns[1]}" {order}'
		return self.run_query(query)

	# agrega un registro en la base de datos de listas (que tiene 2 columnas menos que las visibles)
	def add_record(self, *args):
		columns_list = self.sheet_columns
		self.sheet_columns = self.sheet_columns[0:4]
		super().add_record()
		self.sheet_columns = columns_list

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
				# para editar un item, solo comprueba el "cod y descr", xq algunos precios daban error (ej: 0,20800000001)
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
						self.entry_array[index + 1].insert(END, datetime.now().strftime("%Y/%m/%d %H:%M:%S"))

			query = f'UPDATE {self.table_name} ' \
					f'SET {query_text_column} ' \
					f'WHERE {query_text_item}'

			parameters.extend(param_anterior)  # primero estan los datos actuales, y despues los anteriores
			self.run_query(query, parameters)
			self.message['text'] = f'El elemento {self.entry_array[0].get()} ha sido actualizado'
			self.database_to_tree()
		else:
			messagebox.showwarning('Advertencia', 'Debe seleccionar un registro para editar')
			self.database_to_tree(like='%')

		self.sheet_columns = columns_list

	# funcion override de database para listas, solo impide cuando lista y elemento son identicos a algun registro
	def valid_add(self):
		if maestro.entry_array[0].get() == '' or maestro.entry_array[1].get() == '':
			return False
		query = f'SELECT "{self.sheet_columns[1]}", "{self.sheet_columns[2]}" from {self.table_name}'
		db_rows = self.run_query(query)
		for row in db_rows:
			if self.entry_array[0].get() == row[0] and maestro.entry_array[0].get() == row[1]:
				return False
		return True

	# carga una lista segun el codigo en maestro con doble click
	def load_lista(self, *args):
		self.hab_entry(self.foreign_entry)
		self.entry_array[0].delete(0, 50)
		self.entry_array[0].insert(END, maestro.tree.item(maestro.tree.selection())['text'])
		self.database_to_tree(like='')
		self.deshab_entry(self.foreign_entry)

	#Copia una lista con diferente código
	def copy_list(self, *args):

		# crea el registro en maestro con nuevo código
		if maestro.add_record():
			cursor = super().query_search()

			#arma argumento de query, segun cantidad de columnas
			arg_query = ''
			for column in range(len(self.sheet_columns)-2): #quita las 2 columnas de maestro
				arg_query += '?'

			# arma nuevo dataframe para ingresar en database listas (reemplaza la columna del codigo)
			parameters = []
			index = 0
			for row in cursor.fetchall():
				tupla = (maestro.entry_array[0].get(), row[1], row[2], row[3])
				parameters.insert(index, tupla)
				index += 1

			#Ingresa masivos datos en database de listas
			query = f'INSERT INTO {self.table_name} ' \
					f'VALUES({", ".join(arg_query)})'
			self.run_query_many(query, parameters)

			#hace busqueda de la nueva lista creada
			self.hab_entry(self.foreign_entry)
			self.entry_array[0].delete(0, 100)
			self.entry_array[0].insert(END, maestro.entry_array[0].get())
			self.deshab_entry(self.foreign_entry)
			self.database_to_tree(like='')
			maestro.message['text'] = "Nueva lista creada con éxito"
		else:
			maestro.message['text'] = "No se copió ningún registro"

	# Borra una lista completa de la base de datos
	def delete_list(self, *args):
		try:
			maestro.tree.item(maestro.tree.selection())['text'][0]
		except IndexError as e:
			messagebox.showwarning('Advertencia', 'Debe seleccionar una lista para eliminar')
			return

		#Elimina la lista con mismo código
		cod_sel = maestro.tree.item(maestro.tree.selection())['text']
		query = f'DELETE FROM "{self.table_name}" ' \
				f'WHERE "{self.sheet_columns[0]}" = ?'
		self.run_query(query, (cod_sel, ))  # pone la coma para que se entienda que es una tupla
		self.message['text'] = f'La lista {cod_sel} ha sido eliminada'

		#Por último, borra el registro del maestro
		maestro.delete_record()

		# actualiza la tabla de listas
		self.database_to_tree(like='')

	# encabezado para listas, incluye precio total
	def build_header(self):
		# recorre cada fila del TREE y convierte a DataFrame
		list_row = []
		precio_lista = 0
		for row in self.tree.get_children():
			self.tree.selection_set(row)
			descr_item = self.tree.item(self.tree.selection())['values']
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

	# Carga el registro seleccionado de TREE en los entrys de edicion
	def load_edit_item(self, *args):
		self.hab_entry(self.foreign_entry)
		super().load_edit_item()
		self.deshab_entry(self.foreign_entry)

	#deshabilita varios entrys para evitar su edicion
	def deshab_entry(self, hab_entry=[]):
		for item in hab_entry:
			self.entry_array[item].config(state='readonly')

	#habilita los entry para la edicion o ingreso de datos
	def hab_entry(self, deshab_entry=[]):
		for item in deshab_entry:
			self.entry_array[item].config(state='normal')


#Base de datos especifica para Ordenes de trabajo
class Ordenes_trabajo(Database):
	def __init__(self, *args, table_name='OT'):
		ot_window = Toplevel()
		ot_window.rowconfigure(0, weight=1)
		ot_window.columnconfigure(0, weight=1)
		ot_window.iconbitmap('_LOGOLIE x3.ico')
		ot_window.resizable(True, False)
		super().__init__(ot_window, table_name)


#Clase para manejar los rubros
class Rubros(Database):
	def __init__(self, *args, table_name='RUBROS'):
		rubros_window = Toplevel()
		rubros_window.rowconfigure(0, weight=1)
		rubros_window.columnconfigure(0, weight=1)
		rubros_window.iconbitmap('_LOGOLIE x3.ico')
		rubros_window.resizable(True, False)
		super().__init__(rubros_window, table_name)

	# genera el encabezado genérico de impresión
	def build_header(self):
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
	help_window.iconbitmap('_LOGOLIE x3.ico')
	help_window.resizable(True, True)
	help_window.geometry('800x300')

	# FRAME
	help_frame = LabelFrame(help_window, text='Menu de ayuda')
	help_frame.grid(column=0, row=0, sticky='nsew')

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
	help_two = help_tree.insert('', tk.END, text='Agregar elementos a una lista')
	help_three = help_tree.insert('', tk.END, text='Para mas acciones vea el apartado "comandos rápidos" desde el menú ayuda')

	help_tree.insert(help_one, tk.END, text='Primero debe importar las bases de datos ORIGINALES con Excel: '
											'"MAESTRO.dbf", "LISTAS.dbf", "OT.dbf"')
	help_tree.insert(help_one, tk.END, text='y guardar en la misma carpeta del programa los archivos "*.xls"')
	help_tree.insert(help_one, tk.END, text='Luego debe ir a menú "Herramientas/Crear base de datos".')
	help_tree.insert(help_one, tk.END, text='Y reiniciar el programa." ')

	help_tree.insert(help_two, tk.END, text='Abra la lista a editar, haciendo doble click en la misma desde maestro')
	help_tree.insert(help_two, tk.END, text='Seleccione el item a agregar en Maestro de articulos')
	help_tree.insert(help_two, tk.END, text='y desde la ventana de la lista destino pulsar CTRL-A o menú "Edición/Agregar"')


#Detalle de los HOTKEYS del programa
def hotkeys():
	messagebox.showinfo('Accesos rápidos', 'Enter: filtrar búsqueda por campo\n'
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
										   'Ctrl+P: Imprimir\n'
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

	#ROOT, ventana principal
	root = ThemedTk(theme='clearlooks')
	root.title('GEST2020 | Sistema de Administración de Producción L.I.E. S.R.L.')  # titulo de la ventana
	root.iconbitmap('_LOGOLIE x3.ico')  # icono e la ventana
	root.config(bg="grey")  # config de la ventana bg= back ground color
	root.rowconfigure(0, weight=1)
	root.columnconfigure(0, weight=1)

	# Define la BARRA de menu de MAESTRO de artículos
	bar_menu = tk.Menu(root)
	root.config(menu=bar_menu, width=800, height=200)
	root.resizable(True, True)

	# menu de ARCHIVO
	file_menu = tk.Menu(bar_menu, tearoff=0)
	file_menu.add_command(label="Abrir Listas (ctrl+L)", command=Listas)
	file_menu.add_command(label="Órdenes de trabajo (ctrl+O)", command=Ordenes_trabajo)
	file_menu.add_command(label="Rubros (ctrl+R)", command=Rubros)
	file_menu.add_separator()
	file_menu.add_command(label="Cerrar programa (alt+F4)", command=root.destroy)

	# MENU edicion
	edit_menu = tk.Menu(bar_menu, tearoff=0)
	edit_menu.add_command(label="Copiar Lista (ctrl+C)", command=lambda: Listas('keysym=c'))
	edit_menu.add_command(label="Eliminar Lista (ctrl+F)", command=lambda: Listas('keysym=f'))
	edit_menu.add_separator()
	edit_menu.add_command(label="Agregar Registro (ctrl+A)", command=lambda: maestro.add_record())
	edit_menu.add_command(label="Editar Registro (ctrl+E)", command=lambda: maestro.edit_record())
	edit_menu.add_command(label="Eliminar Registro (ctrl+D)", command=lambda: maestro.delete_record())

	# MENU Herramientas
	tools_menu = tk.Menu(bar_menu, tearoff=0)
	tools_menu.add_command(label="Cambiar apariencia de la ventana", command=lambda: WindowTheme(win_to_mod=root))
	tools_menu.add_separator()
	tools_menu.add_command(label="Crear base de datos GEST2020", command=CreateDatabase)

	#Menu ayuda
	help_menu = tk.Menu(bar_menu, tearoff=0)
	help_menu.add_command(label="Ayuda", command=help_info)
	help_menu.add_command(label="Licencia", command=license)
	help_menu.add_command(label="Comandos rápidos", command=hotkeys)
	help_menu.add_separator()
	help_menu.add_command(label="Acerca de GEST2020", command=help_about)

	# items de la barra de menu
	bar_menu.add_cascade(label="Archivo", menu=file_menu)
	bar_menu.add_cascade(label="Edición", menu=edit_menu)
	bar_menu.add_cascade(label="Herramientas", menu=tools_menu)
	bar_menu.add_cascade(label="Ayuda", menu=help_menu)

	#----INSTANCIAS Y COMANDOS RAPIDOS-------------------------

	#Instancia a Maestro de articulos (database)

	#maestro = Maestro(window=root, table_name='Maestro')
	#'''
	try:
		maestro = Maestro(window=root, table_name='Maestro')
	except:
		messagebox.showwarning('Advertencia', 'No se encuentra la base de datos\n\n'
					'vaya a menú "Herramientas/crear base de datos" para generarla.\n'
					'de no existir los archivos XLS, deberá crearlos previamente con "MS Excel" o "CALC OpenOffice".')
	#'''

	#Instancia Listas con CTRL+L y otras sobre listas
	root.bind('<Control-l>', Listas)    # abre listas
	root.bind('<Control-L>', Listas)    # abre listas
	root.bind('<Control-c>', Listas)    # copia lista con nuevo codigo
	root.bind('<Control-C>', Listas)    # copia lista con nuevo codigo
	root.bind('<Control-f>', Listas)    # Elimina lista seleccionada
	root.bind('<Control-F>', Listas)    # Elimina lista seleccionada

	# Crear database
	root.bind('<Control-b>', CreateDatabase)
	root.bind('<Control-B>', CreateDatabase)

	#LOOP CIERRE
	root.mainloop()
