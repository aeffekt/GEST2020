# Módulo que crea Entry Widget con texto en mayúscula automáticamente
# https://es.stackoverflow.com/questions/356082/como-lograr-poner-mayúscula-en-los-campos-de-tipo-entry-y-formatear-números-en-p

import tkinter as tk
from tkinter import ttk

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