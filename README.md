# GEST 2020 - P.O.S. desarrollado para la empresa L.I.E. S.R.L. por Agustín Arnaiz


## Modulos usados:
Tkinter
numpy
Pandas
pywin32
ttkthemes
tabulator
pyinstaller
win32api
functools

## CREAR INSTALADOR STANDALONE
CMD: 
'''
pyinstaller -w --onefile --icon=./logo.ico GEST2020.py
'''
-w= no muestra ventana CMD al ejecutar
--onefile= crea un ejecutable .exe (es mas lento el programa y no es necesario)
--icon (icono del programa)

## VERSIONES
V0.1.0	02/07/20
Manejo CRUD de base de datos LISTAS.db
Conversion XLS a sqlite3


V0.1.1	08/07/20
Incorpora clase UpperEntry
Herencia de clases para listas
comandos rápidos


V0.1.2	09/07/20	
Desarrollo de metodos específicos para listas
Uso de funciones mágicas (super())
Impresion de listas en archivo XLS
Copia listas completas

V0.1.3	10/07/20	
Uso masivo de run_query (executemany)
Arma encabezado de impresion de listas


V0.1.4	16/07/20
Elimina funciones que llaman a clases hijas
Incorpora informacion de cada elemento (maestro) dentro de las listas de materiales
Imprime en xls el precio total de cada lista y nueva info


V0.1.5	18/07/20	
Define barra menu para cada ventana
Ventana se autoposiciona
Mejora de encabezado y ancho de columnas de impresion de archivo
genera archivo xlsx al imprimir (antes xls)
Agrega items de maestro a listas (no agrega desde lista)

V0.1.6	21/07/20
Menu ayuda mejorado, usa tree
Imprime directo a impresora default
maneja nombres de columnas de base de dato completas ("COD,15")

V0.2.0	06/08/20
Nueva base de datos unica, con 3 tablas. Incorporan NN, AI, FK etc
Usa nuevo conversor de XLS a DB (CLASE integrada al programa)

V0.2.1	08/08/20
no mas A.I. en tables (no recomendado su uso)
Uso de INNER JOIN en la busqueda de info de LISTAS
Se eliminaron columnas sin uso de db (index, comprasug, stkmin, etc)
Uso de ttkthemes (apariencia)
Introduce tabla RUBROS, y maneja rubros con un ComboBox
Las busquedas se pueden hacer desde cualquier columna, no solo de codigo
Elimina (de DB) de precios, cantidades, etc numeros con muchos decimales

v0.2.2	11/08/20
Nuevas clases WindowConfig y ProgramManager
Los widgets se adaptan en X e Y a la ventana
La ventana se adapta a la resolucion del monitor
guarda archivo de configuracion, y lo lee al empezar
crea archivo de configuraciones por defecto si el mismo no existe

v0.2.3	13/08/20
agrupa por plan de trabajo la vista
mejora ediciones de lista
mejora add registro
autocompletado de db en fechas y valores por defecto (solo maestro)

v0.2.4	20/08/20
Triggers fechas_precio y Monto (OT)
Edicion y Add item mejorados 
el tamaño de cada ventana se adapta a la tabla, y no a la resolucion de pantalla
refactor Database <--> ManageTable
se eliminó class read_column (integrada en build_main_view con un solo query)
todos los widgets son ttk
guarda posicion y tamaño de ventana al salir

v0.2.5	24/08/20
Mejoras en copia de listas
Algunas toplevel solo abren una vez (usa variable de metodo)
varios metodos redefinidos a estaticos
Trigger monto redondea a 3 decimales
Si hay un item a la vista, lo edita aun sin estar seleccionado
Edicion multiple simultanea (no permite mod: cod, descr, lista o plan)
Agrega a listas y Orden desde maestro, o desde si mismo, si tiene cargado un código

v0.3.0	04/09/20
no_edit_item atributo de clase 
mejoras en VALID_ADD de listas y Ordenes de trabajo
guarda posicion de las ventanas de Listas y OT ademas de la de Maestro
Genera backup de la base de datos al abrir, hasta 10 copias (reemplaza la mas antigua)
Nueva funcion wrapper en ProgManager, mide el tiempo de ejec de funciones (con decorador)
logging a archivo, debug level
Renombra los NULL de MAESTRO.RUBRO a "NONE" y permite su edicion
Una orden de trabajo carga todos los items de cada lista, y genera una Orden de Compra
mejora en delete_decimal (solo edita numeros con muchos decimales)
metodos de clase implementados (exit_handler OT)
Ordena el TREE por columna al hacer click
La órden de compra unifica los código repetidos y suma su cantidad
Despues de un mensaje en pantalla vuelve a FRONT la ventana que estaba trabajando

v0.3.1	17/09/20
Ordena por todas las columnas del TREE e invierte la busqueda al siguiente click
Manejo de la impresora con Device Context Printer
corregido en todas las tablas build_print, maneja encabezado y dataframe por separado
permite seleccionar los rubros a imprimir en Orden de Compra
uso de wait_window para seleccionar rubros a imprimir
muestra en pantalla el progreso de creacion de base de datos.. update window
instancia maestro con funcion, no hace falta reiniciar programa despues de crear base de datos
al ordenar por columna, tambien ordena por codigo como 2do orden
elimina decimales extras usando pandas astype() y round() en lugar de un FOR
reemplaza 'nan' de dataframe con "Sin_Rubro" en col RUBRO de MAESTRO
no achica las ventanas en eje "Y" de "lista" y "OT" al usar "update()" despues de "super()"
realiza busqueda inversa cambiando la relacion FK entre CODIGO y LISTA
en busqueda inversa muestra en barra de mensajes los datos de la busqueda (codigo y descripción)
fuera de maestro impide busquedas que no sean de LISTA o CODIGO (solo entry_array [0] y [1])
en nombre de archivo PRINT, muestra el codigo en lista inversa

V0.3.2	21/09/20
en órden de trabajo ademas de acumular la cantidad de un item repetido, suma el MONTO
resuelto problema de edicion cuando se autoleccionaba el unico item del tree
archivo *.ico cambia de nombre, pierde el espacio como asi el *.py
Primer ejecutable GEST2020 version 64bits

v1.0.0	13/10/20
Incorpora cantidad de backup en archivo config
los métodos de backup file ya no son estáticos
evita error al ordenar por columna #0 el tree
uso cierre sys.exit() al cierre del programa
logger incluido en ProgManager
el icono de ventana lo toma del archivo EXE (sys.executable)

v1.0.1 28/12/20
resuelto problema que no copiaba bien listas
hotkey copiar lista de "ctrl-c" a "ctrl-m"
corregido, add listas, no toma precio de maestro para editar cantidad
add listas override, limpia el campo de codigo, permite agregar varios items desde maestro
edicion ayuda edicion listas
la busqueda LIKE (%) se hace previo y pos palabra a buscar

V1.0.2  30/12/20
EL MENU DE HOTKEYS, license y acerca de, ya no retorna a maestro de articulos (parent window)
menu de impresion responde a ENTER para imprimir (ademas de barra espaciadora)


v1.0.3  12/03/21
Al abrir una base de datos la carga automaticamente sin cerrar el programa

v1.0.4  26/05/21
Se agrego menu de configuracion de impresion de listas, permite elegir valorizacion y cantidad de listas a imprimir y lo indica en la impresion

v1.0.5  02/06/21
modif desde ayuda hotkey CTRL-M para copiar listas
se agrega "Sin_Rubro" como un rubro en DB, para evitar error al copiar listas

v1.0.6	28/06/21
Edicion multiple actualiza cada item editado en tiempo real en la barra de estado
Copiar lista, funciona si se usa "control-M" (con mayuscula)

v1.0.7  13/10/21
Al agregar items a lista, reemplaza cantidad None por 0, evita errores al imprimir
se agrego (separó del metodo edit) metodo para validar edicion de DB
advertencia edición multiple y masiva
limpieza leve de código en general
Se disminuyó el margen izquierdo de impresión final

v1.0.8  26/01/22
Al crear la base de datos, el precio de maestro es NOT NULL valor por defecto=0
se modificó DB actual: maestro, PRECIO por defecto=0
Corregido error al cargar archivo de configuraciones si el mismo estaba vacío,
O si hay menos configuraciones de las que debería

v1.0.9 08/03/23
corregido error al imprimir 1 sola lista con un solo item, daba error (build_print)

v1.0.10 14/07/23
se agrego boton "Ver" para ver una lista valorizada sin imprimir

v1.1.0 12/02/25
se implementa la opcion de agregar múltiples items maestros a Listas en un solo paso

v1.1.1 13/02/25
Abre ventana con confirmación antes de eliminar uno o varios items