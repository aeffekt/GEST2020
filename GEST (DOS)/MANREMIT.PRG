*........................Mantenimiento de remitos
remicolor=setcolor()
declare arch[2]
arch[1]='remitos'
arch[2]='notas'
marca(arch,2,.t.)

sele a
use maestro index maestro
sele b
use remito index remito,remitore
sele d
use ot index otcodcom
reindex
sele b
panremito=savescreen(2,5,21,75)
do while .t.
sele b
set order to 1
nuop=3
declare mens[nuop]
mens[1]=" Captura   "
mens[2]=" Consultas "
mens[3]=" Impresi¢n "
filini=2
colini=2
fil=1
col=0
fg=19
cg=2
opcion=selin(filini,colini,mens,fg,cg,nuop)
if opcion="salir"
exit
endif

do case

case opcion=mens[1]
sele b
set order to 1
go bottom
nuerem=remito+1
pag=1
go top
set color to w/n
@ 3,7,21,75 box replicate(chr(177),9)
set color to n/w,w/n,n,n,n/w
@ 2,5 clear to 20,73
@ 2,5 to 20,73
@ 2,17 say " Captura de Remito N§:"
@ 2,40 say nuerem picture "##########"
set cursor on
mfecha=date()
mremreal=0
@ 3,7 say "Remito real:" get mremreal picture "##########"
@ 3,45 say "Fecha: " get mfecha
@ 3,64 say "Pag:"+str(pag,2)
read
if lastkey()=27
exit
endif
set cursor off
@ 5,6 say "C¢digo"
@ 5,25 say "      O.T."
@ 5,38 say "Cant."
linea=7

do while .t.

if linea >18
fii=15
coi=10
mensa="Oprima una tecla para continuar..."
m=mensaje(fii,coi,mensa)
@ 7,6 clear to 19,72
pag=pag+1
@ 3,64 say "Pag:"+str(pag,2)
linea=7
endif

set cursor on
mcod=space(15)
@ linea,6 get mcod picture "@!"
read

do case
case  lastkey()=27
exit

case mcod=space(15)
sele d
busca=buscanota(mcod)
if lastkey()=27
loop
endif
mcod=ot->cod
@ linea,6 say mcod picture "@!"

otherwise
sele d
seek mcod
if .not. found()
fii=linea+1
coi=10
mensa="Item inexistente en archivo O.T."
m=mensaje(fii,coi,mensa)
loop
endif
busca=buscanota(mcod)
if .not. busca
loop
endif

endcase

mcant=0
max=(d->cant)-(d->cancum)
@ linea,25 say d->compro
@ linea+1,38 say "(M ximo="+str(max,4)+")"
set cursor on

do while .t.
@ linea,38 get mcant picture "####"
read
if mcant <=max .and. mcant >=0
exit
endif
fii=linea+1
coi=40
mensa="Cantidad fuera de rango"
m=mensaje(fii,coi,mensa)
enddo

set color to n/w
@ linea+1,38 say space(13)
set cursor off
if lastkey()=27
loop
endif
nuop=2
declare mens[nuop]
mens[1]=" Confirma "
mens[2]=" Cancela  "
filini=linea+1
colini=45
fil=1
col=1
fg=19
cg=2
opcion=selin(filini,colini,mens,fg,cg,nuop)
if opcion="salir" .or. opcion=mens[2]
loop
endif
sele b
append blank
replace remito with nuerem
replace remreal with mremreal
replace compro with d->compro
replace cod with mcod
replace cant with mcant
replace fecha with mfecha
sele d
replace cancum with cancum+mcant
if cancum=cant
replace fecum with mfecha
endif
sele b
linea=linea+1
enddo

m=mensaje(18,5," Imprime remito  (S/N) ? ")
if m="S" .or. m="s"
sele b
seek nuerem
nuere="remito"
impremito(nuerem,mremreal,nuerem,nuere)
mensa=" Oprima una tecla para continuar..."
m=mensaje(20,5,mensa)
endif

case opcion=mens[2]      &&..............Consulta
pag=1
go bottom
sele b
set order to 1
nuremax=remito
mcolor=setcolor()
panremito=savescreen(2,5,21,75)
set color to w/n
@ 3,7,21,75 box replicate(chr(177),9)
set color to n/w,w/n,n,n,n/w
@ 2,5 clear to 20,73
@ 2,5 to 20,73
@ 3,7 say " Consulta de Remito N§:"
nuerem=0
do while .t.
nuop=2
declare mens5[nuop]
mens5[1]=" Numeraci¢n Interna "
mens5[2]="      '     Real    "
filini=7
colini=25
fil=1
col=1
fg=19
cg=2
op=selin(filini,colini,mens5,fg,cg,nuop)
if op="salir"
exit
endif

sele b
if op=mens5[1]
set order to 1
cond='remito'
pasi='remreal'
else
set order to 2
cond='remreal'
pasi='remito'
endif

@ 4,6 clear to 19,72
set cursor on
@ 3,30 get nuremax picture "##########"
read
set cursor off
if lastkey()=27
exit
endif
sele b
seek nuremax

if .not. found()
mensa="Remito inexistente. Oprima una tecla para continuar..."
m=mensaje(10,5,mensa)
loop
endif

set color to n/w
@ 3,45 say "Fecha: "+dtoc(fecha)
@ 3,64 say "Pag:"+str(pag,2)
@ 4,28 say "("
@ 4,30 say &pasi
@ 4,40 say ")"
@ 5,6 say "C¢digo"
@ 5,25 say "      O.T."
@ 5,38 say "Cant."
linea=7

do while &cond = nuremax .and. .not. eof()

if linea >18
mensa="Oprima una tecla para continuar..."
m=mensaje(20,5,mensa)
@ 7,6 clear to 19,72
pag=pag+1
@ 3,64 say "Pag:"+str(pag,2)
linea=7
endif

@ linea,6 say cod picture "@!"
@ linea,25 say compro
@ linea,38 say cant
linea=linea+1
skip
enddo

mensa=" Oprima una tecla para continuar..."
m=mensaje(20,5,mensa)
enddo
set color to &mcolor
restscreen(2,5,21,75,panremito)

case opcion=mens[3]     &&................Impresion
sele b
go bottom
nuremax=remito
mcolor=setcolor()
panrem=savescreen(2,5,8,75)
set color to w/n
@ 3,7,8,75 box replicate(chr(177),9)
set color to n/w,w/n,n,n,n/w
@ 2,5 clear to 7,73
@ 2,5 to 7,73
@ 4,7 say " Impresi¢n de Remito N§:"
nuerem=0
do while .t.
nuop=2
declare mens5[nuop]
mens5[1]=" Numeraci¢n Interna "
mens5[2]="      '     Real    "
filini=7
colini=25
fil=1
col=1
fg=19
cg=2
op=selin(filini,colini,mens5,fg,cg,nuop)
if op="salir"
exit
endif

sele b
if op=mens5[1]
set order to 1
cond='remito'
pasi='remreal'
else
set order to 2
cond='remreal'
pasi='remito'
endif
set cursor on
@ 4,30 get nuremax picture "##########"
read
set cursor off
if lastkey()=27
exit
endif
sele b
seek nuremax
if .not. found()
mensa="Remito inexistente. Oprima una tecla para continuar..."
m=mensaje(10,5,mensa)
loop
endif
nurein=remito
nureal=remreal
nure=cond
impremito(remito,remreal,&cond,cond)

mensa=" Oprima una tecla para continuar..."
m=mensaje(20,5,mensa)
enddo
set color to &mcolor
restscreen(2,5,8,75,panrem)

otherwise
endcase
enddo
set color to &remicolor
restscreen(2,5,21,75,panremito)
close all
marca(arch,2,.f.)

*............................................
function buscanota
parameters mcodi
ncolor=setcolor()
panbusca=savescreen(5,25,23,75)
set color to w/n
@ 6,27,23,75 box replicate(chr(177),9)
set color to n/w,w/n
@ 5,25 clear to 22,73
@ 5,25 to 22,73
@ 5,28 say " Seleccione O.T. a cumplir "
private campos[6],tit[6],seph,sepv,sepp

campos[1]='compro'
campos[2]='cod'
campos[3]='cant'
campos[4]='cancum'
campos[5]='feven'
campos[6]='fecum'
tit[1]='    O.T.'
tit[2]='C¢digo'
tit[3]='Sol.'
tit[4]=' Ent.'
tit[5]='Vencim.'
tit[6]='Cumpl.'
seph=" "
sepv="  "
sepp="Ä"
set color to n/w,w/n
dbedit(6,26,21,72,campos,'auxbusca','',tit,seph,sepv,sepp)
restscreen(5,25,23,75,panbusca)
set color to &ncolor
do case
case lastkey()=27
return .f.
case lastkey()=13
if cod=mcodi
return .t.
else
return .f.
endif
endcase
*....................................................
function auxbusca
parameters mode,pos
do case
case mode=0
return(1)
case lastkey()=27 .or. lastkey()=13
return(0)
otherwise
return(1)
endcase
*....................................................
function impremito
parameters nurein,nureal,valcond,campocond
selprint=ajuimp()

do case
case lastkey()=27
return []
case selprint="H" .or. selprint="F"
set device to printer
case selprint="A"
set device to printer
set printer to &mnomarch
endcase

@ prow(),pcol() say chr(27)+chr(15)
@ prow(),pcol() say chr(10)
@ prow(),pcol() say chr(10)
filaini=prow()
pag=1
fila=filaini
@ fila,10 say "L.I.E. S.R.L. -  Sistema de Administraci¢n de Producci¢n GEST "
@ fila,75 say "Remito N§ "
@ fila,85 say nurein picture "##########"
@ fila,100 say "Pag:"
@ fila,104 say pag picture "##"
@ fila,107 say fecha
fila=fila+1
@ fila,75 say "(N§ Real :"
@ fila,85 say nureal
@ fila,95 say ")"
fila=fila+2
@ fila,10 say "Nota de Pedido"
@ fila,28 say "C¢digo"
@ fila,45 say "Descripci¢n"
@ fila,87 say "Cantidad"
@ fila,97 say "UM"
fila=fila+2

do while valcond = &campocond .and. .not. eof()
if fila > 55
if selprint="H"
eject
set device to screen
fii=20
coi=10
mensa="Inserte nueva hoja y oprima una tecla para continuar..."
m=mensaje(fii,coi,mensa)
set device to printer
else
eject
endif
@ prow(),pcol() say chr(10)
@ prow(),pcol() say chr(10)
filaini=prow()
pag=1
fila=filaini
@ fila,10 say "L.I.E. S.R.L. -  Sistema de Administraci¢n de Producci¢n GEST "
@ fila,75 say "Remito N§ "
@ fila,85 say nurein picture "##########"
@ fila,100 say "Pag:"
@ fila,104 say pag picture "##"
@ fila,107 say fecha
fila=fila+1
@ fila,75 say "(N§ Real :"
@ fila,85 say nureal
@ fila,95 say ")"
fila=fila+2
@ fila,10 say "Nota de Pedido"
@ fila,28 say "C¢digo"
@ fila,45 say "Descripci¢n"
@ fila,87 say "Cantidad"
@ fila,97 say "UM"
fila=fila+2
endif

@ fila,10 say compro
@ fila,28 say cod picture "@!"
sele a
seek b->cod
sele b
@ fila,45 say a->des
@ fila,87 say cant
@ fila,97 say a->um1
fila=fila+1
skip
enddo
set device to screen
return []

