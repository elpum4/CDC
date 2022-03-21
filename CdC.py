'''
 Cambios de Celular
'''
from tkinter import Tk, Toplevel, messagebox
from datetime import date
from A5 import A5, A5Tk
import win32com.client
import win32ui
import tempfile
import os
import re

EcoLogo= '''
    R0lGODlhEAAQAHAAACwAAAAAEAAQAIfGxsbFxcXExMTDw8PCwsLAwMC/v7++vr69vb28vLy7u7u5
    ubm4uLi3t7e2tra1tbXJycnIyMjHx8fBwcG6urrMzMzLy8vKysqRpb/Pz8/Ozs7Nzc2otMVgjLzS
    0tLR0dHQ0NC8wst3l7+Am7/Bw8fV1dXU1NTT09OesMiXqsOut8Z9m8HHyMrY2NjX19fW1tasucuv
    ucl0mMK9wsp/nMHb29va2trZ2dnCyNCEoMSuusqotsh2msPe3t7d3d3c3Nza2tnQ09aQqMesusuy
    vc2En8LT0dHYu7Xdl4bSt7Lh4eHg4ODf39/d2djdv7nQvb6drMWrmq7amo7ekn/WuLLk5OTj4+Pi
    4uLexcHds6vVy8nDxs3Dxc3LzdHn5+fm5ubCxtGsssacpsCJlrhhdqpYcKdYcaeiqb7q6urp6em2
    vMy5wtWmsMrJz962wNTDytq1vdKrssTt7e3s7Oy7wc+SoL+ap8Olrsestcylr8irtMuyuMfv7+/u
    7u7o6Ojl5eXy8vLw8PAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAItgArCFRQ
    gaCCgwgRFhxY0GDCgwIFAsAhxmBDhRErnOmwUMGEhD4i4hAhQk3ECREVhAwJRQyHMTgAVADwsaCP
    IE08DMHRYWTGiD6CehiphoOInwKDBg0iZMgQDB6cTEkCNGgTJ0/EREEytUKQiFb8KMWSJUsFHDgq
    NFnZJ6wVHEPEjDFTpoyYlRX69PHRR40aNm3UBD6jlK9eK2rGQGEzBMoQNYX1Stbrx21hw5Pb9vHj
    Z+9eHwEBADs=
'''

Hoy= date.today()
Fecha= Hoy.strftime("%d/%m/%Y")
# Cargo Headers
Param= A5("CdC_Cfg.xlsx","Parametros","ParName")

def Pmt(VarName):
    return Param.D["ParName"][VarName]["ParVal"] 

def GetList(PsvStr,Estan=True):
    # Retorna una lista con las columnas de PsvStr (pipe "|" separated values)
    # Que Estan o NO en la lista total de columnas
    global Heads
    RetDir= []
    Lst= Pmt(PsvStr).split("|")
    if Estan:
        for h in Lst:
            if h in Heads: RetDir.append(h)
    else:
        for h in lHead:
            if (h!="") and (h not in Lst): RetDir.append(h)
    return RetDir


# ------------------------------------------------------------------------------- Carga de listas
Heads={}    # Lista total de encabezados del inventario (resultado= nro columna)
lHead=[]    # Lista ordenada de nombres de columnas como en inventario, sin las ultimas que no deben trabajarse
Invent= {}  # Todas las lineas del inventario (como indice va el numero de linea) (apunta al string con los datos de la linea de inv.)
OkVals={} # Lista de valores existentes en inventario para cada columna
Cust= {}    # Custodios existentes (apunta a una lista de numeros de linea que el  tiene asignadas)
NroTel= {}  # Numeros de telefono existentes (Apunta a Lin de inv.)
Imei= {}    # Idem ant pero con IMEI
Tecm= []    # Lista de tecnicos
Todo= []    # lista total con custodios, telefonos e IMEI
l= 0
nc= 0
EnTi= str(Pmt("EnTi"))
NoTipo= Pmt("NoTipo").split("|")
with open(Pmt("PathName"),errors="backslashreplace") as a:
    for lin in a:
        if l > 1:
            Col= lin[0:-1].split(";")
            for c in range(nc,len(Col)): Col.pop(nc)
            if len(Col) == nc:
                if Col[0] and (Col[Heads[Pmt("Tipo")]] not in NoTipo):
                    Invent[str(l)]= ";".join(Col)
                    c= Col[Heads[Pmt("NroTel")]]
                    if c != "": 
                        NroTel[c]= str(l)
                        Todo.append(c)
                    c= Col[Heads[Pmt("Imei")]]
                    if c != "":
                        Imei[c]= str(l)
                        Todo.append(c)
                    c= Col[Heads[Pmt("Cust")]]
                    if (c != "") and (EnTi not in c):
                        if c not in Cust: Cust[c]= []
                        Cust[c].append(str(l))
                        Todo.append(c)
                    c= Col[Heads[Pmt("Tecm")]]
                    if (c != "") and (c not in Tecm): 
                        Tecm.append(c)
                    for c in lHead:
                        v=Col[Heads[c]]
                        if not re.findall("^[\s]*$", v):
                            if c not in OkVals: OkVals[c]=[v]
                            elif v not in OkVals[c]: OkVals[c].append(v)
            else:
                win32ui.MessageBox("Cantidad de columnas errones en\nTel: "+Col[Heads[Pmt("NroTel")]]+" IMEI: "+Col[Heads[Pmt("Imei")]], "CUIDADO!!!")
        else:
            Col= lin[0:-1].split(";")
            nm= len(Col)-Pmt("Ultimos")
            for h in Col:
                if (nc <  nm) and (h!=""): 
                    Heads[h]= nc
                    lHead.append(h)
                    nc+= 1
        l+=1
    a.close()
# Creo valores nulo para columnas que no tienen ningun valor
for c in lHead:
    if c not in OkVals: OkVals[c]=[]

# Agrego custodios inexistentes en inventarios que si estan en NAEG
#lNae= 0
#kNae= ""
#with open(Pmt("NAEG")+".csv",errors="backslashreplace") as a:
#    for lin in a:
#        Col= lin[0:-1].split(";")
#        if lNae==0:
#            n=0
#            for c in Col:
#                if c == Pmt("kUbica"): kNae= n
#                n+= 1
#        elif (lNae>0) and (EnTi not in Col[kNae]) and (Col[kNae] not in Cust): # Agrego Cust. no existentes en inv, pero si en NAEG
#            OkVals[Pmt("Cust")].append(Col[kNae])
#            print(OkVals)
#        lNae+=1

# Creo linea en blanco para trabajar
vCols= []
for x in range(0,nc): vCols.append("")

ncp= 0  # Proximo indice a usar en cambios
Cambios= {} # Clave: incremental ncp, datos: Tupla [linea vieja,linea nueva]
# Si ya existian cambios, los leo
if os.path.isfile(Pmt("Cambios")):
    with open(Pmt("Cambios"),errors="backslashreplace") as fp:
        for tx in fp:
            lDat= tx[:-1].split("\t")
            Cambios[ncp]= lDat
            ncp+= 1
    fp.close()

# ------------------------------------------------------------------------------- Listo!
def a_Ti(Col):
    # pone a ti en la lista de columnas con datos de personas
    for c in GetList("Persona"): Col[Heads[c]]= EnTi


def GetpCopy(Ori,ParCols=False,Ti=False,Estan=True):
    # Retorna una lista de columnas de la linea de inv. en la que solo carga ParCols 
    # (lo que esta en la lista del parametro ParCols) con los valores de la Orig
    # Estan: Cuando es Verdadero usa la lista negada de ParCols
    # Ti: Verdadero si pasa las personas a TI
    # Si ParCols es vacio, tomara todas las columnas validas
    Col= Ori
    if type(Ori) == str: Col= Ori.split(";")
    nCol= vCols.copy()
    Lista= lHead
    if ParCols: Lista= GetList(ParCols,Estan)
    for h in Lista:
        c= Heads[h]
        nCol[c]= Col[c]
    if Ti: a_Ti(nCol)
    return nCol


def GetLinea(sKeys,sSep="#_$%"):
    # Con telefono o imei retorna el numero de linea de inventario
    # sKeys es la clave o lista de claves separadas (si se quiere) por sSep
    Encontrado={}
    def incEnc(v):
        if v not in Encontrado:Encontrado[v]= 1
        else: Encontrado[v]+= 1
        
    dList= sKeys.split(sSep)
    for k in dList:
        if re.findall("^\d{15}$", k): # es IMEI ?
            if k in Imei: incEnc(Imei[k])
        elif re.findall("^[^a-zA-Z]+$", k): # Es numero de telefono?
            if k in NroTel: incEnc(NroTel[k])
        elif k!="": # Es un custodio?
            if (k in Cust) and (len(Cust[k]) == 1): incEnc(Cust[k][0])
    l=0
    m=0
    for k in Encontrado:
        if Encontrado[k] > m: 
            m= Encontrado[k]
            l= k
    return l


def SacaChip(Linea):
    # Genero linea de chip
    Col= GetpCopy(Invent[Linea], "Chip", True)
    if Col[Heads[Pmt("NroTel")]]:
        # if Pmt("Smart") in Col[Heads[Pmt("Tipo")]]: Col[Heads[Pmt("Tipo")]]= Pmt("SmartLine")
        # else:Col[Heads[Pmt("Tipo")]]= Pmt("Comun")
        Col[Heads[Pmt("Tipo")]]= Pmt("tLine")+" "+re.sub(r"^[^\s]+\s+([^ ]+).*$",r"\1", Col[Heads[Pmt("Tipo")]])
        Col[Heads[Pmt("Uso")]]= Pmt("Asign")
        Cambio("Nuevo", ";".join(Col))


def Celu_a_Ti(Linea,SacaLinea=True):
    Col= Invent[Linea].split(";")
    Col[Heads[Pmt("Tipo")]]= re.sub(" "+Pmt("ConLinea"),"",Col[Heads[Pmt("Tipo")]])
    Col[Heads[Pmt("Uso")]]= Pmt("Asign")
    a_Ti(Col)
    for c in GetList("SoloCelu"): Col[Heads[c]]= ""
    Cambio("Cambio", ";".join(Col),Invent[Linea])
    if SacaLinea:SacaChip(Linea)


def iSelect(event=""):
    # selecciona un item (linea) del inventario y llama Aplico con dicha linea
    xSep= " - "
    k= gui.GetVal("Dato")
    i= False
    if k in Cust:
        if len(Cust[k]) > 1: i= Cust[k]
        else: i= Cust[k][0]
    elif k in Imei:i= Imei[k]
    elif k in NroTel:i= NroTel[k]
    if i:
        def Aplico(Linea=""):
            global ncp
            if gui.GetVal("Tipo") == "Eliminar": Cambio("Eliminar", Invent[Linea])
            elif gui.GetVal("Tipo") == "Desasignar":
                # Genero linea de equipo
                Celu_a_Ti(Linea)
            elif gui.GetVal("Tipo") == "Cambio":
                global mFrame
                Col=""
                if gui.GetVal("Tipo") == "Nuevo":
                    Col= GetpCopy(vCols, "NoModi", True, False)
                else:
                    Col= Invent[Linea].split(";")
                xFrame= Toplevel(mFrame)
                mod= A5Tk(xFrame,gui.GetVal("Tipo"),Icon=EcoLogo)

                def Modificar(event=""):
                    n= mod.GetVal(Pmt("NroTel"))
                    if n != Col[Heads[Pmt("NroTel")]]:
                        if re.findall("^[^A-Za-z]{4,}$", n): # No elimino chip, lo cambio
                            nn= GetLinea(n)
                            nnCol= Invent[nn].split(";")
                            if nnCol[Heads[Pmt("Imei")]] == "": 
                                Cambio("Eliminar", Invent[nn]) # Elimino la linea de inventario para el chip por que la pongo en un equipo
                            else:
                                # Si el chip se mueve a otro celular, dejo datos
                                if (nnCol[Heads[Pmt("Cust")]] != EnTi) and messagebox.askyesno("Aviso", "Desea usar los datos del que poseia el chip anteriormente?"):  
                                    for c in GetList("NoModi", False): 
                                        mod.SetVal(c, nnCol[Heads[c]])
                                Celu_a_Ti(nn,False) # El celular en el que estaba elchip, (queda sin chip) va a TI
                            Col[Heads[Pmt("NroTel")]]= n
                            n= re.sub("\s+15-", "", n)
                            Col[Heads[Pmt("NroProv")]]= re.sub("^0?", "", n)
                            for c in GetList("NoModi", False): Col[Heads[c]]= mod.GetVal(c)
                            Cambio("Cambio", ";".join(Col),Invent[Linea])
                            SacaChip(Linea)
                        else: Celu_a_Ti(Linea) # Le saco el chip
                    else: # Se modificaron datos aleda#os
                        NewCol= GetpCopy(Col)
                        for c in GetList("NoModi", False): NewCol[Heads[c]]= mod.GetVal(c)
                        c= ";".join(NewCol)
                        if c != Invent[Linea]: Cambio("Cambio", c, Invent[Linea])
                            
                    gui.SetFocus("Dato", True)
                    xFrame.destroy()
                
                ph= False
                nl= 0
                for c in GetList("NoModi", False):
                    if not ph: ph= c
                    mod.Create(c, "e", nl, 0, c, Values= OkVals[c])
                    mod.SetVal(c, Col[Heads[c]])
                    mod.GetObj(c).configure(width=30)
                    nl+= 1
                mod.Create("Modi", "b", nl, 0, "Modificar","m",fBind= Modificar, Span= 2)
                mod.GetObj("Modi").configure(width=50)
                mod.SetFocus(ph,True)
                
        if (type(i) == list) and (len(i) > 1):
            dFrame= Toplevel(mFrame)
            def FindData(event=""):
                Aplico(GetLinea(mdat.GetVal("Sel"),xSep))
                dFrame.destroy()
            mdat= A5Tk(dFrame,gui.GetVal("Tipo"),Icon=EcoLogo)
            rbl=[]
            for x in i:
                Col= Invent[x].split(";")
                rbl.append(Col[Heads[Pmt("NroTel")]]+xSep+Col[Heads[Pmt("Imei")]]+xSep+Col[Heads[Pmt("Tipo")]])
            mdat.Create("Sel", "r", 0, 0, k+".- Seleccione opcion", Values=rbl,Horiz=False)
            mdat.Create("Ok", "b",GrdC=1, Text="Proceder",Values="p",fBind=FindData)
            mdat.GetObj("Ok").configure(width=52)
            mdat.SetFocus("Ok", True)
            #dFrame.bind("<FocusOut>", dFrame.destroy)
        else:
            Aplico(i)
            gui.SetVal("Dato", "")

def Envio(event=" "):
    global Cambios
    global nc
    global gui
    if len(Cambios)>0:
        gui.SetVal("Envio", "Sin Cambios")
        gui.On("Envio",False)
        fName=tempfile.gettempdir()+"\\Cambio_Inventario.xlsx"
        Sale=A5(fName,"Inventario",Create=True)
        Sale.SetCell(1, 1, "Movimiento")
        Sale.SetCell(1, 2, lHead)
        Tec= gui.GetVal("Tecnico")
        Items=0
        nl=2
        for k in sorted(Cambios):
            if Cambios[k][1]:
                Sale.SetCell(nl, 1, Pmt("Actual"))
                Sale.SetCell(nl, 2, Cambios[k][1].split(";"))
                nl+=1
            Col= Cambios[k][2].split(";")
            Col[Heads[Pmt("fCambio")]]= Fecha
            Col[Heads[Pmt("Tecm")]]= Tec
            Sale.SetCell(nl, 1, Cambios[k][0])
            Sale.SetCell(nl, 2, Col)
            nl+= 1
            Items+= 1
        Sale.Background("yellow", 1, 1, 1,Title=True)  
        Sale.Save()
        Cambios={}
        # Envio correo
        olMailItem = 0x0
        obj= win32com.client.Dispatch("Outlook.Application")
        mText= "Cambio_Celulares de "+os.getlogin()+" Cantidad "+str(Items)
        Correo = obj.CreateItem(olMailItem)
        Correo.To= Pmt("Correo")
        #Correo.CC= "Operaciones"
        Correo.Subject= mText
        Correo.Attachments.Add(Source=fName)
        Correo.BodyFormat= 2
        Correo.HTMLBody= mText
        Correo.Display(True)
        if os.path.isfile(Pmt("Cambios")): os.remove(Pmt("Cambios"))
    else:
        messagebox.showwarning("Aviso", "No aplico cambios para enviar")


def Cambio(pMovim=False, Cambiado=False ,Original=""):
    # String de linea tipo inventario a cambiar
    if pMovim and Cambiado:
        global ncp
        Movim= Pmt(pMovim)
        ncp+= 1
        Cambios[ncp]= [Movim,Original,Cambiado]
        # Salvo lista de cambios
        if os.path.isfile(Pmt("Cambios")): os.remove(Pmt("Cambios"))
        Sale= open(Pmt("Cambios"),"w",newline="\n")
        for s in Cambios:
            Sale.write("\t".join(Cambios[s])+"\n")
        Sale.close()

    if len(Cambios) > 0:
        gui.SetVal("Envio", "Enviar "+str(len(Cambios))+" cambios")
        gui.On("Envio",True)
    else:
        gui.SetVal("Envio", "Sin Cambios")
        gui.On("Envio",False)
   
mFrame= Tk()
loginname=os.getlogin()[2:]
gui= A5Tk(mFrame, "Manejo de Celulares v1.0.8", TopLeft=[100,100], Icon=EcoLogo)
gui.Create("Tecnico", "e", 0, 0, "Tecnico", Values=Tecm)
gui.GetObj("Tecnico").configure(width=42)
gui.SetVal("Tecnico", loginname)
gui.Create("Tipo", "r", 1, 1, "Cambio", Values=["Eliminar","Cambio","Desasignar"])
gui.Create("Dato", "e", 2, 0, Text="Buscar", Values=Todo,fBind=iSelect)
gui.GetObj("Dato").configure(width=42)
#gui.Create("Cambio","b", 3, 0,"Aplicar","a", fBind=Cambio)
gui.Create("Envio", "b", 3, 1,Text="Sin Cambios", Values="e", fBind=Envio)
gui.GetObj("Envio").configure(width=35,underline=0)
Cambio()
gui.SetFocus("Tecnico", True)
mFrame.mainloop()


