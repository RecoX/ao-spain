Attribute VB_Name = "General"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Sub DarCuerpoDesnudo(ByVal userindex As Integer)

Select Case UCase$(UserList(userindex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 21
                Case "MUJER"
                     UserList(userindex).Char.Body = 39
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 32
                Case "MUJER"
                     UserList(userindex).Char.Body = 40
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 53
                Case "MUJER"
                     UserList(userindex).Char.Body = 60
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 53
                Case "MUJER"
                     UserList(userindex).Char.Body = 60
      End Select
    Case "ORCO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 232
                Case "MUJER"
                     UserList(userindex).Char.Body = 234
      End Select
    Case "HOBBIT"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 239
                Case "MUJER"
                     UserList(userindex).Char.Body = 240
      End Select
      Case Else
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                     UserList(userindex).Char.Body = 21
                Case "MUJER"
                     UserList(userindex).Char.Body = 39
      End Select


    
End Select

UserList(userindex).Flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function




Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer

For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i



End Sub

Sub EnviarSpawnList(ByVal userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(ToIndex, userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal port As Integer)

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = port
Obj.Backlog = 5
Obj.Listen

End Sub




Sub Main()
On Error Resume Next


Call LoadMotd

Prision.Map = 66
Libertad.Map = 66

Prision.X = 75
Prision.Y = 47
Libertad.X = 75
Libertad.Y = 65


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")



ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer


IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "Orco"
ListaRazas(7) = "Hobbit"

ReDim ListaClases(1 To NUMCLASES) As String

ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"


ReDim UserSkills(1 To NUMSKILLS) As Integer

ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"


frmCargando.Show



Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList
Call CargarForbidenWords
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Server.ini"

Call LoadSini

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next LoopC

frmMain.AutoSave.Enabled = True




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets
frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048



Call ConfigListeningSocket(frmMain.Socket1, puerto)


If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #n

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

'ResetThread.CreateNewThread AddressOf ThreadResetActions, tpNormal

'Call MainThread


End Sub



Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

On Error Resume Next
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function
Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid$(Text, LastPos + 1)
End If

End Function
Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub

Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile

Print #nfile, "<stats>"
Print #nfile, "<ao>"
Print #nfile, "<dia>" & Date & "</dia>"
Print #nfile, "<hora>" & Time & "</hora>"
Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
Print #nfile, "</ao>"
Print #nfile, "</stats>"


Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub
Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Integer
  
frmMain.Socket1.Cleanup
frmMain.Socket1.Startup
  
frmMain.Socket2(0).Cleanup
frmMain.Socket2(0).Startup

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos


'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = Val(puerto)
frmMain.Socket1.Listen

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " servidor reiniciado."
Close #n

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal userindex As Integer) As Boolean
    
    If MapInfo(UserList(userindex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger <> 1 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger <> 2 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub EfectoLluvia(ByVal userindex As Integer)
On Error GoTo errhandler


If UserList(userindex).Flags.UserLogged Then
    If Intemperie(userindex) Then
                Dim modifi As Long
                modifi = Porcentaje(UserList(userindex).Stats.MaxSta, 3)
                Call QuitarSta(userindex, modifi)
                Call SendData(ToIndex, userindex, 0, "||¡¡Has perdido stamina, busca pronto refugio de la lluvia!!." & FONTTYPE_INFO)
                Call SendUserStatsBox(userindex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userindex As Integer)

Dim modifi As Integer

If UserList(userindex).Counters.Frio < IntervaloFrio Then
  UserList(userindex).Counters.Frio = UserList(userindex).Counters.Frio + 1
Else
  If MapInfo(UserList(userindex).Pos.Map).Terreno = Nieve Then
    Call SendData(ToIndex, userindex, 0, "||¡¡Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
    modifi = Porcentaje(UserList(userindex).Stats.MaxHP, 5)
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - modifi
    If UserList(userindex).Stats.MinHP < 1 Then
            Call SendData(ToIndex, userindex, 0, "||¡¡Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(userindex).Stats.MinHP = 0
            Call UserDie(userindex)
    End If
  Else
    modifi = Porcentaje(UserList(userindex).Stats.MaxSta, 5)
    Call QuitarSta(userindex, modifi)
    Call SendData(ToIndex, userindex, 0, "||¡¡Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(userindex).Counters.Frio = 0
  Call SendUserStatsBox(userindex)
End If

End Sub


Public Sub EfectoInvisibilidad(ByVal userindex As Integer)

If UserList(userindex).Counters.Invisibilidad < IntervaloInvisible Then
  UserList(userindex).Counters.Invisibilidad = UserList(userindex).Counters.Invisibilidad + 1
Else
  Call SendData(ToIndex, userindex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
  UserList(userindex).Counters.Invisibilidad = 0
  UserList(userindex).Flags.Invisible = 0
  Call SendData(ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
End If
            
End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).Flags.Paralizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal userindex As Integer)

If UserList(userindex).Counters.Ceguera > 0 Then
    UserList(userindex).Counters.Ceguera = UserList(userindex).Counters.Ceguera - 1
Else
    If UserList(userindex).Flags.Ceguera = 1 Then
        UserList(userindex).Flags.Ceguera = 0
        Call SendData(ToIndex, userindex, 0, "NSEGUE")
    Else
        UserList(userindex).Flags.Estupidez = 0
        Call SendData(ToIndex, userindex, 0, "NESTUP")
    End If
    
End If


End Sub


Public Sub EfectoParalisisUser(ByVal userindex As Integer)

If UserList(userindex).Counters.Paralisis > 0 Then
    UserList(userindex).Counters.Paralisis = UserList(userindex).Counters.Paralisis - 1
Else
    UserList(userindex).Flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(ToIndex, userindex, 0, "PARADOK")
End If

End Sub
Public Sub RecStamina(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 1 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 2 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 4 Then Exit Sub
       
      
Dim massta As Integer
If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta Then
   If UserList(userindex).Counters.STACounter < Intervalo Then
       UserList(userindex).Counters.STACounter = UserList(userindex).Counters.STACounter + 1
   Else
       UserList(userindex).Counters.STACounter = 0
       massta = CInt(RandomNumber(1, Porcentaje(UserList(userindex).Stats.MaxSta, 5)))
       UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta + massta
       If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
           Call SendData(ToIndex, userindex, 0, "||Te sentis menos cansado." & FONTTYPE_INFO)
           EnviarStats = True
       End If
End If

End Sub

Public Sub EfectoVeneno(userindex As Integer, EnviarStats As Boolean)
Dim n As Integer

If UserList(userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(userindex).Counters.Veneno = UserList(userindex).Counters.Veneno + 1
Else
  Call SendData(ToIndex, userindex, 0, "||Estas envenenado, si no te curas moriras." & FONTTYPE_VENENO)
  UserList(userindex).Counters.Veneno = 0
  n = RandomNumber(1, 5)
  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - n
  If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
  EnviarStats = True
End If

End Sub

Public Sub DuracionPociones(userindex As Integer)

'Controla la duracion de las pociones
If UserList(userindex).Flags.DuracionEfecto > 0 Then
   UserList(userindex).Flags.DuracionEfecto = UserList(userindex).Flags.DuracionEfecto - 1
   If UserList(userindex).Flags.DuracionEfecto = 0 Then
        UserList(userindex).Flags.TomoPocion = False
        UserList(userindex).Flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(userindex).Stats.UserAtributos(loopX) = UserList(userindex).Stats.UserAtributosBackUP(loopX)
        Next
   End If
End If

End Sub

Public Sub HambreYSed(userindex As Integer, fenviarAyS As Boolean)
'Sed
If UserList(userindex).Stats.MinAGU > 0 Then
    If UserList(userindex).Counters.AGUACounter < IntervaloSed Then
          UserList(userindex).Counters.AGUACounter = UserList(userindex).Counters.AGUACounter + 1
    Else
          UserList(userindex).Counters.AGUACounter = 0
          UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
                            
          If UserList(userindex).Stats.MinAGU <= 0 Then
               UserList(userindex).Stats.MinAGU = 0
               UserList(userindex).Flags.Sed = 1
          End If
                            
          fenviarAyS = True
                            
    End If
End If

'hambre
If UserList(userindex).Stats.MinHam > 0 Then
   If UserList(userindex).Counters.COMCounter < IntervaloHambre Then
        UserList(userindex).Counters.COMCounter = UserList(userindex).Counters.COMCounter + 1
   Else
        UserList(userindex).Counters.COMCounter = 0
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 10
        If UserList(userindex).Stats.MinHam < 0 Then
               UserList(userindex).Stats.MinHam = 0
               UserList(userindex).Flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 1 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 2 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).Trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
   If UserList(userindex).Counters.HPCounter < Intervalo Then
      UserList(userindex).Counters.HPCounter = UserList(userindex).Counters.HPCounter + 1
   Else
      mashit = CInt(RandomNumber(2, Porcentaje(UserList(userindex).Stats.MaxSta, 5)))
                           
      UserList(userindex).Counters.HPCounter = 0
      UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + mashit
      If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
         Call SendData(ToIndex, userindex, 0, "||Has sanado." & FONTTYPE_INFO)
         EnviarStats = True
      End If
End If

End Sub




