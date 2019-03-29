Attribute VB_Name = "mdlCOmercioConUsuario"
'Modulo para comerciar con otro usuario
'Por Alejo (Alejandro Santos)
'
'
'[Alejo]

Option Explicit

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    Cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(Origen As Integer, Destino As Integer)
On Error GoTo errhandler

'Actualiza el inventario del usuario
Call UpdateUserInv(True, Origen, 0)
'Decirle al origen que abra la ventanita.
Call SendData(ToIndex, Origen, 0, "INITCOMUSU")
UserList(Origen).Flags.Comerciando = True

'si es el receptor, enviamos el objeto del otro usu
'If UserList(UserList(Origen).ComUsu.DestUsu).ComUsu.DestUsu = Origen Then
If UserList(Origen).ComUsu.DestUsu = Destino Then
    Call EnviarObjetoTransaccion(Origen)
End If

Exit Sub
errhandler:

End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(AQuien As Integer)
'Dim Object As UserOBJ
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

'Object.Amount = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    'Object.ObjIndex = iORO
    ObjInd = iORO
Else
    'Object.ObjIndex = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

'If Object.ObjIndex > 0 And Object.Amount > 0 Then
'    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
'    & ObjData(Object.ObjIndex).ObjType & "," _
'    & ObjData(Object.ObjIndex).MaxHIT & "," _
'    & ObjData(Object.ObjIndex).MinHIT & "," _
'    & ObjData(Object.ObjIndex).MaxDef & "," _
'    & ObjData(Object.ObjIndex).Valor \ 3)
'End If
If ObjInd > 0 And ObjCant > 0 Then
    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & ObjInd & "," & ObjData(ObjInd).Name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).ObjType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3)
End If

End Sub

Public Sub FinComerciarUsu(UserIndex As Integer)
UserList(UserIndex).ComUsu.Acepto = False
UserList(UserIndex).ComUsu.Cant = 0
UserList(UserIndex).ComUsu.DestUsu = 0
UserList(UserIndex).ComUsu.Objeto = 0

UserList(UserIndex).Flags.Comerciando = False

Call SendData(ToIndex, UserIndex, 0, "FINCOMUSUOK")
End Sub

Public Sub AceptarComercioUsu(UserIndex As Integer)
If UserList(UserIndex).ComUsu.DestUsu <= 0 Or _
    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
    Exit Sub
End If

UserList(UserIndex).ComUsu.Acepto = True

If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    Call SendData(ToIndex, UserIndex, 0, "||El otro usuario aun no ha aceptado tu oferta." & FONTTYPE_TALK)
    Exit Sub
End If

Dim Obj1 As Obj, Obj2 As Obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False
OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu

'[Alejo]: Creo haber podido erradicar el bug de
'         no poder comerciar con mas de 32k de oro.
'         Las lineas comentadas en los siguientes
'         2 grandes bloques IF (4 lineas) son las
'         que originaban el problema.

If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = iORO
    'If Obj1.Amount > UserList(UserIndex).Stats.GLD Then
    If UserList(UserIndex).ComUsu.Cant > UserList(UserIndex).Stats.GLD Then
        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = iORO
    'If Obj2.Amount > UserList(OtroUserIndex).Stats.GLD Then
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

'[CORREGIDO]
'Desde acá corregí el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserStatsBox(OtroUserIndex)
    'y se la doy al otro
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserStatsBox(UserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(UserIndex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If

'pone el oro directamente en la billetera
If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Cant
    Call SendUserStatsBox(UserIndex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Cant
    Call SendUserStatsBox(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, UserIndex)
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]

