Attribute VB_Name = "Comercio"
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
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


Sub UserCompraObj(ByVal UserIndex As Integer, ByVal objIndex As Integer, ByVal NpcIndex As Integer, ByVal Cantidad As Integer)
Dim infla As Integer
Dim Descuento As String
Dim unidad As Long, monto As Long
Dim Slot As Integer
Dim obji As Integer
Dim Encontre As Boolean

If (Npclist(UserList(UserIndex).Flags.TargetNpc).Invent.Object(objIndex).Amount <= 0) Then Exit Sub

obji = Npclist(UserList(UserIndex).Flags.TargetNpc).Invent.Object(objIndex).objIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = obji And _
   UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).objIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
    
    'Le sustraemos el valor en oro del obj comprado
    infla = (Npclist(NpcIndex).Inflacion * ObjData(obji).Valor) \ 100
    Descuento = UserList(UserIndex).Flags.Descuento
    If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
    unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(objIndex).objIndex).Valor + infla) / Descuento)
    monto = unidad * Cantidad
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto
    
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(UserIndex, Comerciar)
    
    If ObjData(obji).ObjType = OBJTYPE_LLAVES Then Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(obji).Name)

'    If UserList(UserIndex).Stats.GLD < 0 Then UserList(UserIndex).Stats.GLD = 0
    
    Call QuitarNpcInvItem(UserList(UserIndex).Flags.TargetNpc, CByte(objIndex), Cantidad)
Else
    Call SendData(ToIndex, UserIndex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
End If


End Sub


Sub NpcCompraObj(ByVal UserIndex As Integer, ByVal objIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer
Dim NpcIndex As Integer
Dim infla As Long
Dim monto As Long
      
If Cantidad < 1 Then Exit Sub

NpcIndex = UserList(UserIndex).Flags.TargetNpc
obji = UserList(UserIndex).Invent.Object(objIndex).objIndex

If ObjData(obji).Newbie = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||No comercio objetos para newbies." & FONTTYPE_INFO)
    Exit Sub
End If

If Npclist(NpcIndex).TipoItems <> OBJTYPE_CUALQUIERA Then
    '¿Son los items con los que comercia el npc?
    If Npclist(NpcIndex).TipoItems <> ObjData(obji).ObjType Then
            Call SendData(ToIndex, UserIndex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
            Exit Sub
    End If
End If

'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until Npclist(NpcIndex).Invent.Object(Slot).objIndex = obji And _
         Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_INVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until Npclist(NpcIndex).Invent.Object(Slot).objIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
'                Call SendData(ToIndex, NpcIndex, 0, "||El npc no puede cargar mas objetos." & FONTTYPE_INFO)
'                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
        
        
End If

If Slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        Npclist(NpcIndex).Invent.Object(Slot).objIndex = obji
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(objIndex), Cantidad)
        'Le sumamos al user el valor en oro del obj vendido
        monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
        Call AddtoVar(UserList(UserIndex).Stats.GLD, monto, MAXORO)
        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(UserIndex, Comerciar)
    
    Else
        Call SendData(ToIndex, UserIndex, 0, "||El npc no puede cargar tantos objetos." & FONTTYPE_INFO)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(objIndex), Cantidad)
    'Le sumamos al user el valor en oro del obj vendido
    monto = ((ObjData(obji).Valor \ 3 + infla) * Cantidad)
    Call AddtoVar(UserList(UserIndex).Stats.GLD, monto, MAXORO)
End If

End Sub

Sub IniciarCOmercioNPC(ByVal UserIndex As Integer)
On Error GoTo errhandler

'Mandamos el Inventario
Call EnviarNpcInv(UserIndex, UserList(UserIndex).Flags.TargetNpc)
'Hacemos un Update del inventario del usuario
Call UpdateUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
Call SendUserStatsBox(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData ToIndex, UserIndex, 0, "INITCOM"
UserList(UserIndex).Flags.Comerciando = True

errhandler:

End Sub

Sub NPCVentaItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)
On Error GoTo errhandler

Dim infla As Long
Dim val As Long
Dim Desc As String

If Cantidad < 1 Then Exit Sub

'NPC VENDE UN OBJ A UN USUARIO
Call SendUserStatsBox(UserIndex)

'Calculamos el valor unitario
infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).Valor) / 100
Desc = Descuento(UserIndex)
If Desc = 0 Then Desc = 1 'evitamos dividir por 0!
val = (ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).Valor + infla) / Desc
        


If UserList(UserIndex).Stats.GLD >= (val * Cantidad) Then
       
       If Npclist(UserList(UserIndex).Flags.TargetNpc).Invent.Object(i).Amount > 0 Then
            If Cantidad > Npclist(UserList(UserIndex).Flags.TargetNpc).Invent.Object(i).Amount Then Cantidad = Npclist(UserList(UserIndex).Flags.TargetNpc).Invent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).Flags.TargetNpc, Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call SendUserStatsBox(UserIndex)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).Flags.TargetNpc)
            Call UpdateVentanaComercio(i, 0, UserIndex)
        
       End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente dinero." & FONTTYPE_INFO)
    Exit Sub
End If


errhandler:

End Sub
Sub NPCCompraItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo errhandler

'NPC COMPRA UN OBJ A UN USUARIO
Call SendUserStatsBox(UserIndex)
   
If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call NpcCompraObj(UserIndex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call SendUserStatsBox(UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).Flags.TargetNpc)
            'Actualizamos la ventana de comercio
            
            Call UpdateVentanaComercio(Item, 1, UserIndex)
            
End If

errhandler:

End Sub


Sub UpdateVentanaComercio(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
 
 
 Call SendData(ToIndex, UserIndex, 0, "TRANSOK" & Slot & "," & NpcInv)
 
End Sub

Function Descuento(ByVal UserIndex As Integer) As String

'Establece el descuento en funcion del skill comercio
Dim PtsComercio As Integer
PtsComercio = UserList(UserIndex).Stats.UserSkills(Comerciar)

If PtsComercio <= 10 And PtsComercio > 5 Then
    UserList(UserIndex).Flags.Descuento = 1.1
    Descuento = 1.1
ElseIf PtsComercio <= 20 And PtsComercio >= 12 Then
    UserList(UserIndex).Flags.Descuento = 1.2
    Descuento = 1.2
ElseIf PtsComercio <= 30 And PtsComercio >= 19 Then
    UserList(UserIndex).Flags.Descuento = 1.3
    Descuento = 1.3
ElseIf PtsComercio <= 40 And PtsComercio >= 29 Then
    UserList(UserIndex).Flags.Descuento = 1.4
    Descuento = 1.4
ElseIf PtsComercio <= 50 And PtsComercio >= 39 Then
    UserList(UserIndex).Flags.Descuento = 1.5
    Descuento = 1.5
ElseIf PtsComercio <= 60 And PtsComercio >= 49 Then
    UserList(UserIndex).Flags.Descuento = 1.6
    Descuento = 1.6
ElseIf PtsComercio <= 70 And PtsComercio >= 59 Then
    UserList(UserIndex).Flags.Descuento = 1.7
    Descuento = 1.7
ElseIf PtsComercio <= 80 And PtsComercio >= 69 Then
    UserList(UserIndex).Flags.Descuento = 1.8
    Descuento = 1.8
ElseIf PtsComercio <= 99 And PtsComercio >= 79 Then
    UserList(UserIndex).Flags.Descuento = 1.9
    Descuento = 1.9
ElseIf PtsComercio <= 999999 And PtsComercio >= 99 Then
    UserList(UserIndex).Flags.Descuento = 2
    Descuento = 2
Else
    UserList(UserIndex).Flags.Descuento = 0
    Descuento = 0
End If

End Function



Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

'Enviamos el inventario del npc con el cual el user va a comerciar...
Dim i As Integer
Dim infla As Long
Dim Desc As String
Dim val As Long
Desc = Descuento(UserIndex)
If Desc = 0 Then Desc = 1 'evitamos dividir por 0!

For i = 1 To MAX_INVENTORY_SLOTS
  If Npclist(NpcIndex).Invent.Object(i).objIndex > 0 Then
        'Calculamos el porc de inflacion del npc
        infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).Valor) / 100
        val = (ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).Valor + infla) / Desc
        SendData ToIndex, UserIndex, 0, "NPCI" & _
        ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).Name _
        & "," & Npclist(NpcIndex).Invent.Object(i).Amount & _
        "," & val _
        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).GrhIndex _
        & "," & Npclist(NpcIndex).Invent.Object(i).objIndex _
        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).ObjType _
        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).MaxHIT _
        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).MinHIT _
        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).objIndex).MaxDef
        
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(1) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(2) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(3) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(4) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(5) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(6) _
'        & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).ClaseProhibida(7)
  Else
        SendData ToIndex, UserIndex, 0, "NPCI" & _
        "Nada" _
        & "," & 0 & _
        "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0 _
        & "," & 0
  End If
  
Next
End Sub


