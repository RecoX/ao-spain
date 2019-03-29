Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim objIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    objIndex = UserList(UserIndex).Invent.Object(i).objIndex
    If objIndex > 0 Then
            If (ObjData(objIndex).ObjType <> OBJTYPE_LLAVES And _
                ObjData(objIndex).ObjType <> OBJTYPE_BARCOS) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If ObjData(objIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(objIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).objIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).objIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).objIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
On Error GoTo errhandler

If Cantidad > 100000 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.objIndex = iORO
            
            If UserList(UserIndex).Flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objIndex).Name)
            
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

'Quita un objeto
UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    UserList(UserIndex).Invent.Object(Slot).objIndex = 0
    UserList(UserIndex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).objIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).objIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.objIndex = 0 Then
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)
        Obj.objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        Obj.Amount = num
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If UserList(UserIndex).Flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.objIndex).Name)
  Else
    Call SendData(ToIndex, UserIndex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.objIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Crea un Objeto
MapData(Map, X, Y).OBJInfo = Obj
Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.objIndex).GrhIndex & "," & X & "," & Y)

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = MiObj.objIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).objIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(ToIndex, UserIndex, 0, "||No podes cargar mas objetos." & FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).objIndex = MiObj.objIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.objIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.objIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.objIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.objIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.objIndex
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If UserList(UserIndex).Flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objIndex).Name)
        End If
        
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No hay nada aqui." & FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario


Dim Obj As ObjData
If Slot = 0 Then Exit Sub
If UserList(UserIndex).Invent.Object(Slot).objIndex = 0 Then Exit Sub

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex)

Select Case Obj.ObjType


    Case OBJTYPE_WEAPON

        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

    Case OBJTYPE_FLECHAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case OBJTYPE_HERRAMIENTAS
    
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
    
    Case OBJTYPE_ARMOUR
        
        Select Case Obj.SubTipo
        
            Case OBJTYPE_ARMADURA
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(UserIndex)
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
            Case OBJTYPE_CASCO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Case OBJTYPE_ESCUDO
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End Select
        
    
End Select

Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub
Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(objIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "HOMBRE"
ElseIf ObjData(objIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean

If ObjData(objIndex).Real = 1 Then
    FaccionPuedeUsarItem = UserList(UserIndex).Faccion.ArmadaReal = 1
Else
    FaccionPuedeUsarItem = 1
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim objIndex As Integer

objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
Obj = ObjData(objIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.ObjType
    Case OBJTYPE_WEAPON
       If ClasePuedeUsarItem(UserIndex, objIndex) And _
          FaccionPuedeUsarItem(UserIndex, objIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    'Animacion por defecto
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                UserList(UserIndex).Invent.WeaponEqpSlot = Slot
                
                'Sonido
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SACARARMA)
        
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                
                Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
    Case OBJTYPE_HERRAMIENTAS
       If ClasePuedeUsarItem(UserIndex, objIndex) And _
          FaccionPuedeUsarItem(UserIndex, objIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = objIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                
       Else
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
    Case OBJTYPE_FLECHAS
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
    
    Case OBJTYPE_ARMOUR
         
         If UserList(UserIndex).Flags.Navegando = 1 Then Exit Sub
         
         Select Case Obj.SubTipo
         
            Case OBJTYPE_ARMADURA
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
                   SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
                   CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) And _
                   FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
                   
                   'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex)
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                    UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                        
                    UserList(UserIndex).Char.Body = Obj.Ropaje
                        
                    UserList(UserIndex).Flags.Desnudo = 0
                        
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    
                    

                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase,genero o raza no puede usar este objeto." & FONTTYPE_INFO)
                End If
            Case OBJTYPE_CASCO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.CascoAnim = NingunCasco
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                    UserList(UserIndex).Invent.CascoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
                End If
            Case OBJTYPE_ESCUDO
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).objIndex) Then
       
                    'Si esta equipado lo quita
                    If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
                    UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                    
                    UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                    
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
                End If
        End Select
End Select

'Actualiza
Call UpdateUserInv(True, UserIndex, 0)


Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

If UserList(UserIndex).Raza = "Humano" Or _
   UserList(UserIndex).Raza = "Elfo" Or _
   UserList(UserIndex).Raza = "Elfo Oscuro" Or _
   UserList(UserIndex).Raza = "Orco" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If



Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'Usa un item del inventario
Dim Obj As ObjData
Dim objIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Exit Sub
End If

objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
UserList(UserIndex).Flags.TargetObjInvIndex = objIndex
UserList(UserIndex).Flags.TargetObjInvSlot = Slot

Select Case Obj.ObjType

    Case OBJTYPE_USEONCE
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        Call AddtoVar(UserList(UserIndex).Stats.MinHam, Obj.MinHam, UserList(UserIndex).Stats.MaxHam)
        UserList(UserIndex).Flags.Hambre = 0
        Call EnviarHambreYsed(UserIndex)
        'Sonido
        SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_COMIDA
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        

            
    Case OBJTYPE_GUITA
    
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).objIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
    Case OBJTYPE_WEAPON
        
        If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
        End If

    
        If ObjData(objIndex).proyectil = 1 Then
            
            
            Call SendData(ToIndex, UserIndex, 0, "T01" & Proyectiles)

           
        Else
        
            If UserList(UserIndex).Flags.TargetObj = 0 Then Exit Sub
            
            TargObj = ObjData(UserList(UserIndex).Flags.TargetObj)
            '¿El target-objeto es leña?
            If TargObj.ObjType = OBJTYPE_LEÑA Then
                    If UserList(UserIndex).Invent.Object(Slot).objIndex = DAGA Then
                        Call TratarDeHacerFogata(UserList(UserIndex).Flags.TargetObjMap _
                             , UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY, UserIndex)
                    
                    Else
             
                    End If
            End If
            
        End If
    Case OBJTYPE_POCIONES
    
        If UserList(UserIndex).Flags.PuedeAtacar = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Debes esperar unos momentos para tomar otra pocion!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Flags.TomoPocion = True
        UserList(UserIndex).Flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(UserIndex).Flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(UserIndex).Flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
            Case 2 'Modif la fuerza
                UserList(UserIndex).Flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), MAXATRIBUTOS)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                AddtoVar UserList(UserIndex).Stats.MinHP, RandomNumber(Obj.MinModificador, Obj.MaxModificador), UserList(UserIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5), UserList(UserIndex).Stats.MaxMAN)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).Flags.Envenenado = 1 Then
                    UserList(UserIndex).Flags.Envenenado = 0
                    Call SendData(ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                   
       End Select
     Case OBJTYPE_BEBIDA
    
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).Flags.Sed = 0
        Call EnviarHambreYsed(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
    
    Case OBJTYPE_LLAVES
        
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).Flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.ObjType = OBJTYPE_PUERTAS Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.Clave = Obj.Clave Then
         
                        MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex _
                        = ObjData(MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex).IndexCerrada
                        UserList(UserIndex).Flags.TargetObj = MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex
                        Call SendData(ToIndex, UserIndex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.Clave = Obj.Clave Then
                        MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex _
                        = ObjData(MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex).IndexCerradaLlave
                        Call SendData(ToIndex, UserIndex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(UserIndex).Flags.TargetObj = MapData(UserList(UserIndex).Flags.TargetObjMap, UserList(UserIndex).Flags.TargetObjX, UserList(UserIndex).Flags.TargetObjY).OBJInfo.objIndex
                        Exit Sub
                     Else
                        Call SendData(ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
    
        Case OBJTYPE_BOTELLAVACIA
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Flags.TargetX, UserList(UserIndex).Flags.TargetY) Then
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.objIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            
        Case OBJTYPE_BOTELLALLENA
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            AddtoVar UserList(UserIndex).Stats.MinAGU, Obj.MinSed, UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).Flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.objIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            
        Case OBJTYPE_HERRAMIENTAS
            
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlProleta, MAXREP)
            
            Select Case objIndex
                Case OBJTYPE_CAÑA
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call SendData(ToIndex, UserIndex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "SFC")

            End Select
        
        Case OBJTYPE_PERGAMINOS
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Flags.Hambre = 0 And _
               UserList(UserIndex).Flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                
            Else
               Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
            End If
       
       Case OBJTYPE_MINERALES
           If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
           End If
           Call SendData(ToIndex, UserIndex, 0, "T01" & FundirMetal)
       
       Case OBJTYPE_INSTRUMENTOS
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Obj.Snd1)
       
       Case OBJTYPE_BARCOS
           UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
           UserList(UserIndex).Invent.BarcoSlot = Slot
           Call DoNavega(UserIndex, Obj)
           
End Select

'Actualiza
Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(True, UserIndex, 0)

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) \ ModHerreriA(UserList(UserIndex).Clase) Then
        If ObjData(ArmasHerrero(i)).ObjType = OBJTYPE_WEAPON Then
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).MinHIT & "/" & ObjData(ArmasHerrero(i)).MaxHIT & ")" & "," & ArmasHerrero(i) & ","
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(ToIndex, UserIndex, 0, "LAH" & cad$)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(UserIndex).Clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(ToIndex, UserIndex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then _
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).MinDef & "/" & ObjData(ArmadurasHerrero(i)).MaxDef & ")" & "," & ArmadurasHerrero(i) & ","
Next i

Call SendData(ToIndex, UserIndex, 0, "LAR" & cad$)

End Sub

                   

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

Call TirarTodosLosItems(UserIndex)
Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

ItemSeCae = ObjData(Index).Real <> 1 And _
            ObjData(Index).Caos <> 1 And _
            ObjData(Index).ObjType <> OBJTYPE_LLAVES And _
            ObjData(Index).ObjType <> OBJTYPE_BARCOS


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

'Call LogTarea("Sub TirarTodosLosItems")

Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS

  ItemIndex = UserList(UserIndex).Invent.Object(i).objIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.objIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
  
Next i

End Sub


Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
  ItemIndex = UserList(UserIndex).Invent.Object(i).objIndex
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                Tilelibre UserList(UserIndex).Pos, NuevaPos
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.objIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
         End If
         
  End If
Next i

End Sub



