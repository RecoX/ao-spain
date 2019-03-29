Attribute VB_Name = "Trabajo"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Exit Sub
End If

If UCase(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res > 9 Then
   UserList(UserIndex).Flags.Oculto = 0
   UserList(UserIndex).Flags.Invisible = 0
   Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
   Call SendData(ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub
Public Sub DoOcultarse(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 91 Then
                    Suerte = 7
End If

If UCase(UserList(UserIndex).Clase) <> "LADRON" And UCase(UserList(UserIndex).Clase) <> "GUERRERO" And UCase(UserList(UserIndex).Clase) <> "CAZADOR" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res <= 5 Then
   UserList(UserIndex).Flags.Oculto = 1
   UserList(UserIndex).Flags.Invisible = 1
   Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
   Call SendData(ToIndex, UserIndex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
   Call SubirSkill(UserIndex, Ocultarse)
Else
   Call SendData(ToIndex, UserIndex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

'[Modificar]
Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData)
If UserList(UserIndex).Flags.Cabalgando = 1 Then UserList(UserIndex).Flags.Cabalgando = 0
Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
    
    If UserList(UserIndex).Flags.Muerto = 0 Then
        UserList(UserIndex).Char.Body = Barco.Ropaje
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).Flags.Navegando = 1
    
Else
    
    UserList(UserIndex).Flags.Navegando = 0
    
    If UserList(UserIndex).Flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If

End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendData(ToIndex, UserIndex, 0, "NAVEG")

End Sub
'[Efestos]
Public Sub DoCabalga(ByVal UserIndex As Integer, ByRef Caballo As ObjData)
If UserList(UserIndex).Flags.Navegando = 1 Then UserList(UserIndex).Flags.Navegando = 0

Dim ModCaballo As Long
ModCaballo = ModEquitacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(Equitacion) / ModCaballo < Caballo.MinSkill Then
    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este caballo." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Para usar este caballo necesitas " & Caballo.MinSkill * ModCaballo & " puntos en equitacion." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Flags.Cabalgando = 0 Then
    
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
    
    If UserList(UserIndex).Flags.Muerto = 0 Then
        UserList(UserIndex).Char.Body = Caballo.Ropaje
'    Else
'        UserList(UserIndex).Char.Body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).Flags.Cabalgando = 1
    
Else
    
    UserList(UserIndex).Flags.Cabalgando = 0
    
    If UserList(UserIndex).Flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If

End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendData(ToIndex, UserIndex, 0, "CABAL")

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(UserIndex).Flags.TargetObjInvIndex > 0 Then
   If ObjData(UserList(UserIndex).Flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call SendData(ToIndex, UserIndex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Long, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        'Si tiene el OBJ Sige por aca y le saca 1
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
        Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
               
        If UserList(UserIndex).Invent.Object(i).Amount = 0 Then
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
                QuitarObjetos = True
                Exit Function
        End If
        
        If UserList(UserIndex).Invent.Object(i).Amount < 1 Then
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        If ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex Then
            Exit For
            'Esto es para que si saca una vez no lo haga de nuevo
        End If
    End If
Next i
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)
'[Efestos]
    Dim c As Integer
    c = ObjData(ItemIndex).Madera * Cantidad
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, c, UserIndex)
'[Efestos]
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer) As Boolean
'[Efestos]
Dim c As Long
    If ObjData(ItemIndex).Madera > 0 Then
            c = val(ObjData(ItemIndex).Madera) * Cantidad
            If Not TieneObjetos(Leña, c, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes madera." & FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True
'[Efestos]
End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de hierro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de plata." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de oro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function



Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(UserIndex, ItemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).ObjType = OBJTYPE_WEAPON Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el arma!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ESCUDO Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el escudo!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_CASCO Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el casco!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).ObjType = OBJTYPE_ARMOUR Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido la armadura!." & FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If

End Sub
Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)

If CarpinteroTieneMateriales(UserIndex, ItemIndex, Cantidad) And _
   UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria Then

    Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, Cantidad)
    Call SendData(ToIndex, UserIndex, 0, "||Has construido el objeto!" & FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = Cantidad            '[Efestos]
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

End Sub

Public Sub DoLingotes(ByVal UserIndex As Integer)
'    Call LogTarea("Sub DoLingotes")
    If UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount < 5 Then
              Call SendData(ToIndex, UserIndex, 0, "||No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
              Exit Sub
    End If
    
    If RandomNumber(1, ObjData(UserList(UserIndex).Flags.TargetObjInvIndex).MinSkill) < 10 Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount - 5
                If UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount < 1 Then
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount = 0
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).ObjIndex = 0
                End If
                Call SendData(ToIndex, UserIndex, 0, "||Has obtenido un lingote!!!" & FONTTYPE_INFO)
                Dim nPos As WorldPos
                Dim MiObj As Obj
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(UserList(UserIndex).Flags.TargetObjInvIndex).LingoteIndex
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
                End If
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Flags.TargetObjInvSlot)
                Call SendData(ToIndex, UserIndex, 0, "||¡Has obtenido un lingote!" & FONTTYPE_INFO)
    Else
        
        UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount - 5
        If UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount < 1 Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).Amount = 0
                UserList(UserIndex).Invent.Object(UserList(UserIndex).Flags.TargetObjInvSlot).ObjIndex = 0
        End If
        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Flags.TargetObjInvSlot)
        Call SendData(ToIndex, UserIndex, 0, "||Los minerales no eran de buena calidad, no has logrado hacer un lingote." & FONTTYPE_INFO)
    End If
    
End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function
'[Efestos]
Function ModEquitacion(ByVal Clase As String) As Integer

Select Case UCase(Clase)
    Case "GLADIADOR"
        ModEquitacion = 1
    Case Else
        ModEquitacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As String) As Integer

Select Case UCase(Clase)
    Case "MINERO"        '[Efestos] No modificar estas cifras sino da error
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As String) As Integer

Select Case UCase(Clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

Select Case UCase(Clase)
    Case "HERRERO"
        ModHerreriA = 1        '[Efestos] NO CAMBIAR!!!!!!!
    Case "MINERO"
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
Select Case UCase(Clase)
    Case "DRUIDA"
        ModDomar = 6
    Case "CAZADOR"
        ModDomar = 6
    Case "CLERIGO"
        ModDomar = 7
    Case Else
        ModDomar = 10
End Select
End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
CalcularPoderDomador = _
UserList(UserIndex).Stats.UserAtributos(Carisma) * _
(UserList(UserIndex).Stats.UserSkills(Domar) / ModDomar(UserList(UserIndex).Clase)) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3)
End Function
Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'Call LogTarea("Sub FreeMascotaIndex")
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) = 0 Then
        FreeMascotaIndex = j
        Exit Function
    End If
Next j
End Function
Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).Flags.Domable <= CalcularPoderDomador(UserIndex) Then
        Dim Index As Integer
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        Index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(Index) = NpcIndex
        UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(ToIndex, UserIndex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
        Call SubirSkill(UserIndex, Domar)
        
    Else
    
        Call SendData(ToIndex, UserIndex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
        
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).Flags.AdminInvisible = 0 Then
        
        UserList(UserIndex).Flags.AdminInvisible = 1
        UserList(UserIndex).Flags.Invisible = 1
        UserList(UserIndex).Flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).Flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).Flags.AdminInvisible = 0
        UserList(UserIndex).Flags.Invisible = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).Flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).Flags.OldHead
        
    End If
    
    
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    
End Sub
Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj

If Not LegalPos(Map, X, Y) Then Exit Sub

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(ToIndex, UserIndex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount / 3
    
    If Obj.Amount > 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Has hecho una fogata." & FONTTYPE_INFO)
    End If
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.X = X
    Fogatita.Y = Y
    Call TrashCollector.Add(Fogatita)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(UserIndex).Clase = "Pescador" Or UserList(UserIndex).Clase = "Aldeano" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Pesca) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
        
    If UserList(UserIndex).Clase = "Pescador" Then
        MiObj.Amount = RandomNumber(1, 5)
    ElseIf UserList(UserIndex).Clase = "Aldeano" Then 'Neptuno
        MiObj.Amount = RandomNumber(1, 4)
    Else
        MiObj.Amount = 1
    End If
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Pesca)


Exit Sub

errhandler:
    Call LogError("Error en DoPescar")

End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub DoRobar")

If MapInfo(UserList(VictimaIndex).Pos.Map).Pk = 1 Then Exit Sub

If UserList(VictimaIndex).Flags.Privilegios < 2 Then
    Dim Suerte As Integer
    Dim res As Integer
    
       
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UCase(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim n As Integer
                
                n = RandomNumber(1, 100)
                
                If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, n, MAXORO)
                
                Call SendData(ToIndex, LadrOnIndex, 0, "||Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(ToIndex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)
    End If

    If Not Criminal(LadrOnIndex) Then
            Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    Call AddtoVar(UserList(LadrOnIndex).Reputacion.LadronesRep, vlLadron, MAXREP)
    Call SubirSkill(LadrOnIndex, Robar)

End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).ObjType <> OBJTYPE_LLAVES And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).ObjType <> OBJTYPE_CABALLOS And _
ObjData(OI).ObjType <> OBJTYPE_BARCOSARMADA
End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 3 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - (daño * 1.5)
        Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & (daño * 2) & FONTTYPE_FIGHT)
        Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(UserIndex).Name & " por " & (daño * 2) & FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - (daño * 2)
        If Npclist(VictimNpcIndex).Stats.MinHP < 0 Then Npclist(VictimNpcIndex).Stats.MinHP = 0
        Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado la criatura por " & (daño * 2) & " (" & Npclist(VictimNpcIndex).Stats.MinHP & "/" & Npclist(VictimNpcIndex).Stats.MaxHP & ")" & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apuñalar)
    End If
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT)
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(UserIndex).Clase = "Leñador" Or UserList(UserIndex).Clase = "Aldeano" Then 'Neptuno
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Talar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(UserIndex).Clase = "Leñador" Then
        MiObj.Amount = RandomNumber(1, 5)
    ElseIf UserList(UserIndex).Clase = "Aldeano" Then 'Neptuno
        MiObj.Amount = RandomNumber(1, 4)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Talar)

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

If UserList(UserIndex).Flags.Privilegios < 2 Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlASALTO, MAXREP)
End Sub


Public Sub DoPlayInstrumento(ByVal UserIndex As Integer)

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(UserIndex).Clase = "Minero" Or UserList(UserIndex).Clase = "Aldeano" Then 'neptuno
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Mineria) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 91 Then
                    Suerte = 7
End If

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).Flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).Flags.TargetObj).MineralIndex
    
    If UserList(UserIndex).Clase = "Minero" Then
        MiObj.Amount = RandomNumber(8, 12)
    ElseIf UserList(UserIndex).Clase = "Aldeano" Then 'Neptuno
        MiObj.Amount = RandomNumber(6, 9)
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call SendData(ToIndex, UserIndex, 0, "||¡Has extraido algunos minerales!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||¡No has conseguido nada!" & FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Mineria)


Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer
Dim pinga As Long
'[Efestos]
If UserList(UserIndex).Stats.UserSkills(Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= -1 Then
                    Suerte = 100
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 11 Then
                    Suerte = 90
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 21 Then
                    Suerte = 80
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 31 Then
                    Suerte = 70
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 41 Then
                    Suerte = 60
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 51 Then
                    Suerte = 50
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 61 Then
                    Suerte = 40
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 71 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 81 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 91 Then
                    Suerte = 10
End If
'[Efestos]
'If UserList(UserIndex).Stats.UserSkills(Meditar) <= 10 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= -1 Then
'                    Suerte = 35
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 20 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 11 Then
'                    Suerte = 30
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 30 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 21 Then
'                    Suerte = 28
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 40 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 31 Then
'                    Suerte = 24
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 50 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 41 Then
'                    Suerte = 22
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 60 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 51 Then
'                    Suerte = 20
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 70 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 61 Then
'                    Suerte = 18
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 80 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 71 Then
'                    Suerte = 15
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 90 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 81 Then
'                    Suerte = 10
'ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 100 _
'   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 91 Then
'                    Suerte = 5
'End If

For pinga = 1 To 5
If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call SendData(ToIndex, UserIndex, 0, "||Has terminado de meditar." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
    UserList(UserIndex).Flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 4)
    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Cant, UserList(UserIndex).Stats.MaxMAN)
    Call SendData(ToIndex, UserIndex, 0, "||¡Has recuperado " & Cant & " puntos de mana!" & FONTTYPE_INFO)
    Call SendUserStatsBox(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If
Next
End Sub




