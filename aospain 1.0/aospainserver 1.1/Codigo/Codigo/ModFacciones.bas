Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos


Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public Const ExpAlUnirse = 100000
Public Const ExpX100 = 100000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 50 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 50 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 18 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 18!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
           If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
              UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.objIndex = TunicaMagoImperialEnanos
           Else
                  MiObj.objIndex = TunicaMagoImperial
           End If
    ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Or _
           UCase$(UserList(UserIndex).Clase) = "PALADIN" Or _
           UCase$(UserList(UserIndex).Clase) = "BANDIDO" Or _
           UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.objIndex = ArmaduraImperial3
              Else
                  MiObj.objIndex = ArmaduraImperial1
              End If
    Else
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.objIndex = ArmaduraImperial3
              Else
                  MiObj.objIndex = ArmaduraImperial2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoReal(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CriminalesMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 crinales mas para recibir la proxima!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.ArmadaReal = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Aprendiz real"
    Case 1
        TituloReal = "Soldado real"
    Case 2
        TituloReal = "Teniente real"
    Case 3
        TituloReal = "Comandante real"
    Case 4
        TituloReal = "General real"
    Case 5
        TituloReal = "Elite real"
    Case 6
        TituloReal = "General Imperial"
    Case 7
        TituloReal = "Elite Imperial"
    Case 8
        TituloReal = "Guardian del bien"
    Case Else
        TituloReal = "Caballero Celestial"
End Select

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas del caos!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 100 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 100 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
                  MiObj.objIndex = TunicaMagoCaos
    ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).Clase) = "CAZADOR" Or _
           UCase$(UserList(UserIndex).Clase) = "PALADIN" Or _
           UCase$(UserList(UserIndex).Clase) = "BANDIDO" Or _
           UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.objIndex = ArmaduraCaos3
              Else
                  MiObj.objIndex = ArmaduraCaos1
              End If
    Else
              If UCase$(UserList(UserIndex).Raza) = "ENANO" Or _
                 UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
                  MiObj.objIndex = ArmaduraCaos3
              Else
                  MiObj.objIndex = ArmaduraCaos2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CiudadanosMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
End If


End Sub

Public Sub ExpulsarCaos(ByVal UserIndex As Integer)
UserList(UserIndex).Faccion.FuerzasCaos = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado del ejercito del caos!!!." & FONTTYPE_FIGHT)
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 1
        TituloCaos = "Esclavo de las sombras"
    Case 2
        TituloCaos = "Guerrero del caos"
    Case 3
        TituloCaos = "Teniente del caos"
    Case 4
        TituloCaos = "Comandante del caos"
    Case 5
        TituloCaos = "General del caos"
    Case 6
        TituloCaos = "Elite caos"
    Case 7
        TituloCaos = "Asolador de las sombras"
    Case 8
        TituloCaos = "Caballero Oscuro"
    Case 9
        TituloCaos = "Asesino del caos"
    Case Else
        TituloCaos = "Alma del demonio"
End Select


End Function

