Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer
DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MAXEXP)
     
If UserList(VictimIndex).Flags.Invisible = 1 Then
    UserList(VictimIndex).Flags.Invisible = 0
    Call SendData(ToMap, 0, UserList(VictimIndex).Pos.Map, "NOVER" & UserList(VictimIndex).Char.CharIndex & ",0")
End If
If UserList(VictimIndex).Flags.Ceguera = 1 Then
    UserList(VictimIndex).Flags.Ceguera = 0
    Call SendData(ToIndex, VictimIndex, 0, "NSEGUE")
End If
If UserList(VictimIndex).Flags.Estupidez = 1 Then
    UserList(VictimIndex).Flags.Estupidez = 0
    Call SendData(ToIndex, VictimIndex, 0, "NESTUP")
End If

'Lo mata
Call SendData(ToIndex, AttackerIndex, 0, "||Has matado " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT)
Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT)
      
Call SendData(ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_FIGHT)

If Not Criminal(VictimIndex) Then
     Call AddtoVar(UserList(AttackerIndex).Reputacion.AsesinoRep, vlASESINO * 2, MAXREP)
     UserList(AttackerIndex).Reputacion.BurguesRep = 0
     UserList(AttackerIndex).Reputacion.NobleRep = 0
     UserList(AttackerIndex).Reputacion.PlebeRep = 0
Else
     Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
End If

Call UserDie(VictimIndex)

Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, 31000)

'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)

End Sub


Sub RevivirUsuario(ByVal UserIndex As Integer)

UserList(UserIndex).Flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 10

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)

End Sub


Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
Call SendData(ToIndex, UserIndex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)

Dim i As Integer
Dim cad$
For i = 1 To NUMSKILLS
   cad$ = cad$ & UserList(UserIndex).Stats.UserSkills(i) & ","
Next
SendData ToIndex, UserIndex, 0, "SKILLS" & cad$
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
Dim cad$
cad$ = cad$ & UserList(UserIndex).Reputacion.AsesinoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BandidoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BurguesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.LadronesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.NobleRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.PlebeRep & ","

Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
l = l / 6

UserList(UserIndex).Reputacion.Promedio = l

cad$ = cad$ & UserList(UserIndex).Reputacion.Promedio

SendData ToIndex, UserIndex, 0, "FAMA" & cad$

End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$
For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call SendData(ToIndex, UserIndex, 0, "ATR" & cad$)
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
    Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
    
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar")

End Sub

Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

Dim CharIndex As Integer

If InMapBounds(Map, X, Y) Then

       'If needed make a new character in list
       If UserList(UserIndex).Char.CharIndex = 0 Then
           CharIndex = NextOpenCharIndex
           UserList(UserIndex).Char.CharIndex = CharIndex
           CharList(CharIndex) = UserIndex
       End If
       
       'Place character on map
       MapData(Map, X, Y).UserIndex = UserIndex
       
       'Send make character command to clients
       Dim klan$
       klan$ = UserList(UserIndex).GuildInfo.GuildName
       Dim bCr As Byte
       bCr = Criminal(UserIndex)
       If klan$ <> "" Then
            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr)
       Else
            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr)
       End If

End If

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
    
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, UserIndex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    
    If UserList(UserIndex).Stats.ELV = 1 Then
      Pts = 10
      
    Else
      Pts = 5
    End If
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
       
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
    UserList(UserIndex).Stats.Exp = 0
    '[Efestos]
    If Not EsNewbie(UserIndex) And WasNewbie Then
        Call QuitarNewbieObj(UserIndex)
        If UserList(UserIndex).Pos.Map = 37 Or UserList(UserIndex).Pos.Map = 166 Then
            Call WarpUserChar(UserIndex, Arcadia.Map, Arcadia.X, Arcadia.Y, True)
        End If
    End If
    '[Efestos]
    If UserList(UserIndex).Stats.ELV < 11 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    ElseIf UserList(UserIndex).Stats.ELV < 25 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    Else
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.2
    End If
    
    Dim AumentoHP As Integer
    Select Case UserList(UserIndex).Clase
        Case "Guerrero"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        
        Case "Cazador"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
            
        Case "Pirata"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
            
        Case "Paladin"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            'Mana
            Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN)
            
            'STA
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            'Golpe
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                        
        Case "Ladron"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLadron
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
            
        Case "Mago"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero / 2
            If AumentoHP < 1 Then AumentoHP = 4
            AumentoST = 15 - AdicionalSTLadron / 2
            If AumentoST < 1 Then AumentoST = 5
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Leñador"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLeñador
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Minero"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTMinero
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Pescador"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTPescador
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                   
        Case "Clerigo"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Druida"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Asesino"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
            
        Case "Bardo"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        
'[Neptuno]
        Case "Gladiador"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        
        Case "Arquero"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
            
        Case "Chaman"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            If AumentoHP < 1 Then AumentoHP = 6
            AumentoST = 15 - AdicionalSTLadron / 2
            If AumentoST < 1 Then AumentoST = 5
            AumentoHIT = 1
            AumentoMANA = 2.5 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        
        Case "Aldeano"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTMinero
            AumentoST = 15 + AdicionalSTLeñador
            AumentoST = 15 + AdicionalSTPescador
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
            
'[/Neptuno]
        Case Else
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
    End Select
    
    'AddtoVar UserList(UserIndex).Stats.MaxHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.MinHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.Def, 2, STAT_MAXDEF
    
    If AumentoHP > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoST > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoST & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, Pts)
   
    SendUserStatsBox UserIndex
    
    
End If


Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub




Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).Flags.Navegando = 1 Or _
  UserList(UserIndex).Flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

On Error Resume Next

Dim nPos As WorldPos

'Move
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)



If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
    
    Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & "," & "1")

    'Update map and user pos
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Pos = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
    
Else
    'else correct user's pos
    Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
End If

End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)


UserList(UserIndex).Invent.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," _
    & ObjData(Object.ObjIndex).Valor \ 3)

Else

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Function NextOpenCharIndex() As Integer

Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer

Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1) Then Exit For
Next LoopC
  
NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp)
End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status:" & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If


Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub


Sub UpdateUserMap(ByVal UserIndex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Map = UserList(UserIndex).Pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
            Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
            If UserList(MapData(Map, X, Y).UserIndex).Flags.Invisible = 1 Then Call SendData(ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
            
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                      Call Bloquear(ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
            End If
        End If
        
    Next X
Next Y

End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).Name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)


'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).Flags.AttackedBy = UserList(UserIndex).Name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                Call VolverCriminal(UserIndex)
       Else
                Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
    End If
    
    'hacemos que el npc se defienda
           Npclist(NpcIndex).Movement = NPCDEFENSA
           Npclist(NpcIndex).Hostile = 1
    
End If


End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).Clase = "Asesino") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

If UserList(UserIndex).Flags.Hambre = 0 And _
   UserList(UserIndex).Flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(UserIndex).Stats.ELV > 3 _
        And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(UserIndex).Stats.ELV >= 6 _
        And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(UserIndex).Stats.ELV >= 10 _
        And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = Int(RandomNumber(1, Prob))
    
    Dim lvl As Integer
    lvl = UserList(UserIndex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
            Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
            Call AddtoVar(UserList(UserIndex).Stats.Exp, 50, MAXEXP)
            Call SendData(ToIndex, UserIndex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_FIGHT)
            Call CheckUserLevel(UserIndex)
    End If

End If

End Sub

Sub UserDie(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler

'Sonido
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)


'Quitar el dialogo del user muerto
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).Flags.AtacadoPorNpc = 0
UserList(UserIndex).Flags.AtacadoPorUser = 0
UserList(UserIndex).Flags.Envenenado = 0
UserList(UserIndex).Flags.Muerto = 1



Dim aN As Integer

aN = UserList(UserIndex).Flags.AtacadoPorNpc

If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).Flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).Flags.OldHostil
      Npclist(aN).Flags.AttackedBy = ""
End If

'<<<< Paralisis >>>>
If UserList(UserIndex).Flags.Paralizado = 1 Then
    UserList(UserIndex).Flags.Paralizado = 0
    Call SendData(ToIndex, UserIndex, 0, "PARADOK")
End If

'<<<< Descansando >>>>
If UserList(UserIndex).Flags.Descansar Then
    UserList(UserIndex).Flags.Descansar = False
    Call SendData(ToIndex, UserIndex, 0, "DOK")
End If

'<<<< Meditando >>>>
If UserList(UserIndex).Flags.Meditando Then
    UserList(UserIndex).Flags.Meditando = False
    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
End If

' << Si es newbie no pierde el inventario >>
If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
    Call TirarTodo(UserIndex)
Else
    
    If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)
    
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
    End If
End If

' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If

'<< Cambiamos la apariencia del char >>
'[Efestos]
If UserList(UserIndex).Flags.Cabalgando = 1 Then _
    UserList(UserIndex).Flags.Cabalgando = 0

If UserList(UserIndex).Flags.Navegando = 0 Then
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
    UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS
    
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).Flags.OldMovement
                Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).Flags.OldHostil
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
           End If
    End If
    
Next i

UserList(UserIndex).NroMacotas = 0


'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
'        Dim MiObj As Obj
'        Dim nPos As WorldPos
'        MiObj.ObjIndex = RandomNumber(554, 555)
'        MiObj.Amount = 1
'        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
'        Dim ManchaSangre As New cGarbage
'        ManchaSangre.Map = nPos.Map
'        ManchaSangre.X = nPos.X
'        ManchaSangre.Y = nPos.Y
'        Call TrashCollector.Add(ManchaSangre)
'End If

'<< Actualizamos clientes >>
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
Call SendUserStatsBox(UserIndex)


Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE")

End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal atacante As Integer)

If EsNewbie(Muerto) Then Exit Sub

If Criminal(Muerto) Then
        If UserList(atacante).Flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(atacante).Flags.LastCrimMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(atacante).Faccion.CriminalesMatados, 1, 65000)
        End If
        
        If UserList(atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(atacante).Faccion.CriminalesMatados = 0
            UserList(atacante).Faccion.RecompensasReal = 0
        End If
Else
        If UserList(atacante).Flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(atacante).Flags.LastCiudMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(atacante).Faccion.CiudadanosMatados, 1, 65000)
        End If
        
        If UserList(atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(atacante).Faccion.CiudadanosMatados = 0
            UserList(atacante).Faccion.RecompensasCaos = 0
        End If
End If


End Sub

Sub Tilelibre(Pos As WorldPos, nPos As WorldPos)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
        
            If LegalPos(nPos.Map, tX, tY) = True Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.X = tX
                     nPos.Y = tY
                     tX = Pos.X + LoopC
                     tY = Pos.Y + LoopC
                End If
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

'Quitar el dialogo
Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

OldMap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y

Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
UserList(UserIndex).Pos.Map = Map

If OldMap <> Map Then
    Call SendData(ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
    Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)

    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

    'Update new Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1

    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If
  
Else
    
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

End If


Call UpdateUserMap(UserIndex)

If FX Then 'FX
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If


Call WarpMascotas(UserIndex)

End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(UserIndex).NroMacotas

For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Flags.Respawn = 0
        PetTypes(i) = UserList(UserIndex).MascotasType(i)
        PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
        Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

For i = 1 To MAXMASCOTAS
    If PetTypes(i) > 0 Then
        UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
        'Controlamos que se sumoneo OK
        If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
        End If
        Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
        Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNpc = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

UserList(UserIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
Next i

If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0


End Sub

