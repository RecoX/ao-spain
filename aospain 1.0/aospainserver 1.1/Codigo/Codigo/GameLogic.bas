Attribute VB_Name = "Extra"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


Option Explicit

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).OBJInfo.objIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.objIndex).ObjType = OBJTYPE_TELEPORT
    End If
    
    If MapData(Map, X, Y).TileExit.Map > 0 Then
        '�Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then
            '�El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '�FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(ToIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                
                Call ClosestLegalPos(UserList(UserIndex).Pos, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.X = tX
                nPos.Y = tY
                '�Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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

Function NameIndex(ByVal Name As String) As Integer

Dim UserIndex As Integer
'�Nombre valido?
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UCase$(Left$(UserList(UserIndex).Name, Len(Name))) = UCase$(Name)
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        UserIndex = 0
        Exit Do
    End If
    
Loop
NameIndex = UserIndex
End Function


Function IP_Index(ByVal inIP As String) As Integer
On Error GoTo local_errHand

Dim UserIndex As Integer
'�Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Do
    End If
    
Loop

local_errHand:
    
    IP_Index = UserIndex

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).Flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).Flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = EAST Then
    nX = X + 1
    nY = Y
End If

If Head = WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nX
Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua = False) As Boolean

'�Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
  
  If Not PuedeAgua Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
  Else
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (HayAgua(Map, X, Y))
  End If
   
End If

End Function



Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).Trigger <> POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).Trigger <> POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(ToIndex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC
End Sub
Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

If Npclist(NpcIndex).NroExpresiones > 0 Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "�" & Npclist(NpcIndex).Expresiones(randomi) & "�" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
End If
                    
End Sub
Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String

'�Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(UserIndex).Flags.TargetMap = Map
    UserList(UserIndex).Flags.TargetX = X
    UserList(UserIndex).Flags.TargetY = Y
    '�Es un obj?
    If MapData(Map, X, Y).OBJInfo.objIndex > 0 Then
        'Informa el nombre
        Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y).OBJInfo.objIndex).Name & FONTTYPE_INFO)
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X, Y).OBJInfo.objIndex
        UserList(UserIndex).Flags.TargetObjMap = Map
        UserList(UserIndex).Flags.TargetObjX = X
        UserList(UserIndex).Flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.objIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.objIndex).ObjType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y).OBJInfo.objIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).Flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.objIndex
            UserList(UserIndex).Flags.TargetObjMap = Map
            UserList(UserIndex).Flags.TargetObjX = X + 1
            UserList(UserIndex).Flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.objIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.objIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.objIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).Flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.objIndex
            UserList(UserIndex).Flags.TargetObjMap = Map
            UserList(UserIndex).Flags.TargetObjX = X + 1
            UserList(UserIndex).Flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.objIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.objIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y + 1).OBJInfo.objIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).Flags.TargetObj = MapData(Map, X, Y).OBJInfo.objIndex
            UserList(UserIndex).Flags.TargetObjMap = Map
            UserList(UserIndex).Flags.TargetObjX = X
            UserList(UserIndex).Flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    '�Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '�Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  �Encontro un Usuario?
            
       If UserList(TempCharIndex).Flags.AdminInvisible = 0 Then

            If EsNewbie(TempCharIndex) Then
                Stat = " <NEWBIE>"
            End If
            
            If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                Stat = Stat & " <Ejercito real> " & "<" & TituloReal(TempCharIndex) & ">"
            ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                Stat = Stat & " <Fuerzas del caos> " & "<" & TituloCaos(TempCharIndex) & ">"
            End If
            
            If UserList(TempCharIndex).GuildInfo.GuildName <> "" Then
                Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
            End If
            
            If Len(UserList(TempCharIndex).Desc) > 1 Then
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ves a " & UserList(TempCharIndex).Name & Stat)
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat
            End If
            
            If Criminal(TempCharIndex) Then
                Stat = Stat & " <CRIMINAL> ~255~0~0~1~0"
            Else
                Stat = Stat & " <CIUDADANO> ~0~0~200~1~0"
            End If
            
            Call SendData(ToIndex, UserIndex, 0, Stat)
                
            
            FoundSomething = 1
            UserList(UserIndex).Flags.TargetUser = TempCharIndex
            UserList(UserIndex).Flags.TargetNpc = 0
            UserList(UserIndex).Flags.TargetNpcTipo = 0
       
       End If
       
    End If
    If FoundChar = 2 Then '�Encontro un NPC?
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & Npclist(TempCharIndex).Desc & "�" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            Else
                
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & "." & FONTTYPE_INFO)
                End If
                
            End If
            FoundSomething = 1
            UserList(UserIndex).Flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).Flags.TargetNpc = TempCharIndex
            UserList(UserIndex).Flags.TargetUser = 0
            UserList(UserIndex).Flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).Flags.TargetNpc = 0
        UserList(UserIndex).Flags.TargetNpcTipo = 0
        UserList(UserIndex).Flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).Flags.TargetNpc = 0
        UserList(UserIndex).Flags.TargetNpcTipo = 0
        UserList(UserIndex).Flags.TargetUser = 0
        UserList(UserIndex).Flags.TargetObj = 0
        UserList(UserIndex).Flags.TargetObjMap = 0
        UserList(UserIndex).Flags.TargetObjX = 0
        UserList(UserIndex).Flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).Flags.TargetNpc = 0
        UserList(UserIndex).Flags.TargetNpcTipo = 0
        UserList(UserIndex).Flags.TargetUser = 0
        UserList(UserIndex).Flags.TargetObj = 0
        UserList(UserIndex).Flags.TargetObjMap = 0
        UserList(UserIndex).Flags.TargetObjX = 0
        UserList(UserIndex).Flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If
End If


End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function



