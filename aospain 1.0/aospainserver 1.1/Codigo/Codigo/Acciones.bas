Attribute VB_Name = "Acciones"
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



Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.objIndex > 0 Then
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X, Y).OBJInfo.objIndex
        
        Select Case ObjData(MapData(Map, X, Y).OBJInfo.objIndex).ObjType
            
            Case OBJTYPE_PUERTAS 'Es una puerta
                Call AccionParaPuerta(Map, X, Y, UserIndex)
            Case OBJTYPE_CARTELES 'Es un cartel
                Call AccionParaCartel(Map, X, Y, UserIndex)
            Case OBJTYPE_FOROS 'Foro
                Call AccionParaForo(Map, X, Y, UserIndex)
            Case OBJTYPE_LEÑA 'Leña
                If MapData(Map, X, Y).OBJInfo.objIndex = FOGATA_APAG Then
                    Call AccionParaRamita(Map, X, Y, UserIndex)
                End If
            
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, X + 1, Y).OBJInfo.objIndex > 0 Then
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.objIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.objIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.objIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.objIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.objIndex > 0 Then
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.objIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.objIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.objIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.objIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.objIndex > 0 Then
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.objIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.objIndex).ObjType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.objIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.objIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
            
        End Select
        
    Else
        UserList(UserIndex).Flags.TargetNpc = 0
        UserList(UserIndex).Flags.TargetNpcTipo = 0
        UserList(UserIndex).Flags.TargetUser = 0
        UserList(UserIndex).Flags.TargetObj = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If
    
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer


If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.objIndex = FOGATA
    Obj.Amount = 1
    
    Call SendData(ToIndex, UserIndex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
    
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0 Then
    Call SubirSkill(UserIndex, Supervivencia)
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

'¿Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.objIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim n As Integer
    For i = 1 To num
        n = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #n
        Input #n, tit
        men = ""
        auxcad = ""
        Do While Not EOF(n)
            Input #n, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #n
        Call SendData(ToIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(ToIndex, UserIndex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.objIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.objIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.objIndex).Llave = 0 Then
                          
                     MapData(Map, X, Y).OBJInfo.objIndex = ObjData(MapData(Map, X, Y).OBJInfo.objIndex).IndexAbierta
                                  
                     Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                     
                     'Desbloquea
                     MapData(Map, X, Y).Blocked = 0
                     MapData(Map, X - 1, Y).Blocked = 0
                     
                     'Bloquea todos los mapas
                     Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
                     Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)
                     
                       
                     'Sonido
                     SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.objIndex = ObjData(MapData(Map, X, Y).OBJInfo.objIndex).IndexCerrada
                
                Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(ToMap, 0, Map, Map, X, Y, 1)
                
                SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(UserIndex).Flags.TargetObj = MapData(Map, X, Y).OBJInfo.objIndex
    Else
        Call SendData(ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, X, Y).OBJInfo.objIndex).ObjType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.objIndex).texto) > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "MCAR" & _
        ObjData(MapData(Map, X, Y).OBJInfo.objIndex).texto & _
        Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.objIndex).GrhSecundario)
  End If
  
End If

End Sub

