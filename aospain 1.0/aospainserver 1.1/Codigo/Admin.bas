Attribute VB_Name = "Admin"
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

Public MaxLines As Integer
Public MOTD() As String

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long


Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloMover As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public MinutosWs As Long
Public puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function


Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As Npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).Flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(i, 0)
        End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(ToAll, 0, 0, "||%%%%GUARDANDO CONFIGURACION DE AOSPAIN, POR FAVOR ESPERE%%%%" & FONTTYPE_INFO)

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.Value = 0

For loopX = 1 To NumMaps
    DoEvents
    
    If MapInfo(loopX).BackUp = 1 Then
    
            Call SaveMapData(loopX)
            FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
    End If

Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).Flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next

Call SendData(ToAll, 0, 0, "||%%%%CONFIGURACION DE AOSPAIN GUARDADA%%%%" & FONTTYPE_INFO)


End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).Flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(ToIndex, i, 0, "||Has sido liberado!" & FONTTYPE_INFO)
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1) 'Or _
(val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NOONE")
End Function
