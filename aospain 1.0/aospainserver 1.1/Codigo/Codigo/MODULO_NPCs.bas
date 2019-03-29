Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

Dim i As Integer
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
     UserList(UserIndex).MascotasIndex(i) = 0
     UserList(UserIndex).MascotasType(i) = 0
     Exit For
  End If
Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)

Dim i As Integer

Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

'For i = 1 To UBound(Npclist(Maestro).Criaturas)
'  If Npclist(Maestro).Criaturas(i).NpcIndex = Mascota Then
'     Npclist(Maestro).Criaturas(i).NpcIndex = 0
'     Npclist(Maestro).Criaturas(i).NpcName = ""
'     Exit For
'  End If
'Next i

End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo errhandler

'   Call LogTarea("Sub MuereNpc")
   
   Dim MiNPC As Npc
   MiNPC = Npclist(NpcIndex)
   
   'Quitamos el npc
   Call QuitarNPC(NpcIndex)
   
   
    
   If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.Flags.Snd3 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MiNPC.Flags.Snd3)
        UserList(UserIndex).Flags.TargetNpc = 0
        UserList(UserIndex).Flags.TargetNpcTipo = 0
        
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMacotas > 0 Then
                Dim t As Integer
                For t = 1 To MAXMASCOTAS
                      If UserList(UserIndex).MascotasIndex(t) > 0 Then
                          If Npclist(UserList(UserIndex).MascotasIndex(t)).TargetNpc = NpcIndex Then
                                  Call FollowAmo(UserList(UserIndex).MascotasIndex(t))
                          End If
                      End If
                Next t
        End If
        
        Call AddtoVar(UserList(UserIndex).Stats.Exp, MiNPC.GiveEXP, MAXEXP)
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & MiNPC.GiveEXP & " puntos de experiencia." & FONTTYPE_FIGHT)
        Call SendData(ToIndex, UserIndex, 0, "||Has matado la criatura!" & FONTTYPE_FIGHT)
        Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)
        
        If MiNPC.Stats.Alineacion = 0 Then
              If MiNPC.Numero = Guardias Then
                    Call VolverCriminal(UserIndex)
              End If
              If Not EsDios(UserList(UserIndex).Name) Then Call AddtoVar(UserList(UserIndex).Reputacion.AsesinoRep, vlASESINO, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 1 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 2 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, vlASESINO / 2, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 4 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        End If
    
        
        'Controla el nivel del usuario
        Call CheckUserLevel(UserIndex)
       
   End If ' Userindex > 0

   
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
   End If
   
   'ReSpawn o no
   Call ReSpawnNpc(MiNPC)
   
Exit Sub

errhandler:
    Call LogError("Error en MuereNpc")
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'Clear the npc's flags

Npclist(NpcIndex).Flags.AfectaParalisis = 0
Npclist(NpcIndex).Flags.AguaValida = 0
Npclist(NpcIndex).Flags.AttackedBy = ""
Npclist(NpcIndex).Flags.Attacking = 0
Npclist(NpcIndex).Flags.BackUp = 0
Npclist(NpcIndex).Flags.Bendicion = 0
Npclist(NpcIndex).Flags.Domable = 0
Npclist(NpcIndex).Flags.Envenenado = 0
Npclist(NpcIndex).Flags.Faccion = 0
Npclist(NpcIndex).Flags.Follow = False
Npclist(NpcIndex).Flags.LanzaSpells = 0
Npclist(NpcIndex).Flags.GolpeExacto = 0
Npclist(NpcIndex).Flags.Invisible = 0
Npclist(NpcIndex).Flags.Maldicion = 0
Npclist(NpcIndex).Flags.OldHostil = 0
Npclist(NpcIndex).Flags.OldMovement = 0
Npclist(NpcIndex).Flags.Paralizado = 0
Npclist(NpcIndex).Flags.Respawn = 0
Npclist(NpcIndex).Flags.RespawnOrigPos = 0
Npclist(NpcIndex).Flags.Snd1 = 0
Npclist(NpcIndex).Flags.Snd2 = 0
Npclist(NpcIndex).Flags.Snd3 = 0
Npclist(NpcIndex).Flags.Snd4 = 0
Npclist(NpcIndex).Flags.TierraInvalida = 0
Npclist(NpcIndex).Flags.UseAINow = False

End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.Heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Attackable = 0
Npclist(NpcIndex).CanAttack = 0
Npclist(NpcIndex).Comercia = 0
Npclist(NpcIndex).GiveEXP = 0
Npclist(NpcIndex).GiveGLD = 0
Npclist(NpcIndex).Hostile = 0
Npclist(NpcIndex).Inflacion = 0
Npclist(NpcIndex).InvReSpawn = 0
Npclist(NpcIndex).Level = 0

If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)

If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

Npclist(NpcIndex).MaestroUser = 0
Npclist(NpcIndex).MaestroNpc = 0

Npclist(NpcIndex).Mascotas = 0
Npclist(NpcIndex).Movement = 0
Npclist(NpcIndex).Name = "NPC SIN INICIAR"
Npclist(NpcIndex).NPCtype = 0
Npclist(NpcIndex).Numero = 0
Npclist(NpcIndex).Orig.Map = 0
Npclist(NpcIndex).Orig.X = 0
Npclist(NpcIndex).Orig.Y = 0
Npclist(NpcIndex).PoderAtaque = 0
Npclist(NpcIndex).PoderEvasion = 0
Npclist(NpcIndex).Pos.Map = 0
Npclist(NpcIndex).Pos.X = 0
Npclist(NpcIndex).Pos.Y = 0
Npclist(NpcIndex).SkillDomar = 0
Npclist(NpcIndex).Target = 0
Npclist(NpcIndex).TargetNpc = 0
Npclist(NpcIndex).TipoItems = 0
Npclist(NpcIndex).Veneno = 0
Npclist(NpcIndex).Desc = ""

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroSpells
    Npclist(NpcIndex).Spells(j) = 0
Next j

Call ResetNpcCharInfo(NpcIndex)
Call ResetNpcCriatures(NpcIndex)
Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

Npclist(NpcIndex).Flags.NPCActive = False


If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
    Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex)
End If

'Nos aseguramos de que el inventario sea removido...
'asi los lobos no volveran a tirar armaduras ;))
Call ResetNpcInv(NpcIndex)
Call ResetNpcFlags(NpcIndex)
Call ResetNpcCounters(NpcIndex)

Call ResetNpcMainInfo(NpcIndex)

If NpcIndex = LastNPC Then
    Do Until Npclist(LastNPC).Flags.NPCActive
        LastNPC = LastNPC - 1
        If LastNPC < 1 Then Exit Do
    Loop
End If
    
  
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If

Exit Sub

errhandler:
    Npclist(NpcIndex).Flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos) As Boolean


If LegalPos(Pos.Map, Pos.X, Pos.Y) Then
    TestSpawnTrigger = _
    MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 3 And _
    MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 2 And _
    MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 1
End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim NIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

NIndex = OpenNPC(NroNPC) 'Conseguimos un indice

If NIndex > MAXNPCS Then Exit Sub

'Necesita ser respawned en un lugar especifico
If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    
    Map = OrigPos.Map
    X = OrigPos.X
    Y = OrigPos.Y
    Npclist(NIndex).Orig = OrigPos
    Npclist(NIndex).Pos = OrigPos
   
Else
    
    Pos.Map = mapa 'mapa
    
    Do While Not PosicionValida
    
        Randomize (Timer)
        Pos.X = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en x
        Pos.Y = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en y
        
        Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
        
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(NIndex).Flags.AguaValida) And _
           Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(NIndex).Pos.Map = newpos.Map
            Npclist(NIndex).Pos.X = newpos.X
            Npclist(NIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        
        End If
            
        'for debug
        Iteraciones = Iteraciones + 1
        If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(NIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                Exit Sub
        End If
    Loop
    
    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = Npclist(NIndex).Pos.X
    Y = Npclist(NIndex).Pos.Y
End If

'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, NIndex, Map, X, Y)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)


Dim CharIndex As Integer

If Npclist(NpcIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    Npclist(NpcIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = NpcIndex
End If


MapData(Map, X, Y).NpcIndex = NpcIndex

Call SendData(sndRoute, sndIndex, sndMap, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y)

End Sub

Sub ChangeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, Body As Integer, Head As Integer, Heading As Byte)

If NpcIndex > 0 Then
    Npclist(NpcIndex).Char.Body = Body
    Npclist(NpcIndex).Char.Head = Head
    Npclist(NpcIndex).Char.Heading = Heading
    Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar < 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los cliente
Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "BP" & Npclist(NpcIndex).Char.CharIndex)

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
            ' es una posicion legal
            If LegalPos(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then
            
                If Npclist(NpcIndex).Flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).Flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
            
                'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            End If
    Else ' No es mascota
            ' Controlamos que la posicion sea legal, los npc que
            ' no son mascotas tienen mas restricciones de movimiento.
            If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).Flags.AguaValida) Then
                
                If Npclist(NpcIndex).Flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).Flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
                'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            Else
                If Npclist(NpcIndex).Movement = NPC_PATHFINDING Then
                    'Someone has blocked the npc's way, we must to seek a new path!
                    Npclist(NpcIndex).PFINFO.PathLenght = 0
                End If
            
            End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).Flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

Dim n As Integer
n = RandomNumber(1, 100)
If n < 30 Then
    UserList(UserIndex).Flags.Envenenado = 1
    Call SendData(ToIndex, UserIndex, 0, "||¡¡La criatura te ha envenenado!!" & FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex
'Call LogTarea("Sub SpawnNpc")

Dim newpos As WorldPos
Dim NIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim it As Integer

NIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0

If NIndex > MAXNPCS Then
    SpawnNpc = NIndex
    Exit Function
End If

Do While Not PosicionValida
        
        Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPos(newpos.Map, newpos.X, newpos.Y) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(NIndex).Pos.Map = newpos.Map
            Npclist(NIndex).Pos.X = newpos.X
            Npclist(NIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(NIndex)
            SpawnNpc = MAXNPCS
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop
    
'asignamos las nuevas coordenas
Map = newpos.Map
X = Npclist(NIndex).Pos.X
Y = Npclist(NIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, NIndex, Map, X, Y)

If FX Then
    Call SendData(ToMap, 0, Map, "TW" & SND_WARP)
    Call SendData(ToMap, 0, Map, "CFX" & Npclist(NIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If

SpawnNpc = NIndex

End Function

Sub ReSpawnNpc(MiNPC As Npc)

If (MiNPC.Flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).Flags.NPCActive _
       And Npclist(NpcIndex).Pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As Npc)

'SI EL NPC TIENE ORO LO TIRAMOS
If MiNPC.GiveGLD > 0 Then
    Dim MiObj As Obj
    MiObj.Amount = MiNPC.GiveGLD
    MiObj.objIndex = iORO
    Call TirarItemAlPiso(MiNPC.Pos, MiObj)
End If

End Sub



Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

Dim NpcIndex As Integer
Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "NPCs.dat"
End If


NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).Flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).Flags.AguaValida = val(GetVar(npcfile, "NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).Flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).Flags.Faccion = val(GetVar(npcfile, "NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).Flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

Npclist(NpcIndex).Veneno = val(GetVar(npcfile, "NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).Flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).PoderAtaque = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEvasion"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))


Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).objIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).Flags.LanzaSpells = val(GetVar(npcfile, "NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).Flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).Flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).Flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = val(GetVar(npcfile, "NPC" & NpcNumber, "Sp" & LoopC))
Next LoopC


If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
    Npclist(NpcIndex).NroCriaturas = val(GetVar(npcfile, "NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = GetVar(npcfile, "NPC" & NpcNumber, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = GetVar(npcfile, "NPC" & NpcNumber, "CN" & LoopC)
    Next LoopC
End If


Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

Npclist(NpcIndex).Flags.NPCActive = True
Npclist(NpcIndex).Flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).Flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).Flags.Respawn = 1
End If

Npclist(NpcIndex).Flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).Flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).Flags.AfectaParalisis = val(GetVar(npcfile, "NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).Flags.GolpeExacto = val(GetVar(npcfile, "NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).Flags.Snd1 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).Flags.Snd2 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).Flags.Snd3 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd3"))
Npclist(NpcIndex).Flags.Snd4 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd4"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = GetVar(npcfile, "NPC" & NpcNumber, "NROEXP")
If aux = "" Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = GetVar(npcfile, "NPC" & NpcNumber, "Exp" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function
Sub EnviarListaCriaturas(ByVal UserIndex As Integer, ByVal NpcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  SD = "LSTCRI" & SD
  Call SendData(ToIndex, UserIndex, 0, SD)
End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).Flags.Follow Then
  Npclist(NpcIndex).Flags.AttackedBy = ""
  Npclist(NpcIndex).Flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).Flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).Flags.OldHostil
Else
  Npclist(NpcIndex).Flags.AttackedBy = UserName
  Npclist(NpcIndex).Flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).Flags.Follow = True
  Npclist(NpcIndex).Movement = SIGUE_AMO 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNpc = 0

End Sub

