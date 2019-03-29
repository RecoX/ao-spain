Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean

Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim retVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim Slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    
    Dim sData As String
    sData = UCase(Rdata)
    
    Select Case sData
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
            logged = True
            UserCiego = False
            EngineRun = True
            IScombate = False
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                   Unload frmPasswd
                   Unload frmCrearPersonaje
                   Unload frmConnect
                   frmMain.Show
            End If
            Call SetConnected
            'Mostramos el Tip
            If tipf = "1" And PrimeraVez Then
                 Call CargarTip
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Call DoFogataFx
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "FINOK" ' Graceful exit ;))
            frmMain.Socket1.Disconnect
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            IScombate = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Call frmMain.StopSound
            frmMain.IsPlaying = plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                CharList(i).invisible = False
            Next i
            bO = 0
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= UBound(UserInventory)
                If UserInventory(i).OBJIndex <> 0 Then
                        frmComerciar.List1(1).AddItem UserInventory(i).Name
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim ii As Integer
            ii = 1
            Do While ii <= UBound(UserInventory)
                If UserInventory(ii).OBJIndex <> 0 Then
                        frmBancoObj.List1(1).AddItem UserInventory(ii).Name
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To UBound(UserInventory)
                If UserInventory(i).OBJIndex <> 0 Then
                        frmComerciarUsu.List1.AddItem UserInventory(i).Name
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = UserInventory(i).Amount
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
            frmMain.Socket1.Disconnect
            Unload frmRecuperar
            Exit Sub
        Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
            frmMain.Socket1.Disconnect
            Unload frmRecuperar
            Exit Sub
        Case "BORROK"
            Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
            frmBorrar.MousePointer = 0
            frmMain.Socket1.Disconnect
            Unload frmBorrar
            Exit Sub
        Case "SFH"
            frmHerrero.Show
            Exit Sub
        Case "SFC"
            frmCarp.Show
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, "La criatura fallo el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha matado!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "Has rechazado el ataque con el escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "El usuario rechazo el ataque con su escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, "Has fallado el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 2)
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            'Call StopSound("lluviain.MP3")
                            'Call StopSound("lluviaout.MP3")
                            '[CODE 001]:MatuX'
                                frmMain.StopSound
                                frmMain.IsPlaying = plNone
                            '[END]'
                        End If
                    End If
                Else
                    'vers incorrecta
                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                    Call LiberarObjetosDX
                    Call UnloadAllForms
                    End
                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            CharList(UserCharIndex).POS = UserPos
            Exit Sub
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a la criatura por " & Rdata & "!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & Rdata & " te ataco y fallo!!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la cabeza por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el torso por " & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
            Else
                  If PuedoQuitarFoco Then _
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
            Exit Sub
        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = CharList(UserCharIndex).POS
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            
            CharList(CharIndex).Fx = Val(ReadField(9, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            CharList(CharIndex).Nombre = ReadField(12, Rdata, 44)
            CharList(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))
            
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            
            Exit Sub
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            
            If Fx = 0 Then
                
                'If Not UserNavegando And Val(ReadField(4, Rdata, 44)) <> 0 Then
                        Call DoPasosFx(CharIndex)
                'Else
                        'FX navegando
                'End If
            
            End If
            
            Call MoveCharbyPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            CharList(CharIndex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            CharList(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            CharList(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            CharList(CharIndex).Fx = Val(ReadField(7, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).Casco = CascoAnimData(tempint)
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            If Musica = 0 Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                If Val(ReadField(1, Rdata, 45)) <> 0 Then
                    'Stop_Midi
                    If Musica = 0 Then
                        CurMidi = Val(ReadField(1, Rdata, 45)) & ".mid"
                        LoopMidi = Val(ReadField(2, Rdata, 45))
                        Call CargarMIDI(DirMidi & CurMidi)
                        Call Play_Midi
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Fx = 0 Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call PlayWaveDS(Rdata & ".wav")
            End If
            Exit Sub
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            '[CODE 001]:MatuX
                If frmMain.IsPlaying <> plFogata Then
                    frmMain.StopSound
                    Call frmMain.Play("fuego.wav", True)
                    frmMain.IsPlaying = plFogata
                End If
            '[END]'
            Exit Sub
    End Select

    Select Case Left(sData, 3)
        Case "VAL"                  ' >>>>> Validar Cliente :: VAL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            If frmBorrar.Visible Then
                Call SendData("BORR" & frmBorrar.txtNombre.Text & "," & frmBorrar.txtPasswd.Text & "," & ValidarLoginMSG(CInt(Rdata)))
            Else
                bK = CLng(ReadField(1, Rdata, Asc(",")))
                bO = 100 'CInt(ReadField(1, Rdata, Asc(",")))
                Call Login(ValidarLoginMSG(CInt(ReadField(2, Rdata, Asc(",")))))
            End If
            Exit Sub
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
            Else
               If bLluvia(UserMap) <> 0 Then
                    If bTecho Then
                        'Call StopSound("lluviain.MP3")
                        'Call PlaySound("lluviainend.MP3")
                        '[CODE 001]:MatuX'
                        Call frmMain.StopSound
                        Call frmMain.Play("lluviainend.wav", False)
                        frmMain.IsPlaying = plNone
                        '[END]'
                   Else
                        'Call StopSound("lluviaout.MP3")
                        'Call PlaySound("lluviaoutend.MP3")
                        '[CODE 001]:MatuX'
                        Call frmMain.StopSound
                        Call frmMain.Play("lluviaoutend.wav", False)
                        frmMain.IsPlaying = plNone
                        '[END]'
                    End If
               End If
               bRain = False
            End If
                        
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).Fx = Val(ReadField(2, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim n As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            n = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg n, n2
            frmMSG.Show
            Exit Sub
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            frmMain.Exp.Caption = "Exp:" & UserExp & "/" & UserPasarNivel
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = ReadField(1, Rdata, 44)
            UserInventory(Slot).OBJIndex = ReadField(2, Rdata, 44)
            UserInventory(Slot).Name = ReadField(3, Rdata, 44)
            UserInventory(Slot).Amount = ReadField(4, Rdata, 44)
            UserInventory(Slot).Equipped = ReadField(5, Rdata, 44)
            UserInventory(Slot).GrhIndex = Val(ReadField(6, Rdata, 44))
            UserInventory(Slot).ObjType = Val(ReadField(7, Rdata, 44))
            UserInventory(Slot).MaxHit = Val(ReadField(8, Rdata, 44))
            UserInventory(Slot).MinHit = Val(ReadField(9, Rdata, 44))
            UserInventory(Slot).Def = Val(ReadField(10, Rdata, 44))
            UserInventory(Slot).Valor = Val(ReadField(11, Rdata, 44))
        
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If
            
            bInvMod = True
            
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = ReadField(1, Rdata, 44)
            UserBancoInventory(Slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(Slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(Slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(Slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(Slot).ObjType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(Slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(Slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(Slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(Slot).Amount & ") " & UserBancoInventory(Slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(Slot).Name
            End If
            
            bInvMod = True
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = ReadField(1, Rdata, 44)
            UserHechizos(Slot) = ReadField(2, Rdata, 44)
            If Slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(Slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmOldPersonaje.MousePointer = 1
            frmPasswd.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then frmMain.Socket1.Disconnect
            MsgBox Rdata
            Exit Sub
    End Select
    
    Select Case Left(sData, 4)
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).ObjType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            
            bInvMod = True
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, "Hay " & Rdata & " npcs.", 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show
            End If
            Exit Sub
    End Select
    
    Select Case Left(sData, 5)
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
        Case "NOVER"             ' >>>>> Invisible :: NOVER
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            Exit Sub
    End Select
    
    Select Case Left(sData, 6)
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show
            Exit Sub
    End Select
    
    Select Case Left(sData, 7)
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "PEACEPR"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            Call frmGuildFoundation.Show(vbModeless, frmMain)
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= UBound(UserInventory)
                    If UserInventory(i).OBJIndex <> 0 Then
                            frmComerciar.List1(1).AddItem UserInventory(i).Name
                    Else
                            frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                        frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= UBound(UserInventory)
                    If UserInventory(i).OBJIndex <> 0 Then
                            frmBancoObj.List1(1).AddItem UserInventory(i).Name
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                ii = 1
                Do While ii <= UBound(UserBancoInventory)
                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
    End Select
    
    '[Alejo]
    Select Case UCase(Left(Rdata, 9))
    Case "COMUSUINV"
        Rdata = Right(Rdata, Len(Rdata) - 9)
        OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
        OtroInventario(1).Name = ReadField(3, Rdata, 44)
        OtroInventario(1).Amount = ReadField(4, Rdata, 44)
        OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
        OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
        OtroInventario(1).ObjType = Val(ReadField(7, Rdata, 44))
        OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
        OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
        OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
        OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
        
        frmComerciarUsu.List2.Clear
        
        frmComerciarUsu.List2.AddItem OtroInventario(1).Name
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
        
        frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
    
End Sub

Sub SendData(ByVal sdData As String)
Dim retcode

Dim AuxCmd As String
AuxCmd = UCase(Left(sdData, 5))

bK = GenCrC(bK, sdData)

bO = bO + 1
If bO > 10000 Then bO = 100


'Agregamos el fin de linea
sdData = sdData & "~" & bK & ENDC

'Para evitar el spamming
If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
    Exit Sub
ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
    Exit Sub
End If

retcode = frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub Login(ByVal valcode As Integer)
Dim Passcliente As String
Passcliente = "orophin"

'Personaje grabado
If SendNewChar = False Then
    SendData ("PASSCL" & Passcliente) 'Comprobar pass del cliente ahora
    SendData ("OLOGIN" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode)
End If

'Crear personaje
If SendNewChar = True Then
    SendData ("PASSCL" & Passcliente) 'Comprobar pass del cliente ahora
    SendData ("NLOGIN" & UserName & "," & UserPassword _
    & "," & 0 & "," & 0 & "," _
    & App.Major & "." & App.Minor & "." & App.Revision & _
    "," & UserRaza & "," & UserSexo & "," & UserClase & "," & _
    UserAtributos(1) & "," & UserAtributos(2) & "," & UserAtributos(3) _
    & "," & UserAtributos(4) & "," & UserAtributos(5) _
     & "," & UserSkills(1) & "," & UserSkills(2) _
     & "," & UserSkills(3) & "," & UserSkills(4) _
     & "," & UserSkills(5) & "," & UserSkills(6) _
     & "," & UserSkills(7) & "," & UserSkills(8) _
     & "," & UserSkills(9) & "," & UserSkills(10) _
     & "," & UserSkills(11) & "," & UserSkills(12) _
     & "," & UserSkills(13) & "," & UserSkills(14) _
     & "," & UserSkills(15) & "," & UserSkills(16) _
     & "," & UserSkills(17) & "," & UserSkills(18) _
     & "," & UserSkills(19) & "," & UserSkills(20) _
     & "," & UserSkills(21) & "," & UserSkills(22) _
     & "," & UserEmail & "," & UserHogar & "," & valcode)          '[Efestos]Nuevo skill 22
End If

End Sub


