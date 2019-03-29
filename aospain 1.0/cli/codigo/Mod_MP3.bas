Attribute VB_Name = "Mod_MP3"
Option Explicit

Public Const SND_SYNC = &H0 ' SINCRONO
Public Const SND_ASYNC = &H1 ' ASINCRONO
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

Public Const SND_CLICK = "click.mp3"
Public Const SND_OVER = "click2.mp3"
Public Const SND_DICE = "cupdice.mp3"
Public Const SND_PASOS1 = "23.mp3"
Public Const SND_PASOS2 = "24.mp3"

Dim CSTRM As New Collection

Public bBass As Boolean

Public Function InitializeBass()
    ' Check that BASS 1.0 was loaded
    If BASS_GetStringVersion <> "1.1" Then
        LogError "Wrong BASS version number -> 1.1"
        bBass = False
    End If
    ' Initialize digital sound - default device, 44100hz, stereo, 16 bits
    If BASS_Init(-1, 44100, BASS_DEVICE_LEAVEVOL, frmMain.hWnd) = BASSFALSE Then
        LogError "Can't init sound system"
        bBass = False
        Exit Function
    End If
    ' Start digital output
    If BASS_Start = BASSFALSE Then
        LogError "Can't start digital output"
        bBass = False
        Exit Function
    End If
    
    bBass = True
End Function

Public Function ShutDownBass()
    If bBass Then
        Dim iX As Integer
        ' Stop digital output
        BASS_Stop
        ' Free the stream
        For iX = 1 To CSTRM.count
            BASS_StreamFree CSTRM(iX)
        Next iX
        ' Close digital sound system
        BASS_Free
        ' Release HStream collection
        Set CSTRM = Nothing
    End If
End Function

Public Sub LoadSounds()
    If bBass Then
        Dim iX As Integer
        Dim StreamHandle As Long
        frmCargando.MP3Files.Pattern = "*.mp3"
        frmCargando.MP3Files.Path = DirSound
        
        'Load MP3s
        For iX = 0 To frmCargando.MP3Files.ListCount - 1
            StreamHandle = BASS_StreamCreateFile(BASSFALSE, DirSound & frmCargando.MP3Files.List(iX), 0, 0, 0)
            If StreamHandle = 0 Then
                LogError "Can't create stream: " & frmCargando.MP3Files.List(iX)
            Else
                Call CSTRM.Add(StreamHandle, frmCargando.MP3Files.List(iX))
            End If
        Next iX
        
        'Load WAVs
        frmCargando.MP3Files.Pattern = "*.wav"
        For iX = 0 To frmCargando.MP3Files.ListCount - 1
            StreamHandle = BASS_StreamCreateFile(BASSFALSE, DirSound & frmCargando.MP3Files.List(iX), 0, 0, 0)
            If StreamHandle = 0 Then
                LogError "Can't create stream: " & frmCargando.MP3Files.List(iX)
            Else
                Call CSTRM.Add(StreamHandle, frmCargando.MP3Files.List(iX))
            End If
        Next iX
    End If
End Sub

Public Function PlaySound(sFile As String, Optional bLoop As Boolean)
    If bBass And Fx = 0 Then
        'Play stream, not flushed
        If bLoop Then
            If BASS_ChannelIsActive(CSTRM(sFile)) = BASSFALSE Then
                If BASS_StreamPlay(CSTRM(sFile), BASSFALSE, BASS_SAMPLE_LOOP) = BASSFALSE Then _
                    LogError "Can't play stream: " & sFile
            End If
        Else
            Call BASS_ChannelStop(CSTRM(sFile))
            If BASS_StreamPlay(CSTRM(sFile), BASSFALSE, 0) = BASSFALSE Then _
                LogError "Can't play stream: " & sFile
        End If
    End If
End Function

Public Function StopSound(sFile As String)
    If bBass Then
        ' Stop the stream
        If BASS_ChannelIsActive(CSTRM(sFile)) = BASSTRUE Then
            Call BASS_ChannelStop(CSTRM(sFile))
        End If
    End If
End Function
