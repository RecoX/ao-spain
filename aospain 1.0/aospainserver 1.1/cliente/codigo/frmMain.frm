VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum Online"
   ClientHeight    =   8625
   ClientLeft      =   390
   ClientTop       =   675
   ClientWidth     =   11910
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6750
      Top             =   1905
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   2790
      Top             =   2760
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1290
      Top             =   2220
   End
   Begin VB.Timer Trabajo 
      Interval        =   600
      Left            =   7200
      Top             =   2760
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   6285
      Top             =   2040
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7605
      Top             =   1905
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   999999
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8145
      Left            =   8235
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   2
      Top             =   -60
      Width           =   3585
      Begin VB.CommandButton DespInv 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   540
         MouseIcon       =   "frmMain.frx":1C77B
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   4800
         Width           =   2430
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         MouseIcon       =   "frmMain.frx":1C8CD
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2160
         Width           =   2430
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   555
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   8
         Top             =   2400
         Width           =   2415
      End
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2790
         Left            =   360
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   4
         Left            =   2040
         MouseIcon       =   "frmMain.frx":1CA1F
         MousePointer    =   99  'Custom
         ToolTipText     =   "Muestra el mapa del continente AOSpain"
         Top             =   6120
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   3
         Left            =   2040
         MouseIcon       =   "frmMain.frx":1CB71
         MousePointer    =   99  'Custom
         ToolTipText     =   "Muestra los comandos del juego"
         Top             =   6480
         Width           =   1290
      End
      Begin VB.Image cmdInfo 
         Height          =   405
         Left            =   2310
         MouseIcon       =   "frmMain.frx":1CCC3
         MousePointer    =   99  'Custom
         Top             =   4830
         Width           =   855
      End
      Begin VB.Image CmdLanzar 
         Height          =   405
         Left            =   450
         MouseIcon       =   "frmMain.frx":1CE15
         MousePointer    =   99  'Custom
         Top             =   4830
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1185
         TabIndex        =   14
         Top             =   435
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label exp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   285
         TabIndex        =   13
         Top             =   675
         Width           =   345
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   0
         Left            =   2085
         Top             =   5955
         Width           =   360
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2595
         TabIndex        =   12
         Top             =   5970
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   2
         Left            =   2025
         MouseIcon       =   "frmMain.frx":1CF67
         MousePointer    =   99  'Custom
         ToolTipText     =   "Muestra la lista de Clanes de AOSpain"
         Top             =   7575
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   1
         Left            =   2025
         MouseIcon       =   "frmMain.frx":1D0B9
         MousePointer    =   99  'Custom
         ToolTipText     =   "Muestra las configuraciones del juego"
         Top             =   7200
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   0
         Left            =   2040
         MouseIcon       =   "frmMain.frx":1D20B
         MousePointer    =   99  'Custom
         ToolTipText     =   "Muestra las estadísticas de tu pj"
         Top             =   6840
         Width           =   1290
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   315
         Top             =   7575
         Width           =   1290
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   315
         Top             =   7245
         Width           =   1290
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   315
         Top             =   6585
         Width           =   1290
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   315
         Top             =   6240
         Width           =   1290
      End
      Begin VB.Shape Hpshp 
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   330
         Top             =   6900
         Width           =   1290
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         MouseIcon       =   "frmMain.frx":1D35D
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1290
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         MouseIcon       =   "frmMain.frx":1D4AF
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1305
         Width           =   1605
      End
      Begin VB.Image InvEqu 
         Height          =   4395
         Left            =   120
         Picture         =   "frmMain.frx":1D601
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   5
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   765
         TabIndex        =   4
         Top             =   450
         Width           =   105
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   285
         TabIndex        =   3
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.Timer Attack 
      Interval        =   2500
      Left            =   7170
      Top             =   1920
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1575
      Visible         =   0   'False
      Width           =   8160
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   45
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":2CF54
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H000000C0&
      Height          =   6165
      Left            =   60
      Top             =   1950
      Width           =   8205
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuDescripcion 
         Caption         =   "Descripcion"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim POS(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long
Implements DirectXEvent

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub


Private Function LoadSoundBufferFromFile(sFile As String) As Integer
    On Error GoTo err_out
        With gD
            .lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
            .lReserved = 0
        End With
        Set gDSB = DirectSound.CreateSoundBufferFromFile(DirSound & sFile, gD, gW)
        With POS(0)
            .hEventNotify = endEvent
            .lOffset = -1
        End With
        DirectX.SetEvent endEvent
        'gDSB.SetNotificationPositions 1, POS()
    Exit Function

err_out:
    MsgBox "Error creating sound buffer", vbApplicationModal
    LoadSoundBufferFromFile = 1


End Function


Public Sub Play(ByVal Nombre As String, Optional ByVal LoopSound As Boolean = False)
    If Fx = 1 Then Exit Sub
    Call LoadSoundBufferFromFile(Nombre)

    If LoopSound Then
        gDSB.Play DSBPLAY_LOOPING
    Else
        gDSB.Play DSBPLAY_DEFAULT
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'[Efestos]
If UserParalizado = True Or UserCiego = True Or UserEstupido = True Then     '65~190~156~0~0
    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes salir porque estas afectado por un hechizo que te lo impide.", 65, 190, 156, 0, 0)
    Cancel = True
    Exit Sub
Else
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End If
'[Efestos]
End Sub

Public Sub StopSound()
    On Local Error Resume Next
    If Not gDSB Is Nothing Then
            gDSB.Stop
            gDSB.SetCurrentPosition 0
    End If
End Sub

Private Sub FPS_Timer()

If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    ActualSecond = Mid(time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     TIMERS                         '
''''''''''''''''''''''''''''''''''''''

Private Sub Trabajo_Timer()
    NoPuedeUsar = False
End Sub

Private Sub Attack_Timer()
    UserCanAttack = 1
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        End If
    End If

    bInvMod = True
End Sub

Private Sub AgarrarItem()
    SendData "AG"
    bInvMod = True
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & ItemElegido
    bInvMod = True
End Sub

Private Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & ItemElegido
    bInvMod = True
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("LH" & hlst.ListIndex + 1)
        Call SendData("UK" & Magia)
        UserCanAttack = 0
    End If
End Sub

Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub DespInv_Click(Index As Integer)
    Select Case Index
        Case 0:
            If OffsetDelInv > 0 Then
                OffsetDelInv = OffsetDelInv - XCantItems
                my = my + 1

            End If
        Case 1:
            If OffsetDelInv < MAX_INVENTORY_SLOTS Then
                OffsetDelInv = OffsetDelInv + XCantItems
                my = my - 1
            End If
    End Select
    bInvMod = True
End Sub

Private Sub Form_Click()

    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If UsingSkill = 0 Then
            SendData "LC" & tX & "," & tY
        Else
            SendData "WLC" & tX & "," & tY & "," & UsingSkill
            UsingSkill = 0
            frmMain.MousePointer = vbDefault
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
If (Not SendTxt.Visible) And _
   ((KeyCode >= 65 And KeyCode <= 90) Or _
   (KeyCode >= 48 And KeyCode <= 57)) Then
        
            Select Case KeyCode
                Case vbKeyM:
                    If Not IsPlayingCheck Then
                        Musica = 0
                        Play_Midi
                    Else
                        Musica = 1
                        Stop_Midi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyC:
                    Call SendData("TAB")
                    IScombate = Not IScombate
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyS:
                    Call SendData("SEG")
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyX:
                    Call SendData("RPU")
                Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
            End Select
        End If
        
        Select Case KeyCode
           'ULISES MACRO HECHIZOS
           Case vbKeySpace:
             If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
               Call SendData("LH" & hlst.ListIndex + 1)
               Call SendData("UK" & Magia)
               UserCanAttack = 0
             End If
            Case vbKeyReturn:
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyF4:
                FPSFLAG = Not FPSFLAG
                If Not FPSFLAG Then _
                    frmMain.Caption = "AOSpain v.1.1"
            Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                End If
            Case vbKeyF5:
                Call frmOpciones.Show(vbModeless, frmMain)
        End Select
End Sub

Private Sub Form_Load()
    frmMain.Caption = "AOSpain v.1.1"
    PanelDer.Picture = LoadPicture(App.Path & _
    "\Graficos\Principalnuevo_sin_energia.jpg")
    
    InvEqu.Picture = LoadPicture(App.Path & _
    "\Graficos\Centronuevoinventario.jpg")
    
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call PlayWaveDS(SND_CLICK)

    Select Case Index
        Case 0
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 1
        '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
        Case 3
            Call frmayuda.Show(vbModeless, frmMain)
        Case 4
            ShellExecute frmMain.hwnd, vbNullString, "http://caratula2000.redtotalonline.com/aospain/aospainmap.php", vbNullString, vbNullString, vbNormalFocus
    End Select
    
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            ItemElegido = FLAGORO
            If UserGLD > 0 Then
                frmCantidad.Show
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show
End Sub

Private Sub Label4_Click()
    Call PlayWaveDS(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

    DespInv(0).Visible = True
    DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
End Sub

Private Sub Label7_Click()
    Call PlayWaveDS(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    DespInv(0).Visible = False
    DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If ItemElegido <> 0 Then SendData "USA" & ItemElegido
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mx As Integer
    Dim my As Integer
    Dim aux As Integer
    mx = X \ 32 + 1
    my = Y \ 32 + 1
    aux = (mx + (my - 1) * 5) + OffsetDelInv
    If aux > 0 And aux < MAX_INVENTORY_SLOTS Then _
        picInv.ToolTipText = UserInventory(aux).Name
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PlayWaveDS(SND_CLICK)

    If (Button = vbRightButton) And (ClicEnItemElegido(CInt(X), CInt(Y))) Then
        PopupMenu mnuObj
    End If

    Call ItemClick(CInt(X), CInt(Y))
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
    stxtbuffer = SendTxt.Text
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j$
                    j$ = MD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    stxtbuffer = "/PASSWD " & j$
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    If frmCrearPersonaje.Visible Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf Not frmRecuperar.Visible Then
        Call SendData("gIvEmEvAlcOde")
    Else
        Dim cmd$
        cmd$ = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        frmMain.Socket1.Write cmd$, Len(cmd$)
    End If
End Sub

Private Sub Socket1_Disconnect()
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
        
    frmConnect.Visible = True
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    bO = 100
    
    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Socket1.Read RD, DataLength

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = Mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = Mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = Mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub



