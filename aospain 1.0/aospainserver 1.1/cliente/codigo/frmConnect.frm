VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   8640
      Width           =   2895
   End
   Begin VB.TextBox DescTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "AOSpain Primario"
      Top             =   2160
      Width           =   2895
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   5715
      Left            =   1350
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2700
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Este Server ->"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   5715
      ItemData        =   "frmConnect.frx":000C
      Left            =   1350
      List            =   "frmConnect.frx":0013
      TabIndex        =   3
      Top             =   2700
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1920
      TabIndex        =   0
      Text            =   "7666"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3840
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   0
      MouseIcon       =   "frmConnect.frx":0024
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   2205
   End
   Begin VB.Image imgServEspana 
      Height          =   435
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   5220
      Width           =   2475
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   4500
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2595
   End
   Begin VB.Image imgGetPass 
      Height          =   615
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   2205
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   9255
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   2205
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      Top             =   -45
      Width           =   12000
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Public Sub CargarLst()

Dim i As Integer

lst_servers.Clear

For i = 1 To UBound(ServersLst)
    lst_servers.AddItem ServersLst(i).desc
Next i

End Sub

Private Sub Command1_Click()
CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
End Sub


Private Sub Form_Activate()
Dim nDirectorio As String
If CurServer <> 0 Then
    IPTxt = ServersLst(CurServer).Ip
    PortTxt = ServersLst(CurServer).Puerto
Else
    IPTxt = IPdelServidor
    PortTxt = PuertoDelServidor
End If

Call CargarLst
DescTxt.Text = ServersLst(CurServer).desc
nDirectorio = Dir(App.Path & "\Web", vbDirectory)
If nDirectorio <> "Web" Then MkDir (App.Path & "\Web")

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
    IPTxt.Text = "localhost"
    IPTxt.Visible = True
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
Dim d As Integer
Dim j
For Each j In Image1()
   j.Tag = "0"
Next
PortTxt.Text = Config_Inicio.Puerto
 
FONDO.Picture = LoadPicture(App.Path & "\Graficos\Conectar.jpg")
'[Efestos]
Do While d <> 5
DescargarTxt(d) = True
d = d + 1
Loop
'[Efestos]
 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Beta: 1"
 '[END]'

End Sub



Private Sub Image1_Click(Index As Integer)

Dim Archivo As String
Dim cadena As String
Dim nArchivo As String
Dim eArchivo As String

If Not IsIp(IPTxt) And CurServer <> 0 Then
    If MsgBox("Atencion, está intentando conectarse a un servidor no oficial, NoLand Studios no se hace responsable de los posibles problemas que estos servidores presenten. ¿Desea continuar?", vbYesNo) = vbNo Then
        If CurServer <> 0 Then
            IPTxt = ServersLst(CurServer).Ip
            PortTxt = ServersLst(CurServer).Puerto
        Else
            IPTxt = IPdelServidor
            PortTxt = PuertoDelServidor
        End If
        Exit Sub
    End If
    CurServer = 0
    IPdelServidor = IPTxt
    PuertoDelServidor = PortTxt
End If


Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        
        If Musica = 0 Then
            CurMidi = DirMidi & "7.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmCrearPersonaje.Show vbModal
    Case 1
        frmOldPersonaje.Show vbModal
    Case 2
        frmBorrar.Show vbModal
    Case 3
        lst_servers.Visible = True
        Text1.Visible = False
    Case 4
        lst_servers.Visible = False
        Text1.Visible = True
        Text1.Text = ""
        eArchivo = Dir(App.Path & "\Web\soynuevo.txt")
        If eArchivo <> "soynuevo.txt" Or DescargarTxt(1) = True Then
            Text2.Text = "Descargando Reglamento..."
            frmMain.Inet1.URL = "http://www.aospain.com/soynuevo.txt"
            Archivo = frmMain.Inet1.OpenURL
            nArchivo = App.Path & "/Web/soynuevo.txt"
            Open nArchivo For Output As #1
                Print #1, Archivo
            Close
            Text2.Text = "Descarga Finalizada"
            DescargarTxt(1) = False
        End If
        eArchivo = App.Path & "\Web\soynuevo.txt"
        Open eArchivo For Input As #1
            While Not EOF(1)
                Line Input #1, cadena
                Text1.Text = Text1.Text + cadena + vbCrLf
            Wend
        Close
        Text1.SetFocus
    Case 5
        lst_servers.Visible = False
        Text1.Visible = True
        Text1.Text = ""
        eArchivo = Dir(App.Path & "\Web\reglamento.txt")
        If eArchivo <> "reglamento.txt" Or DescargarTxt(2) = True Then
            Text2.Text = "Descargando Reglamento..."
            frmMain.Inet1.URL = "http://www.aospain.com/reglamento.txt"
            Archivo = frmMain.Inet1.OpenURL
            nArchivo = App.Path & "/Web/reglamento.txt"
            Open nArchivo For Output As #1
                Print #1, Archivo
            Close
            Text2.Text = "Descarga Finalizada"
            DescargarTxt(2) = False
        End If
        eArchivo = App.Path & "\Web\reglamento.txt"
        Open eArchivo For Input As #1
            While Not EOF(1)
                Line Input #1, cadena
                Text1.Text = Text1.Text + cadena + vbCrLf
            Wend
        Close
        Text1.SetFocus
    Case 6
        lst_servers.Visible = False
        Text1.Visible = True
        Text1.Text = ""
        eArchivo = Dir(App.Path & "\Web\historia.txt")
        If eArchivo <> "historia.txt" Or DescargarTxt(3) = True Then
            Text2.Text = "Descargando Historia..."
            frmMain.Inet1.URL = "http://www.aospain.com/historia.txt"
            Archivo = frmMain.Inet1.OpenURL
            nArchivo = App.Path & "/Web/historia.txt"
            Open nArchivo For Output As #1
                Print #1, Archivo
            Close
            Text2.Text = "Descarga Finalizada"
            DescargarTxt(3) = False
        End If
        eArchivo = App.Path & "\Web\historia.txt"
        Open eArchivo For Input As #1
            While Not EOF(1)
                Line Input #1, cadena
                Text1.Text = Text1.Text + cadena + vbCrLf
            Wend
        Close
        Text1.SetFocus
    Case 7
        'abre la pagina de AOSpain.com
        ShellExecute frmMain.hwnd, vbNullString, "http://www.aospain.com", vbNullString, vbNullString, vbNormalFocus
End Select
End Sub

Private Sub imgGetPass_Click()
    Call PlayWaveDS(SND_CLICK)
    Call frmRecuperar.Show(vbModal, frmConnect)
End Sub

'Private Sub imgServArgentina_Click()
'    Call PlayWaveDS(SND_CLICK)
'    IPTxt.Text = IPdelServidor
'    PortTxt.Text = PuertoDelServidor
'End Sub

'Private Sub imgServEspana_Click()
'    Call PlayWaveDS(SND_CLICK)
'    IPTxt.Text = "62.42.193.233"
'    PortTxt.Text = "7666"
'End Sub



Private Sub lst_servers_Click()
CurServer = lst_servers.ListIndex + 1
DescTxt = ServersLst(CurServer).desc
IPTxt = ServersLst(CurServer).Ip
PortTxt = ServersLst(CurServer).Puerto
End Sub

