VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton desc 
      Caption         =   "Peticion"
      Height          =   495
      Left            =   2100
      MouseIcon       =   "frmCharInfo.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Echar 
      Caption         =   "Echar"
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frmCharInfo.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4200
      MouseIcon       =   "frmCharInfo.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Rechazar 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3120
      MouseIcon       =   "frmCharInfo.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCharInfo.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame rep 
      Caption         =   "Reputacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   5055
      Begin VB.Label reputacion 
         Caption         =   "Reputacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label criminales 
         Caption         =   "Criminales asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Ciudadanos 
         Caption         =   "Ciudadanos asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   5055
      Begin VB.Label faccion 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label integro 
         Caption         =   "Clanes que integro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lider 
         Caption         =   "Veces fue lider de clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label fundo 
         Caption         =   "Fundo el clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label solicitudesRechazadas 
         Caption         =   "Solicitudes rechazadas:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Solicitudes 
         Caption         =   "Solicitudes para ingresar a clanes:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame charinfo 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmCharInfo"
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



Public frmmiembros As Boolean
Public frmsolicitudes As Boolean

Private Sub Aceptar_Click()
frmmiembros = False
frmsolicitudes = False
Call SendData("ACEPTARI" & Right(Nombre, Len(Nombre) - 7))
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Public Sub parseCharInfo(ByVal Rdata As String)

If frmmiembros Then
    Rechazar.Visible = False
    Aceptar.Visible = False
    Echar.Visible = True
    desc.Visible = False
Else
    Rechazar.Visible = True
    Aceptar.Visible = True
    Echar.Visible = False
    desc.Visible = True
End If


'h$ = "CHRINFO" & UserName & ","
'h$ = h$ & MiUser.Raza & ","
'h$ = h$ & MiUser.Clase & ","
'h$ = h$ & MiUser.Genero & ","
'h$ = h$ & MiUser.Stats.ELV & ","
'h$ = h$ & MiUser.Stats.GLD & ","
'h$ = h$ & MiUser.Stats.Banco & ","
'h$ = h$ & MiUser.Reputacion.Promedio & ","


Nombre.Caption = "Nombre:" & ReadField(1, Rdata, 44)
Raza.Caption = "Raza:" & ReadField(2, Rdata, 44)
Clase.Caption = "Clase:" & ReadField(3, Rdata, 44)
Genero.Caption = "Genero:" & ReadField(4, Rdata, 44)
Nivel.Caption = "Nivel:" & ReadField(5, Rdata, 44)
Oro.Caption = "Oro:" & ReadField(6, Rdata, 44)
Banco.Caption = "Banco:" & ReadField(7, Rdata, 44)

Dim Y As Long, k As Long

Y = Val(ReadField(8, Rdata, 44))

If Y > 0 Then
    status.Caption = "Status: Ciudadano"
Else
    status.Caption = "Status: Criminal"
End If


'h$ = h$ & MiUser.GuildInfo.FundoClan & ","9
'h$ = h$ & MiUser.GuildInfo.EsGuildLeader & ","10
'h$ = h$ & MiUser.GuildInfo.Echadas & ","11
'h$ = h$ & MiUser.GuildInfo.Solicitudes & ","12
'h$ = h$ & MiUser.GuildInfo.solicitudesRechazadas & ","13
'h$ = h$ & MiUser.GuildInfo.VecesFueGuildLeader & ","14
'h$ = h$ & MiUser.GuildInfo.ClanesParticipo & ","15

'h$ = h$ & MiUser.GuildInfo.ClanFundado & ","16
'h$ = h$ & MiUser.GuildInfo.GuildName & ","17



Y = Val(ReadField(9, Rdata, 44))

Solicitudes.Caption = "Solicitudes para ingresar a clanes:" & ReadField(12, Rdata, 44)
solicitudesRechazadas.Caption = "Solicitudes rechazadas:" & ReadField(13, Rdata, 44)


If Y = 1 Then
    fundo.Caption = "Fundo el clan: " & ReadField(16, Rdata, 44)
Else
    fundo.Caption = "Fundo el clan: Ninguno"
End If


lider.Caption = "Veces fue lider de clan:" & ReadField(14, Rdata, 44)
integro.Caption = "Clanes que integro:" & ReadField(15, Rdata, 44)

'h$ = h$ & MiUser.faccion.ArmadaReal & "," 18
'h$ = h$ & MiUser.faccion.FuerzasCaos & "," 19
'h$ = h$ & MiUser.faccion.CiudadanosMatados & "," 20
'h$ = h$ & MiUser.faccion.CriminalesMatados 21

Y = Val(ReadField(18, Rdata, 44))

If Y = 1 Then
    faccion.Caption = "Faccion: Ejercito Real"
Else
    k = Val(ReadField(19, Rdata, 44))
    If k = 1 Then
        faccion.Caption = "Faccion: Fuerzas del caos"
    Else
        faccion.Caption = "Faccion: Ninguna"
    End If
End If

Ciudadanos.Caption = "Ciudadanos asesinados:" & ReadField(20, Rdata, 44)
criminales.Caption = "Criminales asesinados:" & ReadField(21, Rdata, 44)
reputacion.Caption = "Reputacion:" & Val(ReadField(8, Rdata, 44))
Me.Show vbModeless, frmMain


End Sub

Private Sub desc_Click()
Call SendData("ENVCOMEN" & Right(Nombre, Len(Nombre) - 7))
End Sub

Private Sub Echar_Click()
Call SendData("ECHARCLA" & Right(Nombre, Len(Nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Rechazar_Click()
Call SendData("RECHAZAR" & Right(Nombre, Len(Nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub
