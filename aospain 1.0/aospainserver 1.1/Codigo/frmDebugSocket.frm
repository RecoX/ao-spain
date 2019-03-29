VERSION 5.00
Begin VB.Form frmDebugSocket 
   Caption         =   "Debug Socket"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "State"
      Height          =   765
      Left            =   45
      TabIndex        =   8
      Top             =   4350
      Width           =   3705
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado."
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2835
      Width           =   3690
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reload Socket"
      Height          =   315
      Left            =   15
      TabIndex        =   4
      Top             =   6135
      Width           =   3750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   300
      Left            =   15
      TabIndex        =   3
      Top             =   5235
      Width           =   3720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start/Stop debug"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   5535
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   300
      Left            =   15
      TabIndex        =   1
      Top             =   5820
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   2280
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   225
      Width           =   3690
   End
   Begin VB.Label Label2 
      Caption         =   "Errores:"
      Height          =   315
      Left            =   30
      TabIndex        =   7
      Top             =   2610
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "Requests"
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   15
      Width           =   2685
   End
End
Attribute VB_Name = "frmDebugSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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



Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Command2_Click()
DebugSocket = Not DebugSocket
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Call ReloadSokcet
End Sub

