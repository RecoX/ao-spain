VERSION 5.00
Begin VB.Form frmayuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmayuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmayuda.frx":000C
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

