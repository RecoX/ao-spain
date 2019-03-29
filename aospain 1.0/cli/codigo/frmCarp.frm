VERSION 5.00
Begin VB.Form frmCarp 
   Caption         =   "Carpintero"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      MouseIcon       =   "frmCarp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2910
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   255
      MouseIcon       =   "frmCarp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2910
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   675
   End
End
Attribute VB_Name = "frmCarp"
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



Private Sub Command3_Click()
'[Efestos]
Dim cant As Integer
Dim l As Integer
If Not IsNumeric(Text1.Text) Then
    MsgBox "Ingresa una cantidad numerica", vbCritical, "Carpinteria"
    Text1.Text = ""
    Text1.SetFocus
    Exit Sub
End If
If Text1.Text = "" Then
    MsgBox "Debes insertar la cantidad de objetos a construir", vbCritical, "Carpinteria"
    Text1.SetFocus
    Exit Sub
End If
If lstArmas.ListIndex = -1 Then
    MsgBox "No has seleccionado un item de la lista", vbCritical, "Carpinteria"
    Exit Sub
End If
cant = CInt(Text1.Text)
Select Case cant
    Case Is < 10
        l = 1
    Case Is < 100
        l = 2
    Case Is < 1000
        l = 3
    Case Else
        l = 4
End Select
On Error Resume Next

Call SendData("CNC" & l & ObjCarpintero(lstArmas.ListIndex) & cant)

Unload Me
'[Efestos]
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

