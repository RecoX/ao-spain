VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   1020
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2715
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FX Activados"
      Height          =   345
      Index           =   1
      Left            =   975
      MouseIcon       =   "frmOpciones.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   345
      Index           =   0
      Left            =   990
      MouseIcon       =   "frmOpciones.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   780
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   180
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
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



Private Sub Command1_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        If Musica = 0 Then
            Musica = 1
            Command1(0).Caption = "Musica Desactivada"
            Stop_Midi
        Else
            Musica = 0
            Command1(0).Caption = "Musica Activada"
            Call CargarMIDI(DirMidi & "2.mid")
            Play_Midi
            
        End If
    Case 1
    
        If Fx = 0 Then
            Fx = 1
            Command1(1).Caption = "FX Desactivados"
            
        Else
            Fx = 0
            Command1(1).Caption = "FX Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub



Private Sub Form_Deactivate()
Me.Visible = False
End Sub

Private Sub Form_Load()
If Musica = 0 Then
    Command1(0).Caption = "Musica Activada"
Else
    Command1(0).Caption = "Musica Desactivada"
End If

If Fx = 0 Then
    Command1(1).Caption = "FX Activados"
Else
    Command1(1).Caption = "FX Desactivados"
End If

End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

End Function
