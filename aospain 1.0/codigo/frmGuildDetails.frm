VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
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
   ScaleHeight     =   6825
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Index           =   1
      Left            =   5160
      MouseIcon       =   "frmGuildDetails.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmGuildDetails.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   6495
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmGuildDetails.frx":02A4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame frmDesc 
      Caption         =   "Descripci�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit


Private Sub Command1_Click(Index As Integer)
Select Case Index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "�", , , vbBinaryCompare)
    
'    If Not AsciiValidos(fdesc$) Then
'        MsgBox "La descripcion contiene caracteres invalidos"
'        Exit Sub
'    End If
    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex1.UBound
'        If Not AsciiValidos(txtCodex1(k)) Then
'            MsgBox "El codex tiene invalidos"
'            Exit Sub
'        End If
        If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
    Next k
    If Cont < 4 Then
            MsgBox "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim chunk$
    
    If CreandoClan Then
        chunk$ = "CIG" & fdesc$
        chunk$ = chunk$ & "�" & ClanName & "�" & Site & "�" & Cont
    Else
        chunk$ = "DESCOD" & fdesc$ & "�" & Cont
    End If
    
    
    
    For k = 0 To txtCodex1.UBound
        chunk$ = chunk$ & "�" & txtCodex1(k)
    Next k
    
    
    Call SendData(chunk$)
    
    CreandoClan = False
    
    Unload Me
    
End Select



End Sub

Private Sub Form_Deactivate()

If Not frmGuildLeader.Visible Then
    Me.SetFocus
Else
    'Unload Me
End If

End Sub

