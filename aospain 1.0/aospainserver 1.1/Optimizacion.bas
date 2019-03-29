Attribute VB_Name = "Optimizacion"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

Public Declare Function SetKernelMaxUsers Lib "aokernel.dll" (ByVal MaxUsers As Long) As Long
Public Declare Function GetMapUsersButIndex Lib "aokernel.dll" (InputMapArray As Any, OutputIndexArray As Any, ByVal MapLocked As Integer, ByVal UserIndex As Integer) As Long
Public Declare Function GetMapUsers Lib "aokernel.dll" (InputMapArray As Any, OutputIndexArray As Any, ByVal MapLocked As Integer) As Long
Public Declare Function SendMapUsers Lib "aokernel.dll" (UsersMapArray As Any, UsersSocketArray As Any, ByVal MapLocked As Integer, ByVal TxtData As String) As Long
Public Declare Function SendMapUsersButIndex Lib "aokernel.dll" (UsersMapArray As Any, UsersSocketArray As Any, ByVal MapLocked As Integer, ByVal TxtData As String, ByVal UserIndex As Integer) As Long
Public Declare Function ReadFieldASM Lib "aokernel.dll" (ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
Public Declare Function send Lib "ws2_32.dll" (ByVal Handle As Long, ByVal Sdata As String, ByVal tam As Long, ByVal NoSe As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Sub CalculaModificadoresClase(ByVal UserIndex As Integer)
    UserList(UserIndex).ClaseModificadorEvasion = ModificadorEvasion(UserList(UserIndex).Clase)
    UserList(UserIndex).ClaseModificadorPoderAtaqueArmas = ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase)
    UserList(UserIndex).ClaseModificadorPoderAtaqueProyectiles = ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase)
    UserList(UserIndex).ClaseModicadorDañoClaseArmas = ModicadorDañoClaseArmas(UserList(UserIndex).Clase)
    UserList(UserIndex).ClaseModificadorPoderAtaqueProyectiles = ModicadorDañoClaseProyectiles(UserList(UserIndex).Clase)
    UserList(UserIndex).ClaseModEvasionDeEscudoClase = ModEvasionDeEscudoClase(UserList(UserIndex).Clase)
End Sub

Sub CalculaNivelModificador(ByVal UserIndex As Integer)
    UserList(UserIndex).NivelModificador = (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0))
End Sub

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
    ReadField = ReadFieldASM(Pos, Text, SepASCII)
End Function


Public Function GenCrc(ByVal CrcKey As Long, ByVal CrcString As String) As Long
GenCrc = 1232
End Function

