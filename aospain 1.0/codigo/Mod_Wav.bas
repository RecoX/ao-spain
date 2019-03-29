Attribute VB_Name = "Mod_Wav"
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

Public Const SND_SYNC = &H0 ' SINCRONO
Public Const SND_ASYNC = &H1 ' ASINCRONO
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?WAVS¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public Const SND_CLICK = "click.Wav"

Public Const SND_PASOS1 = "23.Wav"
Public Const SND_PASOS2 = "24.Wav"
Public Const SND_NAVEGANDO = "50.wav"
Public Const SND_OVER = "click2.Wav"
Public Const SND_DICE = "cupdice.Wav"

Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sFile As String) As Boolean
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX

    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(DirSound & sFile, bufferDesc, waveFormat)

    If Err.Number <> 0 Then
        Exit Function
    End If

    LoadWavetoDSBuffer = True
End Function

Sub PlayWaveDS(File As String)
    If Fx = 1 Then Exit Sub

    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If

    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), File) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If
End Sub
