VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDxSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX As DirectX7
Private DS As DirectSound
Private DDMLoad As DirectMusicLoader
Private DDMPerf As DirectMusicPerformance
Private DSstepformat As WAVEFORMATEX

Private DDMSeg() As DirectMusicSegment
Private nMidi As Long

Private DSWav() As DirectSoundBuffer
Private DSPlayFlag() As CONST_DSBPLAYFLAGS ' controls looping on Waves
Private nWaves As Long

Private Sub Class_Initialize()
    Set DX = New DirectX7
    nWaves = -1
    nMidi = -1
End Sub ' Class_Initialize

Public Sub Initialize(frmParent As Form)
'   Must call this before any other routines
    MidiInit frmParent
    WavInit frmParent
End Sub ' Initialize

Public Sub Terminate()
'   Call when shutting down
    WavTerminate
    MidiTerminate
End Sub ' Terminate


' *******************************
' Routines for playing Wav files
' *******************************
Private Sub WavInit(frmParent As Form)
    Set DS = DX.DirectSoundCreate("")
    
    DS.SetCooperativeLevel frmParent.hWnd, DSSCL_PRIORITY
    
    DSstepformat.nFormatTag = WAVE_FORMAT_PCM
    DSstepformat.nChannels = 2
    DSstepformat.lSamplesPerSec = 22050
    DSstepformat.nBitsPerSample = 16
    DSstepformat.nBlockAlign = DSstepformat.nBitsPerSample / 8 * DSstepformat.nChannels
    DSstepformat.lAvgBytesPerSec = DSstepformat.lSamplesPerSec * DSstepformat.nBlockAlign
    
    If Err.Number <> DD_OK Then
        MsgBox "Error initializing DirectSound", vbExclamation, "Error"
    End If
End Sub ' WavInit

Private Sub WavTerminate()
Dim n As Long
    For n = 0 To nWaves
        WavStop (n)
    Next n
    
    Set DS = Nothing
End Sub ' WavTerminate

Public Function WavLoad(ByVal sfile As String, Optional ByVal fLoop = False) As Long
Dim bufferDesc As DSBUFFERDESC
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    nWaves = nWaves + 1
    ReDim Preserve DSWav(nWaves)
    Set DSWav(nWaves) = DS.CreateSoundBufferFromFile(sfile, bufferDesc, DSstepformat)
    ReDim Preserve DSPlayFlag(nWaves)
    If fLoop Then
        DSPlayFlag(nWaves) = DSBPLAY_LOOPING
    Else
        DSPlayFlag(nWaves) = DSBPLAY_DEFAULT
    End If
    WavLoad = nWaves
End Function ' WavLoad

Public Function WavLoop() As Boolean
'   Used with the WavLoop call
    WavLoop = True
End Function ' WavLoop

Public Sub WavPlay(ByVal n As Long)
    If n < 0 Or n > nWaves Then
        Exit Sub
    End If
    DSWav(n).Play DSPlayFlag(n)
End Sub ' WavPlay

Public Sub WavStop(ByVal n As Long)
    DSWav(n).Stop
    DSWav(n).SetCurrentPosition 0
End Sub ' WavStop



' *******************************
' Routines for playing MIDI files
' *******************************
Private Sub MidiInit(frmParent As Form)
    Set DDMLoad = DX.DirectMusicLoaderCreate
    Set DDMPerf = DX.DirectMusicPerformanceCreate
    
    DDMPerf.Init Nothing, frmParent.hWnd
    DDMPerf.SetPort -1, 1
    DDMPerf.SetMasterAutoDownload True
    
    If Err.Number <> DD_OK Then
        MsgBox "Error initializing DirectMusic", vbExclamation, "Error"
    End If
End Sub ' MidiInit

Private Sub MidiTerminate()
Dim n As Long
    For n = 0 To nMidi
        MidiStop (n)
        Set DDMSeg(n) = Nothing
    Next n
    
    Set DDMLoad = Nothing
    Set DDMPerf = Nothing
End Sub ' MidiTerminate

Public Function MidiLoad(Filename As String) As Long
    nMidi = nMidi + 1
    ReDim Preserve DDMSeg(nMidi)
    Set DDMSeg(nMidi) = DDMLoad.LoadSegment(Filename)
    If Err.Number <> DD_OK Then
        MsgBox "ERROR : Could not load MIDI file!", vbExclamation, "ERROR!"
    End If

    MidiLoad = nMidi
End Function ' MidiLoad

Public Sub MidiPlay(ByVal n As Long)
    If n < 0 Or n > nMidi Then
        Exit Sub
    End If

    DDMPerf.SetMasterVolume 5
    DDMPerf.PlaySegment DDMSeg(n), 0, 0
End Sub ' MidiPlay

Public Sub MidiStop(ByVal n As Long)
    If n < 0 Or n > nMidi Then
        Exit Sub
    End If

    DDMPerf.Stop DDMSeg(n), Nothing, 0, 0
End Sub ' MidiStop
