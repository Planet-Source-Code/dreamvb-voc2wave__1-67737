Attribute VB_Name = "ModWaveTools"
Option Explicit

'This is my Wave Bas, originally made to play sounds from the DOOM WAD Files and also create waves
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Type WavHeadInfo
    RIFF As String * 4
    BufferLen As Long
    WavID As String * 4
    FmtID As String * 4
    FmtLength As Long
    WavFmtTag As Integer
    Channels As Integer
    SampleRate As Long
    BytePerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    DataID As String * 4
    wDataLen As Long
End Type

Public Sub CreateWav(sWaveData As String, WavFile As String, _
Optional mSampleRate As Long = &H2B11, Optional mSamples As Byte = 8, Optional mChannels As Byte = 1)
Dim fp As Long
Dim wInfo As WavHeadInfo

    'Fill in WAV header info
    With wInfo
        .RIFF = "RIFF"
        .BufferLen = Len(sWaveData) + (44 - 8)
        .WavID = "WAVE"
        .FmtID = "fmt "
        .FmtLength = 16
        .WavFmtTag = 1
        .Channels = mChannels
        
        .SampleRate = mSampleRate
        .BytePerSec = (mSampleRate * .Channels) * (mSamples / mSamples)
        .BlockAlign = (.Channels * (mSamples / mSamples))
        .BitsPerSample = mSamples

        .DataID = "data"
        .wDataLen = (wInfo.BufferLen - 44)
    End With
    
    'Create the Wav File
    fp = FreeFile
    
    Open WavFile For Binary As #fp
        Put #fp, , wInfo
        Put #fp, , sWaveData
    Close #fp
    
    ZeroMemory wInfo, Len(wInfo)
    
End Sub

Function IsWavFile(lzFile As String) As Boolean
Dim fp As Long, SIG As String * 4
    fp = FreeFile
    
    Open lzFile For Binary As #fp
        Get #fp, , SIG
    Close #fp
    
    IsWavFile = (SIG = "RIFF") And (GetFileExt(lzFile) = "WAV")
    
End Function

Function Is8BitWav(lzFile As String) As Boolean
Dim wInfo As WavHeadInfo
Dim fp As Long
    
    'Check if the WAV file is 8 bit
    fp = FreeFile
    Open lzFile For Binary As #fp
        Get #fp, , wInfo
    Close #fp
    
    Is8BitWav = (Asc(wInfo.BitsPerSample) = 8)
    ZeroMemory wInfo, Len(wInfo)
    
End Function

Function IsMono(lzFile As String) As Boolean
Dim wInfo As WavHeadInfo
Dim fp As Long
    
    'Check if we have a mono WAV file
    fp = FreeFile
    Open lzFile For Binary As #fp
        Get #fp, , wInfo
    Close #fp
    
    IsMono = (Asc(wInfo.Channels) = 1)
    ZeroMemory wInfo, Len(wInfo)
    
End Function

Function GetWavData(lzFile As String) As String
Dim wInfo As WavHeadInfo
Dim fp As Long
Dim sData As String

    'Extract the WAV Data
    
    fp = FreeFile
    Open lzFile For Binary As #fp
        Get #fp, , wInfo
        sData = Space$(LOF(fp) - Len(wInfo))
        Get #fp, , sData
    Close #fp
    
    GetWavData = sData
    sData = ""
    
End Function

Function GetWavPlayTime(lzFile As String) As String
Dim wInfo As WavHeadInfo
Dim dLen As Long
Dim sPos As Long
Dim PlayLength As Double
Dim fp As Long

    'Check if we have a mono WAV file
    fp = FreeFile
    Open lzFile For Binary As #fp
        'Exit if file is empty
        If LOF(fp) = 0 Then Exit Function
        'Get Wav Header info
        Get #fp, , wInfo
        'Start position of wav data
        sPos = Seek(fp)
        dLen = LOF(fp) - sPos
        PlayLength = (dLen / wInfo.SampleRate / wInfo.BlockAlign)
        GetWavPlayTime = Left(PlayLength, 4)
    Close #fp
    
End Function

Function WavSampleRate(lzFile As String) As Long
Dim wInfo As WavHeadInfo
Dim fp As Long

    fp = FreeFile
    Open lzFile For Binary As #fp
        Get #fp, , wInfo
        WavSampleRate = wInfo.SampleRate
    Close #fp
End Function

