Attribute VB_Name = "ModVocTools"
'VOC TO WAVE Convertor

'This helps in converting 8-bit Stereo/Mono VOC files to wave

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Type VOC_Info
    SIG As String * 19          'SIG must be Creative Voice File
    Termiator As Byte           'End Termiator chr(26)
    Offset  As Integer          'Start of data
    Version As Integer          'Version info
    ID As Integer               'Not sure
    BlockType As Byte           'Block type we only deal with 1 and 9
    BlockLen(1 To 3) As Byte    'Not to sure
    sRate As Byte               'Sample rate
    Reserved As Byte            'Asumed Junk
End Type

Function FindFile(lFile As String) As Boolean
    'Find a file
    FindFile = LenB(Dir(lFile))
End Function

Function VocToWave(VocFile As String, WaveFile As String) As Integer
Dim fp As Long
Dim vfi As VOC_Info
Dim sData As String
Dim SampleRate As Long
Dim m_Offset As Long
Dim Channels As Byte
Dim SampleSize As Byte
    
    VocToWave = -1
    
    If Not FindFile(VocFile) Then
        VocToWave = 3
        Exit Function
    End If
    
    fp = FreeFile
        Open VocFile For Binary As #fp
            Get #fp, , vfi
            
            'Check for vaild signature
            If Not (vfi.SIG = "Creative Voice File") Then
                VocToWave = 0
                GoTo Clean:
            End If
            
            'Check for end Termiator
            If (vfi.Termiator <> 26) Then
                VocToWave = 1
                GoTo Clean:
            End If
            
            Select Case vfi.BlockType
                Case 1
                    'Get the Sample rate
                    SampleRate = Round(1000000 / (256 - vfi.sRate))
                    'Create buffer to hold th data
                    sData = Space(LOF(fp) - Seek(fp))
                    'Get waveform data
                    Get #fp, , sData
                    'Create the wav file, we asume that this for now is 8-bit mono
                    CreateWav sData, WaveFile, SampleRate
                Case 9
                    'Skip over the block length we do not need it for this version
                    m_Offset = Seek(fp) - 2
                    'Get the Sample rate
                    Get #fp, m_Offset, SampleRate
                    'Get sample size
                    Get #fp, Seek(fp), SampleSize
                    
                    If (SampleSize <> 8) Then
                        VocToWave = 2
                        GoTo Clean:
                    End If
                    
                    'Get number of Channels
                    Get #fp, Seek(fp), Channels
                    'Create buffer to hold th data
                    sData = Space(LOF(fp) - Seek(fp))
                    'Get waveform data
                    Get #fp, , sData
                    'Create the wav file
                    CreateWav sData, WaveFile, SampleRate, SampleSize, Channels
                Case Else
                    Exit Function
            End Select
        Close #fp
        
        'Clear garbage
Clean:
        ZeroMemory vfi, Len(vfi)
        sData = vbNullString
        SampleRate = 0
        m_Offset = 0
        Channels = 0
End Function

Public Function VOCErrorCodes(ab_code As Integer) As String
    Select Case ab_code
        Case 0
            VOCErrorCodes = "Incorrect file signature."
        Case 1
            VOCErrorCodes = "Termiator not found."
        Case 2
            VOCErrorCodes = "Only 8-bit VOC files are not supported."
        Case 3
            VOCErrorCodes = "File was not found."
    End Select
End Function
