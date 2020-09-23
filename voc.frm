VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VocToWav"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   420
      Left            =   1995
      TabIndex        =   8
      Top             =   1740
      Width           =   1155
   End
   Begin VB.CommandButton cmdWave 
      Caption         =   "...."
      Enabled         =   0   'False
      Height          =   360
      Left            =   4590
      TabIndex        =   6
      Top             =   1080
      Width           =   555
   End
   Begin VB.TextBox txtWave 
      Height          =   330
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1110
      Width           =   3525
   End
   Begin VB.CommandButton cmdVoc 
      Caption         =   "...."
      Height          =   360
      Left            =   4575
      TabIndex        =   3
      Top             =   315
      Width           =   555
   End
   Begin VB.TextBox txtVoc 
      Height          =   330
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   3525
   End
   Begin MSComDlg.CommonDialog CDC 
      Left            =   5205
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   420
      Left            =   225
      TabIndex        =   0
      Top             =   1740
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save WAVE file to:"
      Height          =   195
      Left            =   1020
      TabIndex        =   7
      Top             =   825
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   195
      Left            =   255
      TabIndex        =   5
      Top             =   1170
      Width           =   675
   End
   Begin VB.Label lblVoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voc File:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   390
      Width           =   615
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Voc2Wav By DreamVB
'website http://www.programming-designs.com/
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Function GetPath(lzPath As String) As String
Dim e_pos As Integer
    e_pos = InStrRev(lzPath, "\", Len(lzPath), vbBinaryCompare)
    
    If (e_pos > 0) Then
        GetPath = Mid(lzPath, 1, e_pos - 1)
    Else
        GetPath = lzPath
    End If
End Function

Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Function ExtractFileTitle(lzFile As String) As String
Dim e_pos As Integer
Dim tmp As String

    tmp = lzFile
    e_pos = InStrRev(tmp, "\", Len(tmp), vbBinaryCompare)
    
    If (e_pos > 0) Then
        tmp = Mid$(tmp, e_pos + 1)
    End If
    
    e_pos = InStrRev(tmp, ".", Len(tmp), vbBinaryCompare)
    If (e_pos > 0) Then
        tmp = Mid(tmp, 1, e_pos - 1)
    End If
    
    ExtractFileTitle = tmp
    tmp = vbNullString
    e_pos = 0
End Function

Private Sub cmdConvert_Click()
Dim iRet As Integer
Dim ans As Integer

    iRet = VocToWave(txtVoc.Text, txtWave.Text)
    If Not (iRet) Then
        MsgBox "Error:" & vbCrLf & VOCErrorCodes(iRet), vbCritical, "Error#" & iRet
    Else
        ans = MsgBox("VOC File has now been converted." _
        & vbCrLf & "Do you want to test the wave file now", vbYesNo Or vbQuestion, frmmain.Caption)
        
        If (ans = vbYes) Then
            PlaySound txtWave.Text, App.hInstance, &H1 Or &H2
        End If
    End If
    
End Sub

Private Sub cmdVoc_Click()
On Error GoTo CanERR:
    With CDC
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "VOC Files (*.voc)|*.voc|"
        .ShowOpen
        txtVoc.Text = .FileName
        txtWave.Text = FixPath(GetPath(.FileName)) & ExtractFileTitle(.FileName) & ".wav"
        cmdWave.Enabled = True
        cmdConvert.Enabled = cmdWave.Enabled
        .FileName = ""
    End With
    
    Exit Sub
CanERR:
End Sub

Private Sub cmdWave_Click()
On Error GoTo CanERR:
    With CDC
        .CancelError = True
        .DialogTitle = "Save As"
        .Filter = "Wave Files (*.wav)|*.wav|"
        .ShowSave
        txtWave.Text = FixPath(GetPath(.FileName)) & ExtractFileTitle(.FileName) & ".wav"
        cmdWave.Enabled = True
        .FileName = ""
    End With
    Exit Sub
CanERR:
End Sub


Private Sub Command1_Click()
    MsgBox frmmain.Caption & vbCrLf & vbTab & "By DreamVB", vbInformation, "Exit..."
    Unload frmmain
End Sub
