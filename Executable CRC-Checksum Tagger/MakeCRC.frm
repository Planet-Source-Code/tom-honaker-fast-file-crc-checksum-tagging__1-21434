VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MakeCRC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Executable CRC-Checksum Tagger"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MakeCRC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6555
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton GoButton 
      Caption         =   "Go!"
      Enabled         =   0   'False
      Height          =   645
      Left            =   5745
      TabIndex        =   5
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "Progress"
      Height          =   780
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6705
      Begin VB.Label ProgressLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I'm waiting... Choose a file!"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   6390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File to CRC-tag"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6705
      Begin VB.CommandButton PickTargetButton 
         Caption         =   "..."
         Height          =   345
         Left            =   5805
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox TargetFileBox 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "• The coder on Planet Source-Code known only as ""Detonate,"" for the idea that inspired this code."
      Height          =   465
      Left            =   195
      TabIndex        =   8
      Top             =   2475
      Width           =   5280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "• Fredrik Qvarfort, for his awesome (and FAST) CRC checksum code;"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   2265
      Width           =   4965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mad ShoutOuts to these people: (As in, the author would like to credit...)"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   2055
      Width           =   5280
   End
End
Attribute VB_Name = "MakeCRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' CRC Verification System - Tagger Portion
' ----------------------------------------
'
' Copyright (C) 2001 by Tom Honaker. All Rights Reserved. Permission granted
' to use, modify, and distribute this code as part of a compiled executable
' file, provided that all Copyright information is preserved and the Author
' is sent a copy of any modifications (E-mail: dmaster@gnt.net).
'
'

Private Sub Form_Load()
    ' Check to see if anything was passed from the command-line, which
    ' makes drag-and-drop file selection easy - drag the file to CRC-tag
    ' onto the compiled version of this code!
    If Command <> "" Then
        TargetFileBox.Text = Command
        GoButton.Enabled = True
        ProgressLabel.Caption = "Cleared for launch - click GO!"
    End If
End Sub

Private Sub GoButton_Click()
    
    Dim FileEnd#
    Dim FileNum%
    Dim TFile$
    Dim FBuffer$
    Dim cCRC As New clsCRC, FileCRC$, EndString$
    Dim N%, X%, A$
    
    ' Jump to a generic error handler if things no workie...
    On Error GoTo FileError
    
    ' Get the file from the textbox...
    TFile = TargetFileBox.Text
    
    ' Get the file's length and set up the file buffer...
    FileEnd = FileLen(TFile)
    FBuffer = Space(FileEnd)
    
    ' Time to open the file and grab the entire contents into the file buffer...
    FileNum = FreeFile
    Open TFile For Binary As #FileNum
    Get #FileNum, , FBuffer
    Close FileNum
    ' Note that we're using binary access and GETting the file, as compared to
    ' using sequential access. Each of the two methods gives a different CRC32
    ' result due to differences in VB's file-to-string handling in each case.
    ' We can use this behavior to our advantage by making it a bit harder to
    ' make the CRC32 if you don't know how to derive it the same way it's done
    ' here.
    
    ' Prep the CRC class and do a CRC check on the file's contents in the buffer...
    cCRC.Algorithm = CRC32
    cCRC.Clear
    FileCRC = Hex(cCRC.CalculateString(FBuffer))
    
    ' BTW, that CRC checksumming code Fredrik Qvarfort wrote (clsCRC) is
    ' FAST... It can easily manage CRC-checksumming an executable several
    ' megabytes in size in a second or so, depending on the computer's
    ' speed.
    
    ' If the CRC32 checksum is less that 8 characters long, pad its leading
    ' edge with zeroes. (Yeah, I know this is a bit of a kludge but hey,
    ' it works!)
    If Len(FileCRC) < 8 Then FileCRC = Left("00000000", 8 - Len(FileCRC)) & Hex(cCRC.CalculateString(FBuffer))
    
    ' Convert the 8-charcter hex value into 4 bytes and store that in the
    ' EndString buffer...
    EndString = ""
    For N = 1 To 4
        A = "&H" & Mid(FileCRC, ((N - 1) * 2) + 1, 1)
        X = Val(A) * 16
        A = "&H" & Mid(FileCRC, ((N - 1) * 2) + 2, 1)
        X = X + Val(A)
        EndString = EndString & Chr(X)
    Next
    
    ' Tack the 4-bite CRC32 tag onto the end of the file's contents...
    FBuffer = FBuffer & EndString
    
    ' Write the buffer, with the 4-byte compressed-CRC32 tag, back to the
    ' file...
    FileNum = FreeFile
    Open TFile For Binary As #FileNum
    Put #FileNum, , FBuffer
    Close FileNum
    ' Note we used the same bianry-access method here too - if you use Print#
    ' to write the data the CRC32 will not be correct!
    
    ' Done deal.
    TargetFileBox.Text = ""
    ProgressLabel.Caption = "DONE - " & Format(FileEnd, "###,###,###,###") & "-byte file - CRC32 is " & FileCRC
    GoButton.Enabled = False
    
    Exit Sub
    
FileError:
    On Error GoTo 0
    ProgressLabel.Caption = "ERROR - I wasn't able to write the CRC tag onto this file!"
End Sub

Private Sub PickTargetButton_Click()
    Dim FileChosen$
    
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.DialogTitle = "Please select a file to CRC-tag..."
    CommonDialog1.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
    CommonDialog1.ShowOpen
    
    FileChosen = CommonDialog1.Filename
    If FileChosen = "" Then Exit Sub
    
    TargetFileBox.Text = FileChosen
End Sub

Private Sub TargetFileBox_Change()
    GoButton.Enabled = True
    ProgressLabel.Caption = "Cleared for launch - click GO!"
End Sub
