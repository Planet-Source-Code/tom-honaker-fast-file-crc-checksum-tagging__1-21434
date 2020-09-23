Attribute VB_Name = "CheckCRC"
Option Explicit
'
' CRC Verification System - Checker Portion
' -----------------------------------------
'
' Copyright (C) 2001 by Tom Honaker. All Rights Reserved. Permission granted
' to use, modify, and distribute this code as part of a compiled executable
' file, provided that all Copyright information is preserved and the Author
' is sent a copy of any modifications (E-mail: dmaster@gnt.net).
'
'

Private Function IDEorEXE() As Boolean
    ' This function detects whether the project is running as a standalone
    ' executable or within the IDE by taking advantage of one fact:
    ' Debug.Print is only executed in the IDE. SO, if we try to force a
    ' simple error with Debug.Print and get that error, we're running from
    ' the IDE. Simple, no?
    '
    ' Returns:
    ' True - App is running as a standalone file.
    ' False - App is running from within the IDE.
    
    ' Set up the local error handler...
    On Local Error GoTo NotInIDE
    
    ' Try to trip a divide by zero error from Debug.Print...
    Debug.Print 1 / 0
    
    ' No error, so we know we're not in the IDE.
    IDEorEXE = True
    
    Exit Function
    
NotInIDE:
    ' BOOM - Divide By Zero error, so we know we're still in the IDE.
    IDEorEXE = False
End Function

Private Function VerifyCRC() As Boolean
    ' This function verifies the CRC32 checksum code appended to the end of
    ' a file. Once the checksum is verified, the program can act on the
    ' results.
    '
    ' Returns:
    ' True - CRC32 checksum verified; this file is unchanged.
    ' False - CRC32 checksum mismatch or no checksum present at the end
    '         of the file.
    
    Dim FileEnd#                ' The file's size.
    Dim FBuffer$                ' This will hold the entire contents of the file.
    Dim TFile$                  ' Our target file to CRC-verify.
    Dim FileNum%                ' Used to derive the next available free file number for the Open command.
    Dim ChecksumBuffer$         ' Temporary holder for the checksum.
    Dim FileChecksum$           ' The file's CRC32 checksum as found at the end of the file.
    Dim N%, X%, A$              ' Temporary variables.
    Dim cCRC As New clsCRC      ' Dim a new instance of class clsCRC.
    Dim FileCRC$                ' The file's current CRC32 checksum.
    
    ' First thing: Check to see if we're in the IDE. If we are, we'll assume
    ' the file is okay and exit gracefully.
'    If IDEorEXE = False Then
'        VerifyCRC = True
'        Exit Function
'    End If
    
    ' Make sure you change the extension! I'm using it to verify the integrity
    ' of an ActiveX control, but it should work for just about anything you
    ' can execute (OCX, EXE, DLL, etc.)
    TFile$ = App.Path & "\" & App.EXEName & ".exe"
    
    ' Get the file's size and set up the file buffer...
    FileEnd = FileLen(TFile)
    FBuffer = Space(FileEnd)
    
    ' Time to open the file and grab the entire contents into the file buffer...
    On Error Resume Next
    FileNum = FreeFile
    Open TFile For Binary As #FileNum
    Get #FileNum, , FBuffer
    Close FileNum
    On Error GoTo 0
    ' Note that we're using binary access and GETting the file, as compared to
    ' using sequential access. Each of the two methods gives a different CRC32
    ' result due to differences in VB's file-to-string handling in each case.
    ' We can use this behavior to our advantage by making it a bit harder to
    ' make the CRC32 if you don't know how to derive it the same way it's done
    ' here.
    
    ' Separate off the last 4 bytes and move them to the ChecksumBuffer...
    ChecksumBuffer = Right(FBuffer, 4)
    FBuffer = Left(FBuffer, Len(FBuffer) - 4)
    
    ' Decode the 4-byte checksum trailer back into an 8-character hex value...
    FileChecksum = ""
    For N = 1 To 4
        A = Mid(ChecksumBuffer, N, 1)
        X = Asc(A)
        FileChecksum = FileChecksum & Hex(X)
    Next
    
    ' Prep the CRC class and do a CRC check on the file's contents in the buffer...
    cCRC.Algorithm = CRC32
    cCRC.Clear
    FileCRC = Hex(cCRC.CalculateString(FBuffer))
    
    ' Compare what the file says its CRC32 checksum should be against what it
    ' actually is at the moment...
    If FileCRC = FileChecksum Then
        ' Yep, they match.
        VerifyCRC = True
    Else
        ' Nope, no soap - something's changed or the file hasn't been tagged
        ' with a 4-byte CRC32 checksum trailer.
        VerifyCRC = False
    End If
End Function



