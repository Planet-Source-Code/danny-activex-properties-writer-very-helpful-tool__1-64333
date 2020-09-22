Attribute VB_Name = "ModMain"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Const APPNAME As String = "Properties Writer"

Private Const ICC_USEREX_CLASSES = &H200

Public bolEnding As Boolean 'Program ending?
Public bolResetWin As Boolean 'Resetting window position?
Public bolResetCol As Boolean 'Resetting columns?

Public Sub Main()
    On Error Resume Next
    
    Dim objRet As tagInitCommonControlsEx
    
    With objRet
        .lngSize = LenB(objRet)
        .lngICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControlsEx objRet
    frmMain.Show
End Sub

Public Sub HCenter(ByRef CenterObject As Object, ByRef CenterOn As Object)
    On Error Resume Next
    
    CenterObject.Left = (CenterOn.Width * 0.5) - (CenterObject.Width * 0.5)
End Sub

Public Sub VCenter(ByRef CenterObject As Object, ByRef CenterOn As Object)
    On Error Resume Next
    
    CenterObject.Top = (CenterOn.Height * 0.5) - (CenterObject.Height * 0.5)
End Sub

Public Sub NumberOnly(ByRef KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Public Sub SelectAll(TextObject As Object)
    On Error Resume Next
    
    With TextObject
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'End the program properly.
Public Sub EndProgram()
    On Error Resume Next 'Do not stop on error.
    
    If bolEnding = False Then
        bolEnding = True
        
        Unload frmAbout
        Unload frmPreferences
        Unload frmCode
        Unload frmMain
    End If
End Sub

'Check if a file exists.
'Faster than using the Dir() method.
Public Function FileExists(ByVal FilePath As String) As Boolean
    Dim lonFF As Long
    
    On Error GoTo ErrorHandler
    
    lonFF = FreeFile
    
    Open FilePath For Input As #lonFF
    Close #lonFF
    
    FileExists = True
    
    Exit Function
    
ErrorHandler:
    '53 - File not found.
    
    If Not Err.Number = 53 And Not Err.Number = 72 Then
        FileExists = True
    End If
    
End Function

Public Sub SafeKill(ByVal FilePath As String)
    On Error GoTo ErrorHandler
    
    Kill FilePath
    
    Exit Sub
    
ErrorHandler:
End Sub

Public Function SafeFileLen(ByVal FilePath As String) As Long
    On Error GoTo ErrorHandler
    
    SafeFileLen = FileLen(FilePath)
    
    Exit Function
    
ErrorHandler:
End Function

Public Function UBoundStr(StringArray() As String) As Long
    On Error GoTo ErrorHandler
    
    UBoundStr = UBound(StringArray())
    
    Exit Function
    
ErrorHandler:
End Function

Public Function GetFileExtension(ByVal FilePath As String) As String
    On Error GoTo ErrorHandler
    
    GetFileExtension = LCase$(Mid$(FilePath, (InStrRev(FilePath, ".") + 1)))
    
    Exit Function
ErrorHandler:
        
End Function
