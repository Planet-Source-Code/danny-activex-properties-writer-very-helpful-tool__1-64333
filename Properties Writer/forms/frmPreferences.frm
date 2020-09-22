VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGen 
      Caption         =   " Code Generator "
      Height          =   2775
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtIndentWidth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkIndentTab 
         Caption         =   "Indent using real tabs."
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   3975
      End
      Begin VB.CheckBox chkIndentWithEnd 
         Caption         =   "Indent code between with/end with statements."
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CheckBox chkIndentSub 
         Caption         =   "Indent code between sub routine/end sub."
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Caption         =   "Add comments to code."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         Caption         =   "Indent width:                  spaces"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   2160
         Width           =   2280
      End
   End
   Begin VB.Frame fraBehave 
      Caption         =   " Program Behavior "
      Height          =   2775
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin PropertiesWriter.XPButton cmdClearWindow 
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "Reset"
      End
      Begin VB.CheckBox chkSaveClear 
         Caption         =   "Save property clear options when closing."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3495
      End
      Begin VB.CheckBox chkSaveListCol 
         Caption         =   "Save list column sizes when closing."
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.CheckBox chkSaveWinPos 
         Caption         =   "Save window position when closing."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin PropertiesWriter.XPButton cmdClearCol 
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "Reset"
      End
   End
   Begin VB.ListBox lstCat 
      Height          =   2790
      ItemData        =   "frmPreferences.frx":0A02
      Left            =   120
      List            =   "frmPreferences.frx":0A0C
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin PropertiesWriter.XPButton cmdApply 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Apply"
      Icon            =   "frmPreferences.frx":0A32
      IconMask        =   "frmPreferences.frx":0D84
   End
   Begin PropertiesWriter.XPButton cmdCancel 
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Cancel"
      Icon            =   "frmPreferences.frx":10D6
      IconMask        =   "frmPreferences.frx":1428
   End
   Begin PropertiesWriter.XPButton cmdDefault 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Defaults..."
      Icon            =   "frmPreferences.frx":177A
      IconMask        =   "frmPreferences.frx":1ACC
   End
   Begin PropertiesWriter.XPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Help"
      Icon            =   "frmPreferences.frx":1E1E
      IconMask        =   "frmPreferences.frx":2170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   7080
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   7080
      Y1              =   3015
      Y2              =   3000
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Used to show the help file.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Used for the ShellExecute() API function.
Private Const SW_NORMAL = 1

Private Enum PREF_CATEGORY
    [Program Behavior] = 0
    [Code Generator] = 1
End Enum

Private Sub ShowCategory(ByVal Category As PREF_CATEGORY)
    If Category = [Program Behavior] Then
        fraGen.Visible = False
        fraBehave.Visible = True
    ElseIf Category = [Code Generator] Then
        fraBehave.Visible = False
        fraGen.Visible = True
    End If
End Sub

Private Sub chkIndentTab_Click()
    If chkIndentTab.Value = 1 Then
        lblDisplay(0).Enabled = False
        txtIndentWidth.Enabled = False
    Else
        lblDisplay(0).Enabled = True
        txtIndentWidth.Enabled = True
    End If
End Sub

Private Sub cmdApply_Click()
    txtIndentWidth.Text = Trim$(txtIndentWidth.Text)
    
    If chkIndentTab.Value = 1 And Len(txtIndentWidth.Text) = 0 Then
        MsgBox "Please enter an indent width or enable the 'Indent using real tabs' option", vbCritical, "Indent Width Required"
        Exit Sub
    End If
    
    If IsNumeric(txtIndentWidth.Text) = False Then
        MsgBox "Please enter an integer for the indent width", vbCritical, "Invalid Value"
        ShowCategory [Code Generator]
        SelectAll txtIndentWidth
        Exit Sub
    End If
    
    With udtPref
        .intComment = chkComment.Value
        .intIndentSub = chkIndentSub.Value
        .intIndentTab = chkIndentTab.Value
        .intIndentWidth = Val(txtIndentWidth.Text)
        .intSaveClear = chkSaveClear.Value
        .intSaveListCol = chkSaveListCol.Value
        .intSaveWinPos = chkSaveWinPos.Value
    End With
    
    SavePref
    ReadPref True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearCol_Click()
    On Error Resume Next
    
    Dim intLoop As Integer
    
    For intLoop = 1 To 6
        DeleteSetting APPNAME, "Main", "Col" & intLoop
    Next intLoop
    
    bolResetCol = True
    
    MsgBox "The column sizes have been reset. These settings will not take effect until you restart the program", vbInformation, "Columns Reset"
End Sub

Private Sub cmdClearWindow_Click()
    On Error Resume Next
    
    DeleteSetting APPNAME, "Main", "Left"
    DeleteSetting APPNAME, "Main", "Top"
    DeleteSetting APPNAME, "Main", "Width"
    DeleteSetting APPNAME, "Main", "Height"
    DeleteSetting APPNAME, "Main", "WS"
    
    bolResetWin = True
    
    MsgBox "The window position has been reset. These settings will not take effect until you restart the program", vbInformation, "Settings Reset"
End Sub

Private Sub cmdDefault_Click()
    Dim msgResult As VbMsgBoxResult
    
    msgResult = MsgBox("Are you sure you want to load the default settings?", vbQuestion + vbYesNo, "Load Defaults")
    
    If msgResult = vbYes Then
        LoadPrefDefaults
        SavePref
        ReadPref True
    End If
End Sub

Private Sub cmdHelp_Click()
    Dim lonRet As Long
    
    lonRet = ShellExecute(Me.hwnd, "open", App.Path & "\help.htm", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Form_Load()
    lstCat.ListIndex = 0
    ReadPref True
    lblDisplay(0).Enabled = Not chkIndentTab.Value
    
    If chkIndentTab.Value = 1 Then
        lblDisplay(0).Enabled = False
        txtIndentWidth.Enabled = False
    Else
        lblDisplay(0).Enabled = True
        txtIndentWidth.Enabled = True
    End If
End Sub

Private Sub lstCat_Click()
    With lstCat
        
        If Len(.List(.ListIndex)) > 0 Then
            
            Select Case .ListIndex
                
                Case 0
                    ShowCategory [Program Behavior]
                
                Case 1
                    ShowCategory [Code Generator]
                
                Case Else
                    ShowCategory [Program Behavior]
            End Select
        
        End If
    
    End With
End Sub

Private Sub txtIndentWidth_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
End Sub
