VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Properties Writer"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog objCD 
      Left            =   1200
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Ready."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lvList 
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Container"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Default value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Object"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Read-only"
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.CheckBox chkDefault 
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      ToolTipText     =   "Clear default value after adding"
      Top             =   1200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkCont 
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      ToolTipText     =   "Clear container after adding"
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkType 
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      ToolTipText     =   "Clear data type after adding"
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkName 
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Clear property name after adding"
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin PropertiesWriter.XPButton cmdAdd 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Add"
      Icon            =   "frmMain.frx":0A02
      IconMask        =   "frmMain.frx":0D54
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Property is read-only"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox chkObject 
      Caption         =   "Property is object"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtDefault 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtCont 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmMain.frx":10A6
      Left            =   1800
      List            =   "frmMain.frx":10C5
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin PropertiesWriter.XPButton cmdRemove 
      Height          =   375
      Left            =   1733
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Remove"
      Icon            =   "frmMain.frx":1117
      IconMask        =   "frmMain.frx":1469
   End
   Begin PropertiesWriter.XPButton cmdEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Edit"
      Icon            =   "frmMain.frx":17BB
      IconMask        =   "frmMain.frx":1B0D
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   120
      X2              =   4800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":1E5F
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Default value:"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   4800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Container:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   765
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":2861
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Data type:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   780
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":3263
      Top             =   480
      Width           =   240
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Property name:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":3C65
      Top             =   120
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   4800
      Y1              =   1695
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   135
      X2              =   4800
      Y1              =   2175
      Y2              =   2160
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "&Generate"
      Begin VB.Menu mnuGenerateCode 
         Caption         =   "Generate Code"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Require variable declaration.

'Used to show the help file.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Used for the ShellExecute() API function.
Private Const SW_NORMAL = 1

'File extension for property list files.
'Stands for 'User Control Properties'.
Private Const PW_EXT As String = "ucp"

'Seperator (delimeter) for property list file.
Private Const PW_DELIM As String = "" 'Chr$(1) & Chr$(22) & Chr$(17)

'End of property information for property list file.
Private Const PW_EOP As String = "???"

'Section in registry to store this window's properties.
Private Const SEC As String = "Main"

Private strCurFile As String 'Current property list file being used.

Private bolChanged As Boolean 'User has made changes to the file?

'Read window settings.
Private Sub ReadWinSettings()
    On Error Resume Next 'Do not stop on error.
    
    Dim lonCX As Long, lonCY As Long
    Dim intLoop As Integer
    
    lonCX = (Screen.Width * 0.5) - (Me.Width * 0.5)
    lonCY = (Screen.Height * 0.5) - (Me.Height * 0.5)
    
    If udtPref.intSaveWinPos = 1 Then
        Me.Left = GetSetting(APPNAME, SEC, "Left", lonCX)
        Me.Top = GetSetting(APPNAME, SEC, "Top", lonCY)
        Me.Width = GetSetting(APPNAME, SEC, "Width", 5040)
        Me.Height = GetSetting(APPNAME, SEC, "Height", 6105)
        Me.WindowState = GetSetting(APPNAME, SEC, "WS", 0)
    Else
        DeleteSetting APPNAME, SEC, "Left"
        DeleteSetting APPNAME, SEC, "Top"
        DeleteSetting APPNAME, SEC, "Width"
        DeleteSetting APPNAME, SEC, "Height"
        DeleteSetting APPNAME, SEC, "WS"
        
        Me.Left = lonCX
        Me.Top = lonCY
        Me.Width = 5040
        Me.Height = 6105
        Me.WindowState = 0
    End If
    
    If udtPref.intSaveClear = 1 Then
        chkName.Value = GetSetting(APPNAME, SEC, "CName", 1)
        chkType.Value = GetSetting(APPNAME, SEC, "CType", 0)
        chkCont.Value = GetSetting(APPNAME, SEC, "CCont", 1)
        chkDefault.Value = GetSetting(APPNAME, SEC, "CDef", 1)
    Else
        DeleteSetting APPNAME, SEC, "CName"
        DeleteSetting APPNAME, SEC, "CType"
        DeleteSetting APPNAME, SEC, "CCont"
        DeleteSetting APPNAME, SEC, "CDef"
        
        chkName.Value = 1
        chkType.Value = 0
        chkCont.Value = 1
        chkDefault.Value = 1
    End If
    
    If udtPref.intSaveListCol = 1 Then
        
        With lvList
            .ColumnHeaders(1).Width = GetSetting(APPNAME, SEC, "Col1", 2440.06)
            .ColumnHeaders(2).Width = GetSetting(APPNAME, SEC, "Col2", 1840.25)
            .ColumnHeaders(3).Width = GetSetting(APPNAME, SEC, "Col3", 1440)
            .ColumnHeaders(4).Width = GetSetting(APPNAME, SEC, "Col4", 1440)
            .ColumnHeaders(5).Width = GetSetting(APPNAME, SEC, "Col5", 840.18)
            .ColumnHeaders(6).Width = GetSetting(APPNAME, SEC, "Col6", 840.18)
        End With
    
    Else
        
        For intLoop = 1 To 6
            DeleteSetting APPNAME, SEC, "Col" & intLoop
        Next intLoop
        
        With lvList
            .ColumnHeaders(1).Width = 2440.06
            .ColumnHeaders(2).Width = 1840.25
            .ColumnHeaders(3).Width = 1440
            .ColumnHeaders(4).Width = 1440
            .ColumnHeaders(5).Width = 840.18
            .ColumnHeaders(6).Width = 840.18
        End With
    
    End If
End Sub
'Save window settings.
Private Sub SaveWinSettings()
    On Error Resume Next 'Do not want to stop on error.
    
    Dim intLoop As Integer 'Loop counter.
    
    'Save window dimensions.
    If udtPref.intSaveWinPos = 1 And bolResetWin = False Then
        SaveSetting APPNAME, SEC, "Left", Me.Left
        SaveSetting APPNAME, SEC, "Top", Me.Top
        SaveSetting APPNAME, SEC, "Width", Me.Width
        SaveSetting APPNAME, SEC, "Height", Me.Height
        SaveSetting APPNAME, SEC, "WS", Me.WindowState
    Else
        DeleteSetting APPNAME, SEC, "Left"
        DeleteSetting APPNAME, SEC, "Top"
        DeleteSetting APPNAME, SEC, "Width"
        DeleteSetting APPNAME, SEC, "Height"
        DeleteSetting APPNAME, SEC, "WS"
    End If
    
    'Save clear options.
    If udtPref.intSaveClear = 1 Then
        SaveSetting APPNAME, SEC, "CName", chkName.Value
        SaveSetting APPNAME, SEC, "CType", chkType.Value
        SaveSetting APPNAME, SEC, "CCont", chkCont.Value
        SaveSetting APPNAME, SEC, "CDef", chkDefault.Value
    Else
        DeleteSetting APPNAME, SEC, "CName"
        DeleteSetting APPNAME, SEC, "CType"
        DeleteSetting APPNAME, SEC, "CCont"
        DeleteSetting APPNAME, SEC, "CDef"
    End If
    
    'Save column headers.
    If udtPref.intSaveListCol = 1 And bolResetCol = False Then
        
        With lvList
            
            For intLoop = 1 To .ColumnHeaders.Count
                SaveSetting APPNAME, SEC, "Col" & intLoop, .ColumnHeaders(intLoop).Width
            Next intLoop
        
        End With
    
    Else
        
        For intLoop = 1 To 6
            DeleteSetting APPNAME, SEC, "Col" & intLoop
        Next intLoop
        
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim strPropName As String 'Property name.
    Dim strType As String 'Data type.
    Dim strCont As String 'Container variable.
    Dim strDefault As String 'Default value.
    Dim lonCount As Long 'Total number of items in list.
    
    'Check user input.
    strPropName = Trim$(txtName.Text)
    strType = Trim$(cmbType.Text)
    strCont = Trim$(txtCont.Text)
    strDefault = Trim$(txtDefault.Text)
    
    If Len(strPropName) = 0 Then
        'User did not enter a property name.
        MsgBox "Please enter a property name", vbCritical, "Property Name Required"
        txtName.SetFocus
        Exit Sub
    ElseIf Len(strType) = 0 Then
        'User did not enter a data type.
        MsgBox "Please enter a data type", vbCritical, "Data Type Required"
        cmbType.SetFocus
        Exit Sub
    ElseIf Len(strCont) = 0 Then
        'User did not enter a container variable.
        MsgBox "Please enter a container variable", vbCritical, "Container Required"
        txtCont.SetFocus
        Exit Sub
    ElseIf Len(strDefault) = 0 And chkReadOnly.Value = 0 Then
        'User did not enter a default value.
        MsgBox "Please enter a default value", vbCritical, "Default Value Required"
        txtDefault.SetFocus
        Exit Sub
    End If
    
    'Check if this property name has already been added.
    If PropertyAdded(strPropName) = True Then
        MsgBox "This property name has already been added", vbCritical, "Property Already Added"
        SelectAll txtName
        Exit Sub
    End If
    
    With lvList
        .ListItems.Add , , strPropName
        lonCount = .ListItems.Count
        .ListItems(lonCount).SubItems(1) = strType
        .ListItems(lonCount).SubItems(2) = strCont
        .ListItems(lonCount).SubItems(3) = strDefault
        .ListItems(lonCount).SubItems(4) = CBool(chkObject.Value)
        .ListItems(lonCount).SubItems(5) = CBool(chkReadOnly.Value)
    End With
    
    If chkName.Value = 1 Then txtName.Text = ""
    If chkType.Value = 1 Then cmbType.Text = ""
    If chkCont.Value = 1 Then txtCont.Text = ""
    If chkDefault.Value = 1 Then txtDefault.Text = ""
    
    chkObject.Value = 0
    chkReadOnly.Value = 0
    
    If Len(txtName.Text) = 0 Then
        txtName.SetFocus
    ElseIf Len(cmbType.Text) = 0 Then
        cmbType.SetFocus
    ElseIf Len(txtCont.Text) = 0 Then
        txtCont.SetFocus
    ElseIf Len(txtDefault.Text) = 0 Then
        txtDefault.SetFocus
    Else
        SelectAll txtName
    End If
    
    bolChanged = True
End Sub

Private Sub cmdEdit_Click()
    If LCase$(cmdEdit.Caption) = "edit" Then
        Dim lonSel As Long, strName As String
        Dim strType As String, strCont As String
        Dim strDef As String, bolObj As Boolean
        Dim bolReadOnly As Boolean, lonFind As Long
        
        With lvList
            
            If .ListItems.Count > 0 Then
                lonSel = .SelectedItem.Index
                
                If lonSel > 0 Then
                    PropertyInfoByIndex lonSel, strName, strType, strCont, strDef, bolObj, bolReadOnly
                    
                    txtName.Text = strName
                    cmbType.Text = strType
                    txtCont.Text = strCont
                    txtDefault.Text = strDef
                    chkObject.Value = Abs(CInt(bolObj))
                    chkReadOnly.Value = Abs(CInt(bolReadOnly))
                    
                    cmdEdit.Caption = "Update"
                    bolChanged = True
                Else
                    MsgBox "Please select a property to edit", vbCritical, "No Property Selected"
                    Exit Sub
                End If
            
            End If
        
        End With
    
    ElseIf LCase$(cmdEdit.Caption) = "update" Then
        lonFind = FindProperty(txtName.Text)
        
        If lonFind > 0 Then
            
            With lvList
                .ListItems(lonFind).Text = txtName.Text
                .ListItems(lonFind).SubItems(1) = cmbType.Text
                .ListItems(lonFind).SubItems(2) = txtCont.Text
                .ListItems(lonFind).SubItems(3) = txtDefault.Text
                .ListItems(lonFind).SubItems(4) = CBool(chkObject.Value)
                .ListItems(lonFind).SubItems(5) = CBool(chkReadOnly.Value)
            End With
            
            If chkName.Value = 1 Then txtName.Text = ""
            If chkType.Value = 1 Then cmbType.Text = ""
            If chkCont.Value = 1 Then txtCont.Text = ""
            If chkDefault.Value = 1 Then txtDefault.Text = ""
            
            chkObject.Value = 0
            chkReadOnly.Value = 0
            
            If Len(txtName.Text) = 0 Then
                txtName.SetFocus
            ElseIf Len(cmbType.Text) = 0 Then
                cmbType.SetFocus
            ElseIf Len(txtCont.Text) = 0 Then
                txtCont.SetFocus
            ElseIf Len(txtDefault.Text) = 0 Then
                txtDefault.SetFocus
            Else
                SelectAll txtName
            End If
            
            bolChanged = True
            cmdEdit.Caption = "Edit"
        Else
            MsgBox "Unable to update property", vbCritical, "Property Not Found"
        End If
    
    End If
End Sub

'Find a property in the list and return its index.
Private Function FindProperty(ByVal PropertyName As String) As Long
    Dim lonLoop As Long, strCompName As String
    Dim lonCount As Long
    
    With lvList
        lonCount = .ListItems.Count
        
        If lonCount > 0 Then
            strCompName = LCase$(PropertyName)
            
            For lonLoop = 1 To lonCount
                
                If LCase$(.ListItems(lonLoop).Text) = strCompName Then
                    FindProperty = lonLoop
                    Exit For
                End If
            
            Next lonLoop
        
        End If
    
    End With
End Function

'Get a property's information from the list.
Private Sub PropertyInfoByIndex(ByVal Index As Long, Optional ByRef Name As String, _
    Optional ByRef DataType As String, Optional ByRef Container As String, _
    Optional ByRef Default As String, Optional ByRef IsObject As Boolean, _
    Optional ByRef ReadOnly As Boolean)

    Dim lonSel As Long
    
    If Index > 0 Then
        
        With lvList
            Name = .ListItems(Index).Text
            DataType = .ListItems(Index).SubItems(1)
            Container = .ListItems(Index).SubItems(2)
            Default = .ListItems(Index).SubItems(3)
            
            If LCase$(.ListItems(Index).SubItems(4)) = "true" Then
                IsObject = True
            Else
                IsObject = False
            End If
            
            If LCase$(.ListItems(Index).SubItems(5)) = "true" Then
                ReadOnly = True
            Else
                ReadOnly = False
            End If
            
        End With
    
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lonSel As Long
    Dim msgRes As VbMsgBoxResult
    
    With lvList
        
        If .ListItems.Count > 0 Then
            lonSel = .SelectedItem.Index
            
            If lonSel > 0 Then
                msgRes = MsgBox("Are you sure you want to remove '" & .ListItems(lonSel).Text & "'?", vbQuestion + vbYesNo, "Remove Property")
                
                If msgRes = vbYes Then
                    .ListItems.Remove lonSel
                    bolChanged = True
                End If
            
            Else
                MsgBox "Please select a property to remove", vbCritical, "No Property Selected"
            End If
        
        End If
    
    End With
End Sub

Private Sub Form_Load()
    'Read the preferences from the registry.
    'See the ModPreferences module.
    ReadPref
    ReadWinSettings
    strCurFile = "***"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bolEnding = False Then
        SaveWinSettings
        EndProgram
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Dim intLoop As Integer
    
    For intLoop = 0 To 3
        Line1(intLoop).X2 = Me.Width - 240
    Next intLoop
    
    lvList.Width = Me.Width - 345
    lvList.Height = Me.Height - 3930
    
    cmdEdit.Top = lvList.Height + 2385
    cmdRemove.Top = cmdEdit.Top
    cmdAdd.Top = cmdRemove.Top
    HCenter cmdRemove, Me
    cmdAdd.Left = lvList.Width - 1335
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Click()
    Dim lonCount As Long
    
    lonCount = lvList.ListItems.Count
    
    mnuSave.Enabled = (lonCount > 0) And (bolChanged = True) And (InStr(1, strCurFile, "***") = 0)
    mnuSaveAs.Enabled = (lonCount > 0)
End Sub

Private Sub mnuGenerate_Click()
    mnuGenerateCode.Enabled = lvList.ListItems.Count
End Sub

'Gets the indent width to use (from the preferences).
'Returns -1 for a standard 'real' tab (vbTab).
'Returns length of spaces (if set).
Private Function GetIndentWidth() As Integer
    If udtPref.intIndentTab = 1 Then
        GetIndentWidth = -1
    Else
        
        If udtPref.intIndentWidth = 0 Then
            GetIndentWidth = -1
        Else
            GetIndentWidth = udtPref.intIndentWidth
        End If
    
    End If
End Function

Private Sub mnuGenerateCode_Click()
    Dim lonLoop As Long, strName As String
    Dim strType As String, strCont As String
    Dim strDef As String, bolObject As Boolean
    Dim bolReadOnly As Boolean, strRet As clsConcat
    Dim lonCount As Long, intWidth As Integer
    
    Set strRet = New clsConcat
    strRet.ReInit
    
    With lvList
        lonCount = .ListItems.Count
        
        If lonCount > 0 Then
            
            'Generate private declarations.
            
            'Check if we need to add comments.
            If udtPref.intComment = 1 Then
                'Add comments.
                strRet.SConcat "'Private declarations for property containers." & vbNewLine
            End If
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObject, bolReadOnly
                
                'Make sure the property isn't read-only.
                If bolReadOnly = False Then
                    strRet.SConcat "Private " & strCont & " As " & strType & vbNewLine
                End If
                
            Next lonLoop
            
            'Generate property get, let/set routines.
            strRet.SConcat vbNewLine
            
            'Check if we need to add comments.
            If udtPref.intComment = 1 Then
                'Add comments.
                strRet.SConcat "'Property get, set/let routines." & vbNewLine
            End If
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObject, bolReadOnly
                
                strRet.SConcat "Public Property Get " & strName & "() As " & strType & vbNewLine
                
                'Check if we need to indent the code.
                If udtPref.intIndentSub = 1 Then
                    intWidth = GetIndentWidth
                    
                    If intWidth = -1 Then
                        strRet.SConcat vbTab
                    Else
                        strRet.SConcat String$(intWidth, Chr$(32))
                    End If
                    
                End If
                
                'Add code for property get routine.
                'Check if the property is an object.
                If bolObject = True Then
                    strRet.SConcat "Set "
                Else
                    strRet.SConcat "Let "
                End If
                
                strRet.SConcat strName & " = " & strCont & vbNewLine
                strRet.SConcat "End Property" & vbNewLine & vbNewLine
                
                'Generate property let/set routine (unless it is read-only).
                If bolReadOnly = False Then
                    strRet.SConcat "Public Property "
                    
                    'Check if the property is an object.
                    If bolObject = True Then
                        strRet.SConcat "Set "
                    Else
                        strRet.SConcat "Let "
                    End If
                    
                    strRet.SConcat strName & "("
                    
                    'Check if the property is an object.
                    'Objects are passed ByRef, others are passed ByVal to save memory.
                    If bolObject = True Then
                        strRet.SConcat "ByRef "
                    Else
                        strRet.SConcat "ByVal "
                    End If
                    
                    strRet.SConcat "NewValue As " & strType & ")" & vbNewLine
                    
                    'Check if we need to indent the code.
                    If udtPref.intIndentSub = 1 Then
                        intWidth = GetIndentWidth
                        
                        If intWidth = -1 Then
                            strRet.SConcat vbTab
                        Else
                            strRet.SConcat String$(intWidth, Chr$(32))
                        End If
                        
                    End If
                    
                    'Check if property is an object.
                    If bolObject = True Then
                        strRet.SConcat "Set "
                    Else
                        strRet.SConcat "Let "
                    End If
                    
                    strRet.SConcat strCont & " = NewValue" & vbNewLine
                    
                    'Check if we need to indent the code.
                    If udtPref.intIndentSub = 1 Then
                        intWidth = GetIndentWidth
                        
                        If intWidth = -1 Then
                            strRet.SConcat vbTab
                        Else
                            strRet.SConcat String$(intWidth, Chr$(32))
                        End If
                        
                    End If
                    
                    'Generate PropertyChanged() statement.
                    strRet.SConcat "PropertyChanged(" & Chr$(34) & strName & Chr$(34) & ")" & vbNewLine
                    strRet.SConcat "End Property" & vbNewLine & vbNewLine
                End If
            Next lonLoop
            
            'Generate UserControl_InitProperties() event.
            'Check if we need to add comments.
            If udtPref.intComment = 1 Then
                strRet.SConcat "'UserControl is intializing the properties." & vbNewLine
                strRet.SConcat "'This is where you set the default values." & vbNewLine
            End If
            
            strRet.SConcat "Private Sub UserControl_InitProperties()" & vbNewLine
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObject, bolReadOnly
                
                'Make sure property isn't read-only.
                If bolReadOnly = False Then
                    
                    'Check if we need to indent the code.
                    If udtPref.intIndentSub = 1 Then
                        intWidth = GetIndentWidth
                        
                        If intWidth = -1 Then
                            strRet.SConcat vbTab
                        Else
                            strRet.SConcat String$(intWidth, Chr$(32))
                        End If
                    
                    End If
                    
                    'Check if the property is an object.
                    If bolObject = True Then
                        strRet.SConcat "Set "
                    Else
                        strRet.SConcat "Let "
                    End If
                    
                    strRet.SConcat strName & " = " & strDef & vbNewLine
                
                End If
            
            Next lonLoop
            
            strRet.SConcat "End Sub" & vbNewLine & vbNewLine
            
            'Generate UserControl_ReadProperties() event.
            'Check if we need to add comments.
            If udtPref.intComment = 1 Then
                strRet.SConcat "'UserControl_ReadProperties() event." & vbNewLine
                strRet.SConcat "'The UserControl is reading the properties from the property bag." & vbNewLine
            End If
            
            strRet.SConcat "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine
            
            'Check if we need to indent the code.
            If udtPref.intIndentSub = 1 Then
                intWidth = GetIndentWidth
                
                If intWidth = -1 Then
                    strRet.SConcat vbTab
                Else
                    strRet.SConcat String$(intWidth, Chr$(32))
                End If
            
            End If
            
            strRet.SConcat "With PropBag" & vbNewLine
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObject, bolReadOnly
                
                'Make sure the property is not read-only.
                If bolReadOnly = False Then
                    'Check if we need to indent the code.
                    If udtPref.intIndentWithEnd = 1 Then
                        intWidth = GetIndentWidth
                        
                        If intWidth = -1 Then
                            strRet.SConcat vbTab & vbTab
                        Else
                            strRet.SConcat String$((intWidth * 2), Chr$(32))
                        End If
                    
                    End If
                    
                    'Check if the property is an object.
                    If bolObject = True Then
                        strRet.SConcat "Set "
                    Else
                        strRet.SConcat "Let "
                    End If
                    
                    strRet.SConcat strName & " = .ReadProperty(" & Chr$(34) & strName & Chr$(34) & ", " & strDef & ")" & vbNewLine
                End If
                
            Next lonLoop
            
            'Check if we need to indent the code.
            If udtPref.intIndentSub = 1 Then
                intWidth = GetIndentWidth
                
                If intWidth = -1 Then
                    strRet.SConcat vbTab
                Else
                    strRet.SConcat String$(intWidth, Chr$(32))
                End If
            
            End If
            
            strRet.SConcat "End With" & vbNewLine
            strRet.SConcat "End Sub" & vbNewLine & vbNewLine
            
            'Generate UserControl_WriteProperties() event.
            
            'Check if we need comments.
            If udtPref.intComment = 1 Then
                strRet.SConcat "'UserControl_WriteProperties() event." & vbNewLine
                strRet.SConcat "'The UserControl is saving the properties' values." & vbNewLine
            End If
            
            strRet.SConcat "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)" & vbNewLine
            
            'Check if we need to indent the code.
            If udtPref.intIndentSub = 1 Then
                intWidth = GetIndentWidth
                
                If intWidth = -1 Then
                    strRet.SConcat vbTab
                Else
                    strRet.SConcat String$(intWidth, Chr$(32))
                End If
            
            End If
            
            strRet.SConcat "With PropBag" & vbNewLine
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObject, bolReadOnly
                
                'Make sure the property is not read-only.
                If bolReadOnly = False Then
                    'Check if we need to indent the code.
                    If udtPref.intIndentWithEnd = 1 Then
                        intWidth = GetIndentWidth
                        
                        If intWidth = -1 Then
                            strRet.SConcat vbTab & vbTab
                        Else
                            strRet.SConcat String$((intWidth * 2), Chr$(32))
                        End If
                    
                    End If
                    
                    strRet.SConcat ".WriteProperty " & Chr$(34) & strName & Chr$(34)
                    strRet.SConcat ", " & strCont & ", " & strDef & vbNewLine
                End If
            
            Next lonLoop
            
            'Check if we need to indent the code.
            If udtPref.intIndentSub = 1 Then
                intWidth = GetIndentWidth
                
                If intWidth = -1 Then
                    strRet.SConcat vbTab
                Else
                    strRet.SConcat String$(intWidth, Chr$(32))
                End If
            
            End If
            
            strRet.SConcat "End With" & vbNewLine
            strRet.SConcat "End Sub"
            
            frmCode.Show
            frmCode.txtCode.Text = strRet.GetString
            frmCode.stbStatus.SimpleText = Len(frmCode.txtCode.Text) & " byte(s)."
            strRet.ReInit
        End If
    
    End With
End Sub

Private Sub mnuHelpContents_Click()
    Dim lonRet As Long
    
    lonRet = ShellExecute(Me.hWnd, "open", App.Path & "\help.htm", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub NewForm()
    lvList.ListItems.Clear
    txtName.Text = ""
    cmbType.Text = ""
    txtCont.Text = ""
    txtDefault.Text = ""
    chkObject.Value = 0
    chkReadOnly.Value = 0
End Sub

Private Sub mnuNew_Click()
    On Error GoTo ErrorHandler
    
    Dim msgRes As VbMsgBoxResult
    Dim strMsg As String
    
    If bolChanged = True Then
        strMsg = "The properties in this file have been changed since they were last saved."
        strMsg = strMsg & vbNewLine & vbNewLine & "Do you want to save the changes?"
        
        msgRes = MsgBox(strMsg, vbExclamation + vbYesNo, "Save Changes")
        
        If msgRes = vbYes Then
            
            If Len(strCurFile) = 0 Or InStr(1, strCurFile, "***") > 0 Then
                
                With objCD
                    .CancelError = True
                    .DialogTitle = "Save Property List"
                    .Filter = "Property List (*." & PW_EXT & ")|*." & PW_EXT
                    .ShowSave
                    
                    If Len(.FileName) > 0 Then
                        stbStatus.SimpleText = "Saving file..."
                        SaveProperties strCurFile
                        stbStatus.SimpleText = "File saved."
                        strCurFile = "***"
                        NewForm
                    Else
                        NewForm
                        strCurFile = "***"
                    End If
                
                End With
            
            Else
                stbStatus.SimpleText = "Saving file..."
                SaveProperties strCurFile
                strCurFile = "***"
                NewForm
                stbStatus.SimpleText = "File saved."
            End If
        
        Else
            strCurFile = "***"
            NewForm
        End If
    
    Else
        strCurFile = "***"
        NewForm
    End If
    
    Exit Sub
    
ErrorHandler:
    '32755 - User hit cancel.
    If Not Err.Number = 32755 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
End Sub

'Load a property list file into the list.
Private Sub LoadPropertyFile(ByVal FilePath As String)
    Dim lonFF As Long, strBuff() As String
    Dim bytData() As Byte, strData As String
    Dim lonLoop As Long, lonBnd As Long
    Dim strInfo() As String, lonCount As Long
    
    lonFF = FreeFile
    
    ReDim bytData(1 To SafeFileLen(FilePath)) As Byte
    'Load data into a byte array first (faster than loading it into a string).
    Open FilePath For Binary Access Read As #lonFF
        Get #lonFF, 1, bytData()
    Close #lonFF
    
    'Quickly convert the byte array to a string.
    strData = StrConv(bytData(), vbUnicode)
    
    If InStr(1, strData, PW_EOP) > 0 Then
        strBuff() = Split(strData, PW_EOP)
        strData = "" 'Clean up.
        lonBnd = UBoundStr(strBuff())
        
        If lonBnd > 0 Then
            
            With lvList
                .ListItems.Clear
                
                For lonLoop = 0 To lonBnd
                    If Len(strBuff(lonLoop)) > 0 Then
                        strInfo() = Split(strBuff(lonLoop), PW_DELIM)
                        
                        'Name
                        .ListItems.Add , , strInfo(0)
                        lonCount = .ListItems.Count
                        'Type
                        .ListItems(lonCount).SubItems(1) = strInfo(1)
                        'Container
                        .ListItems(lonCount).SubItems(2) = strInfo(2)
                        'Default value
                        .ListItems(lonCount).SubItems(3) = strInfo(3)
                        'Is object
                        If strInfo(4) = "0" Then
                            .ListItems(lonCount).SubItems(4) = "False"
                        ElseIf strInfo(4) = "1" Then
                            .ListItems(lonCount).SubItems(4) = "True"
                        End If
                        'Is read-only
                        If strInfo(5) = "0" Then
                            .ListItems(lonCount).SubItems(5) = "False"
                        ElseIf strInfo(5) = "1" Then
                            .ListItems(lonCount).SubItems(5) = "True"
                        End If
                    
                    End If
                    
                Next lonLoop
            
            End With
        
        End If
    
    Else
        '
    End If
End Sub

Private Sub mnuOpen_Click()
    On Error GoTo ErrorHandler
    
    With objCD
        .CancelError = True
        .DialogTitle = "Open Property List"
        .Filter = "Property List (*." & PW_EXT & ")|*." & PW_EXT
        .ShowOpen
        
        If Len(.FileName) > 0 Then
            
            If Not GetFileExtension(.FileName) = LCase$(PW_EXT) Then
                MsgBox "You did not select a valid Property List file", vbCritical, "Invalid File Format"
                Exit Sub
            End If
            
            If FileExists(.FileName) = False Then
                MsgBox "The selected file could not be found", vbCritical, "File Not Found"
                Exit Sub
            Else
                
                If SafeFileLen(.FileName) = 0 Then
                    MsgBox "The property list file you selected is empty", vbCritical, "Empty File"
                    Exit Sub
                End If
            
            End If
            
            stbStatus.SimpleText = "Loading property list..."
            LoadPropertyFile .FileName
            stbStatus.SimpleText = "Property list loaded."
            strCurFile = .FileName
        End If
    
    
    
    End With
    
    Exit Sub
    
ErrorHandler:
    
    '32755 - User hit cancel.
    If Not Err.Number = 32755 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
        
End Sub

Private Sub mnuPreferences_Click()
    frmPreferences.Show vbModal 'Show the preferences form.
End Sub

'Save the list of properties to a file.
Private Sub SaveProperties(ByVal FilePath As String)
    Dim lonLoop As Long, lonFF As Long
    Dim strAll As clsConcat, lonCount As Long
    Dim strName As String, strType As String
    Dim strCont As String, strDef As String
    Dim bolObj As Boolean, bolReadOnly As Boolean
    Dim intObj As Integer, intReadOnly As Integer
    
    Set strAll = New clsConcat
    strAll.ReInit
    
    With lvList
        lonCount = .ListItems.Count
        
        If lonCount > 0 Then
            
            For lonLoop = 1 To lonCount
                PropertyInfoByIndex lonLoop, strName, strType, strCont, strDef, bolObj, bolReadOnly
                intObj = Abs(bolObj)
                intReadOnly = Abs(bolReadOnly)
                
                'Avoid using &'s.
                strAll.SConcat strName
                strAll.SConcat PW_DELIM
                strAll.SConcat strType
                strAll.SConcat PW_DELIM
                strAll.SConcat strCont
                strAll.SConcat PW_DELIM
                strAll.SConcat strDef
                strAll.SConcat PW_DELIM
                strAll.SConcat CStr(intObj)
                strAll.SConcat PW_DELIM
                strAll.SConcat CStr(intReadOnly)
                strAll.SConcat PW_EOP
            Next lonLoop
            
            'Get an available file handle to use.
            lonFF = FreeFile
            
            Open FilePath For Binary Access Write As #lonFF
                Put #lonFF, 1, strAll.GetString
            Close #lonFF
            
            DoEvents
            
            strAll.ReInit
        End If
    
    End With
        
End Sub

'Check if a property name exists in the list.
Private Function PropertyAdded(ByVal PropertyName As String) As Boolean
    Dim lonCount As Long, lonLoop As Long
    Dim strQuery As String
    
    strQuery = LCase$(PropertyName)
    
    With lvList
        lonCount = .ListItems.Count
        
        If lonCount > 0 Then
            
            For lonLoop = 1 To lonCount
                
                If LCase$(.ListItems(lonLoop).Text) = strQuery Then
                    PropertyAdded = True
                    Exit For
                End If
            
            Next lonLoop
        
        End If
    
    End With
End Function

Private Sub mnuSave_Click()
    On Error GoTo ErrorHandler
    
    If Len(strCurFile) > 0 Then
        
        If InStr(1, strCurFile, "***") > 0 Then
            
            With objCD
                .CancelError = True
                .DialogTitle = "Save Properties"
                .Filter = "Property List (*." & PW_EXT & ")|*." & PW_EXT
                .ShowSave
                
                If Len(.FileName) > 0 Then
                    stbStatus.SimpleText = "Saving file..."
                    SafeKill .FileName
                    DoEvents
                    SaveProperties .FileName
                    stbStatus.SimpleText = "File saved."
                    strCurFile = .FileName
                    Exit Sub
                End If
            
            End With
        
        Else
            stbStatus.SimpleText = "Saving file..."
            SafeKill strCurFile
            DoEvents
            SaveProperties strCurFile
            stbStatus.SimpleText = "File saved."
        End If
        
    End If
    
    Exit Sub
    
ErrorHandler:
    '32755 - User hit cancel.
    If Not Err.Number = 32755 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error GoTo ErrorHandler
    
    With objCD
        .CancelError = True
        .DialogTitle = "Save Properties"
        .Filter = "Property List (*." & PW_EXT & ")|*." & PW_EXT
        .ShowSave
        
        If Len(.FileName) > 0 Then
            stbStatus.SimpleText = "Saving file..."
            SafeKill .FileName
            DoEvents
            SaveProperties .FileName
            stbStatus.SimpleText = "File saved."
            strCurFile = .FileName
        End If
    
    End With
    
    Exit Sub
    
ErrorHandler:
    '32755 - User hit cancel.
    If Not Err.Number = 32755 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If
End Sub
