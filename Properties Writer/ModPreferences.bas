Attribute VB_Name = "ModPreferences"
Option Explicit

Public Type PW_PREF
    intSaveWinPos As Integer
    intSaveListCol As Integer
    intSaveClear As Integer
    intComment As Integer
    intIndentSub As Integer
    intIndentWithEnd As Integer
    intIndentTab As Integer
    intIndentWidth As Integer
End Type

Public udtPref As PW_PREF

Public Const PREF_SEC As String = "Preferences"

Public Sub SavePref()
    With udtPref
        SaveSetting APPNAME, PREF_SEC, "Comment", .intComment
        SaveSetting APPNAME, PREF_SEC, "IndentSub", .intIndentSub
        SaveSetting APPNAME, PREF_SEC, "IndentTab", .intIndentTab
        SaveSetting APPNAME, PREF_SEC, "IndentWidth", .intIndentWidth
        SaveSetting APPNAME, PREF_SEC, "IndentWithEnd", .intIndentWithEnd
        SaveSetting APPNAME, PREF_SEC, "SaveClear", .intSaveClear
        SaveSetting APPNAME, PREF_SEC, "SaveListCol", .intSaveListCol
        SaveSetting APPNAME, PREF_SEC, "SaveWinPos", .intSaveWinPos
    End With
End Sub

Public Sub ReadPref(Optional ByVal Apply As Boolean = False)
    With udtPref
        .intComment = GetSetting(APPNAME, PREF_SEC, "Comment", 0)
        .intIndentSub = GetSetting(APPNAME, PREF_SEC, "IndentSub", 1)
        .intIndentTab = GetSetting(APPNAME, PREF_SEC, "IndentTab", 1)
        .intIndentWidth = GetSetting(APPNAME, PREF_SEC, "IndentWidth", 4)
        .intIndentWithEnd = GetSetting(APPNAME, PREF_SEC, "IndentWithEnd", 1)
        .intSaveClear = GetSetting(APPNAME, PREF_SEC, "SaveClear", 1)
        .intSaveListCol = GetSetting(APPNAME, PREF_SEC, "SaveListCol", 1)
        .intSaveWinPos = GetSetting(APPNAME, PREF_SEC, "SaveWinPos", 1)
    End With
    
    If Apply = True Then
        
        With frmPreferences
            .chkComment.Value = udtPref.intComment
            .chkIndentSub.Value = udtPref.intIndentSub
            .chkIndentTab.Value = udtPref.intIndentTab
            .txtIndentWidth.Text = udtPref.intIndentWidth
            .chkIndentWithEnd.Value = udtPref.intIndentWithEnd
            .chkSaveClear.Value = udtPref.intSaveClear
            .chkSaveListCol.Value = udtPref.intSaveListCol
            .chkSaveWinPos.Value = udtPref.intSaveWinPos
        End With
    
    End If
End Sub

Public Sub LoadPrefDefaults()
    With udtPref
        .intComment = 0
        .intIndentSub = 1
        .intIndentTab = 1
        .intIndentWidth = 4
        .intIndentWithEnd = 1
        .intSaveClear = 1
        .intSaveListCol = 1
        .intSaveWinPos = 1
    End With
End Sub
