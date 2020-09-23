Attribute VB_Name = "ModAssoc"
Option Explicit
'Registry API
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0&
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0


'Just used in this demo - not needed in real-life usage
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long



'**********************Standard Registry Functions*************************************
Public Sub CreateKey(hKey As Long, strPath As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Private Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
    Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
End Sub
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim RKey As Long
    Dim Result As Long
    Result = RegOpenKey(hKey, strPath, RKey)
    Result = RegDeleteValue(RKey, strValue)
    Result = RegCloseKey(RKey)
End Sub
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Or REG_EXPAND_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
            If lValueType = REG_EXPAND_SZ Then GetSettingString = StripTerminator(ExpandEnvStr(GetSettingString))
        End If
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function





'**************************String Functions*****************************************
Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    s = ""
    c = ExpandEnvironmentStrings(sData, s, c)
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
    'removes quotes from each end of a string
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = Chr(34) Then
        If Right$(strQuotedString, 1) = Chr(34) Then
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function
Public Function GetDosPath(mPath As String) As String
    'returns the short path name
    Dim s As String
    Dim i As Long
    Dim PathLength As Long
    i = Len(mPath) + 1
    s = String(i, 0)
    PathLength = GetShortPathName(mPath, s, i)
    GetDosPath = Left$(s, PathLength)
End Function
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    'returns the long path name
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function
Public Function APPstr() As String
    'get the correct path with a backslash - just in case the user installed
    'our app to the root directory which would already place a back slash
    APPstr = App.Path
    If Right(APPstr, 1) <> "\" Then APPstr = APPstr + "\"
End Function



'**********************************File Association Routines****************************
Public Sub InitFileTypes(mFileTypeName As String, mFileTypeDescript As String, mFileTypeVerb As String, Optional mDefIcon As String)
    'This sets up a FileType in registry which we can link an Extension to later
    'Example call...
    'InitFileTypes "MrBobo.textfile", "Plain text file", "Open"
    SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName, "", mFileTypeDescript
    SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName & "\shell", "", ""
    SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName & "\shell\" & mFileTypeVerb, "", ""
    SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName & "\shell\" & mFileTypeVerb & "\command", "", GetDosPath(APPstr + App.EXEName + ".exe") + " %1"
    If Len(mDefIcon) > 0 Then
        SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName & "\DefaultIcon", "", mDefIcon
    Else
        SaveSettingString HKEY_CLASSES_ROOT, mFileTypeName & "\DefaultIcon", "", GetDosPath(APPstr + App.EXEName + ".exe")
    End If

End Sub
Public Sub AssFile(mExt As String, mFileTypeName As String)
    Dim CurAss As String
    If IsAssFile(mExt, mFileTypeName) = 1 Then Exit Sub 'we're already associated - bail out here
    'who's currently associated?
    CurAss = GetSettingString(HKEY_CLASSES_ROOT, mExt, "")
    'remember current association so we can revert to it later if asked to do so
    'We can store this info under the same key.
    'This is standard practice by "professional" applications
    SaveSettingString HKEY_CLASSES_ROOT, mExt, "OldAss", CurAss
    'OK, now put our FileType in as the new associated type
    SaveSettingString HKEY_CLASSES_ROOT, mExt, "", mFileTypeName
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Sub DisAssFile(mExt As String, mFileTypeName As String)
    Dim CurAss As String
    If IsAssFile(mExt, mFileTypeName) = 0 Then Exit Sub 'we're not associated - bail out here
    'Who, if anyone was associated to this extension before us?
    'We saved this info when we first associated ourselves
    CurAss = GetSettingString(HKEY_CLASSES_ROOT, mExt, "OldAss")
    'return the old association
    SaveSettingString HKEY_CLASSES_ROOT, mExt, "", CurAss
    'remove the saved data - first make a dummy entry to avoid errors...
    SaveSettingString HKEY_CLASSES_ROOT, mExt, "OldAss", "Dummy"
    'then total the entry
    DeleteValue HKEY_CLASSES_ROOT, mExt, "OldAss"
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

Public Function IsAssFile(mExt As String, mFileTypeName As String) As Long
    Dim CurAss As String
    'Are we currently associated to this extension?
    CurAss = GetSettingString(HKEY_CLASSES_ROOT, mExt, "")
    If CurAss = mFileTypeName Then IsAssFile = 1
End Function


'Just used in this demo - not needed in real-life usage

Public Sub GetCurrentAssociation()
    Dim CurAss As String
    Dim tmpKeyVal As String
    Dim hIcon As Long
    Dim fType As String
    Dim oWith As String
    Dim fDescript As String
    Form1.PicIcon.Picture = LoadPicture()
    CurAss = GetSettingString(HKEY_CLASSES_ROOT, ".txt", "")
    fType = GetSettingString(HKEY_CLASSES_ROOT, ".txt", "Content Type")
    If FileExists(CurAss) Then
        hIcon = ExtractAssociatedIcon(App.hInstance, CurAss, 2)
        fDescript = "Unspecified"
        oWith = CurAss
    Else
        tmpKeyVal = GetSettingString(HKEY_CLASSES_ROOT, CurAss & "\DefaultIcon", "")
        fDescript = GetSettingString(HKEY_CLASSES_ROOT, CurAss, "")
        If Len(fDescript) = 0 Then fDescript = "Unspecified"
        If Len(tmpKeyVal) <> 0 Then
            hIcon = ExtractAssociatedIcon(App.hInstance, tmpKeyVal, 2)
        Else
            tmpKeyVal = GetSettingString(HKEY_CLASSES_ROOT, CurAss & "\shell\open\command", "")
            If Len(tmpKeyVal) <> 0 Then
                 hIcon = ExtractAssociatedIcon(App.hInstance, tmpKeyVal, 2)
            Else
                Exit Sub
            End If
        End If
        CurAss = GetSettingString(HKEY_CLASSES_ROOT, CurAss & "\shell\open\command", "")
        oWith = CurAss
    End If
    If hIcon <> 0 Then
        DrawIconEx Form1.PicIcon.hdc, 0, 0, hIcon, 32, 32, 0, 0, DI_NORMAL
        DestroyIcon hIcon
    End If
    If Right(oWith, 2) = "%1" Then oWith = Trim(Left(oWith, Len(oWith) - 2))
    oWith = FileOnly(GetLongFilename(oWith))
    Form1.lblDescription.Caption = fDescript
    Form1.lblFileType.Caption = fType
    Form1.lblOpensWith.Caption = oWith
End Sub
