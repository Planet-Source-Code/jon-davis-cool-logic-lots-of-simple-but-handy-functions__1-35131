VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoolLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ------------------------------------------------------------------
' COOL LOGIC v0.92
'   a compilation of reusable VB5 logic
'   compiled by Jon Davis, jon@jondavis.net
'
'   NOTE: some of this code was borrowed from contributions
'   found elsewhere, and such methods have been noted in comments
'
' Instructions: Just drop this code into your Visual Basic 5/6 app,
'   or you can compile this code as a class module to an ActiveX DLL
' -------------------------------------------------------------------

Option Explicit

' ----------------------------------------
'            DECLARATIONS
' ----------------------------------------

' For ShortPathName
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' For LongPathName
Private Declare Function OSGetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' For WinSysDir
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' For WinDir
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' For GetTempDir
Private Declare Function GetTempPath Lib "kernel32" _
Alias "GetTempPathA" (ByVal nBufferLength As Long, _
ByVal lpBuffer As String) As Long
Const MAX_PATH = 255


' For PlaySound
Private Declare Function sndPlaySound Lib "winmm" Alias _
    "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
    ' flag values for uFlags paramaeter

' For dialogue boxes
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

' For OpenWebbrowser
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Cache variable for default editor, i.e. Notepad.exe
Private vDefEditor As String


' --

' For MonthString
Private clnMonths As New Collection

' --

' For WinSysDir
Const gintMAX_SIZE% = 255                        'Maximum buffer size

'For PlaySound
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

'For dialogue boxes
Const OFN_READONLY = &H1
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_CREATEPROMPT = &H2000
Const OFN_SHAREAWARE = &H4000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOTESTFILECREATE = &H10000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Const OFN_EXPLORER = &H80000                         '  new look commdlg
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
'
Const OFN_SHAREFALLTHROUGH = 2
Const OFN_SHARENOWARN = 1
Const OFN_SHAREWARN = 0

' More for dialogue boxes
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'For DateString
Enum DateType
    MMDDYYYY
    MMDDYY
    DDMMYYYY
    DDMMYY
    YYYYMMDD
    YYMMDD
    YYYYDDMM
    YYDDMM
End Enum
' --------------------------------------->



' ----------------------------------------
'               FUNCTIONS
' ----------------------------------------



' <---------------------------------------
'     CommonDialogue boxes without ocx
' ----------------------------------------
' from http://www.kather.net/VisualBasicSource/vb5commondialognoocx.txt
' "thanks to  Rick Ratayczak of Future Works Media (rickr@execpc.com)"
' Note: Additional declarations at the top.
Function SaveDialog(FormName, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = FormName.HWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        A = GetSaveFileName(ofn)

        If (A) Then
            SaveDialog = Trim$(ofn.lpstrFile)
        Else
            SaveDialog = ""
        End If

End Function

Function OpenDialog(FormName, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = FormName.HWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        A = GetOpenFileName(ofn)

        If (A) Then
            OpenDialog = Trim$(ofn.lpstrFile)
        Else
            OpenDialog = ""
        End If

End Function
' --------------------------------------->


' <---------------------------------------
'             OpenWebBrowser
' ----------------------------------------
' Opens a given URL in the default web browser
' Note: Additional declarations at the top.
Sub OpenWebBrowser(URL As String)
    ShellExecute 0, "Open", URL, "", App.Path, 1
End Sub
' --------------------------------------->


' <---------------------------------------
'              Find & Replace
' ----------------------------------------
' Takes a source string and returns the string with the
' search string replaced.  Faster than Replace().
' Example: Text1.text = FindReplace(Text1.text)
Function ReplaceAll(SourceString, SearchString, ReplaceString) As String
    Dim tmpString1 As String, tmpString2 As String, vFixed As Boolean
    tmpString1 = SourceString
    Do Until vFixed
        tmpString2 = tmpString1
        tmpString1 = ReplaceFirst(tmpString1, SearchString, ReplaceString)
        If tmpString1 = tmpString2 Then vFixed = True
    Loop
    ReplaceAll = tmpString1
End Function

Function ReplaceFirst(SourceString, SearchString, ReplaceString) As String
    Dim FoundLoc As Integer
    FoundLoc = InStr(1, SourceString, SearchString)
    If FoundLoc <> 0 Then
        ReplaceFirst = Left(SourceString, FoundLoc - 1) & _
        ReplaceString & Right(SourceString, _
        Len(SourceString) - (FoundLoc - 1) - Len(SearchString))
    Else
        ReplaceFirst = SourceString
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'               IsLeapYear
' ----------------------------------------
' Checks to see if a year is a leap-year
' from http://www.kather.net/VisualBasicSource/vbLeapYear.txt
Function IsLeapYear(strYear As String) As Boolean
    Dim strDatum As String
    If strYear = "" Then Exit Function
    strDatum = ("29-2-" & strYear)
    If IsDate(strDatum) Then IsLeapYear = True _
        Else IsLeapYear = False
End Function
' --------------------------------------->


' <---------------------------------------
'              isEven/isOdd
' ----------------------------------------
' Returns true if even or odd
Function IsEven(Value As Variant)
    If (Value / 2) <> Int(Value / 2) Then _
        IsEven = False Else IsEven = True
End Function

Function IsOdd(Value As Variant)
    If (Value / 2) <> Int(Value / 2) Then _
        IsOdd = True Else IsOdd = False
End Function

' --------------------------------------->


' <---------------------------------------
'               CenterForm
' ----------------------------------------
' centers a form on the screen
Sub CenterForm(vForm)
    vForm.Move (Screen.Width / 2) - (vForm.Width / 2), (Screen.Height / 2) - (vForm.Height / 2)
End Sub
' --------------------------------------->


' <---------------------------------------
'              ConvertTime
' ----------------------------------------
' Takes nearly any string and converts it into a time.
' From http://www.kather.net/VisualBasicSource/vbconvertingtexttotime.txt
'
' Giving the input in Value => result
' 1      => 01:00
' 12     => 12:00
' 1.1    => 01:10
' 915    => 9:15
' 9.15   => 9:15
' 1015   => 10:15
' 10.15  => 10:15
' 12:15  => 12:15
' giving an impossible time => result
' 26 => 00:00
' 1976   => 00:00
' giving just nothing (TAB/ENTER) will give no result
Public Function CTime(Value As String) As Date
    Dim t%
    'checking on digits and seperator
    Const Getal$ = "1234567890.:"
    For t% = 1 To Len(Value)
        If InStr(Getal$, Mid$(Value, t%, 1)) = 0 Then Exit Function
        If Mid$(Value, t%, 1) = "." Then Value = Left$(Value, t% - 1) & ":" & Right$(Value, Len(Value) - t%)
    Next t%

    Select Case Len(Value)
    Case 0
        Exit Function
    Case 1
        Value = "0" & Value & ":00"
    Case 2
        Value = Value & ":00"
    Case 3
        t% = InStr(Value, ":")
        If t% = 0 Then _
            Value = Left$(Value, 1) & ":" & Right$(Value, 2)
    Case 4
        t% = InStr(Value, ":")
        If t% = 0 Then _
            Value = Left$(Value, 2) & ":" & Right$(Value, 2)
    Case 5
        Value = Left$(Value, 2) & ":" & Right$(Value, 2)
    End Select
    On Error Resume Next
    CTime = TimeValue(Value)
End Function
' --------------------------------------->

' <---------------------------------------
'           ASCII-encode/decode
' ----------------------------------------
' Here 's a little doo-dad that moves the ascii value of each character up (encode) or down
' (decode), with a default change of 40. So, for example, chr(233) might become chr(21), or
' chr(74) might become chr(114).
' Anyone who knows what they're doing could easily decode it, but at least it's not legible
' if you open the file in Notepad, etc. Try it!
'
' Syntax is:  encode(mytext, [depth]) and decode(mytext, [depth]).
Public Function ASCIIEncode(ByVal Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Integer
Dim NewData As String
Dim vChar As Integer

For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40 'DEFAULT DEPTH
        If Depth > 254 Then Depth = 254

        TempAsc = TempAsc + Depth
        If TempAsc > 255 Then TempAsc = TempAsc - 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
ASCIIEncode = NewData

End Function
    

Public Function ASCIIDecode(ByVal Data As String, Optional Depth As Integer) As String
Dim TempChar As String
Dim TempAsc As Integer
Dim NewData As String
Dim vChar As Integer

For vChar = 1 To Len(Data)
    TempChar = Mid$(Data, vChar, 1)
        TempAsc = Asc(TempChar)
        If Depth = 0 Then Depth = 40 'DEFAULT DEPTH
        If Depth > 254 Then Depth = 254
    TempAsc = TempAsc - Depth
        If TempAsc < 0 Then TempAsc = TempAsc + 255
        TempChar = Chr(TempAsc)
        NewData = NewData & TempChar
Next vChar
ASCIIDecode = NewData

End Function
' --------------------------------------->


' <<--------------------------------------
'               DateString
' ----------------------------------------
' Converts a date to a specified format
' Example: February 14, 1999 becomes "02/14/99"
' Example: February 14, 1999 becomes "14021999"
Public Function DateString(Optional ByVal vType As DateType = 0, Optional Value As Variant, Optional Delimiter As String = "/") As String
    If IsDate(Value) Then
        Value = CDate(Value)
    Else
        DateString = Value
        Exit Function
    End If
    
    'MMDDYYYY = 0
    'MMDDYY   = 1
    'DDMMYYYY = 2
    'DDMMYY   = 3
    'YYYYMMDD = 4
    'YYMMDD   = 5
    'YYYYDDMM = 6
    'YYDDMM   = 7
    
    Select Case vType
    Case 0
        DateString = Month(Value)
        If Month(Value) < 10 Then
            DateString = "0" & DateString
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        DateString = DateString & Year(Value)
    Case 1
        DateString = Month(Value)
        If Month(Value) < 10 Then
            DateString = "0" & DateString
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        DateString = DateString & Right(Year(Value), 2)
    Case 2
        If Day(Value) < 10 Then
            DateString = "0" & Day(Value)
        Else
            DateString = Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        DateString = DateString & Year(Value)
    Case 3
        If Day(Value) < 10 Then
            DateString = "0" & Day(Value)
        Else
            DateString = Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        DateString = DateString & Right(Year(Value), 2)
    Case 4
        DateString = Year(Value)
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
    Case 5
        DateString = Right(Year(Value), 2)
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
    Case 6
        DateString = Year(Value)
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
    Case 7
        DateString = Right(Year(Value), 2)
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Day(Value) < 10 Then
            DateString = DateString & "0" & Day(Value)
        Else
            DateString = DateString & Day(Value)
        End If
        If Not IsMissing(Delimiter) And Delimiter <> "" Then DateString = DateString & Delimiter
        If Month(Value) < 10 Then
            DateString = DateString & "0" & Month(Value)
        Else
            DateString = DateString & Month(Value)
        End If
    End Select
End Function
' --------------------------------------->


' <---------------------------------------
'                PlaySound
' ----------------------------------------
' Yep.. plays a sound.  Only WAV files--other media
' types require system-specific support.
' Note: Additional declarations at the top.
Public Function PlaySound(ByVal SoundFile As String, Optional Synchronous As Boolean, Optional vLoop As Boolean, Optional NoStop As Boolean, Optional vMemory As Boolean)
    Dim Snd_Options
    If Not FileExists(SoundFile) Then
        If FileExists(WinDir & "media\" & SoundFile) Then
            SoundFile = WinDir & "media\" & SoundFile
        Else
            If FileExists(SoundFile & ".wav") Then
                SoundFile = SoundFile & ".wav"
            Else
                If FileExists(WinDir & "media\" & SoundFile & ".wav") Then
                    SoundFile = WinDir & "media\" & SoundFile & ".wav"
                End If
            End If
        End If
    End If
    If Synchronous Then Snd_Options = SND_SYNC _
     Else Snd_Options = SND_ASYNC
    If vLoop Then Snd_Options = Snd_Options + SND_LOOP
    If NoStop Then Snd_Options = Snd_Options + SND_NOSTOP
    If vMemory Then Snd_Options = Snd_Options + SND_MEMORY
    PlaySound = sndPlaySound(SoundFile, Snd_Options)
End Function
' --------------------------------------->


' <---------------------------------------
'               ImportFile
' ----------------------------------------
' Imports and returns a file's contents
Public Function ImportFile(ByVal Path As String)
    Dim FileNum
    FileNum = FreeFile
    Open Path For Binary Access Read As FileNum
    ImportFile = Input(LOF(FileNum), FileNum)
    Close FileNum
End Function
' --------------------------------------->


' <---------------------------------------
'            GetFileExtension
' ----------------------------------------
' Checks to see if something already exists on the hard drive
Public Function GetFileExtension(ByVal Filename As String, Optional IncludeDot As Boolean) As String
    Dim tmpStr As String, curLoc As Integer, prevLoc As Integer, dotLoc As Integer
    curLoc = 1: prevLoc = 1
    Do Until curLoc = 0
        curLoc = InStr(prevLoc + 1, Filename, ".")
        If curLoc <> 0 Then prevLoc = curLoc
    Loop
    dotLoc = prevLoc
    curLoc = 1: prevLoc = 1
    Do Until curLoc = 0
        curLoc = InStr(prevLoc + 1, Filename, "\")
        If curLoc <> 0 Then prevLoc = curLoc
    Loop
    If prevLoc < dotLoc Then
        GetFileExtension = Right(Filename, Len(Filename) - dotLoc - IncludeDot)
    Else
        GetFileExtension = ""
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'               FileExists
' ----------------------------------------
' Checks to see if something already exists on the hard drive
Public Function FileExists(ByVal parmPath As String) As Boolean
    FileExists = Not (Dir(parmPath) = "")
End Function
' --------------------------------------->


' <---------------------------------------
'                DirExists
' ----------------------------------------
' Determines whether the specified directory name exists.

Public Function DirExists(ByVal strDirName As String) As Boolean
    On Error Resume Next
    Dim strDummy
    AddDirSep strDirName
    strDummy = Dir$(strDirName & "*.*", vbDirectory)
    DirExists = Not (strDummy = "")
    Err = 0
End Function
' --------------------------------------->


' <---------------------------------------
'               AddDirSep
' ----------------------------------------
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists

Public Function AddDirSep(strPathName As String)
    If Right(Trim(strPathName), 1) <> "/" And _
       Right(Trim(strPathName), 1) <> "\" Then
        strPathName = RTrim$(strPathName) & "\"
    End If
    AddDirSep = strPathName
End Function
' --------------------------------------->


' <---------------------------------------
'              GetFileName
' ----------------------------------------
' Returns the filename of a given file path.
' Like GetPathFile, but executes differently
' and offers bStripExtension, an option to
' remove the file extension from the filename.

Public Function GetFileName(ByVal vPath As String, _
 Optional ByVal bStripExtension As Boolean, _
 Optional ByVal DirSep As String = "\") As String
    Dim dotLoc As Integer, prevLoc As Integer
    Dim slashLoc As Integer, iLen
    prevLoc = 0: dotLoc = 1: slashLoc = 1
    'InStrRev() wasn't supported in VB5, so we have to
    'do it manually.
    Do Until dotLoc = 0
        dotLoc = InStr(prevLoc + 1, vPath, ".")
        If dotLoc > 0 Then prevLoc = dotLoc
    Loop
    dotLoc = prevLoc
    prevLoc = 0
    Do Until slashLoc = 0
        slashLoc = InStr(prevLoc + 1, vPath, DirSep)
        If slashLoc > 0 Then prevLoc = slashLoc
    Loop
    slashLoc = prevLoc
    If dotLoc > slashLoc Then
        iLen = Len(vPath) - slashLoc
    Else
        iLen = Len(vPath)
    End If
    GetFileName = Right(vPath, iLen)
    dotLoc = 1: prevLoc = 0
    Do Until dotLoc = 0
        dotLoc = InStr(prevLoc + 1, GetFileName, ".")
        If dotLoc > 0 Then prevLoc = dotLoc
    Loop
    dotLoc = prevLoc
    If bStripExtension Then _
     GetFileName = Left(GetFileName, dotLoc - 1)
End Function
' --------------------------------------->


' <---------------------------------------
'             ShortPathName
' ----------------------------------------
' Retrieve the short pathname version of a path possibly
'   containing long subdirectory and/or file names

 Public Function ShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String
     Dim lResult As Long

     On Error GoTo 0
     strShortPath = String(cchBuffer, Chr$(0))
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         Error 53 ' File not found
     Else
         ShortPathName = StripTerminator(strShortPath)
     End If
 End Function
' --------------------------------------->


 ' <---------------------------------------
 '             LongPathName
 ' ----------------------------------------
 ' Retrieve the long pathname version of a path possibly
 '   containing short subdirectory and/or file names
 ' ----------------------------------------
 
 Public Function LongPathName(ByVal strShortPath As String) As String
    Const cchBuffer = 300
    Dim strLongPath As String
    Dim lResult As Long
    
    On Error GoTo 0
    
    strLongPath = String(cchBuffer, Chr$(0))
    lResult = OSGetLongPathName(strShortPath, strLongPath, cchBuffer)
    If lResult = 0 Then
        Error 53 ' File not found
    Else
        LongPathName = StripTerminator(strLongPath)
    End If
 End Function
' --------------------------------------->


' <---------------------------------------
'                WinDir
' ----------------------------------------
' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present

Public Function WinDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        WinDir = strBuf
    Else
        WinDir = ""
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'                TempDir
' ----------------------------------------
' Calls the windows API to get the temp directory and
' ensures that a trailing dir separator is present

Public Function GetTempDir() As String
    Dim sRet As String, lngLen As Long
    
    'create buffer
    sRet = String(MAX_PATH, 0)

    lngLen = GetTempPath(MAX_PATH, sRet)
    If lngLen = 0 Then Err.Raise Err.LastDllError
    GetTempDir = Left$(sRet, lngLen)
End Function
' --------------------------------------->


' <---------------------------------------
'               WinSysDir
' ----------------------------------------
' Calls the windows API to get the windows\SYSTEM directory
' and ensures that a trailing dir separator is present
' Returns: The windows\SYSTEM directory

Function WinSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        WinSysDir = strBuf
    Else
        WinSysDir = ""
    End If
End Function

' <---------------------------------------
'             StripTerminator
' ----------------------------------------
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'               SaveString
' ----------------------------------------
' Writes a string to a file
Public Function SaveString(strValue As String, strFilename As String)
    'On Error Resume Next
    
    'Dim FileNum As Integer
    'FileNum = FreeFile
    'Open strFilename For Output As FileNum
    'Print #FileNum, strValue
    'Close #FileNum
    Dim fso As New FileSystemObject, ts As TextStream
    Set ts = fso.OpenTextFile(strFilename, ForAppending, True)
    ts.Write strValue
    
    If Err.Number = 0 Then
        SaveString = 0
    Else
        SaveString = Err.Number
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'               SetFormFont
' ----------------------------------------
' Walks through all controls on specified form and
' sets Font properties.

Public Sub SetFormFont(frm, FontName As String, FontSize As Integer, FontBold As Boolean)
    Dim ctl As Control
    '
    ' this procedure will set the font of every control to
    ' a font listed in the resource file
    '
    ' some controls may fail, so we will do a resume next...
    '
    On Error Resume Next
    '
    ' set the form's fontname (don't set the size and bold until later)
    '
    frm.FontName = FontName
    '
    ' Now set the FontSize and FontBold.  This needs to be set
    ' after the FontName has been set because not all fonts
    ' support all sizes, etc.
    '
    frm.FontSize = FontSize
    frm.FontBold = FontBold
    '
    ' loop through each control and try to set its font property
    ' this may fail, but our error handling is shut off
    '
    For Each ctl In frm.Controls
        ctl.FontName = FontName
        ctl.FontSize = FontSize
        ctl.FontBold = FontBold
    Next
    '
    ' get out, reset error handling
    '
    Set ctl = Nothing
    On Error GoTo 0
    Exit Sub
       
End Sub
' --------------------------------------->


' <<--------------------------------------
'               QuoteString
' ----------------------------------------
' Adds quotation marks around an unquoted string, by default.  If the string is already quoted
' it returns without making any changes unless vForce is set to True (vForce defaults to False) except that white
' space before and after the quotes will be removed unless vTrim is False.  If the string contains leading or
' trailing white space it is trimmed unless vTrim is set to False (vTrim defaults to True).

Public Function QuoteString(strUnQuotedString As String, Optional vForce As Boolean, Optional vTrim As Boolean)
    Dim strQuotedString As String
    Const gstrQUOTE = """"
    
    If IsMissing(vForce) Then
        vForce = False
    End If
    If IsMissing(vTrim) Then
        vTrim = True
    End If
    
    strQuotedString = strUnQuotedString
    '
    ' Trim the string if necessary
    '
    If vTrim Then
        strQuotedString = Trim(strQuotedString)
    End If
    '
    ' See if the string is already quoted
    '
    If Not vForce Then
        If (Left(strQuotedString, 1) = gstrQUOTE) And (Right(strQuotedString, 1) = gstrQUOTE) Then
            '
            ' String is already quoted.  We are done.
            '
            GoTo DoneQuoteString
        End If
    End If
    '
    ' Add the quotes
    '
    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
DoneQuoteString:
    QuoteString = strQuotedString
End Function

' ----------------------------------------
'              UnQuoteString
' ----------------------------------------
' Tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, removes them.

Public Function UnQuoteString(ByVal strQuotedString As String)
    Const gstrQUOTE = """"
    strQuotedString = Trim(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    UnQuoteString = strQuotedString
End Function
' --------------------------------------->

' <---------------------------------------
'              ParseString
' ----------------------------------------
' Takes a string of delimited data and parses it into a collection.
' RetainStrings means that a value surrounded with quotes will not be
' unquoted, whereas the default function will unquote the string
' and attempt to convert numbers to longs and dates to dates.
Public Function ParseString(ByVal tmpString As String, Optional DelimiterCharacter As String, Optional RetainStrings As Boolean) As Collection
    On Error GoTo ErrorHandler
    Dim FieldValue As String, tmpCol As New Collection
    If IsMissing(DelimiterCharacter) Or DelimiterCharacter = "" Then DelimiterCharacter = ","
    Dim NoMore As Boolean, LastDelim As Integer, Delimiter As Integer, _
        quotDelimiter As Integer, commaDelimiter As Integer
    Do Until NoMore
        LastDelim = 0
        Delimiter = InStr(LastDelim + 1, tmpString, DelimiterCharacter)
        If Delimiter > 0 Then
            FieldValue = Left(tmpString, Delimiter - 1)
            If Left(tmpString, 1) = """" Then
ReQuery:
                quotDelimiter = InStr(LastDelim + 2, tmpString, """")
                commaDelimiter = InStr(LastDelim + 1, tmpString, DelimiterCharacter)
                If commaDelimiter < quotDelimiter Then
                    If commaDelimiter > 0 Then
                        'There's a quoted comma
                        LastDelim = commaDelimiter
                        GoTo ReQuery
                    Else
                        'It's ok -- it's the last field on the line...
                        FieldValue = tmpString
                        NoMore = True
                    End If
                Else
                    If commaDelimiter = quotDelimiter + 1 Then
                        'the comma is right after the close-quote, which means it's the end of the field
                        Delimiter = commaDelimiter
                        FieldValue = Left(tmpString, Delimiter - 1)
                        GoTo OKQuery
                    Else
                        'a quoted quote.
                        'quotdelimiter becomes the minimum starting point
                        LastDelim = quotDelimiter
                        GoTo ReQuery
                    End If
                End If
            End If
OKQuery:
            tmpString = Right(tmpString, Len(tmpString) - Delimiter)
        Else
            FieldValue = tmpString
            NoMore = True
        End If
        

        If Not RetainStrings Then
            tmpCol.Add CVal(UnQuoteString(FieldValue))
        Else
            tmpCol.Add FieldValue
        End If
    Loop
    Set ParseString = tmpCol
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
' --------------------------------------->


' <---------------------------------------
'                  CVal
' ----------------------------------------
' Takes a string and converts it to a variant
' and converts numbers to longs and date strings
' to dates

Public Function CVal(FieldValue As String) As Variant
    On Error GoTo ErrorHandler
    If IsDate(FieldValue) Then
        FieldValue = CDate(FieldValue)
    Else
        If IsNumeric(FieldValue) Then
            FieldValue = CLng(FieldValue)
        End If
    End If
    CVal = FieldValue
    Exit Function
ErrorHandler:
    On Error Resume Next
    CVal = FieldValue
End Function
' --------------------------------------->


' <---------------------------------------
'               ImportData
' ----------------------------------------
' Imports a comma-delimited file into a
' collection (rows) of collections (columns)
Public Function ImportData(ByVal Path As String, Optional DelimiterCharacter As String, Optional RetainStrings As Boolean) As Collection
    Dim FieldValues As New Collection
    Set ImportData = New Collection
    Dim CSVFileNum As Integer, tmpString As String
    CSVFileNum = FreeFile
    Open Path For Input As #CSVFileNum
    Do Until EOF(CSVFileNum)
        Line Input #CSVFileNum, tmpString
        Set FieldValues = ParseString(tmpString, DelimiterCharacter, RetainStrings)
        ImportData.Add FieldValues
    Loop
    Close #CSVFileNum
End Function
' --------------------------------------->


' <---------------------------------------
'               GetFileLineCount
' ----------------------------------------
' Returns the number of lines in a text file
Public Function GetFileLineCount(ByVal Path As String, Optional IgnoreFirstRecord As Boolean) As Integer
    Dim FN, strDummy
    FN = FreeFile
    Open Path For Input As FN
    Do Until EOF(FN)
        Line Input #FN, strDummy
        GetFileLineCount = GetFileLineCount + 1
    Loop
    Close #FN
End Function
' --------------------------------------->

' <---------------------------------------
'                GetPathDir
' ----------------------------------------
' Returns the directory of a path
Public Function GetPathDir(Path As String, Optional DirDelimeter As String)
    On Error Resume Next
    If IsMissing(DirDelimeter) Or IsNull(DirDelimeter) Or DirDelimeter = "" Then DirDelimeter = "\"
    If DirExists(Path) Then
        GetPathDir = Path
        If Right(GetPathDir, 1) <> DirDelimeter Then GetPathDir = GetPathDir & DirDelimeter
        Exit Function
    End If
    If Right(Path, 1) = DirDelimeter Then
        GetPathDir = Path
    Else
        Do
            Dim delim, vFN
            delim = InStr(delim + 1, Path, DirDelimeter)
            If delim <> 0 Then
                vFN = delim
            Else
                Exit Do
            End If
        Loop Until delim = 0
        GetPathDir = Left(Path, vFN)
    End If
End Function

' <---------------------------------------
'                GetPathFile
' ----------------------------------------
' Returns the filename of a path
Public Function GetPathFile(Path As String, Optional DirDelimeter)
    On Error Resume Next
    If IsMissing(DirDelimeter) Or IsNull(DirDelimeter) Or DirDelimeter = "" Then DirDelimeter = "\"
    If DirExists(Path) Then
        GetPathFile = ""
        Exit Function
    End If
    Dim vDelim
    If Right(Path, 1) = DirDelimeter Then
        GetPathFile = ""
    Else
        Do
            Dim delim, vFN
            delim = InStr(delim + 1, Path, DirDelimeter)
            If delim <> 0 Then
                vFN = delim
            Else
                Exit Do
            End If
        Loop Until delim = 0
        GetPathFile = Right(Path, Len(Path) - vFN)
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'            ExportDataCollection
' ----------------------------------------
' Writes a file from a collection (rows) of collections (columns)
' Uses comma by default.
' Encapsulates strings in quotes by default.
' Does not overwrite, by default.
Public Function ExportDataCollection(ByVal Path As String, Data As Collection, Optional DelimiterCharacter As _
String, Optional QuoteEncapsulate As Boolean, Optional Overwrite As Boolean) As Integer
    On Error GoTo ErrorHandler
Dim expFileNum As Integer, tmpString As String
Dim tmpRow As String, tmpCol As String, Row As Variant, _
 c As Integer, Column As Variant
Dim tmpData
    If FileExists(Path) Then
        If Overwrite Then
            Kill (Path)
        End If
    Else
        ExportDataCollection = -1
        Exit Function
    End If
    
    expFileNum = FreeFile
    Open Path For Output As #expFileNum
        For Each Row In Data
            tmpRow = ""
            c = 0
            For Each Column In Row
                c = c + 1 'Index of Column
                If IsMissing(QuoteEncapsulate) Or QuoteEncapsulate = True Then _
                    Column = UnQuoteString(Column)
                    Column = ReplaceAll(Column, """", "``")
                    Column = QuoteString(CStr(Column))
                    tmpCol = CStr(Column)
                If c > 1 Then
                    tmpRow = tmpRow & "," & tmpCol
                Else
                    tmpRow = tmpCol
                End If
            Next
            'If Right(tmpRow, 1) = "," Then tmpRow = Left(tmpRow, Len(tmpRow) - 1)
            tmpRow = tmpRow & vbCrLf
            tmpData = tmpData & tmpRow
        Next
        'MsgBox tmpdata
        Print #expFileNum, tmpData
'    End If
    Close #expFileNum
ErrorHandler:
    ExportDataCollection = Err.Number
End Function
' --------------------------------------->


' <---------------------------------------
'              RandomCharacter
' ----------------------------------------
' Generates a random character; "AllowNumber" makes it alphanumeric
Public Function RandomCharacter(Optional AllowNumber As Boolean = False) As String
    Dim newchar, thing
    
    'number or letter?
    Randomize
    thing = Int(2 * Rnd)
    If thing Or Not AllowNumber Then 'it's a letter
        'lowercase or uppercase?
        Randomize
        thing = Int(2 * Rnd)
        If thing Then 'it's upper-case
            Randomize
            newchar = Chr(Int(26 * Rnd) + 65)
        Else 'it's lower-case
            Randomize
            newchar = Chr(Int(26 * Rnd) + 97)
        End If
    Else 'it's a number
        Randomize
        newchar = CStr(Int(10 * Rnd))
    End If
    RandomCharacter = newchar
End Function
' --------------------------------------->


' <---------------------------------------
'               RandomString
' ----------------------------------------
' Generates a random string
Public Function RandomString(length As Integer, Optional Nums As Boolean) As String
    Dim charpos As Integer
    For charpos = 1 To length
        RandomString = RandomString & RandomCharacter(Nums)
    Next
End Function
' --------------------------------------->

' <---------------------------------------
'               IsEmptyDir
' ----------------------------------------
' Checks to see if there are any files in the directory
Public Function IsEmptyDir(Path As String) As Boolean
    IsEmptyDir = Not FileExists(AddDirSep(Path) & "*.*")
End Function
' --------------------------------------->


' <---------------------------------------
'                 LnTrim
' ----------------------------------------
' Like trim but also trims linebreaks
Public Function LnTrim(str As String)
    Dim dirty As Boolean
reexec:
    str = Trim(str)
    If Left(str, 1) = vbCr Or Left(str, 1) = vbLf _
     Or Left(str, 1) = vbCrLf Then
        str = Right(str, Len(str) - 1)
        dirty = True
    Else
        dirty = False
    End If
    If Right(str, 1) = vbCr Or Right(str, 1) = vbLf _
     Or Right(str, 1) = vbCrLf Then
        str = Left(str, Len(str) - 1)
        dirty = True
    End If
    If dirty Then GoTo reexec
    LnTrim = str
End Function
' --------------------------------------->


' <---------------------------------------
'               EscapeQuotes
' ----------------------------------------
' Sterilizes "deadly" single-quotes in a string
' (for SQL statements)

' Abbreviated form: EscQ(str)
Function EscQ(str As String)
    EscQ = EscapeQuotes(str)
End Function

Function EscapeQuotes(str As String)
    Dim r
    Dim i
    Dim c
    str = Trim(str)
    r = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c = "'" Then
            r = r + "''"
        Else
            r = r + c
        End If
    Next
    EscapeQuotes = r
End Function

' --------------------------------------->


' <---------------------------------------
'               Bool
' ----------------------------------------
' Returns true or false from a variant based
' on a positive string, i.e. "YES", "TRUE"

Function Bool(Value As Variant) As Boolean
    On Local Error Resume Next
    If IsNull(Value) Then
        Bool = False
        Exit Function
    End If
    Select Case UCase(CStr(Value))
        Case "YES", "Y", "TRUE"
            Bool = True
        Case Else
            Bool = False
    End Select
    If IsNumeric(Value) Then Bool = Int(Value) 'If Int(Value) > 0 Or Int(Value) = vbTrue Then Bool = True
End Function
' --------------------------------------->


' <---------------------------------------
'               ReadBinFile
' ----------------------------------------
' Returns a byte array for the binary data of a binary file

Function ReadBinFile(ByVal strFilename As String)
       Dim fl As Long
       Dim binByte() As Byte
       Dim FileID As Integer
       
       FileID = FreeFile
       Open strFilename For Binary Access Read As #FileID

       fl = FileLen(strFilename)
       ReDim binByte(fl)

       Get #FileID, , binByte

       Close #FileID

       ReadBinFile = binByte

End Function
' --------------------------------------->


' <---------------------------------------
'               WriteBinFile
' ----------------------------------------
' Writes a byte array of binary data to a binary file

Function WriteBinFile(ByVal binData As Variant, ByVal strFilename As String)

       Dim FileID As Integer
       Dim binByte() As Byte

       binByte = binData
       
       FileID = FreeFile
       Open strFilename For Binary Access Write As #FileID

       Put #FileID, , binByte

       Close #FileID

End Function
' --------------------------------------->

' <---------------------------------------
'               xMLSpanStart
' ----------------------------------------
' Identifies the beginning location of an XML,
' HTML, or SGML span (crude)

Function xMLSpanStart(strDocument As String, Tag As String, Optional Start As Integer = 1, Optional IncludeTags As Boolean = False) As Integer
    Dim tmpStr As String, tagStart As Integer, tagEnd As Integer, tagKill As Integer
    tmpStr = "<" & Tag
    tagStart = InStr(Start, strDocument, tmpStr)
    If tagStart = 0 Then
        xMLSpanStart = 0
        Exit Function
    End If
    If IncludeTags Then
        xMLSpanStart = tagStart
        Exit Function
    End If
    tagKill = InStr(tagStart, strDocument, "/")
    tagEnd = InStr(tagStart, strDocument, ">")
    If tagKill < tagEnd Then
        xMLSpanStart = 0
    Else
        xMLSpanStart = tagEnd + 1
    End If
End Function
' --------------------------------------->


' <---------------------------------------
'                 xMLSpanEnd
' ----------------------------------------
' Identifies the ending location of an XML,
' HTML, or SGML span (crude)

Function xMLSpanEnd(strDocument As String, Tag As String, spnStart As Integer, Optional IncludeTags As Boolean = False) As Integer
    Dim tmpStr As String, tagStart As Integer
    tmpStr = "</" & Tag
    tagStart = InStr(spnStart, strDocument, tmpStr)
    If tagStart <> 0 Then xMLSpanEnd = tagStart - 1 Else xMLSpanEnd = spnStart
    If IncludeTags Then xMLSpanEnd = InStr(tagStart, strDocument, ">")
End Function
' --------------------------------------->


' <---------------------------------------
'               xMLSpanLength
' ----------------------------------------
' Returns the length of an XML, HTML, or SGML span (crude)

Function xMLSpanLength(strDocument As String, Tag As String, Optional spnStart As Integer = 0, Optional IncludeTags As Boolean = False)
    Dim spnEnd
    If spnStart = 0 Then spnStart = xMLSpanStart(strDocument, Tag, 1, IncludeTags)
    spnEnd = xMLSpanEnd(strDocument, Tag, spnStart, IncludeTags)
    xMLSpanLength = spnEnd - spnStart + 1 ' + IncludeTags
End Function
' --------------------------------------->


' <---------------------------------------
'                  xMLSpan
' ----------------------------------------
' Returns the first XML, HTML, or SGML span found (crude)

Function xMLSpan(strDocument As String, Tag As String, Optional Start As Integer = 1, Optional IncludeTags As Boolean = False)
    Dim spnStart As Integer
    spnStart = xMLSpanStart(strDocument, Tag, Start, IncludeTags)
    xMLSpan = Mid(strDocument, xMLSpanStart(strDocument, Tag, Start, IncludeTags), xMLSpanLength(strDocument, Tag, spnStart, IncludeTags))
End Function
' --------------------------------------->

' <---------------------------------------
'                  MonthString
' ----------------------------------------
' Returns the string value of the month (i.e. 1 = "January") or the integer
' of a month string (i.e. "January" = 1)
Public Function MonthString(Optional MonthID, Optional vDate) As String
    If IsMissing(vDate) And IsDate(MonthID) And InStr(CStr(MonthID), "/") > 0 Then vDate = MonthID
    If IsDate(vDate) Then MonthID = Month(vDate)
    MonthInit
    MonthString = clnMonths(MonthID)
End Function

Public Property Get Months() As Variant
    MonthInit
    Set Months = clnMonths
End Property

Public Sub MonthInit()
    ClearCln clnMonths
    clnMonths.Add "January", "1"
    clnMonths.Add "February", "2"
    clnMonths.Add "March", "3"
    clnMonths.Add "April", "4"
    clnMonths.Add "May", "5"
    clnMonths.Add "June", "6"
    clnMonths.Add "July", "7"
    clnMonths.Add "August", "8"
    clnMonths.Add "September", "9"
    clnMonths.Add "October", "10"
    clnMonths.Add "November", "11"
    clnMonths.Add "December", "12"
    clnMonths.Add 1, "January"
    clnMonths.Add 2, "February"
    clnMonths.Add 3, "March"
    clnMonths.Add 4, "April"
    clnMonths.Add 5, "May"
    clnMonths.Add 6, "June"
    clnMonths.Add 7, "July"
    clnMonths.Add 8, "August"
    clnMonths.Add 9, "September"
    clnMonths.Add 10, "October"
    clnMonths.Add 11, "November"
    clnMonths.Add 12, "December"
End Sub
' --------------------------------------->

' <---------------------------------------
'                ClearCln
' ----------------------------------------
' Clears the contents of a collection
Public Function ClearCln(cln As Collection) As Collection
    Dim cnt As Integer, i As Integer
    cnt = cln.Count
    For i = 1 To cnt
        cln.Remove 1
    Next
    Set ClearCln = cln
End Function
' --------------------------------------->

' <---------------------------------------
'                 EditBox
' ----------------------------------------
' Like MsgBox but loads Notepad instead
Public Function EditBox(vStr, Optional vStrLoadFile, Optional vStrSaveFile, Optional bSave)
    Dim strDestFile As String
    Dim index As Integer
    Dim str As String: str = CStr(vStr)
    Dim strLoadFile As String: If Not IsMissing(vStrLoadFile) Then strLoadFile = CStr(vStrLoadFile)
    Dim strSaveFile As String: If Not IsMissing(vStrSaveFile) Then strSaveFile = CStr(vStrSaveFile)
    Dim nploc
    strDestFile = strSaveFile
    
    If IsMissing(bSave) And ((str = "" And strLoadFile <> "") Or strSaveFile <> "") Then _
      bSave = True Else bSave = False
    
    If str <> "" Then
        If strDestFile = "" Then
            strDestFile = GetTempDir & "CL_" & Left(PrepFileString(str), 50) & ".txt"
verifyNotExist:
            If FileExists(strDestFile) Then GoTo addIndex
        End If
        SaveString str, strDestFile
    ElseIf strLoadFile <> "" Then
        If strDestFile = "" Then
            strDestFile = strSaveFile
            SaveString ImportFile(strLoadFile), strDestFile
        Else
            strDestFile = strLoadFile
        End If
    Else
        Err.Raise 9001, "CoolLogic.EditBox: ", "'vStr' or 'vStrLoadFile' is a required parameter"
    End If
    Shell DefaultEditor & " " & strDestFile, vbNormalFocus
    Dim t
    For t = 1 To 10000
        DoEvents
    Next
    On Error Resume Next
tryAgain:
    If Not bSave Then Kill strDestFile
    If Err.Number <> 0 Then
        '.. the file is still being opened... wait...
        Err.Clear
        DoEvents
        GoTo tryAgain
    End If
    Exit Function
addIndex:
    If index > 0 Then
        strDestFile = Left(strDestFile, Len(strDestFile) - 6 - Len(CStr(index))) & "(" & CStr(index) & ").txt"
        index = index + 1
    Else
        strDestFile = Left(strDestFile, Len(strDestFile) - 4) & "(" & CStr(index) & ").txt"
        index = index + 1
    End If
    GoTo verifyNotExist
End Function
' --------------------------------------->

' <---------------------------------------
'              PrepFileString
' ----------------------------------------
' Clears invalid characters from a file path
Public Function PrepFileString(str As String, Optional mask As String = "_")
    Dim vStr As String
    vStr = StripTerminator(str)
    vStr = ReplaceAll(vStr, "\", mask)
    vStr = ReplaceAll(vStr, "/", mask)
    vStr = ReplaceAll(vStr, "*", mask)
    vStr = ReplaceAll(vStr, ":", mask)
    vStr = ReplaceAll(vStr, """", mask)
    vStr = ReplaceAll(vStr, "?", mask)
    vStr = ReplaceAll(vStr, "<", mask)
    vStr = ReplaceAll(vStr, ">", mask)
    vStr = ReplaceAll(vStr, "|", mask)
    PrepFileString = vStr
End Function
' --------------------------------------->

' <---------------------------------------
'              DefaultEditor
' ----------------------------------------
' Defaults to Notepad
Public Property Get DefaultEditor() As String
    If vDefEditor <> "" Then
        DefaultEditor = vDefEditor
    Else
        DefaultEditor = GetSetting("CoolLogic", "Settings", "DefaultEditor")
        If DefaultEditor = "" Then
            DefaultEditor = WinDir & "notepad.exe"
            If Not FileExists(DefaultEditor) Then DefaultEditor = WinSysDir & "notepad.exe"
        End If
        vDefEditor = DefaultEditor
    End If
End Property

Public Property Let DefaultEditor(ByVal vNewValue As String)
    Dim vResults
    If Not FileExists(vNewValue) Then
        vResults = MsgBox("The file '" & vNewValue & "' does not exist.  Are you sure you wish to set " & _
          "this as your default editor?", vbYesNo + vbCritical)
        If vResults = vbNo Then Exit Property
    End If
    SaveSetting "CoolLogic", "Settings", "DefaultEditor", vNewValue
    vDefEditor = vNewValue
End Property
' --------------------------------------->


