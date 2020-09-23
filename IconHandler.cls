VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const S_FALSE As Long = 1
Private Const CP_ACP As Long = 0
Private Const IMAGE_ICON As Long = 1
Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2

Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long


' *****************************************************************************
' * Icons4VB.IconHandler                                                      *
' * By grigri, 2004                                                           *
' *                                                                           *
' * Purpose:      Implement an Icon Handler Shell Extension for VB files      *
' *                                                                           *
' * How it Works: Since this class is registered as the icon handler for the  *
' *               filetypes in the registry, when Explorer wants an icon it   *
' *               creates an instance of it.                                  *
' *               First it calls IPersistFile::Load() to inform us of the     *
' *               filename, then it calls IExtractIcon::GetIconLocation() to  *
' *               retrieve the path for the icon. If we return a value with   *
' *               GIL_NOTFILENAME from this method then Explorer will call    *
' *               IExtractIcon::Extract() for us to create the icon, passing  *
' *               us back the parameters we specified in GetIconLocation()    *
' *               When the class is released, we must destroy the icons we've *
' *               created to avoid a resource leak.                           *
' *                                                                           *
' * Note:         I've only implemented the ANSI version (IExtractIconA) here *
' *               On NT/2k/XP systems, Explorer will first query for the      *
' *               UNICODE version (IExtractIconW). It didn't seem worth it    *
' *               to implement it here, as it could obfuscate the purpose and *
' *               I wanted to keep it simple (that's also why I did't cover   *
' *               asynchronous extraction... maybe another time)              *
' *****************************************************************************

' These are precalculated codes for file extensions
' They represent the ascii bytes for the extensions,
' packed into a Long variable. Note that the order
' you see is reversed (they all end in &h2E which is
' the ascii code for ".", the first character of the
' extension). These are used to speed up the identification
' of the file type.
Private Enum ExtensionCodes
    ExtCodeVBG = &H6762762E
    ExtCodeVBP = &H7062762E
    ExtCodeFRM = &H6D72662E
End Enum

' The Shell Extension interfaces we must implement
Implements IPersistFile
Implements IExtractIconA

' The file we want the icon for
Private sSelectedFile As String
' Its code (calculated in IPersistFile::Load)
Private eSelectedExtCode As ExtensionCodes
' Our generated icon handles
Private hIconLarge As Long
Private hIconSmall As Long

Private Sub Class_Terminate()
    ' Delete the icons
    DestroyIcon hIconSmall
    DestroyIcon hIconLarge
End Sub

Private Function IExtractIconA_GetIconLocation(ByVal uFlags As Icons4VBHelper.GIL_INFLAGS, ByVal szIconFile As Long, ByVal cchMax As Long, piIndex As Long) As Icons4VBHelper.GIL_OUTFLAGS
    Dim nFile As Integer, sLine As String
    
    On Error Resume Next
    
    ' Copy the dll's filename to the buffer, converting it to ANSI as we go
    WideCharToMultiByte CP_ACP, 0, StrPtr(App.Path & "\" & App.EXEName & ".dll" & vbNullChar), -1, szIconFile, cchMax, 0, 0
    
    ' Setup return flags
    ' GIL_NOTFILENAME means "don't try and extract by yourself, call Extract() afterwards)
    ' GIL_DONTCACHE means "this icon might change later, so don't assume it's fixed"
    IExtractIconA_GetIconLocation = GIL_NOTFILENAME Or GIL_DONTCACHE
    
    ' Branch on Extension Code
    Select Case eSelectedExtCode
    Case ExtCodeVBG
        ' VB Project Group File
        ' ---------------------
        piIndex = 101
        ' TODO : Add number of projects as blob-overlay on icon
    Case ExtCodeVBP
        ' VB Project File
        ' ---------------
        ' Default to "Old" project icon
        piIndex = 102
        ' Get a free file index and open the file for reading
        Err.Clear
        nFile = FreeFile
        Open sSelectedFile For Input Access Read Shared As nFile
        If Err Then
            ' Couldn't open the file
            ' It might be in a zip file or something
            ' Just bug out with the default icon
            Exit Function
        End If
        ' Read the first line (contains the project type)
        Line Input #nFile, sLine
        ' Close the file
        Close nFile
        ' Branch on the project type
        If Left$(sLine, 5) = "Type=" Then
            Select Case Mid$(sLine, 6)
            Case "Exe"
                piIndex = 103
            Case "OleDll"
                piIndex = 104
            Case "OleExe"
                piIndex = 105
            Case "Control"
                piIndex = 106
            End Select
        End If
    Case ExtCodeFRM
        ' VB Form File
        ' ------------
        ' Default to normal form icon
        piIndex = 107
        ' Get a free file index and open the file for reading
        Err.Clear
        nFile = FreeFile
        Open sSelectedFile For Input Access Read Shared As nFile
        If Err Then
            ' Couldn't open the file
            ' It might be in a zip file or something
            ' Just bug out with the default icon
            Exit Function
        End If
        ' Read the file line by line
        Do While Not EOF(nFile)
            Line Input #nFile, sLine
            ' Check for end of properties - might as well bug out now if it is
            If sLine = "End" Then Exit Do
            ' Check for MDI Client
            If Left$(sLine, 16) = "Begin VB.MDIForm" Then
                piIndex = 109
                Exit Do
            End If
            ' Check for "MDIChild" property
            If Left$(LTrim$(sLine), 8) = "MDIChild" Then
                If InStr(9, sLine, "-1") > 0 Then piIndex = 108
                Exit Do
            End If
        Loop
        Close nFile
    Case Else
        Err.Raise S_FALSE
    End Select
End Function

Private Sub IExtractIconA_Extract(ByVal pszFile As Long, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIconSize As Long)
    Dim hLib As Long
    Dim nLargeIconSize As Long, nSmallIconSize As Long
    
    ' LOWORD of nIconSize contains the size of the large icon
    nLargeIconSize = nIconSize And &HFFFF&
    ' HIWORD of nIconSize contains the size of the small icon
    nSmallIconSize = (nIconSize) \ &H10000
    
    ' Load the required file (in memory, for its resources)
    hLib = LoadLibraryEx(pszFile, 0, LOAD_LIBRARY_AS_DATAFILE)
    
    ' Load the icons at the specified size
    
    hIconLarge = LoadImage(hLib, nIconIndex, IMAGE_ICON, nLargeIconSize, nLargeIconSize, 0)
    hIconSmall = LoadImage(hLib, nIconIndex, IMAGE_ICON, nSmallIconSize, nSmallIconSize, 0)
    phiconLarge = hIconLarge
    phiconSmall = hIconSmall
    
    ' Free the library
    FreeLibrary hLib
End Sub

Private Sub IPersistFile_Load(ByVal pszFileName As Long, ByVal dwMode As Long)
    ' Retrieve the file whose icon we're going to supply
    sSelectedFile = StrFromPtr(pszFileName, True)
    ' Calculate the extension code for this file
    eSelectedExtCode = CalcExtCode(sSelectedFile)
End Sub

Private Function IPersistFile_GetCurFile() As Long
    ' Not used in Icon Handlers
End Function
Private Sub IPersistFile_IsDirty()
    ' Not used in Icon Handlers
End Sub
Private Sub IPersistFile_Save(ByVal pszFileName As Long, ByVal fRemember As Long)
    ' Not used in Icon Handlers
End Sub
Private Sub IPersistFile_SaveCompleted(ByVal pszFileName As Long)
    ' Not used in Icon Handlers
End Sub

Private Function StrFromPtr(ByVal lpsz As Long, Optional Unicode As Boolean) As String
    ' Convert a C-style string to a VB string
    If Unicode Then
        StrFromPtr = String$(lstrlenW(lpsz), 0)
        MoveMemory ByVal StrPtr(StrFromPtr), ByVal lpsz, LenB(StrFromPtr)
    Else
        StrFromPtr = String$(lstrlenA(lpsz), 0)
        lstrcpyA StrFromPtr, lpsz
    End If
End Function

Private Function CalcExtCode(sFile As String) As Long
    ' This is a nasty hack :)
    ' Translate the last 4 characters of sFile to ANSI,
    ' and store them in a long variable
    ' This enables MUCH faster checking
    WideCharToMultiByte CP_ACP, 0, StrPtr(sFile) + LenB(sFile) - 8, 4, VarPtr(CalcExtCode), 4, 0, 0
    ' Make the letters all uppercase
    ' (this only works with characters a-z and A-Z, no accents or "stooges")
    CalcExtCode = CalcExtCode Or &H20202000
End Function

