VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLoadResPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Property Get TempDir() As String
Dim sRet As String, c As Long
Dim lErr As Long
   sRet = String$(MAX_PATH, 0)
   c = GetTempPath(MAX_PATH, sRet)
   lErr = Err.LastDllError
   If c = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   TempDir = Left$(sRet, c)
End Property

Private Property Get TempFileName( _
        Optional ByVal sPrefix As String, _
        Optional ByVal sPathName As String) As String
Dim lErr As Long
Dim iPos As Long

   If sPrefix = "" Then sPrefix = ""
   If sPathName = "" Then sPathName = TempDir
   
   Dim sRet As String
   sRet = String(MAX_PATH, 0)
   GetTempFileName sPathName, sPrefix, 0, sRet
   lErr = Err.LastDllError
   If Not lErr = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   iPos = InStr(sRet, vbNullChar)
   If Not iPos = 0 Then
      TempFileName = Left$(sRet, iPos - 1)
   End If

End Property

Private Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
   
   ' Return the error message associated with LastDLLError:
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If
   
End Function

Public Property Get LoadResPicture(ByVal ID As Variant, ByVal Format As Variant) As IPicture
Dim sFile As String
Dim B() As Byte
Dim iFile As Integer

On Error GoTo ErrorHandler
   
   B = LoadResData(ID, Format)
   sFile = TempFileName("LRP")
   iFile = FreeFile
   Open sFile For Binary Access Write Lock Read As #iFile
   Put #iFile, , B
   Close #iFile
   iFile = 0
   Set LoadResPicture = LoadPicture(sFile)
   KillFile sFile
   Exit Property

ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number:   sErr = Err.Description
   If Not iFile = 0 Then Close #iFile
   KillFile sFile
   Err.Raise Err.Number, App.EXEName & ".cLoadResPicture", Err.Description
   Exit Property
   
End Property

Private Sub KillFile(ByVal sFile As String)
   
   On Error Resume Next
   Kill sFile

End Sub
