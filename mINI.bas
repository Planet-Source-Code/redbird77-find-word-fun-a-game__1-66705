Attribute VB_Name = "mINI"
' mINI.bas \ redbird77@earthlink.net \ 2006.10.02

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub PutValue(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal File As String)

Dim Ret As Long

    Ret = WritePrivateProfileString(Section, Key, Value, File)

End Sub

Public Sub DelValue(ByVal Section As String, ByVal Key As String, ByVal File As String)

Dim Ret As Long

    Ret = WritePrivateProfileString(Section, Key, vbNullString, File)
    
End Sub

Public Function GetValue(ByVal Section As String, ByVal Key As String, ByVal File As String, Optional ByVal Default As String = "") As String

Dim Buf As String
Dim Ret As Long
Dim Pos As Long
    
    Buf = String$(1024, vbNullChar)
    Ret = GetPrivateProfileString(Section, Key, Default, Buf, Len(Buf), File)
    Pos = InStr(Buf, vbNullChar)
    
    If Pos Then GetValue = Left$(Buf, Pos - 1)
    
End Function

Public Sub PutSettings()

' Write settings from the global Settings user-defined type to the
' Settings INI file.

Dim File    As String
Dim Section As String

    File = App.Path & "\Settings.ini"
    Section = "Settings"
    
    Call PutValue(Section, "SquareSize", Game.Settings.SquareSize, File)
    Call PutValue(Section, "SquareCount", Game.Settings.SquareCount, File)
    
    Call PutValue(Section, "SelectColor", Game.Settings.SelectColor, File)
    Call PutValue(Section, "SolutionColor", Game.Settings.SolutionColor, File)
    
    Call PutValue(Section, "Backwards", Game.Settings.Backwards, File)
    
End Sub

Public Sub GetSettings()

' Place settings from file (using default values, if file not found) into
' the global Game.Settings user-defined type.

Dim File    As String
Dim Section As String

    File = App.Path & "\Settings.ini"
    Section = "Settings"
    
    With Game.Settings
        
        .SquareSize = GetValue(Section, "SquareSize", File, 50)
        .SquareCount = GetValue(Section, "SquareCount", File, 15)
    
        .SelectColor = GetValue(Section, "SelectColor", File, vbRed)
        .SolutionColor = GetValue(Section, "SolutionColor", File, &HEEEEEE)
    
        .Backwards = GetValue(Section, "Backwards", File, 1)
        
    End With
    
End Sub
