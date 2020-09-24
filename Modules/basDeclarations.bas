Attribute VB_Name = "basDeclarations"
Option Explicit
Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Public Const PortToListen As Long = 13754
Public Const MaxSockets As Long = 3000

Public Const FiltersPath As String = "%APP_PATH%\Filters\"
Public Const FiltersExt As String = "*.*"
Public Const MainClassName As String = "MainClass"


Public Type WEB_FILTER
    ClassName As String
    Name As String
    Description As String
    Version As String
    Path As String
    MainClass As Object
    FilterType As enumFilterType
    FilterPort As Long
End Type

Public Enum enumFilterType
    ProtocolHandler = 1
    ScriptParser = 2
End Enum

Public ArrFilterType() As String
Public WebFilters() As WEB_FILTER
Public WebFilterCtr As Long
