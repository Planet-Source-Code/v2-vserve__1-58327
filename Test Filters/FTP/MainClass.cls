VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const FILTER_NAME As String = "FTP"
Private Const FILTER_DESC As String = "FTP Protocol Handler"
Private Const FILTER_TYPE As Byte = 1
Private Const FILTER_PORT As Long = 21

Public Function Parser(ByVal strInput As String, Optional ByVal LongReserved As Long, Optional ByVal StrReserved As String) As String
Dim sTmp As String

    Select Case UCase(Left(strInput, 4))
        Case "USER"
            sTmp = "331 SUCCESS"
        Case "PASS"
            sTmp = "331 SUCCESS"
    End Select
    
    Parser = sTmp
    
End Function

Public Function GetInfo() As String
    GetInfo = FILTER_NAME & Chr(0) & _
                FILTER_DESC & Chr(0) & _
                GetVersion & Chr(0) & _
                FILTER_TYPE & Chr(0) & _
                FILTER_PORT
End Function


Private Function GetVersion() As String
    GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


Public Function Configure() As Boolean
    MsgBox "Nothing To Configure Yet...!!!"
End Function

Public Function About() As Boolean
    MsgBox "About"
End Function


