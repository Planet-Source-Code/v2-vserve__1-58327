Attribute VB_Name = "basFilters"
Option Explicit

Public Sub ScanFilters()
Dim sTmp As String
Dim nCtr As Long
Dim nCtr1 As Long

    sTmp = FiltersPath
    sTmp = Replace(UCase(sTmp), "%APP_PATH%", App.Path)
    sTmp = Replace(UCase(sTmp), "\\", "\")
    
    With frmMain.lstFile
        .Path = sTmp
        .Pattern = FiltersExt
        .Refresh
    End With
    
    For nCtr = 0 To frmMain.lstFile.ListCount - 1
        sTmp = Replace(frmMain.lstFile.Path & "\" & frmMain.lstFile.List(nCtr), "\\", "\")
        AddFilter sTmp
    Next
    
    frmMain.lblStatus.Caption = "Total " & WebFilterCtr & " Filters Loaded."

End Sub

Public Function AddFilter(strPath As String) As Boolean
On Error GoTo errHandler


Dim objTmp As Object
Dim ClassName As String
Dim sTmp() As String
Dim nCtr As Long
Dim SckID As Long

    ClassName = Mid(strPath, InStrRev(strPath, "\") + 1)
    
    If InStr(ClassName, ".") > 0 Then
        ClassName = UCase(Mid(ClassName, 1, InStrRev(ClassName, "."))) & MainClassName
    Else
        ClassName = UCase(ClassName) & "." & MainClassName
    End If
    
    
    For nCtr = 0 To WebFilterCtr - 1
        If WebFilters(nCtr).ClassName = ClassName Then
            AddFilter = False
            Exit Function
        End If
    Next
    
    
    RegisterModule (strPath)
    ReDim Preserve WebFilters(WebFilterCtr) As WEB_FILTER
    Set WebFilters(WebFilterCtr).MainClass = CreateObject(ClassName)
    
    If TypeName(WebFilters(WebFilterCtr).MainClass) <> MainClassName Then
        Set WebFilters(WebFilterCtr).MainClass = Nothing
        AddFilter = False
    Else
        sTmp = Split(WebFilters(WebFilterCtr).MainClass.GetInfo, Chr(0))
        WebFilters(WebFilterCtr).ClassName = ClassName
        WebFilters(WebFilterCtr).Name = sTmp(0)
        WebFilters(WebFilterCtr).Description = sTmp(1)
        WebFilters(WebFilterCtr).Version = sTmp(2)
        WebFilters(WebFilterCtr).FilterType = CByte(sTmp(3))
        WebFilters(WebFilterCtr).Path = strPath

        If WebFilters(WebFilterCtr).FilterType = ProtocolHandler Then
            WebFilters(WebFilterCtr).FilterPort = CLng(sTmp(4))
            SckID = GetFreeSocket(True)
            frmMain.sckServer(SckID).Close
            frmMain.sckServer(SckID).LocalPort = WebFilters(WebFilterCtr).FilterPort
            frmMain.sckServer(SckID).Listen
        End If
        WebFilterCtr = WebFilterCtr + 1
        AddFilter = True
    End If
    
errHandler:
    If Err.Number = 429 Then
        AddFilter = False
        Exit Function
    End If
    
End Function

Public Function ProtocolHandlerRoutine(ByVal strInput As String, ByRef strOutput As String, Optional ByVal ServerPort As Long = 0) As Boolean

Dim nCtr As Long
Dim sTmp As String

    
    For nCtr = 0 To WebFilterCtr - 1
        If WebFilters(nCtr).FilterType = ProtocolHandler Then
            If WebFilters(nCtr).FilterPort = ServerPort Then
                sTmp = WebFilters(nCtr).MainClass.Parser(strInput, vbNull, vbNullString)
                Exit For
            End If
        End If
    Next
    
    strOutput = sTmp
    ProtocolHandlerRoutine = True
    
End Function

Public Function ScriptHandlerRoutine(ByVal strInput As String, ByRef strOutput As String) As Boolean

Dim nCtr As Long
Dim sTmp As String
    
    sTmp = strInput
    For nCtr = 0 To WebFilterCtr - 1
        If WebFilters(nCtr).FilterType = ScriptParser Then
            sTmp = WebFilters(nCtr).MainClass.Parser(sTmp, vbNull, vbNullString)
        End If
    Next
    
    strOutput = sTmp
    ScriptHandlerRoutine = True
    
End Function


Public Function RegisterModule(sPath As String) As Long
Dim hModule As Long
Dim hProc As Long
Dim lRet As Long

    lRet = -1
    
    hModule = LoadLibrary(sPath)
    If hModule <= 0 Then
        Exit Function
    End If
    
    hProc = GetProcAddress(hModule, "DllRegisterServer")
    If hProc <= 0 Then
        FreeLibrary hModule
        Exit Function
    End If
    
    
    lRet = CallWindowProc(hProc, 0, 0, 0, 0)
    FreeLibrary hModule
    RegisterModule = lRet
    
End Function

Public Sub LoadFilters(lstView As ListView, FilterType As enumFilterType)
Dim nCtr As Long

    lstView.ListItems.Clear
    
    For nCtr = 0 To WebFilterCtr - 1
        If WebFilters(nCtr).FilterType = FilterType Then
            lstView.ListItems.Add , , WebFilters(nCtr).Name
            lstView.ListItems(lstView.ListItems.Count).Tag = nCtr
            lstView.ListItems(lstView.ListItems.Count).SubItems(1) = WebFilters(nCtr).Version
            lstView.ListItems(lstView.ListItems.Count).SubItems(2) = WebFilters(nCtr).Description
            lstView.ListItems(lstView.ListItems.Count).SubItems(3) = WebFilters(nCtr).FilterPort
        End If
    Next
    
End Sub
