Attribute VB_Name = "basFunctions"
Option Explicit



Public Function GetFreeSocket(Optional ServerSocket As Boolean = False) As Long
Dim nCtr As Long
    
    If ServerSocket = True Then
        For nCtr = 0 To frmMain.sckServer.Count - 1
            If frmMain.sckServer(nCtr).State = sckClosed Then
                GetFreeSocket = nCtr
                Exit Function
            End If
        Next
    
        If frmMain.sckServer.Count < MaxSockets Then
            Load frmMain.sckServer(frmMain.sckServer.Count)
            GetFreeSocket = frmMain.sckServer.Count - 1
        End If
    Else
        For nCtr = 0 To frmMain.sckPool.Count - 1
            If frmMain.sckPool(nCtr).State = sckClosed Then
                GetFreeSocket = nCtr
                Exit Function
            End If
        Next
    
        If frmMain.sckPool.Count < MaxSockets Then
            Load frmMain.sckPool(frmMain.sckPool.Count)
            GetFreeSocket = frmMain.sckPool.Count - 1
        End If
    End If

End Function



