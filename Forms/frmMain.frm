VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vServe"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox lstFile 
      Height          =   480
      Left            =   7500
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSWinsockLib.Winsock sckPool 
      Index           =   0
      Left            =   8730
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   8730
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnScan 
      Caption         =   "Scan For Filters"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3690
      TabIndex        =   0
      Top             =   60
      Width           =   1950
   End
   Begin MSComctlLib.ListView lstFilters 
      Height          =   4875
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8599
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Port"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5445
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9604
      MultiRow        =   -1  'True
      Style           =   1
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Protocol Handler"
            Key             =   "PROTOCOL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script Handler"
            Key             =   "SCRIPT"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   5490
      Width           =   60
   End
   Begin VB.Menu mnuFilter 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnScan_Click()
    ScanFilters
    TabStrip.Tabs(1).Selected = True
End Sub

Private Sub lstFilters_DblClick()
Dim sTmp As String

    sTmp = sTmp & "Name: " & WebFilters(lstFilters.SelectedItem.Tag).Name & vbCrLf
    sTmp = sTmp & "Description: " & WebFilters(lstFilters.SelectedItem.Tag).Description & vbCrLf
    sTmp = sTmp & "Version: " & WebFilters(lstFilters.SelectedItem.Tag).Version & vbCrLf
    sTmp = sTmp & "Type: " & ArrFilterType(WebFilters(lstFilters.SelectedItem.Tag).FilterType - 1) & vbCrLf
    sTmp = sTmp & "Path: " & WebFilters(lstFilters.SelectedItem.Tag).Path
    
    
    MsgBox sTmp, vbInformation
    
End Sub

Private Sub lstFilters_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuFilter
    End If
End Sub

Private Sub mnuAbout_Click()
    WebFilters(lstFilters.SelectedItem.Tag).MainClass.About
End Sub

Private Sub mnuConfigure_Click()
    WebFilters(lstFilters.SelectedItem.Tag).MainClass.Configure
End Sub

Private Sub sckPool_Close(Index As Integer)
    sckPool(Index).Close
End Sub

Private Sub sckPool_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim sBuf As String
Dim sTmp As String

    sBuf = String(bytesTotal, Chr(0))
    sckPool(Index).GetData sBuf
        
    If Len(sBuf) > 0 Then
        If ProtocolHandlerRoutine(sBuf, sTmp, Val(sckPool(Index).Tag)) = True Then
            Call ScriptHandlerRoutine(sTmp, sTmp)
            If Len(sTmp) > 0 And sckPool(Index).State = sckConnected Then sckPool(Index).SendData sTmp
        Else
            sckPool(Index).Close
        End If
    Else
        sckPool(Index).Close
    End If
    
End Sub

Private Sub sckPool_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckPool(Index).Close
End Sub

Private Sub sckPool_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If bytesRemaining <= 0 Then sckPool(Index).Close
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim SckID As Long
    
    SckID = GetFreeSocket
    If SckID >= 0 Then
        sckPool(SckID).Tag = sckServer(Index).LocalPort
        sckPool(SckID).Accept requestID
    End If
    
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    sckServer(Index).Close
    
    While Not sckServer(Index).State = sckClosed
        DoEvents
    Wend
    
    sckServer(Index).Listen
    
End Sub


Private Sub TabStrip_Click()
    LoadFilters lstFilters, TabStrip.SelectedItem.Index
End Sub

