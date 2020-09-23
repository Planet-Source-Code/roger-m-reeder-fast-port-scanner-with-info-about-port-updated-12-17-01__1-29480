VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Port Scanner"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Timeout in Sec"
      Height          =   615
      Index           =   2
      Left            =   6480
      TabIndex        =   19
      Top             =   780
      Width           =   1335
      Begin VB.TextBox txtTimeout 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "0.125"
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar pgbrPorts 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4680
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11139
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "Info"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/11/2001"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:54 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "# of Winsocks"
      Height          =   615
      Index           =   1
      Left            =   6480
      TabIndex        =   13
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtWinsocks 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "200"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame framePortsToScan 
      Caption         =   "What Ports To Check"
      Height          =   1335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optPortOptions 
         Caption         =   "Manual"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optPortOptions 
         Caption         =   "Known Trojan/Backdoors"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtToPort 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "32000"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtFromPort 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optPortOptions 
         Caption         =   "Registered"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   960
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   6600
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Index           =   0
      Left            =   6720
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP Range"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3195
      Begin VB.CommandButton cmdFindHost 
         Caption         =   "?"
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         ToolTipText     =   "Look up by Host and Domain"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtEndIP 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkJustMyComputer 
         Caption         =   "Just My Computer"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtStartIP 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Start Scan"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7920
      TabIndex        =   1
      Top             =   360
      Width           =   915
   End
   Begin MSComctlLib.TreeView tvwScans 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NumberOfWinsocks = 1000
Dim sngTimeOut As Single   'Timeout in seconds
Dim blRefreshing As Boolean
Dim iphDNS As New IPHostResolver

Dim cn As ADODB.Connection
Dim rsPorts As ADODB.Recordset
Dim rsTroj As ADODB.Recordset
Dim rsReg As ADODB.Recordset

Dim blStop As Boolean
Dim lngCurrentWinsock As Long
Dim objPorts As New clsPorts
Dim lngX As Single
Dim lngY As Single


Private Sub chkJustMyComputer_Click()
    If Me.chkJustMyComputer Then
        Me.txtStartIP.Text = Me.tcpClient(0).LocalIP
        Me.txtEndIP.Text = Me.tcpClient(0).LocalIP
    End If
    Me.txtStartIP.Enabled = IIf(Me.chkJustMyComputer, False, True)
    Me.txtEndIP.Enabled = IIf(Me.chkJustMyComputer, False, True)
End Sub

Private Sub cmdFindHost_Click()
    'Find by Host or Domain.
    Dim strHost As String
    strHost = InputBox("Enter Hostname and Domain", "Find Host")
    If strHost <> "" Then
        Me.txtStartIP.Text = iphDNS.NameToAddress(strHost)
        Me.txtEndIP.Text = Me.txtStartIP
        If Me.chkJustMyComputer Then Me.chkJustMyComputer = False
    End If
End Sub

Private Sub Form_Load()
    
    Me.txtStartIP.Text = Left(Me.tcpClient(0).LocalIP, InStrRev(Me.tcpClient(0).LocalIP, ".")) & "0"
    Me.txtEndIP.Text = Left(Me.tcpClient(0).LocalIP, InStrRev(Me.tcpClient(0).LocalIP, ".")) & "255"
    Set cn = MakeConnection
    Set rsTroj = New ADODB.Recordset
    Set rsPorts = New ADODB.Recordset
    Set rsReg = New ADODB.Recordset
    
    rsTroj.CursorLocation = adUseClient
    rsTroj.CursorType = adOpenDynamic
    rsTroj.Open "SELECT DISTINCT fldID,fldPort, fldTrojanName FROM tblTrojanPorts WHERE fldType = 'TCP' ORDER BY fldPort, fldTrojanName", cn, adOpenDynamic, adLockReadOnly
    
    rsPorts.CursorLocation = adUseClient
    rsPorts.CursorType = adOpenDynamic
    
    rsReg.CursorLocation = adUseClient
    rsReg.CursorType = adOpenDynamic
    rsReg.Open "SELECT DISTINCT fldID, fldPort, fldRegisterName FROM tblRegisteredPorts WHERE fldType = 'TCP' ORDER BY fldPort, fldRegisterName", cn, adOpenDynamic, adLockReadOnly
    With Me.tvwScans.Nodes  'Add Four Categories to Root in Tree View.
        .Clear
        .Add , , "H", "Open Hosts"
        .Add , , "P", "Open Ports"
        .Add , , "R", "Possible Registed Ports"
        .Add , , "T", "Possible Trojans/Backdoors"
    End With
    chkJustMyComputer_Click 'Initial Setting is just my computer
End Sub

Private Sub ScanPorts()

    Dim strSub As String
    Dim lngHex As Long
    Dim lngStartTime As Long
    Dim lngPort As Long
    Dim lstItems As ListItems
    Dim strKey As String
    Dim lngTime As Long
    Dim lngPortCount As Long
    Dim lngPortMax As Long
    Dim lngWinSocks As Long
    Dim lngStartHex As Long
    Dim lngEndHex As Long
    Dim lngStartPort As Long
    Dim lngEndPort As Long
    Dim rs As ADODB.Recordset
    
    
    If blRefreshing Then Exit Sub
    strSub = Left(Me.txtStartIP, InStrRev(Me.txtStartIP, "."))
    lngStartHex = Val(Mid(Me.txtStartIP, Len(strSub) + 1))
    lngEndHex = Val(Mid(Me.txtEndIP, Len(strSub) + 1))
    lngStartPort = Val(Me.txtFromPort.Text)
    lngEndPort = Val(Me.txtToPort.Text)
    sngTimeOut = Val(Me.txtTimeout)
    If lngStartHex > lngEndHex Then Exit Sub
    blRefreshing = True
    Me.Frame1(0).Enabled = False
    Me.Frame1(1).Enabled = False
    Me.framePortsToScan.Enabled = False
    Me.Frame1(2).Enabled = False
    Progbar True
    Me.cmdScan.Caption = "Halt"
    lngWinSocks = Val(Me.txtWinsocks)
    Me.MousePointer = vbArrowHourglass
    StatusBar "Creating " & lngWinSocks & " Winsocks..."
    Me.pgbrPorts.Min = 0
    Me.pgbrPorts.Max = lngWinSocks
    Me.pgbrPorts.Value = 0
    Do Until Me.tcpClient.Count = lngWinSocks
        Load Me.tcpClient(Me.tcpClient.Count)
        Me.tcpClient(Me.tcpClient.Count - 1).Tag = Timer
        Me.pgbrPorts.Value = Me.tcpClient.Count
        DoEvents
        If blStop Then Exit Do
    Loop
    lngCurrentWinsock = 0
    StatusBar "Scanning Ports..."
    Me.pgbrPorts.Max = lngEndPort
    Me.pgbrPorts.Min = lngStartPort
    Me.pgbrPorts.Value = lngStartPort
    With Me.tvwScans.Nodes
        .Clear
        .Add , , "H", "Open Hosts"
        .Add , , "P", "Open Ports"
        .Add , , "R", "Possible Registed Ports"
        .Add , , "T", "Possible Trojans/Backdoors"
    End With
    'On Error GoTo Error_ScanPorts
    If Me.optPortOptions(2).Value = True Then   'Manual
        lngStartTime = Timer()
        Me.pgbrPorts.Max = lngEndPort
        Me.pgbrPorts.Min = lngStartPort
        Me.pgbrPorts.Value = lngStartPort
        For lngPort = lngStartPort To lngEndPort
            Me.pgbrPorts.Value = lngPort
            If lngTime > 0 And lngTime < Timer Then Exit For
    
            For lngHex = lngStartHex To lngEndHex
                If lngTime > 0 And lngTime < Timer Then Exit For
                Do
                    DoEvents
                    If Me.tcpClient(lngCurrentWinsock).State = sckClosed Then
                        'Doubles scan time.
                        'StatusBar "winsock(" & lngCurrentWinsock & ") " & strSub & lngHex & ":" & lngPort
                        Me.tcpClient(lngCurrentWinsock).RemoteHost = strSub & lngHex
                        Me.tcpClient(lngCurrentWinsock).RemotePort = lngPort
                        Me.tcpClient(lngCurrentWinsock).LocalPort = 0
                        Me.tcpClient(lngCurrentWinsock).Tag = "" & Timer + sngTimeOut
                        Me.tcpClient(lngCurrentWinsock).Connect
                        If iHex = 255 And lngPort = 200 Then blStop = True
                        Exit Do
                    End If
                    If tcpClient(lngCurrentWinsock).State = sckConnected Then
                        AddPortToTree Me.tcpClient(lngCurrentWinsock).RemoteHostIP, Me.tcpClient(lngCurrentWinsock).RemotePort
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    If tcpClient(lngCurrentWinsock).State = sckError Then
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    If tcpClient(lngCurrentWinsock).State >= sckResolvingHost And tcpClient(lngCurrentWinsock).State <= sckConnecting And tcpClient(lngCurrentWinsock).Tag < Timer Then
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    lngCurrentWinsock = lngCurrentWinsock + 1
                    If lngCurrentWinsock = lngWinSocks Then lngCurrentWinsock = 0
                    If blStop Then
                        If lngTime = 0 Then lngTime = Timer + 2
                    End If
                    If lngTime > 0 And lngTime < Timer Then Exit Do
                Loop
            Next lngHex
        Next lngPort
        MsgBox "Scan took " & Timer() - lngStartTime & " seconds", vbInformation, "Scan Finished"
    Else
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        If Me.optPortOptions(1).Value = True Then
            rs.Open "SELECT DISTINCT fldPort FROM tblTrojanPorts WHERE fldType = 'TCP' ORDER BY fldPort", cn, adOpenDynamic, adLockReadOnly
        Else
            rs.Open "SELECT DISTINCT fldPort FROM tblRegisteredPorts WHERE fldType = 'TCP' ORDER BY fldPort", cn, adOpenDynamic, adLockReadOnly
        End If
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
        End If
        Me.pgbrPorts.Max = rs.RecordCount
        Me.pgbrPorts.Min = 0
        lngStartTime = Timer
        Do Until rs.EOF
            Me.pgbrPorts.Value = rs.AbsolutePosition
            lngPort = rs.Fields("fldPort")
            If lngTime > 0 And lngTime < Timer Then Exit Do
    
            For lngHex = lngStartHex To lngEndHex
                If lngTime > 0 And lngTime < Timer Then Exit For
                Do
                    DoEvents
                    If Me.tcpClient(lngCurrentWinsock).State = sckClosed Then
                        'Doubles scan time.
                        'StatusBar "winsock(" & lngCurrentWinsock & ") " & strSub & lngHex & ":" & lngPort
                        Me.tcpClient(lngCurrentWinsock).RemoteHost = strSub & lngHex
                        Me.tcpClient(lngCurrentWinsock).RemotePort = lngPort
                        Me.tcpClient(lngCurrentWinsock).LocalPort = 0
                        Me.tcpClient(lngCurrentWinsock).Tag = "" & Timer + sngTimeOut
                        Me.tcpClient(lngCurrentWinsock).Connect
                        If iHex = 255 And lngPort = 200 Then blStop = True
                        Exit Do
                    End If
                    If tcpClient(lngCurrentWinsock).State = sckConnected Then
                        AddPortToTree Me.tcpClient(lngCurrentWinsock).RemoteHostIP, Me.tcpClient(lngCurrentWinsock).RemotePort
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    If tcpClient(lngCurrentWinsock).State = sckError Then
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    If tcpClient(lngCurrentWinsock).State >= sckResolvingHost And tcpClient(lngCurrentWinsock).State <= sckConnecting And tcpClient(lngCurrentWinsock).Tag < Timer Then
                        tcpClient(lngCurrentWinsock).Close
                    End If
                    lngCurrentWinsock = lngCurrentWinsock + 1
                    If lngCurrentWinsock = lngWinSocks Then lngCurrentWinsock = 0
                    If blStop Then
                        If lngTime = 0 Then lngTime = Timer + 2
                    End If
                    If lngTime > 0 And lngTime < Timer Then Exit Do
                Loop
            Next lngHex
            rs.MoveNext
        Loop
        MsgBox "Scan took " & Timer() - lngStartTime & " seconds", vbInformation, "Scan Finished"
    End If
    StatusBar "Disolving " & lngWinSocks & " Winsocks..."
    Me.pgbrPorts.Min = 0
    Me.pgbrPorts.Max = Me.tcpClient.Count
    Me.pgbrPorts.Value = Me.tcpClient.Count
    Do Until Me.tcpClient.Count = 1
        Unload Me.tcpClient(Me.tcpClient.Count - 1)
        Me.pgbrPorts.Value = Me.tcpClient.Count
        DoEvents
    Loop
    blRefreshing = False
    blStop = False
    Me.Caption = "Port Scanner"
    Me.MousePointer = vbNormal
    Me.Frame1(0).Enabled = True
    Me.Frame1(1).Enabled = True
    Me.Frame1(2).Enabled = True
    Me.framePortsToScan.Enabled = True
    Progbar False
    StatusBar "Ready"
    Me.cmdScan.Caption = "Start Scan"
    Exit Sub
    
Error_ScanPorts:
    blStop = True
    Debug.Print Err.Number; ":" & Err.Description
    Resume Next
End Sub

Private Sub AddPortToTree(strIP As String, lngOpenPort As Long)
    Dim strHost As String
    Dim strPartialKey As String
    Dim objNode As Node
    Dim objNodes As Nodes
    Dim objParentNode As Node
    Dim strTemp As String
    Dim objPort As clsPort
    Dim strTrojans As String
    Dim strRegistered As String
    
    strHost = strIP & " " & iphDNS.AddressToName(strIP)
    strPartialKey = strIP & ":" & lngOpenPort
    Set objPort = objPorts("" & lngOpenPort)    'Find port in collection, if not create it.
    If objPort Is Nothing Then
        If rsReg.State = adStateOpen Then
            rsReg.Filter = "fldPort = " & lngOpenPort
            Do Until rsReg.EOF
                If strTemp <> rsReg.Fields("fldRegisterName") Then
                    
                    strRegistered = strRegistered & IIf(Len(strRegistered) > 0, ", ", "") & rsReg.Fields("fldRegisterName")
                    strTemp = rsReg.Fields("fldRegisterName")
                End If
    '            AddRegisteredToTree rsReg.Fields("fldID"), rsReg.Fields("fldRegisterName"), strHost, lngOpenPort
                rsReg.MoveNext
            Loop
        End If
        strTemp = ""
        If rsTroj.State = adStateOpen Then
            rsTroj.Filter = "fldPort = " & lngOpenPort
            Do Until rsTroj.EOF
                If strTemp <> rsTroj.Fields("fldTrojanName") Then
                    strTrojans = strTrojans & IIf(Len(strTrojans) > 0, ", ", "") & rsTroj.Fields("fldTrojanName")
                    strTemp = rsTroj.Fields("fldTrojanName")
                End If
    '            AddTrojanToTree rsTroj.Fields("fldID"), rsTroj.Fields("fldTrojanName"), strHost, lngOpenPort
                rsTroj.MoveNext
            Loop
        End If
        Set objPort = objPorts.Add(lngOpenPort, strTrojans, strRegistered, "" & lngOpenPort)
    End If
    Set objNodes = Me.tvwScans.Nodes
    
    Set objParentNode = AddNodeToParent(objNodes("H"), "H" & strIP, strHost & " (1 ports)")
    objParentNode.Tag = Val(objParentNode.Tag) + 1
    objParentNode.Text = Left(objParentNode.Text, InStr(1, objParentNode.Text, " (") - 1) & " (" & objParentNode.Tag & " ports)"
    Set objNode = AddNodeToParent(objNodes("H" & strIP), strPartialKey, Format(objPort.PortNumber, "00000") & ":" & IIf(objPort.PossibleRegister <> "", "(Reg: " & objPort.PossibleRegister & ") ", "") & IIf(objPort.PossibleTrojans <> "", "(Troj: " & objPort.PossibleTrojans & ")", ""))
    
    Set objParentNode = AddNodeToParent(objNodes("P"), "P" & Format(lngOpenPort, "00000"), Format(objPort.PortNumber, "00000") & ":" & IIf(objPort.PossibleRegister <> "", "(Reg: " & objPort.PossibleRegister & ") ", "") & IIf(objPort.PossibleTrojans <> "", "(Troj: " & objPort.PossibleTrojans & ")", ""))
    objParentNode.Tag = Val(objParentNode.Tag) + 1
    objParentNode.Text = Format(objPort.PortNumber, "00000") & ":" & IIf(objPort.PossibleRegister <> "", "(Reg: " & objPort.PossibleRegister & ") ", "") & IIf(objPort.PossibleTrojans <> "", "(Troj: " & objPort.PossibleTrojans & ")", "") & " (" & objParentNode.Tag & " hosts)"

    Set objNode = AddNodeToParent(objNodes("P" & Format(lngOpenPort, "00000")), "P" & Format(lngOpenPort, "00000") & strPartialKey, strHost)
    If objPort.PossibleRegister <> "" Then  'Registered, Add to treeview
        Set objParentNode = AddNodeToParent(objNodes("R"), "R" & Format(objPort.PortNumber, "00000"), Format(objPort.PortNumber, "00000") & ":" & objPort.PossibleRegister)
        objParentNode.Tag = Val(objParentNode.Tag) + 1
        objParentNode.Text = Format(objPort.PortNumber, "00000") & ":" & objPort.PossibleRegister & " (" & objParentNode.Tag & " hosts)"
        Set objNode = AddNodeToParent(objNodes("R" & Format(objPort.PortNumber, "00000")), "R" & Format(objPort.PortNumber, "00000") & strPartialKey, strHost)
    End If
    If objPort.PossibleTrojans <> "" Then 'Possible Trojan, Add To Treeview
        Set objParentNode = AddNodeToParent(objNodes("T"), "T" & Format(objPort.PortNumber, "00000"), Format(objPort.PortNumber, "00000") & ":" & objPort.PossibleTrojans)
        objParentNode.Tag = Val(objParentNode.Tag) + 1
        objParentNode.Text = Format(objPort.PortNumber, "00000") & ":" & objPort.PossibleTrojans & " (" & objParentNode.Tag & " hosts)"
        Set objNode = AddNodeToParent(objNodes("T" & Format(objPort.PortNumber, "00000")), "T" & Format(objPort.PortNumber, "00000") & strPartialKey, strHost)
    End If


End Sub

Private Function AddNodeToParent(objParentNode As Node, strKey As String, strText As String) As Node
    Dim nodeParent As Node
    On Error Resume Next
    Set nodeParent = Me.tvwScans.Nodes(strKey)
    If nodeParent Is Nothing Then
        Set AddNodeToParent = Me.tvwScans.Nodes.Add(objParentNode, tvwChild, strKey, strText)
    Else
        Set AddNodeToParent = Me.tvwScans.Nodes(strKey)
    End If
End Function

Private Sub AddTrojanToTree(lngKey As Long, strTrojan As String, strHost As String, lngOpenPort As Long)

End Sub

Private Sub AddRegisteredToTree(lngKey As Long, strRegistered As String, strHost As String, lngOpenPort As Long)

End Sub

Private Sub cmdScan_Click()
    If blRefreshing Then
        blStop = True
    Else
        ScanPorts
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.tvwScans.Width = Me.ScaleWidth - Me.tvwScans.Left * 2
    Me.tvwScans.Height = Me.ScaleHeight - Me.framePortsToScan.Top - Me.framePortsToScan.Height - (Me.tvwScans.Top - Me.framePortsToScan.Top - Me.framePortsToScan.Height) * 2 - Me.sbMain.Height
    If Me.pgbrPorts.Visible Then Progbar True
'    Me.pgbrPorts.Width = Me.ScaleWidth - Me.pgbrPorts.Left * 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blRefreshing Then
        MsgBox "You must halt the scan before closing"
        Cancel = -1
        Exit Sub
    End If
    
    If rsPorts.State = adStateOpen Then rsPorts.Close
    Set rsPorts = Nothing
    
    If rsTroj.State = adStateOpen Then rsTroj.Close
    Set rsTroj = Nothing
    
    If rsReg.State = adStateOpen Then rsReg.Close
    Set rsReg = Nothing
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
End Sub

Private Function getRegisteredForPort(Port As Long) As String
    If rsReg.State <> adStateOpen Then Exit Function
    rsReg.Filter = "fldPort = " & Port
    Do Until rsReg.EOF
        getRegisteredForPort = getRegisteredForPort & IIf(getRegisteredForPort = "", "", ", ") & rsReg.Fields("fldRegisterName")
        rsReg.MoveNext
    Loop
End Function

Private Function getTrojansForPort(Port As Long) As String
    If rsTroj.State <> adStateOpen Then Exit Function
    rsTroj.Filter = "fldPort = " & Port
    Do Until rsTroj.EOF
        getTrojansForPort = getTrojansForPort & IIf(getTrojansForPort = "", "", ", ") & rsTroj.Fields("fldTrojanName")
        rsTroj.MoveNext
    Loop
End Function

Private Function GetConnectionString() As String
    GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\scanner.mdb"
End Function

Private Function MakeConnection() As ADODB.Connection
    Dim cn As ADODB.Connection
    Dim strCon As String
    On Error GoTo Error_MakeConnection
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    strCon = GetConnectionString()
    cn.Open strCon
    Set MakeConnection = cn
    Exit Function

Error_MakeConnection:
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Sub Label2_Click()

End Sub

Private Sub optPortOptions_Click(Index As Integer)
    Me.txtFromPort.Enabled = Me.optPortOptions(2)
    Me.txtToPort.Enabled = Me.optPortOptions(2)
End Sub

Private Sub txtStartIP_Validate(Cancel As Boolean)
    Me.txtEndIP.Text = Left(Me.txtStartIP, InStrRev(Me.txtStartIP, ".")) & "254"
End Sub

Private Sub StatusBar(Info As String)
    Me.sbMain.Panels("Info").Text = Info
End Sub

Private Sub Progbar(ProgBarVisible As Boolean)
    If ProgBarVisible Then
        If lngX = 0 Then
            Me.pgbrPorts.Left = Me.sbMain.Left + Me.sbMain.Panels("Date").Left
            Me.pgbrPorts.Top = Me.sbMain.Top
            Me.pgbrPorts.Height = Me.sbMain.Height
            Me.pgbrPorts.Width = Me.sbMain.Panels("Time").Left + Me.sbMain.Panels("Time").Width - Me.sbMain.Panels("Date").Left
            lngX = Me.pgbrPorts.Left - Me.ScaleWidth
            lngY = Me.pgbrPorts.Top - Me.ScaleHeight
        Else
            Me.pgbrPorts.Move Me.ScaleWidth + lngX, Me.ScaleHeight + lngY
        End If
        Me.pgbrPorts.Visible = True
    Else
        Me.pgbrPorts.Visible = False
    End If
End Sub

