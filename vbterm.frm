VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTerminal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Octane CarrierDetect"
   ClientHeight    =   8280
   ClientLeft      =   2925
   ClientTop       =   2040
   ClientWidth     =   9075
   ForeColor       =   &H00000000&
   Icon            =   "vbterm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9075
   Begin VB.Timer tmrTimeOut 
      Interval        =   2000
      Left            =   9090
      Top             =   225
   End
   Begin VB.Frame Frame4 
      Caption         =   "Terminal Window"
      Height          =   7485
      Left            =   0
      TabIndex        =   7
      Top             =   450
      Width           =   4830
      Begin VB.TextBox txtTerm 
         Height          =   7110
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   270
         Width           =   4545
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Connected"
      Height          =   7485
      Left            =   6975
      TabIndex        =   5
      Top             =   450
      Width           =   2040
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Log"
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   7065
         Width           =   1815
      End
      Begin VB.ListBox lstCarriers 
         Height          =   6690
         ItemData        =   "vbterm.frx":030A
         Left            =   135
         List            =   "vbterm.frx":030C
         TabIndex        =   6
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Numbers Dialed"
      Height          =   7485
      Left            =   4860
      TabIndex        =   3
      Top             =   450
      Width           =   2040
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Log"
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   7065
         Width           =   1815
      End
      Begin VB.ListBox lstDialed 
         Height          =   6690
         ItemData        =   "vbterm.frx":030E
         Left            =   135
         List            =   "vbterm.frx":0310
         TabIndex        =   4
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   270
      Top             =   1800
   End
   Begin MSComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenLogFile"
            Description     =   "Open Log File..."
            Object.ToolTipText     =   "Open Log File..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CloseLogFile"
            Description     =   "Close Log File"
            Object.ToolTipText     =   "Close Log File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DialPhoneNumber"
            Description     =   "Dial Phone Number..."
            Object.ToolTipText     =   "Start Dialing..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "HangUpPhone"
            Description     =   "Hang Up Phone"
            Object.ToolTipText     =   "Hang Up Phone"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Description     =   "Properties..."
            Object.ToolTipText     =   "Properties..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "TransmitTextFile"
            Description     =   "Transmit Text File..."
            Object.ToolTipText     =   "Transmit Text File..."
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4000
         TabIndex        =   2
         Top             =   75
         Width           =   240
         Begin VB.Image imgConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":0312
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgNotConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":045C
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   45
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   105
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "LOG"
      FileName        =   "Open Communications Log File"
      Filter          =   "Log File (*.log)|*.log;"
      FilterIndex     =   501
      FontSize        =   9.02458e-38
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7965
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Status:"
            TextSave        =   "Status:"
            Key             =   "Status"
            Object.ToolTipText     =   "Communications Port Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12146
            MinWidth        =   2
            Text            =   "Settings:"
            TextSave        =   "Settings:"
            Key             =   "Settings"
            Object.ToolTipText     =   "Communications Port Settings"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1244
            Key             =   "ConnectTime"
            Object.ToolTipText     =   "Connect Time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":08C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":0BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":0EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":120E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":1528
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenLog 
         Caption         =   "&Open Log File..."
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "&Close Log File"
         Enabled         =   0   'False
      End
      Begin VB.Menu M3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendText 
         Caption         =   "&Transmit Text File..."
         Enabled         =   0   'False
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "&CommPort"
      Begin VB.Menu mnuOpen 
         Caption         =   "Port &Open"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
   End
   Begin VB.Menu mnuMSComm 
      Caption         =   "&MSComm"
      Begin VB.Menu mnuInputLen 
         Caption         =   "&InputLen..."
      End
      Begin VB.Menu mnuRThreshold 
         Caption         =   "&RThreshold..."
      End
      Begin VB.Menu mnuSThreshold 
         Caption         =   "&SThreshold..."
      End
      Begin VB.Menu mnuParRep 
         Caption         =   "P&arityReplace..."
      End
      Begin VB.Menu mnuDTREnable 
         Caption         =   "&DTREnable"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCD 
         Caption         =   "&CDHolding..."
      End
      Begin VB.Menu mnuHCTS 
         Caption         =   "CTSH&olding..."
      End
      Begin VB.Menu mnuHDSR 
         Caption         =   "DSRHo&lding..."
      End
   End
   Begin VB.Menu mnuCall 
      Caption         =   "C&all"
      Begin VB.Menu mnuRnd 
         Caption         =   "&Dialing Options"
      End
      Begin VB.Menu mnuDial 
         Caption         =   "&Start Dialing Numbers..."
      End
      Begin VB.Menu mnuHangUp 
         Caption         =   "&Hang Up Phone"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                        
Dim Ret As Integer
Dim Temp As String
Dim hLogFile As Integer ' Handle of open log file.
Dim StartTime As Date   ' Stores starting time for port timer

Private Sub Command1_Click()
Dim Resp As VbMsgBoxResult

Resp = MsgBox("Are you sure?", vbQuestion + vbYesNo, "OctaCD")

If Resp = vbYes Then
   Me.lstDialed.Clear
   If Dir(App.Path & "\Dailed.log") <> "" Then
      Kill App.Path & "\Dailed.log"
   End If
End If

End Sub

Private Sub Command2_Click()
Dim Resp As VbMsgBoxResult

Resp = MsgBox("Are you sure?", vbQuestion + vbYesNo, "OctaCD")
If Resp = vbYes Then
   Me.lstCarriers.Clear
   If Dir(App.Path & "\Carriers.log") <> "" Then
      Kill App.Path & "\Carriers.log"
   End If
End If

End Sub

Private Sub Form_Load()
    Dim CommPort As String, Handshaking As String, Settings As String
    Dim tmp As String
    
    On Error Resume Next
    
    ' Set the default color for the terminal
    txtTerm.SelLength = Len(txtTerm)
    txtTerm.SelText = ""
    txtTerm.ForeColor = vbBlue
       
    ' Set Title
    App.Title = "Octane CarrierDetect"
    
    ' Set up status indicator light
    imgNotConnected.ZOrder
       
    ' Center Form
    frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    ' Load Registry Settings
    
    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
    If Settings <> "" Then
        MSComm1.Settings = Settings
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' frmTerminal.MSComm1.CommPort
    If CommPort <> "" Then MSComm1.CommPort = CommPort
    
    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
    If Handshaking <> "" Then
        MSComm1.Handshaking = Handshaking
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
    On Error GoTo 0
    
    If Dir(App.Path & "\Dialed.log") <> "" Then
       Open App.Path & "\Dialed.log" For Input As #2
         While Not EOF(2)
            Line Input #2, tmp
            lstDialed.AddItem tmp
         Wend
       Close #2
    End If
    If Dir(App.Path & "\Carriers.log") <> "" Then
       Open App.Path & "\Carriers.log" For Input As #2
         While Not EOF(2)
            Line Input #2, tmp
            lstCarriers.AddItem tmp
         Wend
       Close #2
    End If
    HangUp

End Sub

Private Sub Form_Resize()
   ' Resize the Term (display) control
   'txtTerm.Move 0, tbrToolBar.Height, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - sbrStatus.Height - tbrToolBar.Height
   
   ' Position the status indicator light
   Frame1.Left = ScaleWidth - Frame1.Width * 1.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

    If MSComm1.PortOpen Then
       ' Wait 10 seconds for data to be transmitted.
       Counter = Timer + 10
       Do While MSComm1.OutBufferCount
          Ret = DoEvents()
          If Timer > Counter Then
             Select Case MsgBox("Data cannot be sent", 34)
                ' Cancel.
                Case 3
                   Cancel = True
                   Exit Sub
                ' Retry.
                Case 4
                   Counter = Timer + 10
                ' Ignore.
                Case 5
                   Exit Do
             End Select
          End If
       Loop

       MSComm1.PortOpen = 0
    End If

    ' If the log file is open, flush and close it.
    If hLogFile Then mnuCloseLog_Click
    End
End Sub

Private Sub imgConnected_Click()
    ' Call the mnuOpen_Click routine to toggle connect and disconnect
    Call mnuOpen_Click
End Sub

Private Sub imgNotConnected_Click()
    ' Call the mnuOpen_Click routine to toggle connect and disconnect
    Call mnuOpen_Click
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuCloseLog_Click()
    ' Close the log file.
    Close hLogFile
    hLogFile = 0
    mnuOpenLog.Enabled = True
    tbrToolBar.Buttons("OpenLogFile").Enabled = True
    mnuCloseLog.Enabled = False
    tbrToolBar.Buttons("CloseLogFile").Enabled = False
    frmTerminal.Caption = "Visual Basic Terminal"
End Sub

Private Sub mnuDial_Click()
    On Local Error Resume Next
    
    Dim x As Double
    
    Num = ""
    Num = RndNum
    
    For x = 0 To frmRndOpt.lstDontCall.ListCount - 1
        If Right(frmRndOpt.lstDontCall.List(x), 6) = Num Then
           mnuDial_Click
        End If
    Next
    For x = 0 To Me.lstDialed.ListCount - 1
        If Right(Me.lstDialed.List(x), 6) = Num Then
           mnuDial_Click
        End If
    Next x
    
    
    
    ' Get a number from the user.
    'Num = InputBox$("Enter Phone Number:", "Dial Number", Num)
     Num = frmRndOpt.txtPreFix & Num
    
    ' Open the port if it isn't already open.
    If Not MSComm1.PortOpen Then
       mnuOpen_Click
       If Err Then Exit Sub
    End If
      
    ' Enable hang up button and menu item
    mnuHangUp.Enabled = True
    tbrToolBar.Buttons("HangUpPhone").Enabled = True
              
    'Add Number to dialed list
    lstDialed.AddItem Num
    Open App.Path & "\Dialed.log" For Output As #2
         For x = 0 To lstDialed.ListCount - 1
             Print #2, lstDialed.List(x)
         Next x
    Close #2
    
    DoEvents
    
   For x = 1 To 4000
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
   Next x
   
   tmrTimeOut.Enabled = True
   For x = 1 To 4000
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
   Next x
   
   While Waiting = True
       DoEvents
   Wend
    
    ' Dial the number.
    DoEvents
    MSComm1.Output = "ATDT" & Num & vbCrLf
    
    ' Start the port timer
    StartTiming
End Sub

' Toggle the DTREnabled property.
Private Sub mnuDTREnable_Click()
    ' Toggle DTREnable property
    MSComm1.DTREnable = Not MSComm1.DTREnable
    mnuDTREnable.Checked = MSComm1.DTREnable
End Sub


Private Sub mnuFileExit_Click()
    ' Use Form_Unload since it has code to check for unsent data and an open log file.
    Form_Unload Ret
End Sub



' Toggle the DTREnable property to hang up the line.
Private Sub mnuHangup_Click()
    On Error Resume Next
    HangUp
    MSComm1.Output = "ATH"      ' Send hangup string
    Ret = MSComm1.DTREnable     ' Save the current setting.
    MSComm1.DTREnable = True    ' Turn DTR on.
    MSComm1.DTREnable = False   ' Turn DTR off.
    MSComm1.DTREnable = Ret     ' Restore the old setting.
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' If port is actually still open, then close it
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    MSComm1.PortOpen = False
    
    ' Notify user of error
    'If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Settings: "
    
    ' Turn off indicator light and uncheck open menu
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Stop the port timer
    StopTiming
    sbrStatus.Panels("Status").Text = "Status: "
    
    mnuDial_Click
    
    On Error GoTo 0
End Sub

' Display the value of the CDHolding property.
Private Sub mnuHCD_Click()
    If MSComm1.CDHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CDHolding = " + Temp
End Sub

' Display the value of the CTSHolding property.
Private Sub mnuHCTS_Click()
    If MSComm1.CTSHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CTSHolding = " + Temp
End Sub

' Display the value of the DSRHolding property.
Private Sub mnuHDSR_Click()
    If MSComm1.DSRHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "DSRHolding = " + Temp
End Sub

' This procedure sets the InputLen property, which determines how
' many bytes of data are read each time Input is used
' to retreive data from the input buffer.
' Setting InputLen to 0 specifies that
' the entire contents of the buffer should be read.
Private Sub mnuInputLen_Click()
    On Error Resume Next

    Temp = InputBox$("Enter New InputLen:", "InputLen", Str$(MSComm1.InputLen))
    If Len(Temp) Then
        MSComm1.InputLen = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

Private Sub mnuProperties_Click()
  ' Show the CommPort properties form
  frmProperties.Show vbModal
  
End Sub

' Toggles the state of the port (open or closed).
Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim OpenFlag

    MSComm1.PortOpen = Not MSComm1.PortOpen
    If Err Then MsgBox Error$, 48
    
    OpenFlag = MSComm1.PortOpen
    
    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag
        
    If MSComm1.PortOpen Then
        ' Enable dial button and menu item
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Enable hang up button and menu item
        mnuHangUp.Enabled = True
        tbrToolBar.Buttons("HangUpPhone").Enabled = True
        
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Settings: " & MSComm1.Settings
        StartTiming
    Else
        ' Enable dial button and menu item
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Disable hang up button and menu item
        mnuHangUp.Enabled = False
        tbrToolBar.Buttons("HangUpPhone").Enabled = False
        
        imgNotConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Settings: "
        StopTiming
    End If
    
End Sub

Private Sub mnuOpenLog_Click()
   Dim replace
   On Error Resume Next
   OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
   OpenLog.CancelError = True
      
   ' Get the log filename from the user.
   OpenLog.DialogTitle = "Open Communications Log File"
   OpenLog.Filter = "Log Files (*.LOG)|*.log|All Files (*.*)|*.*"
   
   Do
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then Exit Sub
      Temp = OpenLog.FileName

      ' If the file already exists, ask if the user wants to overwrite the file or add to it.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
      If Ret Then
         replace = MsgBox("Replace existing file - " + Temp + "?", 35)
      Else
         replace = 0
      End If
   Loop While replace = 2

   ' User clicked the Yes button, so delete the file.
   If replace = 6 Then
      Kill Temp
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
   End If

   ' Open the log file.
   hLogFile = FreeFile
   Open Temp For Binary Access Write As hLogFile
   If Err Then
      MsgBox Error$, 48
      Close hLogFile
      hLogFile = 0
      Exit Sub
   Else
      ' Go to the end of the file so that new data can be appended.
      Seek hLogFile, LOF(hLogFile) + 1
   End If

   frmTerminal.Caption = "Visual Basic Terminal - " + OpenLog.FileTitle
   mnuOpenLog.Enabled = False
   tbrToolBar.Buttons("OpenLogFile").Enabled = False
   mnuCloseLog.Enabled = True
   tbrToolBar.Buttons("CloseLogFile").Enabled = True
End Sub

' This procedure sets the ParityReplace property, which holds the
' character that will replace any incorrect characters
' that are received because of a parity error.
Private Sub mnuParRep_Click()
    On Error Resume Next

    Temp = InputBox$("Enter Replace Character", "ParityReplace", frmTerminal.MSComm1.ParityReplace)
    frmTerminal.MSComm1.ParityReplace = Left$(Temp, 1)
    If Err Then MsgBox Error$, 48
End Sub

Private Sub mnuRnd_Click()
frmRndOpt.Show
End Sub

' This procedure sets the RThreshold property, which determines
' how many bytes can arrive at the receive buffer before the OnComm
' event is triggered and the CommEvent property is set to comEvReceive.
Private Sub mnuRThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Enter New RThreshold:", "RThreshold", Str$(MSComm1.RThreshold))
    If Len(Temp) Then
        MSComm1.RThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If

End Sub




' The OnComm event is used for trapping communications events and errors.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    
    ' Branch according to the CommEvent property.
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            'Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            ShowData txtTerm, (StrConv(Buffer, vbUnicode))
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Change in CTS Detected"
            
            lstCarriers.AddItem Num
        Case comEvDSR
            EVMsg$ = "Change in DSR Detected"
        Case comEvCD
            EVMsg$ = "Change in CD Detected"
        Case comEvRing
            EVMsg$ = "The Phone is Ringing"
        Case comEvEOF
            EVMsg$ = "End of File Detected"

        ' Error messages.
        Case comBreak
            ERMsg$ = "Break Received"
            mnuHangup_Click
        Case comCDTO
            ERMsg$ = "Carrier Detect Timeout"
            mnuHangup_Click
        Case comCTSTO
            ERMsg$ = "CTS Timeout"
            mnuHangup_Click
        Case comDCB
            ERMsg$ = "Error retrieving DCB"
            mnuHangup_Click
        Case comDSRTO
            ERMsg$ = "DSR Timeout"
            mnuHangup_Click
        Case comFrame
            ERMsg$ = "Framing Error"
            mnuHangup_Click
        Case comOverrun
            ERMsg$ = "Overrun Error"
            mnuHangup_Click
        Case comRxOver
            ERMsg$ = "Receive Buffer Overflow"
            mnuHangup_Click
        Case comRxParity
            ERMsg$ = "Parity Error"
            mnuHangup_Click
        Case comTxFull
            ERMsg$ = "Transmit Buffer Full"
            mnuHangup_Click
        Case Else
            ERMsg$ = "Unknown error or event"
            mnuHangup_Click
    End Select
    
    If Len(EVMsg$) Then
        ' Display event messages in the status bar.
        sbrStatus.Panels("Status").Text = "Status: " & EVMsg$
                
        ' Enable timer so that the message in the status bar
        ' is cleared after 2 seconds
        Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Display event messages in the status bar.
        sbrStatus.Panels("Status").Text = "Status: " & ERMsg$
        
        ' Display error messages in an alert message box.
        Beep
        'Ret = MsgBox(ERMsg$, 1, "Click Cancel to quit, OK to ignore.")
        
        ' If the user clicks Cancel (2)...
        If Ret = 2 Then
            MSComm1.PortOpen = False    ' Close the port and quit.
        End If
        
        ' Enable timer so that the message in the status bar
        ' is cleared after 2 seconds
        Timer2.Enabled = True
    End If
End Sub

Private Sub mnuSendText_Click()
   Dim hSend, BSize, LF&
   
   On Error Resume Next
   
   mnuSendText.Enabled = False
   tbrToolBar.Buttons("TransmitTextFile").Enabled = False
   
   ' Get the text filename from the user.
   OpenLog.DialogTitle = "Send Text File"
   OpenLog.Filter = "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
   Do
      OpenLog.CancelError = True
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then
        mnuSendText.Enabled = True
        tbrToolBar.Buttons("TransmitTextFile").Enabled = True
        Exit Sub
      End If
      Temp = OpenLog.FileName

      ' If the file doesn't exist, go back.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         mnuSendText.Enabled = True
         tbrToolBar.Buttons("TransmitTextFile").Enabled = True
         Exit Sub
      End If
      If Ret Then
         Exit Do
      Else
         MsgBox Temp + " not found!", 48
      End If
   Loop

   ' Open the log file.
   hSend = FreeFile
   Open Temp For Binary Access Read As hSend
   If Err Then
      MsgBox Error$, 48
   Else
      ' Display the Cancel dialog box.
      CancelSend = False
      frmCancelSend.Label1.Caption = "Transmitting Text File - " + Temp
      frmCancelSend.Show
      
      ' Read the file in blocks the size of the transmit buffer.
      BSize = MSComm1.OutBufferSize
      LF& = LOF(hSend)
      Do Until EOF(hSend) Or CancelSend
         ' Don't read too much at the end.
         If LF& - Loc(hSend) <= BSize Then
            BSize = LF& - Loc(hSend) + 1
         End If
      
         ' Read a block of data.
         Temp = Space$(BSize)
         Get hSend, , Temp
      
         ' Transmit the block.
         MSComm1.Output = Temp
         If Err Then
            MsgBox Error$, 48
            Exit Do
         End If
      
         ' Wait for all the data to be sent.
         Do
            Ret = DoEvents()
         Loop Until MSComm1.OutBufferCount = 0 Or CancelSend
      Loop
   End If
   
   Close hSend
   mnuSendText.Enabled = True
   tbrToolBar.Buttons("TransmitTextFile").Enabled = True
   CancelSend = True
   frmCancelSend.Hide
End Sub


' This procedure sets the SThreshold property, which determines
' how many characters (at most) have to be waiting
' in the output buffer before the CommEvent property
' is set to comEvSend and the OnComm event is triggered.
Private Sub mnuSThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Enter New SThreshold Value", "SThreshold", Str$(MSComm1.SThreshold))
    If Len(Temp) Then
        MSComm1.SThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

' This procedure adds data to the Term control's Text property.
' It also filters control characters, such as BACKSPACE,
' carriage return, and line feeds, and writes data to
' an open log file.
' BACKSPACE characters delete the character to the left,
' either in the Text property, or the passed string.
' Line feed characters are appended to all carriage
' returns.  The size of the Term control's Text
' property is also monitored so that it never
' exceeds MAXTERMSIZE characters.
Private Static Sub ShowData(Term As Control, Data As String)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i
    Dim z As Integer
    
    ' Make sure the existing text doesn't get too large.
    TermSize = Len(Term.Text)
    If TermSize > MAXTERMSIZE Then
       Term.Text = Mid$(Term.Text, 4097)
       TermSize = Len(Term.Text)
    End If

    ' Point to the end of Term's data.
    Term.SelStart = TermSize

    ' Filter/handle BACKSPACE characters.
    Do
       i = InStr(Data, Chr$(8))
       If i Then
          If i = 1 Then
             Term.SelStart = TermSize - 1
             Term.SelLength = 1
             Data = Mid$(Data, i + 1)
          Else
             Data = Left$(Data, i - 2) & Mid$(Data, i + 1)
          End If
       End If
    Loop While i

    ' Eliminate line feeds.
    Do
       i = InStr(Data, Chr$(10))
       If i Then
          Data = Left$(Data, i - 1) & Mid$(Data, i + 1)
       End If
    Loop While i

    ' Make sure all carriage returns have a line feed.
    i = 1
    Do
       i = InStr(i, Data, Chr$(13))
       If i Then
          Data = Left$(Data, i) & Chr$(10) & Mid$(Data, i + 1)
          i = i + 1
       End If
    Loop While i

    ' Add the filtered data to the SelText property.
    Term.SelText = Data
    z = InStr(1, Data, "ATDT", vbTextCompare)
    If z = 0 Then
       If InStr(1, Data, "BUSY", vbTextCompare) Then mnuHangup_Click
       If InStr(1, Data, "NO CARRIER", vbTextCompare) Then mnuHangup_Click
       If InStr(1, Data, "NO DIALTONE", vbTextCompare) Then
          HangUp
          MSComm1.PortOpen = False
          MSComm1.PortOpen = False
          MSComm1.PortOpen = False
          MSComm1.PortOpen = False
          MSComm1.PortOpen = False
          MSComm1.PortOpen = False
          MSComm1.PortOpen = True
          HangUp
          DoEvents
          mnuHangup_Click
       End If
       If InStr(1, Data, "CONNECT", vbTextCompare) Then
          lstCarriers.AddItem Num
          Open App.Path & "\Carriers.log" For Output As #2
             For z = 0 To lstCarriers.ListCount - 1
                 Print #2, lstCarriers.List(z)
             Next z
             MSComm1.PortOpen = False
             MSComm1.PortOpen = True
             mnuHangup_Click
          Close #2
       End If
     End If
      
    ' Log data to file if requested.
    If hLogFile Then
       i = 2
       Do
          Err = 0
          Put hLogFile, , Data
          If Err Then
             i = MsgBox(Error$, 21)
             If i = 2 Then
                mnuCloseLog_Click
             End If
          End If
       Loop While i <> 2
    End If
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
If Err.Number <> 8012 Then
   MsgBox Error$
End If
    Resume Next
End Sub

Private Sub Timer2_Timer()
sbrStatus.Panels("Status").Text = "Status: "
Timer2.Enabled = False

End Sub

Private Sub tmrTimeOut_Timer()
Waiting = True
Dim x As Integer
HangUp
For x = 1 To 3500
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
Next x
HangUp
tmrTimeOut.Enabled = False
Waiting = False
End Sub

' Keystrokes trapped here are sent to the MSComm
' control where they are echoed back via the
' OnComm (comEvReceive) event, and displayed
' with the ShowData procedure.
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    ' If the port is opened...
    If MSComm1.PortOpen Then
        ' Send the keystroke to the port.
        MSComm1.Output = Chr$(KeyAscii)
        
        ' Unless Echo is on, there is no need to
        ' let the text control display the key.
        ' A modem usually echos back a character
        If Not Echo Then
            ' Place position at end of terminal
            txtTerm.SelStart = Len(txtTerm)
            KeyAscii = 0
        End If
    End If
     
End Sub




Private Sub tbrToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
Select Case Button.Key
Case "OpenLogFile"
    Call mnuOpenLog_Click
Case "CloseLogFile"
    Call mnuCloseLog_Click
Case "DialPhoneNumber"
    Call mnuDial_Click
Case "HangUpPhone"
    Call mnuHangup_Click
Case "Properties"
    Call mnuProperties_Click
Case "TransmitTextFile"
    Call mnuSendText_Click
End Select
End Sub

Private Sub Timer1_Timer()
    ' Display the Connect Time
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
End Sub
' Call this function to start the Connect Time timer
Private Sub StartTiming()
    StartTime = Now
    Timer1.Enabled = True
End Sub
' Call this function to stop timing
Private Sub StopTiming()
    Timer1.Enabled = False
    sbrStatus.Panels("ConnectTime").Text = ""
End Sub

Sub HangUp()
Dim x As Integer
    
For x = 1 To 5
    On Error Resume Next
    MSComm1.Output = "ATH"      ' Send hangup string
    Ret = MSComm1.DTREnable     ' Save the current setting.
    MSComm1.DTREnable = True    ' Turn DTR on.
    MSComm1.DTREnable = False   ' Turn DTR off.
    MSComm1.DTREnable = Ret     ' Restore the old setting.
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' If port is actually still open, then close it
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    MSComm1.PortOpen = False
    
    ' Notify user of error
    'If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Settings: "
    
    ' Turn off indicator light and uncheck open menu
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Stop the port timer
    StopTiming
    sbrStatus.Panels("Status").Text = "Status: "
    On Error GoTo 0
Next x
End Sub

