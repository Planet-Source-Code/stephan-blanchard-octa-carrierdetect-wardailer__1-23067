VERSION 5.00
Begin VB.Form frmRndOpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Dialing options"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmRndOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   2805
      Begin VB.TextBox txtPreFix 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   8
         Text            =   "*31*0800"
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1305
         TabIndex        =   7
         Text            =   "######"
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Mask:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   225
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   3105
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Not to be dialed...."
      Height          =   3300
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   2805
      Begin VB.CommandButton cmdRem 
         Caption         =   "&Remove"
         Height          =   285
         Left            =   1845
         TabIndex        =   5
         Top             =   585
         Width           =   780
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   285
         Left            =   1845
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox txtDontCall 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
      Begin VB.ListBox lstDontCall 
         Height          =   2595
         ItemData        =   "frmRndOpt.frx":27A2
         Left            =   90
         List            =   "frmRndOpt.frx":27AC
         TabIndex        =   1
         Top             =   585
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmRndOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

    Me.lstDontCall.AddItem txtDontCall.Text
    txtDontCall.Text = ""
    Open App.Path & "\DontCall.log" For Output As #2
         For x = 0 To lstDontCall.ListCount - 1
             Print #2, lstDontCall.List(x)
         Next x
    Close #2

End Sub

Private Sub cmdOK_Click()
Me.Visible = False
Me.Hide
frmTerminal.Show
End Sub


Private Sub cmdRem_Click()
Dim i As Integer
Dim x As Integer


If lstDontCall.Text <> "" Then
   i = lstDontCall.ListIndex
   lstDontCall.RemoveItem i
    Open App.Path & "\DontCall.log" For Output As #2
         For x = 0 To lstDontCall.ListCount - 1
             Print #2, lstDontCall.List(x)
         Next x
    Close #2
End If


End Sub

Private Sub Form_Load()
Dim x As Integer
Dim tmp As String

If Dir(App.Path & "\DontCall.log") <> "" Then
   Open App.Path & "\DontCall.log" For Input As #2
      lstDontCall.Clear
      While Not EOF(2)
        Line Input #2, tmp
        lstDontCall.AddItem tmp
      Wend
   Close #2
End If
   
   

End Sub

Private Sub lstDontCall_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyDelete Then
   'lstDontCall.RemoveItem
End If

End Sub

Private Sub txtDontCall_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Me.lstDontCall.AddItem txtDontCall.Text
    txtDontCall.Text = ""
    Open App.Path & "\DontCall.log" For Output As #2
         For x = 0 To lstDontCall.ListCount - 1
             Print #2, lstDontCall.List(x)
         Next x
    Close #2
End If

End Sub
