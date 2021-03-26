VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H009F700B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client (Use this Form to Connect to Remote Host as a Client)."
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8520
   Begin VB.Frame Frame1 
      BackColor       =   &H009F700B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   6015
      Begin VB.Frame Frame5 
         BackColor       =   &H009F700B&
         Caption         =   "Reciving End"
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   5775
         Begin VB.CommandButton Command13 
            Caption         =   "&Clear Text"
            Height          =   495
            Left            =   3360
            TabIndex        =   24
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtOutput 
            Height          =   1815
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Text            =   "frmClient.frx":0000
            Top             =   240
            Width           =   5535
         End
         Begin VB.CommandButton Command12 
            Caption         =   "&Save File"
            Height          =   495
            Left            =   1080
            TabIndex        =   21
            Top             =   2160
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H009F700B&
         Caption         =   "Sending End"
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton Command3 
            Caption         =   "&Clear Text"
            Height          =   495
            Index           =   0
            Left            =   4320
            TabIndex        =   23
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtSendData 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "frmClient.frx":000E
            Top             =   240
            Width           =   5535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Send &Data"
            Height          =   495
            Left            =   1320
            TabIndex        =   18
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            Caption         =   " &Open File"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Open's File for your "
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Send &File Name"
            Height          =   495
            Left            =   2760
            TabIndex        =   16
            Top             =   2160
            Width           =   1215
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   6120
         Width           =   525
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Bcck To Main Menu"
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Connec&tion Alive/Dead"
      Height          =   615
      Left            =   6240
      TabIndex        =   11
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   6240
      TabIndex        =   10
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Help"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F700B&
      Height          =   1815
      Left            =   6120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Disconnect "
         Height          =   735
         Left            =   1200
         Picture         =   "frmClient.frx":0077
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Connect"
         Height          =   735
         Left            =   120
         Picture         =   "frmClient.frx":0381
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox HostName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "adahmed"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H009F700B&
         Caption         =   "Host Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F700B&
      Caption         =   "Computer Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   6120
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
      Begin VB.CommandButton cmdhostname 
         Caption         =   "&Host Name"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdIP 
         Caption         =   "&IP Adress"
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmndlg1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
   End
   Begin VB.Label Label3 
      BackColor       =   &H009F700B&
      Caption         =   "Chat Client "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3113
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhostname_Click()
Dim X As String
    X = tcpClient.LocalHostName
    MsgBox X
End Sub

Private Sub cmdIP_Click()
Dim X As String
    X = tcpClient.LocalIP
    MsgBox X


End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandler

If HostName = vbNullString Then
 MsgBox "Enter Remote Host Name First!", vbInformation
 Exit Sub
 End If
 
  
If tcpClient.State = 7 Then
 MsgBox "Click Disconnect button first and then connect ", vbExclamation
Exit Sub
End If
 
 
 
tcpClient.RemoteHost = HostName
tcpClient.RemotePort = 1001
 tcpClient.Connect
 
     
ErrHandler:
  Exit Sub
 

End Sub



Private Sub Command10_Click()
On Error GoTo ErrorFile
Dim Size As Integer, nfilenum As String
Dim X As Byte
nfilenum = FreeFile
cmndlg1.InitDir = "C:\My Documents\"
cmndlg1.Filter = "Text Files(*.txt)|*.txt"
cmndlg1.ShowOpen
txtSendData = vbNullString
Open cmndlg1.FileName For Input As #nfilenum
txtSendData.Text = Input$(LOF(nfilenum), #nfilenum)
If (LOF(nfilenum) = 0) Then 'Check file is empty
MsgBox "File Empty"
Exit Sub
End If
'If Len(txtSendData.Text) > 8000 Then
'  MsgBox "You can't send data more than 8KB. It is Network" & vbCrLf _
'     & "data cpacity limit and it depneds on the Network" & vbCrLf _
'    & "    architecture installted. Hit ok to ...", vbExclamation
'   txtSendData.Text = vbNullString
'  Exit Sub
'End If

Close #nfilenum
Command2.Enabled = False
Command11.Enabled = True

Exit Sub

ErrorFile:
MsgBox "Unable to open", vbExclamation, "File Error!"
Exit Sub
End Sub

Private Sub Command11_Click()
Dim FNM As String
' Sends text to other Computer
Dim Temp As String
If tcpClient.State = 0 Then
 MsgBox "Not Conneted"
 Exit Sub
End If

 FNM = cmndlg1.FileName
 tcpClient.SendData FNM & vbCrLf _
 & vbCrLf & "Copy the above path name and save the coming" & vbCrLf _
 & "data of file with this copied path. [Just Paste path in" & vbCrLf _
 & "file name field and then click Save.]"
 txtSendData.SetFocus
 Command2.Enabled = True
 'Label6.Caption = "Data Send: " & cmndlg1.FileTitle
End Sub

Private Sub Command12_Click()
' -------------------------------
' Save file in binary format
'---------------------------------
On Error GoTo ErrorFile
nfilenum = FreeFile 'Get free file number
cmndlg1.Flags = &H80000 Or &H4 Or &H2 Or &H200000
cmndlg1.DialogTitle = "Save File"
cmndlg1.InitDir = "C:\My Documents\" 'Path to save

cmndlg1.Filter = "Text Files(*.txt)|*.txt"
 cmndlg1.ShowSave
 Open cmndlg1.FileName For Binary As #nfilenum
  Put #nfilenum, , txtOutput.Text  'txtOutput.Text
  Close #nfilenum
 

ErrorFile:
Exit Sub

End Sub

Private Sub Command13_Click()
txtOutput = vbNullString
End Sub

Private Sub Command2_Click()
Dim a1 As Long, Txt As String, ss As String

If tcpClient.State = 0 Then
 MsgBox "Not Conneted", vbExclamation
 Exit Sub
End If

If (txtSendData.Text) = vbNullString Then
Label6.Caption = "Please Write a Message First"
Else
 'Label6.Caption = "Data send"
'End If
'Txt = txtSendData
ss = tcpClient.LocalHostName
 tcpClient.SendData ss & ":" & Space(2) & txtSendData
 txtSendData.SetFocus
 ' Label6.Caption = "Data send"
End If
End Sub

Private Sub Command3_Click(Index As Integer)
    
txtSendData = vbNullString
End Sub

Private Sub Command4_Click()
txtOutput = vbNullString
End Sub

Private Sub Command5_Click()
MsgBox "Use this form to send or recive message to or from intended reciptent " & vbCrLf _
& "on the Network as well as for Genral chating.Note whenever you open" & vbCrLf _
& "a required secret text file you first have to send that file's name to" & vbCrLf _
& "the intended recipitent and after that you will able to send file's data.", vbInformation
End Sub

Private Sub Command6_Click()
If tcpClient.State = 0 Then
 MsgBox "First click Connect and then disconnect", vbExclamation
Exit Sub
End If

tcpClient.Close
Label6.Caption = "Connection terminated"
End Sub

Private Sub Command7_Click()
Dim Response As Integer
If tcpClient.State = 7 Then
 Response = MsgBox("If you exit your connection to remote host" & vbCrLf _
 & "will be lost. Are you sure to exit(Y/N)", vbYesNo)
 If Response = vbYes Then
   tcpClient.Close
   Unload frmClient1
   Else
   Exit Sub
 End If
  End If
Unload frmClient1
End Sub

Private Sub Command8_Click()
If tcpClient.State = 7 Then
 Label6.Caption = "Connected to Remote Host"
Else
 Label6.Caption = "Not Connected"
 End If
 Exit Sub
End Sub

Private Sub Form_Load()

 Left = (Screen.Width - Width) / 2   ' Center form horizontally.
 Top = (Screen.Height - Height) / 2   ' Center form vertically.
Command11.Enabled = False
HostName.Text = tcpClient.LocalHostName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Response As Integer, MSG As String
If tcpClient.State = 7 Then
 Response = MsgBox("If you exit your connection yo remote host" & vbCrLf _
 & "will be lost. Are you sure to exit(Y/N/C)", vbQuestion + vbYesNo)
  Select Case Response
      Case vbNo ' Don't allow close.
         Cancel = -1
         Exit Sub
         MSG = "Connection not terminated."
      Case vbYes
       tcpClient.Close
       Unload frmClient
    End Select
     ' Display message if.
End If

End Sub

Private Sub HostName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
 
End Sub

Private Sub tcpClient_Close()
 Label6.Caption = "Connection not accepted. Now first click disconnect" & vbCrLf _
  & "button inorder to connect to any remote host on the network"
Command2.Enabled = False

End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    tcpClient.GetData strData
     'txtOutput.Text = vbNullString
     If frmClient.WindowState = 1 Then
     frmClient.WindowState = 0
    End If
    txtOutput.Text = strData
Command2.Enabled = True
Label6.Caption = "Data Received " & "and " & "Bytes received: " & bytesTotal
txtSendData.SetFocus

End Sub


Private Sub tcpClient_SendProgress(ByVal bytesSent As Long, _
ByVal bytesRemaining As Long)
Label6.Caption = "Bytes send: " & bytesSent & vbCrLf _
              & "Remaning: " & bytesRemaining
                 
End Sub
