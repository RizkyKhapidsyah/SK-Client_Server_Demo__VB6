VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main1 
   BackColor       =   &H009F700B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Connection Server (Use this form to Accept Multiple Connections)."
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8940
   Icon            =   "main1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H009F700B&
      Caption         =   "Your Computer Information"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6480
      TabIndex        =   21
      Top             =   2400
      Width           =   2415
      Begin VB.CommandButton Command3 
         Caption         =   "&Host Name"
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&IP Adress"
         Height          =   615
         Index           =   0
         Left            =   1320
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Back To Main Form"
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   20
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6960
      TabIndex        =   19
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Rest Server"
      Height          =   495
      Left            =   6960
      TabIndex        =   18
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Help"
      Height          =   495
      Left            =   6960
      TabIndex        =   17
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Frame Frame2 
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
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6255
      Begin VB.Frame Frame3 
         BackColor       =   &H009F700B&
         Caption         =   "Sending End"
         ForeColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6015
         Begin VB.CommandButton Command1 
            Caption         =   "Send to current host"
            Height          =   495
            Index           =   2
            Left            =   960
            TabIndex        =   25
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton Send 
            Caption         =   "Send &File Name"
            Height          =   495
            Index           =   1
            Left            =   3360
            TabIndex        =   24
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtSendData 
            ForeColor       =   &H00000000&
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Text            =   "main1.frx":000C
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton Command10 
            Caption         =   "&Open File"
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
            TabIndex        =   15
            ToolTipText     =   "Open's File for your "
            Top             =   2040
            Width           =   615
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&Clear Text"
            Height          =   495
            Index           =   1
            Left            =   4920
            TabIndex        =   14
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton Command13 
            Caption         =   "   Send to    &all  hosts"
            Height          =   495
            Index           =   0
            Left            =   2280
            TabIndex        =   13
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H009F700B&
         Caption         =   "Receiving End"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   6015
         Begin VB.CommandButton Command11 
            Caption         =   "&Save File "
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtOutput 
            Height          =   1815
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "main1.frx":0075
            Top             =   240
            Width           =   4815
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Clear Text "
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   450
      End
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   9360
      Tag             =   "0"
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock disallow 
      Left            =   9240
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F700B&
      Caption         =   " Information "
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton advanced 
         Caption         =   "Client Information"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label live_connections_info 
         BackColor       =   &H009F700B&
         Caption         =   "N/A"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Clients Connected :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label sockets_loaded_info 
         BackColor       =   &H009F700B&
         Caption         =   "N/A"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Sockets Loaded :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.ListBox status 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   2280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog cmndlg1 
      Left            =   9240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu servertab 
      Caption         =   "Server"
      Begin VB.Menu settingstab 
         Caption         =   "Settings"
      End
      Begin VB.Menu tabthang 
         Caption         =   "-"
      End
      Begin VB.Menu reset_server_tab 
         Caption         =   "Reset"
      End
      Begin VB.Menu tabthing 
         Caption         =   "-"
      End
      Begin VB.Menu shutdown_server 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu debugtab 
      Caption         =   "Debug"
      Begin VB.Menu start_telnet 
         Caption         =   "TelNet Connect"
      End
   End
End
Attribute VB_Name = "main1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Intmax As Integer
Private Sub advanced_Click()
wininfo.Show
End Sub

Private Sub Command1_Click(Index As Integer)
Dim ss As String

If sock(Intmax).State = 0 Then
 MsgBox "Not Conneted", vbExclamation
 Exit Sub
 End If


ss = sock(Intmax).LocalHostName
send_data Intmax, ss & " says:" & Space(2) & txtSendData
'sock(intMax).SendData ss & " says:" & Space(2) & txtSendData, 0

End Sub

Private Sub Command10_Click()
On Error GoTo ErrorFile
Dim Size As Integer
Dim X As Byte
nfilenum = FreeFile
cmndlg1.InitDir = "C:\My Documents\"
cmndlg1.Filter = "Text Files(*.txt)|*.txt"
cmndlg1.ShowOpen
txtSendData = vbNullString
Open cmndlg1.FileName For Input As #nfilenum
txtSendData.Text = Input$(LOF(nfilenum), #nfilenum)
Send(1).Enabled = True
Command13(0).Enabled = False

If (LOF(nfilenum) = 0) Then 'Check file is empty
MsgBox "file empty"
Exit Sub
End If
 
Close #nfilenum
Exit Sub

ErrorFile:
Exit Sub

End Sub

Private Sub Command13_Click(Index As Integer)
Dim ss As String

If sock(Intmax).State = 0 Then
 MsgBox "Not Conneted", vbExclamation
 Exit Sub
 End If


ss = sock(Intmax).LocalHostName
'sock(intMax).SendData ss & ":" & Space(2) & txtSendData
mass_send ss & " says:" & Space(2) & txtSendData, 0
'Label6.Caption = "Data send"

End Sub

Private Sub Command2_Click(Index As Integer)
Dim a As String
a = sock(Intmax).LocalIP
MsgBox a

End Sub

Private Sub Command3_Click(Index As Integer)
Dim a As String
a = sock(Intmax).LocalHostName
MsgBox a
End Sub

Private Sub Command5_Click()
MsgBox "Use this form to send or receive message to or from intended reciptent " & vbCrLf _
& "on the Network as well as for Genral chating.Note whenever you open" & vbCrLf _
& "a required secret text file you first have to send that file's name to" & vbCrLf _
& "the intended recipitent and after that you will able to send file's data.", vbInformation

End Sub

Private Sub Command6_Click()
reset_server
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
Dim Response As Integer, MSG As String
If sock(Intmax).State = 7 Then
 Response = MsgBox("If you exit your connection yo remote host" & vbCrLf _
 & "will be lost. Are you sure to exit(Y/N/C)", vbQuestion + vbYesNo)
  Select Case Response
      Case vbNo ' Don't allow close.
         Cancel = -1
         Exit Sub
       Case vbYes
       'when the program ends, close all the sockets.
      close_all_sockets
      Unload wininfo
      Unload Me
 
    End Select
     ' Display message if.
End If
End Sub

Private Sub Command9_Click(Index As Integer)
txtSendData = vbNullString
End Sub

Private Sub Form_Load()

'==========================================
 ''       adahmed911@hotmail.com
 '
 '==========================================


Intmax = 0
'start up the server
start_server
Command13(0).Enabled = False
'update info
update_info
Send(1).Enabled = False
Command1(2).Enabled = False
frmClient.Show
Client1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Response As Integer, MSG As String
If sock(Intmax).State = 7 Then
 Response = MsgBox("If you exit your connection yo remote host" & vbCrLf _
 & "will be lost. Are you sure to exit(Y/N/C)", vbQuestion + vbYesNo)
  Select Case Response
      Case vbNo ' Don't allow close.
         Cancel = -1
         Exit Sub
       Case vbYes
       'tcpClient.Close
       'Unload frmClient
       'when the program ends, close all the sockets.
      close_all_sockets
      Unload wininfo
      Unload Me
 
    End Select
     ' Display message if.
End If
End Sub

Private Sub reset_server_tab_Click()
reset_server
End Sub

Private Sub Send_Click(Index As Integer)
If sock(Intmax).State = 0 Then
 MsgBox "not Conneted"
 Exit Sub
 End If

If (txtSendData.Text) = vbNullString Then
Label6.Caption = "Please Write a Message First"
Else
Label6.Caption = "Data send"
End If
 FNM = cmndlg1.FileName
 sock(Intmax).SendData FNM & vbCrLf _
 & vbCrLf & "Copy the above path name and save the coming" & vbCrLf _
 & "data of file with this copied path. [Just Paste path in" & vbCrLf _
 & "file name field and then click Save.]"
 txtSendData.SetFocus
Command13(0).Enabled = True
'Label6.Caption = "Data Send: " & cmndlg1.FileTitle
 
End Sub

Private Sub settingstab_Click()
'show the settings window
settings_window.Show
End Sub

Private Sub shutdown_server_Click()
Unload Me
End Sub

Private Sub sock_Close(Index As Integer)
'Log out clients once they have quit
logout_client Index, "Connection long"
If sock(Intmax).State = 7 Then
' Just exit
Else
Label6.Caption = "Connetion terminated"
Command13(0).Enabled = False
End If

End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'incomming data,to recive it and send it to get decoded
Dim new_data As String, Response As Integer
sock(Index).GetData new_data
DoEvents
decode_data new_data, Index
 If main1.WindowState = 1 Then
    main1.WindowState = vbNormal
  End If
 txtOutput.Text = new_data
If Intmax > 1 Then
Response = MsgBox("Accept Data from Another host! ", vbYesNo)
If Response = vbYes Then
 txtOutput.Text = new_data
Else
 Exit Sub
End If
End If
Label6.Caption = "Data Received. " & "and " & "Bytes received: " & bytesTotal

End Sub

Private Sub sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 'Log out clients if error on port
logout_client Index, "Error - " & Description
End Sub

Private Sub sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'Login a new user on a connection request
Dim Response As Integer

If Index = "0" Then
Intmax = Intmax + 1
'show in status
'update_status ">> Incomming Connection Request <<"
'login new user
     Response = MsgBox("Remote host wants Connection, Proceed?", vbQuestion + vbYesNo)
       If Response = vbYes Then
         If main1.WindowState = 1 Then
          main1.WindowState = 0
          End If
         Beep
         new_connection requestID
        DoEvents
       '' sock(Intmax).SendData "Connection accepted. " & sock(Intmax).LocalHostName
         Label6.Caption = "Connected to Remote hostname: " & sock(Intmax).RemoteHost
        Command13(0).Enabled = True
        Command1(2).Enabled = True
          Exit Sub
           Else
           Intmax = Intmax - 1
            Exit Sub
         End If





End If

End Sub



Private Sub sock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Label6.Caption = "Btes send: " & bytesSent & vbCrLf _
                & "Remaning: " & bytesRemaining

End Sub

Private Sub start_telnet_Click()

AppActivate Shell("telnet 127.0.0.1 " & server_port, vbNormalNoFocus)



End Sub

