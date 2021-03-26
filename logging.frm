VERSION 5.00
Begin VB.Form settings_window 
   BackColor       =   &H009F700B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Settings"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "logging.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2910
   Begin VB.Frame Frame2 
      BackColor       =   &H009F700B&
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin VB.TextBox info_max 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "TBA"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox info_port 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "TBA"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Maximum Clients : "
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Server Port : "
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1155
      End
   End
   Begin VB.CommandButton default_settings 
      Caption         =   "Default"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton apply_settings 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H009F700B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You can change Maximum clients up to 120, but server port  is not allowed to change in this version of application."
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   2895
   End
End
Attribute VB_Name = "settings_window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub apply_settings_Click()
'change the settings

'change the max number of clients
If info_max <> "" Then
Let max_clients = info_max
If live_connections > max_clients Then MsgBox "Their are currently more users connect than the limit.", vbOKOnly, "Settings Notice"
End If

'change the server port
If info_port <> server_port And info_port <> "" Then change_server_port info_port



Unload Me
End Sub

Private Sub default_settings_Click()
info_max = default_max_clients
info_port = default_server_port
End Sub

Private Sub Form_Load()
'show the current settings
 Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2   ' Center form vertically.

info_max = max_clients
info_port = server_port

End Sub

Private Sub info_max_Change()
'change max number of users

If info_max = "" Then Exit Sub

On Error GoTo ec
If info_max < 0 Then Let info_max = 0
If info_max > server_max_clients Then
MsgBox "Max 200 clients are allowed", vbInformation
info_max = server_max_clients
End If
Exit Sub
ec:
info_max.Text = max_clients

End Sub

Private Sub info_port_Change()


'change max number of users

If info_port = "" Then Exit Sub

On Error GoTo ec
If info_port < 1 Then Let info_port = 1
If info_port > 65535 Then info_port = 65535

Exit Sub
ec:
info_port = server_port


End Sub

Private Sub SSTab1_DblClick()

End Sub
