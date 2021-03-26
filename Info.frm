VERSION 5.00
Begin VB.Form wininfo 
   BackColor       =   &H009F700B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Information"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5175
   Begin VB.CommandButton refresh_list 
      Caption         =   "Refresh &List"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox selected_id 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F700B&
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      Begin VB.Frame Frame2 
         BackColor       =   &H009F700B&
         Caption         =   "Chat with Selected client"
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
         Height          =   2175
         Left            =   2520
         TabIndex        =   13
         Top             =   960
         Width           =   2415
         Begin VB.CommandButton Command1 
            Caption         =   "&Apply"
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H009F700B&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"Info.frx":030A
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
            Height          =   1335
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.TextBox info_ip 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton refresh_button 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox info_idle_since 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox info_connected_at 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox info_socket 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "IP:"
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
         TabIndex        =   12
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Idle Since:"
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
         TabIndex        =   8
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Connected At:"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H009F700B&
         Caption         =   "Socket:"
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
         Width           =   675
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H009F700B&
      Caption         =   "Client ID :"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "wininfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If selected_id = "" Then Exit Sub
If selected_id < 1 Then Exit Sub

n2 = selected_id.Text
n1 = n2
Unload wininfo
End Sub

Private Sub Form_Load()
 Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2   ' Center form vertically.
n2 = 0
n1 = 0
'get the list on load
refresh_list_Click

End Sub

Private Sub refresh_button_Click()
'update info
selected_id_Click
End Sub

Private Sub refresh_list_Click()

selected_id.Clear

'fill the selection list with the possible clientIDs
For i = 1 To max_clients
If Client(i).socket <> "0" Then selected_id.AddItem i
Next i

'select the first item
If selected_id.ListCount > 0 Then selected_id.ListIndex = "0"


End Sub

Private Sub selected_id_Click()
'show the info on the selected client

If selected_id = "" Then Exit Sub
If selected_id < 1 Then Exit Sub

info_connected_at = Client(selected_id).connected_at
info_idle_since = Client(selected_id).idle_since
info_socket = Client(selected_id).socket
info_ip = main1.sock(Client(selected_id).socket).RemoteHostIP

End Sub
