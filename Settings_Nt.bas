Attribute VB_Name = "Settings_Nt"

'maximum ammount of clients

'the maximum clients the server will handle
Public Const server_max_clients = 200
Public n1 As Integer
Public n2 As Integer
'the default maximum number of clients
Public Const default_max_clients = 20
Public max_clients As Integer

'port for clients to connect to
Public Const default_server_port = "1001" '"6000"
Public server_port As Long

Public live_connections As Integer

'this is the data-type for each client.
'it keeps a record of everybody connected
'and also stores data on what socket they
'are using, customize for your needs.

Type client_type

'socket they are using, 0 if not used
socket As Integer

'time they connected
connected_at As String

'remember when his last command was
idle_since As String


End Type

'this creates an array for each possible client
Public Client(server_max_clients) As client_type


Sub send_data(socket As Integer, data As String)
'use this to send data out to 1 socket.
'all of my server code will use this.
If data = "" Then Exit Sub
'socket = socket - 1
main1.sock(socket).SendData data
DoEvents

End Sub

Sub mass_send(data As String, exception_socket As Integer)
'this sends data out to EVERY client connected,
'except for the 'exception_socket' socket. leave
'exception_socket' as '0' if you want no exceptions.

'send data to every connected socket
Dim i As Integer
For i = 1 To (main1.sock.Count - 1)
If main1.sock(i).State = sckConnected And i <> exception_socket Then send_data i, data
Next i

End Sub

Sub send_data_to_clientid(clientid As Integer, data As String)
'use this to send data to a clientid, saves you having
'to find out their socket.

'simple, but saves time
send_data get_socket(clientid), data

End Sub



Public Sub update_status(message As String)
'this updates the status window with any new messages
main1.status.AddItem "[" & f_time & "] " & message
main1.status.ListIndex = (main1.status.ListCount - 1)
End Sub

