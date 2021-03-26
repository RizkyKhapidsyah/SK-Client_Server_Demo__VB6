Attribute VB_Name = "LogIN_Nt"
Option Compare Text 'lose cap sensitivity for this module
'in the engine i used some preset message's.
'this stores all the messages so its easyier to edit them.
'if you dont want it to send the message then just leave it
'blank and it wont send it.

'Server Full
Public Const message_1 = "Server Full"
Public Intmax As String
'On Connection Message
Public Const message_2 = "Server says : Hello Everbody!(adahmed911@hotmail.com)"


Public Sub decode_data(data As String, socket As Integer)
'a socket has sent some data to the server, write your code
'to translate the data here..

'first update the idle information
Client(get_clientid(socket)).idle_since = f_time

'now decode the data

'if mid(data,1,3) = "SAY" then... etc etc


End Sub

'this bas file contains the connection information which
'allows more than 1 user to join, and leave. also handles
'their accounts.

Public Sub new_connection(requestID As Long)
'new connection, so give them a socket

'socket for new user to have
Dim use_socket As Integer

'check if the server is full (with clients) or not
If live_connections >= max_clients Then disallow_connection requestID, message_1: Exit Sub

'search the loaded sockets to see if any are long
For i = 1 To (main1.sock.Count - 1)
If main1.sock(i).Tag = "0" Then
use_socket = i
GoTo found_sock
End If
Next i

'no sockets free so create a new socket
Dim socket_to_create As Integer
socket_to_create = main1.sock.Count
Load main1.sock(socket_to_create)
use_socket = socket_to_create


found_sock:

'log them in (if no socket found then act as if it were full)
If login_client(use_socket, requestID) = False Then disallow_connection requestID, message_1: Exit Sub
'update info
update_info

End Sub

Public Function login_client(socket As Integer, requestID As Long) As Boolean
'client connected, so now find him a clientid and setup
'his own account, returns if he managed to log in or not
Dim aa As String
For i = 1 To max_clients
If Client(i).socket = "0" Then
'found an empty client

'set client settings
Client(i).connected_at = f_time
Client(i).idle_since = f_time
Client(i).socket = socket

'tag the socket to remember the clientID
main1.sock(socket).Tag = i

'connect them on the chosen socket
main1.sock(socket).Close
main1.sock(socket).Accept requestID
'User logged in ok (show in status)
update_status "Client " & i & " Logged In (" & main1.sock(user_socket).RemoteHostIP & ")"

'recount live-connections
live_connections = live_connections + 1

'send welcome message
send_data socket, message_2

login_client = True
Exit Function
End If
Next i
'All sockets are in use, so return as false

End Function

Public Sub kickout_client(socket As Integer, notice As String)
'if you log them out and what them to know the reason.

send_data socket, notice
logout_client socket, notice


End Sub


Public Sub logout_client(socket As Integer, reason As String)
'client has disconnected, so close
'his socket, and blank out his clientid
'so sombody else can use it.
'the reason is simply their for status purposes.

'disconnect him
main1.sock(socket).Close

'clear his account (remember its the SOCKET, not clientID)
Client(main1.sock(socket).Tag).socket = "0"


'User logged out (show in status)
update_status "Client " & main1.sock(socket).Tag & " Logged Out (" & reason & ")"
'Unasign his socket
main1.sock(socket).Tag = "0"
main1.Intmax = main1.Intmax - 1
'recount live-connections
live_connections = live_connections - 1

'update info
update_info

End Sub

Public Sub disallow_connection(requestID As Long, reason As String)
'if you dont want sombody to be allowed to connect,
'instead of just not envoking the new_connection command
'envoke this as it lets them connect to a special socket,
'which'll then tell them the reason they cannot connect
'and then disconnect them from intself.
'ideal for 'server full' style messages

'User logged in ok (show in status)
update_status "Client Rejected (" & reason & ")"

'if no reason given, dont try to tell him it
If reason = "" Then Exit Sub


main1.disallow.Close
main1.disallow.Accept requestID
DoEvents

main1.disallow.SendData reason
DoEvents

main1.disallow.Close

End Sub


Public Function count_sockets() As Integer
'show the number of sockets loaded

count_sockets = main1.sock.Count

End Function


Public Sub count_live_connections()
'recount the connections (not used anymore)


'count how many are connected
Dim Temp As Integer

For i = 1 To (main1.sock.Count - 1)
If main1.sock(i).State <> scklong Then Temp = Temp + 1
Next i

'set it
live_connections = Temp

End Sub


