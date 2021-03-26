Attribute VB_Name = "Misc_Nt"

Public Function get_socket(clientid As Integer) As Integer
'returns the socket which the specified clientid is using

get_socket = Client(clientid).socket

End Function

Public Function get_clientid(socket As Integer) As Integer
'returns the clientid of the client using the specified socket

get_clientid = main1.sock(socket).Tag

End Function

Public Function f_time() As String
'returns time in a nice format

f_time = Format(Time, "hh:mm:ss")

End Function


Public Sub update_info()

'updates the information on the main1 form
main1.sockets_loaded_info.Caption = count_sockets
main1.live_connections_info.Caption = live_connections
'If live_connections = 0 Then
'Do nothing
' Exit Sub
' End If
'main1.List1.AddItem live_connections

End Sub


Public Sub close_all_sockets()
'close down every socket
'(not designed for restart, deseigned for when sombody closes the program)

Dim i As Integer
For i = 0 To (count_sockets - 1)
main1.sock(i).Close
Next i

'show its been shutdown.
update_status "*** Server ShutDown ***"



End Sub

Public Sub reset_server()
'this totally resets the server.

'show its reset in status
update_status "*** Server Reset ***"

'turn off the main1 connection socket
main1.sock(0).Close

'disconnect all the users
For i = 1 To max_clients
If Client(i).socket <> 0 Then logout_client Client(i).socket, "Server Reset"
Next i

'start up the main1 socket again
main1.sock(0).Listen

End Sub

Public Sub start_server()
On Error GoTo ec
'this just starts the main1 connection socket up to listen

'load settings
set_up_settings
'Dim aa As String
'aa = main1.sock(0).RemoteHost

main1.sock(0).LocalPort = server_port
main1.sock(0).Listen

'show its started in the status
update_status "*** Server Started *** (" & main1.sock(0).LocalIP & ":" & server_port & ")"

Exit Sub
ec:
MsgBox "Unable To Start Server - Port In Use", vbExclamation + vbOKOnly, "Error Starting Server"

End Sub

Public Sub set_up_settings()
'this simply sets up all the settings

'set the maxmimum number of clients
max_clients = default_max_clients
server_port = default_server_port

End Sub
Public Sub change_server_port(port As Long)
On Error GoTo ec


main1.sock(0).Close
main1.sock(0).LocalPort = port
main1.sock(0).Listen
server_port = info_port
update_status "*** Server Port Changed To - " & server_port

'error controll
Exit Sub
ec:
MsgBox "Error Changing Server Port."
update_status "*** Error Forced Server Port To Remain1 At - " & server_port
main1.sock(0).Close
main1.sock(0).LocalPort = server_port
main1.sock(0).Listen

End Sub

