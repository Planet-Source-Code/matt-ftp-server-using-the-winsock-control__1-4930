Attribute VB_Name = "client_control"
Option Explicit

Public Function login_client(socket As Integer)

    'Once user control is implemented, put code for logging in here.  duh.
    client(socket).homeDir = "D:\"  'Users homedir local path
    client(socket).currentLocalDir = client(socket).homeDir
    client(socket).currentDir = "/"

    'Remember where in the list they were added.
    client(socket).listIndex = frmMain.lstConnectedClients.ListCount

    'Add user to the list if connected clients
    frmMain.lstConnectedClients.AddItem client(socket).userName

    client(socket).IPAddress = frmMain.commandSock(socket).RemoteHostIP    'Save clients IP Address.

    login_client = 1 'Send 1 to show login was successful.

End Function

Public Sub logout_client(socket As Integer)

    If client(socket).userName = "" Then Exit Sub

    'Remove the user name from the clients connected list.
    frmMain.lstConnectedClients.RemoveItem client(socket).listIndex

    'Show the user logged out.
    update_server_log "User " & client(socket).userName & " logged out."

    'Clear out thier info from the client array
    With client(socket)
        .connectedAt = ""
        .currentDir = ""
        .currentLocalDir = ""
        .fFile = 0
        .homeDir = ""
        .idleSince = ""
        .IPAddress = ""
        .listIndex = 0
        .remotePort = 0
        .userName = ""
        .userPassword = ""
    End With

    'Unload the winsock instances that were created for this user to save memory.
    'Unload frmMain.commandSock(socket)
    'Unload frmMain.dataSock(socket)

End Sub
