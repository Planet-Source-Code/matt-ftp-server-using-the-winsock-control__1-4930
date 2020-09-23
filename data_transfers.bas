Attribute VB_Name = "data_transfers"
Option Explicit

''''''''''''''''''''''''''''
'SEND A MESSAGE TO THE CLIENT
''''''''''''''''''''''''''''

Public Sub send_response(socket As Integer, message As String)

    'With each message sent to the client you MUST be sure to include
    'a trailing vbCrLF or the client will ignore it.
    frmMain.commandSock(socket).SendData message & vbCrLf 'Send message.
    DoEvents

End Sub

''''''''''''''''''''''''''''
'DATA CONNECTION
''''''''''''''''''''''''''''

Public Sub make_connection(socket As Integer)

    With frmMain.dataSock(socket)
        .Close  'Just to be sure.
        .RemoteHost = client(socket).IPAddress   'Address to connect to.
        .remotePort = client(socket).remotePort 'Port the client is listening on.
        .Connect
    End With

    Do
        DoEvents
    Loop Until frmMain.dataSock(socket).State = sckConnected    'Make sure a connection is made.

End Sub

''''''''''''''''''''''''''''
'SEND DATA TO THE CLIENT
''''''''''''''''''''''''''''

Public Sub send_data(socket As Integer, raw_data As String)

    If raw_data = "" Then Exit Sub  'Dont send empty packets
    
    frmMain.dataSock(socket).SendData raw_data  'Send data.
    DoEvents

End Sub
