Attribute VB_Name = "client_settings"
Option Explicit

''''''''''''''''''''''''''''
'CLIENT SETTINGS
''''''''''''''''''''''''''''
Type FTPclient

    userName As String          'User name
    userPassword As String      'User password
    IPAddress As String            'Users IP Address, just because...
    connectedAt As String       'Time at which they connected.
    idleSince As String         'Time at which the last command was received.
    listIndex As Integer        'Reference number to where they are listed.
    homeDir As String           'Users home directory, directory they will start in.
    currentDir As String        'The directory the client is currently in.
    currentLocalDir As String   'Current local path.
    remotePort As Integer       'The data port specified by the client for the server to connect to.
    fFile As Long               'Holds the file number that the client is using.
    currentFile As String       'Name and path of file being transfered
    transferTotalBytes As Long  'Holds the length of the file currently being transfered
    transferBytesSent As Long   'How many bytes of the current file have been sent

End Type

'Client information array
Public client(Server_Max_Clients) As FTPclient
