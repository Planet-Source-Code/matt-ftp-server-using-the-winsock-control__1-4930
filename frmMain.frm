VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "VB FTP Server ver .1a"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock dataSock 
      Index           =   0
      Left            =   8880
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock commandSock 
      Index           =   0
      Left            =   8280
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.ListBox lstServerLog 
         Height          =   1620
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   5415
      End
      Begin VB.ListBox lstConnectedClients 
         Height          =   1620
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblServerLog 
         Caption         =   "Server Log"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblConnectedClients 
         Caption         =   "Connected Clients"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROGRAM:       VB FTP Server Version .1a
'AUTHOR:        Matt Thomas
'CONTACT:       mthomas@aspire.com
'LAST UPDATED:  12/14/99
'FEATURES:      Navigate directories
'               Upload/download files
'               Create/delete directorys
'               Create/delete files.
'NOTE:          If you add any features or make improvments on my existing code
'               please let me know so I can update my code as well...
'               I'd like this project to kind of be a calaboritive effort towards a
'               decent ftp server.

Option Explicit

Private Sub Form_Load()

    start_server    'Starts the server and stuff...

End Sub

''''''''''''''''''''''''''''
'WINSOCK CONNECTION SUBS
''''''''''''''''''''''''''''

Private Sub commandSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Index = 0 Then

        update_server_log "Connection requested..."
        new_connection requestID
        DoEvents

    End If

End Sub

''''''''''''''''''''''''''''
'WINSOCK DATA ARRIVAL
''''''''''''''''''''''''''''

Private Sub commandSock_DataArrival(Index As Integer, ByVal totalBytes As Long)

    Dim raw_data As String
    commandSock(Index).GetData raw_data 'Get the data the client has sent.

    handle_commands Index, raw_data    'Now do something with it.

End Sub

Private Sub dataSock_DataArrival(Index As Integer, ByVal totalBytes As Long)

    Dim data As String

    dataSock(Index).GetData data
    Put client(Index).fFile, , data

End Sub

Private Sub dataSock_SendComplete(Index As Integer)

    If client(Index).transferTotalBytes > 0 Then

        If client(Index).transferTotalBytes = client(Index).transferBytesSent Then
            dataSock(Index).Close
            Close #client(Index).fFile
            send_response Index, "226 Transfer complete."
            client(Index).transferBytesSent = 0
            client(Index).transferTotalBytes = 0
        Else
            send_fileData Index
        End If
    End If

End Sub

Private Sub dataSock_Close(Index As Integer)

    'Once the file is sent the client will close the data connection.
    'When the client closes the connect this sub will be fired.

    Close #client(Index).fFile  'Close the file.
    dataSock(Index).Close

    'Tell the client you have successuflly received the file.
    send_response Index, "226 Transfer complete."

End Sub

Private Sub commandSock_Close(Index As Integer)

    'Logout the client
    logout_client Index

End Sub

''''''''''''''''''''''''''''
'WINSOCK ERROR SUBS
''''''''''''''''''''''''''''

Private Sub commandSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    update_server_log "ERROR - commandSock: " & Description & " (" & Number & ")"

End Sub

Private Sub dataSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    update_server_log "ERROR - dataSock: " & Description & " (" & Number & ")"

End Sub

Private Sub errorSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    update_server_log "ERROR - errorSock: " & Description & " (" & Number & ")"

End Sub

