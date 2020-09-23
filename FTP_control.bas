Attribute VB_Name = "FTP_control"
Option Explicit

''''''''''''''''''''''''''''
'HANDLE FTP COMMANDS FROM CLIENT
''''''''''''''''''''''''''''

Public Function handle_commands(socket As Integer, raw_data As String)

    Dim data
    Dim ftpCommand As String
    Dim ftpArgs As String

    data = Replace(raw_data, vbCrLf, "")    'Remove carriage return and line feed.

    If InStr(data, " ") = 0 Then
        ftpCommand = data 'Command has no arguments...
    Else
        ftpCommand = Left(data, (InStr(data, " ") - 1))    'Get the command.
        ftpArgs = Right(data, (Len(data) - InStr(data, " ")))    'Get the command arguments.
    End If

    'This is where we will response do the clients command
    'Depending on the request, if its simple I will just put the
    'code here in the select case, otherwise, it will be in a sub below.

    'Hrmmm, I dont think I like all these If Then statments...Ill consolodate them into the functions
    'instead of putting them here later.  Always puttin things off  :)

    Select Case UCase(ftpCommand)

        Case "USER" 'Client sends user name
            'For right now the server will accept anything until
            'user control is implemented.
            'Im planning on implementing NT authentication hopfully.
            'If you would like a shot at that yourself you can get the docs
            'for it here: http://msdn.microsoft.com/library/psdk/winbase/accclsrv_9cfm.htm
            client(socket).userName = ftpArgs
            send_response socket, "331 User name ok, need password."

        Case "PASS" 'Client sends password
            'For right now the server will accept anything until
            'user control is implemented.
            client(socket).userPassword = ftpArgs
            If login_client(socket) = 1 Then
                send_response socket, "230 User logged in, proceed."
                update_server_log "User " & client(socket).userName & " logged in." 'Update log
            Else
                send_response socket, "530 Not logged in."
            End If

        Case "TYPE" 'Is usually TYPE I (IMAGE) or TYPE A (ASCII)
            send_response socket, "200 Type set to " & ftpArgs

        Case "REST"
            '"The argument field represents the server marker at which
            'file transfer is to be restarted.  This command does not
            'cause file transfer but skips over the file to the specified
            'data checkpoint.  This command shall be immediately followed
            'by the appropriate FTP service command which shall cause
            'file transfer to resume." - RFC 959

            If ftpArgs > 0 Then
                send_response socket, "504 Resuming is currently unsupported"
            Else
                send_response socket, "350 Restarting at 0"
            End If

        Case "PWD" 'Print Working Directory
            'When sending the current working directory to the client, it must be contained
            'in double quotes or it will be ignored.
            send_response socket, "257 " & Chr(34) & client(socket).currentDir & Chr(34) & " is current directory."

        Case "CWD"  'Change Working Directory
            If changeDirectory(socket, ftpArgs) = 1 Then
                send_response socket, "250 Directory changed to " & client(socket).currentDir
            Else
                send_response socket, "550 " & ftpArgs & ": No such file or directory"
            End If

        Case "CDUP" 'Go up one directory
            If CDUP(socket) = 1 Then
                send_response socket, "250 Directory changed to " & client(socket).currentDir
            Else
                send_response socket, "550 " & ftpArgs & ": No such file or directory"
            End If

        Case "PORT" 'The data port specified by the client for the server to connect to.
            Dim tmpArray() As String 'Six slots required (0 - 5)
            tmpArray = Split(ftpArgs, ",")
            client(socket).remotePort = tmpArray(4) * 256 Or tmpArray(5)
            send_response socket, "200 Port command successful."

        Case "LIST" 'Client asks for a directory listing of the current directory.
            'Tell the client the data is coming.
            send_response socket, "150 Opening ASCII mode data connection for /bin/ls."
            'Send the data.
            send_data socket, getDirectoryList(socket)
            'Tell the client you are done sending data.
            send_response socket, "226 Transfer complete."

        Case "STOR" 'Client sends a file.
            send_response socket, "150 Opening BINARY mode data connection for " & ftpArgs
            open_file socket, ftpArgs

        Case "RETR" 'Client asks server to send a file.
            If send_file(socket, ftpArgs) = 1 Then
                send_response socket, "550 " & ftpArgs & ": No such file or directory"
            End If

        Case "NOOP" 'Dumb little command that keeps the client from timing out due to inactivity.
            send_response socket, "200 Command ok."

        Case "DELE" 'Client asks to delete a file.  Becareful with this as there are no user permission in the server yet.
            If delete_file(socket, ftpArgs) = 1 Then
                send_response socket, "250 DELE command successful."
            Else
                send_response socket, "550 Permission denied."
            End If

        Case "RMD"  'Client asks to delete a directory.
            If delete_directory(socket, ftpArgs) = 1 Then
                send_response socket, "250 RMD command successful."
            Else
                send_response socket, "550 Delete directory failed."
            End If

        Case "MKD"  'Client asks to create a directory.
            If make_directory(socket, ftpArgs) = 1 Then
                send_response socket, "257 " & ftpArgs & " directory created."
            Else
                send_response socket, "550 Failed to create " & ftpArgs
            End If

    End Select

End Function
