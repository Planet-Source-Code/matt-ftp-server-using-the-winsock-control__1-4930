Attribute VB_Name = "file_control"
Option Explicit

''''''''''''''''''''''''''''
'GET DIRECTORY LISTING
''''''''''''''''''''''''''''

Public Function getDirectoryList(socket As Integer)

    make_connection socket  'Make the data connection
    DoEvents

    Dim fso As New Scripting.FileSystemObject
    Dim f, f1, fc

    'This method of getting directory and file properties is pretty damn inefficient.
    'If you have a directory with many files there is a pretty noticable delay
    'as it loops through each file.  Eats too much CPU doing it...
    'If anyone can find a more efficient way of doing this please let me know.

    'Get list of directorys
    Set f = fso.GetFolder(client(socket).currentLocalDir).SubFolders

    For Each f1 In f
        'Send one directory at a time.
        send_data socket, "drwx------ 1 user group 0 " & Format(f1.DateLastModified, " mmm dd hh:mm ") & f1.Name & vbCrLf
    Next

    Set f = Nothing

    'Get list of files.
    Set f = fso.GetFolder(client(socket).currentLocalDir)
    Set fc = f.Files
    For Each f1 In fc
        'Send one file at a time.
        send_data socket, "-rwx------ 1 user group " & fso.GetFile(client(socket).currentLocalDir & f1.Name).Size & Format(fso.GetFile(client(socket).currentLocalDir & f1.Name).DateLastModified, " mmm dd hh:mm ") & f1.Name & vbCrLf
    Next

    Set f = Nothing
    Set fc = Nothing
    Set fso = Nothing

    frmMain.dataSock(socket).Close  'Clost the data connection when transfer is complete.

End Function

''''''''''''''''''''''''''''
'CHANGE WORKING DIRECTORY
''''''''''''''''''''''''''''

Public Function changeDirectory(socket As Integer, dir As String)

    Dim fso As New Scripting.FileSystemObject

    If Left(dir, 1) = "/" Then 'Client specified a specific dir starting at the root.
        Dim LocalPath As String
        
        'Convert client path to local path
        LocalPath = Right(dir, (Len(dir) - 1)) 'Strip off that "/" at the beginning.
        LocalPath = Replace(LocalPath, "/", "\")    'Change / to \ for local compatibility

        If fso.FolderExists(client(socket).homeDir & LocalPath & "\") = True Then
            client(socket).currentLocalDir = client(socket).homeDir & LocalPath & "\"
        
            'Path the client sees.
            client(socket).currentDir = dir

            changeDirectory = 1 'Directory found.
        Else
            changeDirectory = 0 'Directory doesnt exist
        End If
        
    Else    'Change to one of the subdirectories in the current dir.

        If fso.FolderExists(client(socket).currentLocalDir & dir & "\") = True Then
            Dim Exception As String
    
            'This is to prevent a path like "//Blah"
            'This really only effect dirs in the root.
            If client(socket).currentDir = "/" Then
                Exception = ""
            Else
                Exception = "/"
            End If
    
            'Easy enough eh?
            client(socket).currentDir = client(socket).currentDir & Exception & dir
            client(socket).currentLocalDir = client(socket).currentLocalDir & dir & "\"

            changeDirectory = 1 'Directory found.
        Else
            changeDirectory = 0 'Directory doesnt exist
        End If
    End If

    Set fso = Nothing

End Function

''''''''''''''''''''''''''''
'GO UP ONE DIRECTORY
''''''''''''''''''''''''''''

Public Function CDUP(socket As Integer)

    If client(socket).currentLocalDir = client(socket).homeDir Then 'Cant go up anymore.
        CDUP = 1
        Exit Function
    End If

    Dim fso As New Scripting.FileSystemObject

    'If parent directory exists
    If fso.FolderExists(fso.GetFolder(client(socket).currentLocalDir).ParentFolder) = True Then
        Dim path As String

        If fso.GetFolder(client(socket).currentLocalDir).ParentFolder = client(socket).homeDir Then
            client(socket).currentLocalDir = fso.GetFolder(client(socket).currentLocalDir).ParentFolder
            path = client(socket).currentLocalDir
        Else
            client(socket).currentLocalDir = fso.GetFolder(client(socket).currentLocalDir).ParentFolder & "\"
            path = Left(client(socket).currentLocalDir, (Len(client(socket).currentLocalDir) - 1))
        End If

        'Convert local path to client friendly path.
        path = Replace(path, "D:\", "/")
        path = Replace(path, "\", "/")

        client(socket).currentDir = path
        CDUP = 1
    Else
        CDUP = 0    'Parent directory doesnt exist.  Uh oh.
    End If

End Function

''''''''''''''''''''''''''''
'OPEN FILE TO SAVE
''''''''''''''''''''''''''''

Public Sub open_file(socket As Integer, fileName As String)

    client(socket).fFile = FreeFile

    make_connection socket  'connect to client

    'Open the file and wait for data to be sent.
    Open (client(socket).currentLocalDir & fileName) For Binary Access Write As #client(socket).fFile
    DoEvents

End Sub

''''''''''''''''''''''''''''
'OPEN FILE TO SAVE
''''''''''''''''''''''''''''

Public Function send_file(socket As Integer, fileName As String)

    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(client(socket).currentLocalDir & fileName) = False Then
        send_file = 1   'File doesnt exsist.
        Exit Function
    End If

    Set fso = Nothing

    send_response socket, "150 Opening BINARY mode data connection for " & fileName

    make_connection socket  'connect to clients data port

    client(socket).currentFile = client(socket).currentLocalDir & fileName   'Just to save some typing.

    client(socket).transferTotalBytes = FileLen(client(socket).currentFile)

    client(socket).fFile = FreeFile

    Open (client(socket).currentFile) For Binary Access Read As #client(socket).fFile 'Open the file to send
    DoEvents

    send_fileData socket

End Function

Public Sub send_fileData(socket As Integer)

    Dim data As String
    Dim BlockSize As Integer
        BlockSize = 1024    '1K blocks.

    If BlockSize > (client(socket).transferTotalBytes - client(socket).transferBytesSent) Then
        BlockSize = (client(socket).transferTotalBytes - client(socket).transferBytesSent)
    End If

    data = Space$(BlockSize) 'allocate space to store data.
    Get client(socket).fFile, , data    'get data

    client(socket).transferBytesSent = client(socket).transferBytesSent + BlockSize

    frmMain.dataSock(socket).SendData data

End Sub

Public Function delete_file(socket As Integer, fileName As String)

    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(client(socket).currentLocalDir & fileName) = True Then
        'File exists, delete it.
        fso.GetFile(client(socket).currentLocalDir & fileName).Delete
        delete_file = 1
    Else
        delete_file = 0
    End If

    Set fso = Nothing

End Function

Public Function delete_directory(socket As Integer, dirName As String)

    Dim fso As New Scripting.FileSystemObject

    If fso.FolderExists(client(socket).currentLocalDir & dirName) = True Then
        'Folder exists, delete it.
        fso.DeleteFolder (client(socket).currentLocalDir & dirName)
        delete_directory = 1
    Else
        delete_directory = 0
    End If

    Set fso = Nothing

End Function

Public Function make_directory(socket As Integer, dirName As String)

    Dim fso As New Scripting.FileSystemObject

    If fso.FolderExists(client(socket).currentLocalDir & dirName) = True Then
        'Directory already exists, cant create.
        make_directory = 0
    Else
        fso.CreateFolder (client(socket).currentLocalDir & dirName)
        make_directory = 1
    End If

    Set fso = Nothing

End Function
