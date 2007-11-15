Attribute VB_Name = "shFile"
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                 SHARED MODULE                       [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                                                     [[
'[[         File/Folder Utility Functions               [[
'[[                                                     [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Option Explicit

Public g_sFileLastError As String

Public Function PathExists(szPath As String) As Boolean
    PathExists = False
    On Error GoTo Error_Handler
    PathExists = (Dir(szPath, vbDirectory) <> "")
Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage "Error in PathExists(" & szPath & ")"
    Resume Fin
End Function

Public Function FileExists(szFile As String) As Boolean
    FileExists = False
    On Error GoTo LExit
    FileExists = (Dir(szFile) <> "")
LExit:
End Function

Public Sub CreateFolder(szFolderName As String)
On Error GoTo Error_Handler
   If PathExists(szFolderName) = False Then
        MkDir szFolderName
   End If
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "Error in CreateFolder(" & szFolderName & ")"
    Resume Fin
End Sub

'szSourceFolder and szDestFolder must NOT have a finishing backslash
' GOOD = C:\TEMP
' BAD = C:\TEMP\
Public Sub MirrorFolder(ByVal szSourceFolder As String, ByVal szDestFolder As String, Optional ByVal bOverWrite As Boolean = False)
On Error GoTo ErrorHandler

    Dim oFS As New Scripting.FileSystemObject
    Dim dirSource As Scripting.Folder
    Dim dirDest As Scripting.Folder
    Dim dirSub As Scripting.Folder
    Dim oFile As Scripting.File
    Dim sDestName As String
    
    If Not oFS.FolderExists(szSourceFolder) Then
        PrintErrorMessage "MirrorFolder : Source Folder doesnt exist - " & szSourceFolder
    Else
        'MyDebug "MirrorFolder: Source: " & szSourceFolder
        'MyDebug "MorrorFolder: Dest: " & szDestFolder
        
        Set dirSource = oFS.GetFolder(szSourceFolder)
        If Not oFS.FolderExists(szDestFolder) Then
            MyDebug "MirrorFolder - Creating new folder : " & szDestFolder
            Set dirDest = oFS.CreateFolder(szDestFolder)
        Else
            Set dirDest = oFS.GetFolder(szDestFolder)
        End If
        
        For Each oFile In dirSource.Files
            sDestName = dirDest.Path & "\" & oFile.Name
            If bOverWrite Or (Not oFS.FileExists(sDestName)) Then
                MyDebug "...Copied " & oFile.Path & " to " & sDestName
                Call oFS.CopyFile(oFile.Path, sDestName)
            End If
        Next oFile
        
        'Recursive call for subfolders
        For Each dirSub In dirSource.SubFolders
            Call MirrorFolder(dirSub.Path, dirDest.Path & "\" & dirSub.Name, bOverWrite)
        Next dirSub
        
    End If
    
Fin:
    Set oFS = Nothing
    Set dirSource = Nothing
    Set dirDest = Nothing
    Set dirSub = Nothing
    Set oFile = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "MirrorFolder(" & szSourceFolder & ", " & szDestFolder & " ) - " & Err.Description
    Resume Fin
End Sub


Public Sub CreateFile(szFilename As String)
On Error GoTo Error_Handler
    Dim fileNum As Integer

    fileNum = FreeFile(0)
    Open szFilename For Output As #fileNum
    Close #fileNum
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "Error in CreateFile(" & szFilename & ")"
    Resume Fin
End Sub
