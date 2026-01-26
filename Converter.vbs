Option Explicit

Dim objArgs, objFSO 'Setup the variables the whole code accesses

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments

Dim Count, Exports
Count = 0
Exports = 0

Function Determine_Type(ThisArg)
    'Figures out what kind of object its reading, A Folder, A file, and if it's a file, determine if it's a valid format to export with.
    Dim sourcePath, ext
    sourcePath = objFSO.GetAbsolutePathName(ThisArg)
    If objFSO.FolderExists(sourcePath) Then
        Determine_Type = "FOLDER"
        Exit Function
    ElseIf objFSO.FileExists(sourcePath) Then
        ext = LCase(objFSO.GetExtensionName(ThisArg))
        If (ext <> "pptx") And (ext <> "ppt") And (ext <> "ppsx") And (ext <> "pdf") Then
            Determine_Type = "N/A"
            Exit Function
        ElseIf ext = "pdf" Then
            Determine_Type = "IGNORE"
            Exit Function
        End If
        Determine_Type = "PowerPoint" 'There is a fairly significant oversight here, without the above "Exit Function"s, all files are still treated as if they're pdfs
    End If
    Exit Function
End Function

Dim flaggedFiles 'List of files flagged to be moved later on

Sub ConvertPowerPoint(ThisArg)
    ' Takes the file path, 
    Dim sourcePath, destPath
    If Determine_Type(ThisArg) <> "PowerPoint" Then Exit Sub
    sourcePath = objFSO.GetAbsolutePathName(ThisArg)
    destPath = objFSO.GetParentFolderName(sourcePath) & "\" & _
            objFSO.GetBaseName(sourcePath) & ".pdf"

    Dim objPPT, objPres
    On Error Resume Next
    Set objPPT = CreateObject("PowerPoint.Application")

    Set objPres = objPPT.Presentations.Open(sourcePath, -1, 0, -1)

    If Err.Number <> 0 Then
        MsgBox "Error opening file: " & sourcePath & " | " & Err.Description, 16, "Debug Info"
        objPPT.Close
        objPPT.Quit 'Make sure to close the applicaiton, otherwise it'll hang for minutes.
        WScript.Quit
    End If

    'The following code is by far the most iritating thing on earth and god knows I don't understand it.
    objPres.ExportAsFixedFormat _
        destPath, _
        CLng(2), _
        CLng(0), _
        CLng(-1), _
        CLng(1), _
        CLng(2), _
        CLng(0), _
        Nothing

    If Err.Number <> 0 Then
        MsgBox "Export error: " & Err.Number & " - " & Err.Description, 16, "Debug Info"
        'An error he
    End If

    objPres.Close
    objPPT.Quit

    Exports = Exports + 1

    Set objPres = Nothing
    Set objPPT = Nothing
End Sub

Sub Tidy_Folder(Arg)
    'Ensure we are only considering a directory
    If not Determine_Type(Arg) = "FOLDER" Then
        Exit Sub
    End If
    dim TargetDirectory
    'Find a subdirectory called "PowerPoint", If no subdirectory exists, create one
    if Not objFSO.FolderExists(Arg&"\PowerPoint") Then
        objFSO.CreateFolder Arg&"\PowerPoint"
    end if
    TargetDirectory = objFSO.GetFolder(Arg&"\PowerPoint")
    'Iterate through the flaggedFiles and move them into the new destination
    Dim f_file
    for Each f_file in flaggedFiles
        'Move directory
        objFSO.MoveFile objFSO.GetAbsolutePathName(f_file), TargetDirectory&"\"&objFSO.GetFileName(f_file)
    Next
End Sub 

' vba arrays are of fixed length, we are creating a new +1 length array for every new element, because we cannot natively increase it's size.
Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

' Initialize our array
flaggedFiles = Array()

Sub Next_Step(Arg)
    Dim FolderArg, Result
    Result = Determine_Type(Arg)
    If Result = "PowerPoint" Then
        Count = Count + 1
        ConvertPowerPoint Arg
        flaggedFiles = AddItem(flaggedFiles, Arg) 'Tag the appropriate folder
    ElseIf Result = "FOLDER" Then
        For Each FolderArg in objFSO.GetFolder(Arg).Files
            Call Next_Step(FolderArg)
        Next
    ElseIf Result = "N/A" Then
        Count = Count + 1
        MsgBox "The given file isn't compatible. Filename; " & Arg & vbNewLine & "Make sure the file has a valid extension: PPT, PPTX, PPSX", 32, "Converter"
    ElseIf Result = "IGNORE" Then
        Count = Count + 1
    End If
End Sub

Sub RunLoop(ThisArg)
    Dim Arg
    For Each Arg in ThisArg
        Next_Step Arg
        Tidy_Folder Arg
    Next
End Sub

RunLoop objArgs

if Exports = 0 Then
    MsgBox "No files were converted. Make sure there are valid PowerPoint files!", 32, "Converter"
ElseIf Count = 1 And Exports = 1 Then
    MsgBox "Your file has been converted", 32, "Converter"
ElseIf Exports >= 1 Then
    MsgBox "Succesfully exported " & Exports & " of " & Count & " Files.", 32, "Converter"
End If