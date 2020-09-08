Const ForReading = 1
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("D:\7.5_Robot_Test_Automation_Sets\7.5_CIT\ConfigFile\GlobalConfig.py", ForReading)
Do Until objFile.AtEndOfStream
    strLine = objFile.Readline
    strLine = Trim(strLine)
    If Len(strLine) > 0 Then
        strNewContents = strNewContents & strLine & vbCrLf
    End If
Loop
objFile.Close
Set objFile = objFSO.OpenTextFile("D:\7.5_Robot_Test_Automation_Sets\7.5_CIT\ConfigFile\GlobalConfig.py", ForWriting)
objFile.Write strNewContents
objFile.Close