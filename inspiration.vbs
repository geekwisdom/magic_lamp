' **************************************************************************************
' Script Name: inspiration.vbs
' **************************************************************************************
' @(#)    Purpose:
' @(#)    Improve your mood and thinking process !!
' @(#)    Show a motivational quote each time your unlock your workstation
' @(#)    
' **************************************************************************************
'  Written By: GeekWisdom.org / Brad D.
'
' Created:     2018-08-02
' **************************************************************************************

Function RandomLineFromFile(FileName)

Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOpen = objFSO.OpenTextFile(FileName, ForReading)
Dim arrFileName()

i = 0
Do While objOpen.AtEndOfStream <> true
    text = objOpen.ReadLine
        ReDim Preserve arrFileName(i)
    arrFileName(i) = text
    i=i+1
loop
objOpen.Close
Set objOpen = Nothing
Set objFSO = Nothing

min=0
max=UBound(arrFileName)-1

Randomize
Num=Int((max-min+1)*Rnd+min)
Result=arrFileName(Num)
RandomLineFromFile=Result
End Function
strPath=WScript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile)
Result=RandomLineFromFile(strFolder + ".\inspiration.txt")
If Result = "" Then Result = "Way to GO!, You can do this!"
Msgbox Result,0,"Thought of the Day"
