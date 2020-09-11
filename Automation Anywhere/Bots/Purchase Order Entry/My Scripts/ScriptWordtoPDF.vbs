Set oShell = CreateObject("WScript.Shell")
'strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
sFolder = Wscript.Arguments.Item(0) 
'strHomeFolder & "\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\MasterCard_Chargeback_EPX\Output\Success\03388805\"
'msgbox sFolder 

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWord = CreateObject("Word.Application")
oWord.Visible = True

Set oFolder = oFSO.GetFolder(sFolder)
ConvertFolder(oFolder)
oWord.Quit

Sub ConvertFolder(oFldr)
  For Each oFile In oFldr.Files
    If LCase(oFSO.GetExtensionName(oFile.Name)) = "doc" OR LCase(oFSO.GetExtensionName(oFile.Name)) = "docx"Then
        Set oDoc = oWord.Documents.Open(oFile.path)
        'msgbox "path" & oFile.path
        Str = left(oFile,instr(1,oFile,".")-1) 
        oWord.ActiveDocument.SaveAs Str & ".pdf", 17
        oDoc.Close
    End If
   Next
End Sub