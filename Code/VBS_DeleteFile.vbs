Sub DeleteFile_OnLoad

	
	EKOManager.StatusMessage("Deleting File: " & fileName)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile("C:\Nuance\CompleteDemo\Settings\ToApprove\"  & UserName & "\" & fileName & ".txt")


	EKOManager.StatusMessage("Adding Form File: " & fileName)

	Set OutputKDocument = KnowledgeObject.AddDocument    
	If Not(OutputKDocument Is Nothing) Then
		OutputKDocument.FilePath = "C:\Nuance\CompleteDemo\Settings\ToApprove\" & UserName & "\"  & fileName & ".PDF"
		Set OutputKOptions = OutputKDocument.AddOption
		OutputKOptions.OnFailure = 2 	'FILEOPTION_DELETE
		OutputKOptions.OnSuccess = 2	'FILEOPTION_DELETE
	End If
End Sub

Sub DeleteFile_OnUnload

End Sub
