Sub GetPictureImage_OnLoad


	
	vsSelected = "C:\Nuance\MobileDemo\Settings\BookedVisits\" & UserName & "\" &  SelectedVisitor & ".txt"
	EKOManager.StatusMessage("Deleting File: " & vsSelected)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(vsSelected)
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		PictureFile = FilePath
		KnowledgeDocument.Visible = False
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend
		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::Picture~", PictureFile
	End If
	
End Sub

Sub GetPictureImage_OnUnload

End Sub
