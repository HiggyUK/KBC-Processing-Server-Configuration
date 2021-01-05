Sub GetPhotoID_OnLoad

	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		EKOManager.StatusMessage "FileName: " + FileName
		EKOManager.StatusMessage "FilePath: " + FilePath
		
		PhotoFile = FilePath
		EKOManager.StatusMessage "Photo File Path: " & FilePath
		KnowledgeDocument.Visible = False

		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	
	Wend
		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::Photo~", PhotoFile
	End If
	
End Sub

Sub GetPhotoID_OnUnload

End Sub
