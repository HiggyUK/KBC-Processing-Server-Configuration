Sub GetPhotos_OnLoad
	
		
	' Get the Signature Images and Remove from KBO
	EKOManager.StatusMessage "Starting File Count: " & FileCount 
	
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		If ucase(left(FileName,10)) = "SIGNATURE_" Then
			SignatureFile = FilePath
			EKOManager.StatusMessage "New Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
			FileCount = FileCount - 1
		end if
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend

		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::Signature~", SignatureFile
	End If
	
End Sub

Sub GetSignatures_OnUnload

End Sub
