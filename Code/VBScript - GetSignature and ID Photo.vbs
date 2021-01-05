Sub GetSignature_OnLoad
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		EKOManager.StatusMessage "FileName: " + FileName
		EKOManager.StatusMessage "FilePath: " + FilePath
		
		If ucase(Left(FileName,10)) = "SIGNATURE_" Then
			SignatureFile = FilePath
			EKOManager.StatusMessage "Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
		Else
			PhotoFile = FilePath
			EKOManager.StatusMessage "Photo File Path: " & FilePath
			KnowledgeDocument.Visible = False
		End If		
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend
		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::Signature~", SignatureFile
		PTopic.Replace "~USR::Photo~", PhotoFile
	End If
End Sub

Sub GetSignature_OnUnload

End Sub
