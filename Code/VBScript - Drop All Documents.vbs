Sub DropAllDocuments_OnLoad

	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		KnowledgeDocument.Visible = False
			
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend
	
End Sub

Sub DropAllDocuments_OnUnload

End Sub
