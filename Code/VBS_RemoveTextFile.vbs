Sub RemoveTextFile_OnLoad
	
	EKOManager.StatusMessage "Removing the Text File"
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()
	While Not(KnowledgeDocument Is Nothing)
		
		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		If ucase(right(FileName,3)) = "TXT" Then
			
			KnowledgeDocument.Visible = False
		End If
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
		
	Wend

End Sub

Sub RemoveTextFile_OnUnload

End Sub
