Sub GetSignatures_OnLoad
	
	
	EKOManager.StatusMessage "Getting Signatures"
	
	' Get the Signature Images and Remove from KBO
	photos = ""
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		EKOManager.StatusMessage "Checking: " & FileName
		                               
		If ucase(left(FileName,18)) = "INJURYEDSIGNATURE_" Then
			CareworkerSignatureFile = FilePath
			EKOManager.StatusMessage "New Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
			FileCount = FileCount - 1
		end if

		If ucase(left(FileName,12)) = "COMPLETEDBY_" Then
			PatientSignatureFile = FilePath
			EKOManager.StatusMessage "New Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
			FileCount = FileCount - 1
		end if
		
		If  ucase(left(FileName,18)) <> "INJURYEDSIGNATURE_" And ucase(left(FileName,12)) <> "COMPLETEDBY_" Then
			photos = photos + FilePath + "|"
			KnowledgeDocument.Visible = False
		
		End If
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend
		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::InjuredSignature~", CareworkerSignatureFile
		PTopic.Replace "~USR::CompletedBySignature~", PatientSignatureFile
		PTopic.Replace "~USR::Photos~", photos
		PTopic.Replace "~USR::FileCount~", FileCount
	End If
	
End Sub

Sub GetSignatures_OnUnload

End Sub
