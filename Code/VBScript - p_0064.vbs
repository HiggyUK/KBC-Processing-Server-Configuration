Sub GetPhotos_OnLoad
	
		
	' Get the Signature Images and Remove from KBO
	photos = ""
	EKOManager.StatusMessage "Starting File Count: " & FileCount 
	
	
	Set KnowledgeDocument = KnowledgeObject.GetFirstDocument()	
	While Not(KnowledgeDocument Is Nothing)

		FileName = KnowledgeDocument.GetFileName
		FilePath = KnowledgeDocument.FilePath
		
		If ucase(left(FileName,20)) = "COLLECTIONSIGNATURE_" Then
			CollectionSignatureFile = FilePath
			EKOManager.StatusMessage "New Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
			FileCount = FileCount - 1
		end if
		If ucase(left(FileName,16)) = "DRIVERSIGNATURE_" Then
			DriverSignatureFile = FilePath
			EKOManager.StatusMessage "New Signature File Path: " & FilePath 
			KnowledgeDocument.Visible = False
			FileCount = FileCount - 1
		end if
		
		If  ucase(left(FileName,16)) <> "DRIVERSIGNATURE_" And ucase(left(FileName,20)) <> "COLLECTIONSIGNATURE_"  Then
			photos = photos  + FilePath + "|"
			KnowledgeDocument.Visible = False
		
		End If
		
		Set KnowledgeDocument = KnowledgeObject.GetNextDocument
	Wend

	'photos = right(photos,len(photos)-1)
	
	EKOManager.StatusMessage "File Count: " & FileCount 
		
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::CollectionSignature~", CollectionSignatureFile
		PTopic.Replace "~USR::DriverSignature~", DriverSignatureFile
		PTopic.Replace "~USR::Photos~", photos
		PTopic.Replace "~USR::PhotoCount~", FileCount
	End If
	
End Sub

Sub GetSignatures_OnUnload

End Sub
