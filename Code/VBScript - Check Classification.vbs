Sub DocumentClassificationCheck_OnLoad
	
	DocClass = "Not Found"
	
	EKOManager.StatusMessage "Searching: " + OCRZone
	
	If instr(ucase(OCRZone),"CONFIDENTIAL") > 0 Then
		DocClass="CONFIDENTIAL"
	End If
	
	If instr(ucase(OCRZone),"RESTRICTED") > 0 Then
		DocClass="RESTRICTED"
	End If

	If instr(ucase(OCRZone),"INTERNAL USE") > 0 Then
		DocClass="INTERNAL"
	End If

	If instr(ucase(OCRZone),"PUBLIC") > 0 Then
		DocClass="PUBLIC"
	End If
	
	EKOManager.StatusMessage "Doc Class is: " + DocClass
	
	Set Topic = KnowledgeContent.GetTopicInterface
	If Topic Is Nothing Then
		KnowledgeObject.Status = 2
		EKOManager.ErrorMessage "Cannot retrieve Topic Interface"
		Exit Sub
	End If
	Call Topic.Replace("~USR::DocClass~", DocClass)
	
End Sub

Sub DocumentClassificationCheck_OnUnload

End Sub
