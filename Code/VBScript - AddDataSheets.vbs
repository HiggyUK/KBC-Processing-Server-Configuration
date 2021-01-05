Sub AddDataSheets_OnLoad
	
	EKOManager.StatusMessage "Adding DataSheets"
	products = split(productList,";")
	
	EQ = "No"
	EC = "No"
	SC = "No"
	OM = "No"
	BC = "No"
	OF = "No"
	AC = "No"
	AX = "No"
	WC = "No"
	QC = "No"
	AW = "No"
	
	For Each product In products

		Select Case product
			Case "1" : EQ = "Yes"
			Case "2" : EC = "Yes"
			Case "3" : SC = "Yes"
			Case "4" : OM = "Yes"
			Case "5" : BC = "Yes"
			Case "6" : OF = "Yes"
			Case "7" : AC = "Yes"
			Case "8" : AX = "Yes"
			Case "9" : WC = "Yes"
			Case "10" : QC = "Yes"
			Case "11" : AW = "Yes"
		End Select
		
		EKOManager.StatusMessage "Adding " + product

		Set NewKnowledgeDocument = KnowledgeObject.AddDocument
		If Not(NewKnowledgeDocument Is Nothing) Then
			NewKnowledgeDocument.FilePath = "C:\Nuance\MobileDemo\Settings\SalesInfo\" + product + ".pdf"
			NewKnowledgeDocument.Status = 1 'DOC_STATUS_OK
			Set KDocumentOptions = NewKnowledgeDocument.AddOption
			'If Not(KDocumentOptions Is Nothing) Then
			'	KDocumentOptions.OnSuccess = 2 'FILEOPTION_DELETE
			'	KDocumentOptions.OnFailure = 1 'FILEOPTION_MOVE
			'End If
		End If
	Next
	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::EQ~", EQ
		PTopic.Replace "~USR::EC~", EC
		PTopic.Replace "~USR::SC~", SC
		PTopic.Replace "~USR::OM~", OM
		PTopic.Replace "~USR::BC~", BC
		PTopic.Replace "~USR::OF~", OF
		PTopic.Replace "~USR::AC~", AC
		PTopic.Replace "~USR::AX~", AX
		PTopic.Replace "~USR::WC~", WC
		PTopic.Replace "~USR::QC~", QC
		PTopic.Replace "~USR::AW~", AW
	End If
	
End Sub

Sub AddDataSheets_OnUnload

End Sub
