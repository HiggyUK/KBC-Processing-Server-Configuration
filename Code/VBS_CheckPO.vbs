Sub CheckPO_OnLoad

	ErrorMessge = "Purhase Order Acceptable"
	Route = "Accepted"
	Status = "Accepted"
	
	
	' Check the Signature
	If PO_Signed = "FALSE" Or PO_Signed = "" Then
		ErrorMessage = "Purchase Order Not Signed"
		Route = "Error"
		Status = "Error"
	End If
	
	
	
	Set Topic = KnowledgeContent.GetTopicInterface
	If Topic Is Nothing Then
		KnowledgeObject.Status = 2
		EKOManager.ErrorMessage "Cannot retrieve Topic Interface"
		Exit Sub
	End If
	Call Topic.Replace("~USR::Message~", ErrorMessage)
	Call Topic.Replace("~USR::Route~", Route)
	Call Topic.Replace("~USR::Status~",Status)
	EKOManager.StatusMessage ("Status: " & ErrorMessage)

End Sub

Sub CheckPO_OnUnload

End Sub
