Sub CheckUserLevel_OnLoad

	sendAllowed = "rejected"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile("C:\Nuance\MobileDemo\Settings\Security\DocClass\"+ userName + "\" + DocClass +".txt", 1)
	Do Until (objTextFile.AtEndOfStream)
		the_line = objTextFile.Readline
		arr=Split(the_line, "=")
		sendType = arr(0)
		sendAllowed = arr(1)
		If sendType = requestType Then
			requestAllowed = sendAllowed
		End If
	Loop
	objTextFile.Close
	
	If requestType = "Folder" And requestAllowed = "Allow" Then
		routeNext = "p_0022"
	End If
	If requestType = "Email" And requestAllowed = "Allow" Then
		routeNext = "p_0027"
	End If
	If requestType = "Folder" And requestAllowed = "Reject" Then
		routeNext = "p_0026"
	End If
	If requestType = "Email" And requestAllowed = "Reject" Then
		routeNext = "p_0026"
	End If
	
	
	Set Topic = KnowledgeContent.GetTopicInterface
	If Topic Is Nothing Then
		KnowledgeObject.Status = 2
		EKOManager.ErrorMessage "Cannot retrieve Topic Interface"
		Exit Sub
	End If
	
	Call Topic.Replace("~USR::SendType~", requestType)
	Call Topic.Replace("~USR::RequestAllowed~", requestAllowed)
	Call Topic.Replace("~USR::RouteNext~",routeNext)
	
	EKOManager.StatusMessage ("Send Type " & requestType)
	EKOManager.StatusMessage ("Request Allowed " & requestAllowed)
	EKOManager.StatusMessage ("Route Next " & routeNext)

End Sub

Sub CheckUserLevel_OnUnload

End Sub


