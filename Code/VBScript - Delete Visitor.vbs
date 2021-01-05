Sub DeleteVisitor_OnLoad
	
	vsSelected = "C:\Nuance\MobileDemo\Settings\BookedVisits\" & UserName & "\" &  SelectedVisitor & ".txt"
	EKOManager.StatusMessage("Deleting File: " & vsSelected)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(vsSelected)
		
	
End Sub

Sub DeleteVisitor_OnUnload

End Sub
