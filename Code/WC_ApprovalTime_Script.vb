Option Strict Off

Imports System
Imports System.IO
Imports Microsoft.VisualBasic

Imports NSi.AutoStore.WebCapture.Workflow

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)


		Dim tsList As ListField = eventData.Form.Fields.GetField("TimeSheet")
		
		Dim tsInformation As String()
		Dim tsTimeSheetID As String
		Dim tsFileName As String
		Dim tsContractor As String
		Dim tsDate As String
		Dim UserName As String = eventData.User.UserName
		
		tsList.Items.Clear
		
		Dim fileEntries As String() = Directory.GetFiles("C:\Nuance\MobileDemo\Settings\TimeSheet\ToApprove\" & UserName,"*.txt")
		Dim fileName As String
		For Each fileName In fileEntries
			If fileName <> "-DO NOT DELETE-.txt" Then
				fileName = left(fileName,len(fileName)-4)
				tsFileName = right(fileName,len(fileName)-len("C:\Nuance\MobileDemo\Settings\TimeSheet\ToApprove\" & UserName)-1)
				tsInformation = split(tsFileName,"_")
				tsTimeSheetID = tsInformation(0)			
				tsContractor = tsInformation(1)
				tsDate = tsInformation(2)
	
			
			Dim listItem As listItem = New ListItem(tsContractor & " (" & tsDate & ")", tsFileName)
			tsList.Items.Add(listItem)
			End If
			
		Next fileName

		'TODO: add code here to execute when the form is first shown
	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub TimeSheet_OnChange(ByVal eventData As ClientEventData)
	
		Dim UserName As String = eventData.User.UserName
		Dim tsSelected As String
		Dim tsLine As String = ""
		Dim tsRead As String = ""
		Dim newLine As String = chr(13) & chr(10)

		Dim tsList As ListField = eventData.Form.Fields.GetField("TimeSheet")
		Dim tsDetails As BaseField = eventData.Form.Fields.GetField("TimeSheetDetails")

		tsSelected = "C:\Nuance\CompleteDemo\Settings\TimeSheet\ToApprove\" & UserName & "\" &  tsList.Value & ".txt"
		
		If (File.Exists(tsSelected)) Then
			Dim objReader As StreamReader = New System.IO.StreamReader(tsSelected)
			Do 
				tsRead = objReader.ReadLine 	
				tsLine = tsLine & newLine & tsRead 
			Loop Until tsRead Is Nothing
			objReader.Close()
			tsLine = tsLine & newLine & "---END---"
			tsDetails.Value = tsLine
			
		Else
			tsDetails.Value = tsSelected
		End If
		
	End Sub
	Sub Action_OnChange(ByVal eventData As ClientEventData)
		
		If eventData.Form.Fields.GetField("Action").Value = "p_0009" Then
			eventData.Form.Fields.GetField("RejectReason").IsHidden = True
		Else
			eventData.Form.Fields.GetField("RejectReason").IsHidden = False
		End If
		
	End Sub
		
End Module
