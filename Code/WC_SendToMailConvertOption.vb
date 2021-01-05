Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		eventData.Form.Fields.GetField("ConvertTo").IsHidden = True
		eventData.Form.Fields.GetField("Next").Value = "p_0027"
	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

	Sub Form_OnSubmit(ByVal eventData As ClientEventData)
		If eventData.Form.Fields.GetField("Convert").Value = "Yes" Then
			eventData.Form.Fields.GetField("Next").Value = "p_0024"
		Else
			eventData.Form.Fields.GetField("Next").Value = "p_0027"
		End If
		'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub
	
	Sub Convert_OnChange(ByVal eventData As ClientEventData)
		
		If eventData.Form.Fields.GetField("Convert").Value = "Yes" Then
			eventData.Form.Fields.GetField("ConvertTo").IsHidden = False
			eventData.Form.Fields.GetField("Next").Value = "p_0024"
		Else
			eventData.Form.Fields.GetField("ConvertTo").IsHidden = True
			eventData.Form.Fields.GetField("Next").Value = "p_0027"
		End If
		
	End Sub
		
End Module
