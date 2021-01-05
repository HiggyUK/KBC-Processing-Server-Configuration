Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		eventData.Form.Fields.GetField("EnteredPassword").IsHidden = False

	
	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub
	
	Sub EnterPassword_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("EnterPassword").Value = 1 Then
			eventData.Form.Fields.GetField("EnteredPassword").IsHidden = False
		Else
			eventData.Form.Fields.GetField("EnteredPassword").IsHidden = True
		End If
		
	End Sub
		
	
	
End Module
