Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = True
		
      'TODO: add code here to execute when the form is first shown
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub VisitParking_OnChange(ByVal eventData As ClientEventData)
		If eventData.Form.Fields.GetField("VisitParking").Value = "Yes" Then
			eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = False
		Else
			eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = True
		End If
		
	End Sub
		
End Module
