Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)
		
		eventData.Form.Fields.GetField("formNo").Value = 1
		setForm(eventData)
		
		eventData.Form.Fields.GetField("Label8").Value = "Notification of Insurance Excess for Accidental Damage and Theft in which the Driver is to blame: £250 (25 & Over), £400 (21 - 24), £500 (Under 21)."
		eventData.Form.Fields.GetField("Label9").Value = "Any accidents must be reported to the Fleet Manager as soon as possible and an Accident Report Form must be filled in no matter how trivial the incident is."
		eventData.Form.Fields.GetField("Label10").Value = "I received and accept the above vehicle, which is now in my care and I have been informed that it is my duty to ensure regular servicing is carried out and the vehicle is kept CLEAN and TIDY to our high standards"
		eventData.Form.Fields.GetField("Label11").Value = "I have been notified of Insurance Excess and understand that the vehicle can be inspected at any time"
		
      'TODO: add code here to execute when the form is first shown
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub setForm(ByVal eventData As ClientEventData)
	
		Dim totalFields As Integer = eventData.Form.Fields.Count
		Dim forms() As String = { "1,2,3,4,5,6,7,38" , "8,9,10,11,12,13,38,39", "14,15,16,17,18,19,20,38,39", "21,22,23,24,38,39", "25,26,27,28,29,30,38,39", "31,32,33,34,35,36,37,39"}
		Dim currentForm As BaseField = eventData.Form.Fields.GetField("formNo")
		Dim fieldItem As String
		Dim fieldItems As String()
		Dim fields As Integer
		Dim items As Integer
		Dim fieldNo As Integer

			
		'Turn all fields off
		For fields = 0 To totalFields - 1
			eventData.Form.Fields.Item(fields).IsHidden = True
		Next
		
		'Turn on fields for the current form
		fieldItems = split(Forms(CInt(currentForm.Value)-1),",")
		For items = 0 To fieldItems.Length -1
			fieldNo = CInt(fieldItems(items))-1
			eventData.Form.Fields.Item(fieldNo).IsHidden=False
		Next
			
	End Sub
	Sub Button_OnClick(ByVal eventData As ClientEventData)
	
	
		Dim button As ButtonField = eventData.Form.Fields.GetField(eventData.EventSource)
		Dim formNo As NumericField = eventData.Form.Fields.GetField("formNo")
		
		If button.Name = "btnNext"
			formNo.value = formNo.Value + 1
			setForm(eventData)
		Else If button.Name = "btnBack"
			formNo.value = formNo.value - 1
			setForm(eventData)
		End If    

		

	End Sub
End Module
