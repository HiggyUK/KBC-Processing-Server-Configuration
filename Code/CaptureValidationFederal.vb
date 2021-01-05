Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports Microsoft.VisualBasic

Module Script
    Sub Form_OnLoad(ByVal eventData As ClientEventData)
        'TODO add code here to execute when the form is first shown

		Dim DateListField As New DateField
		DateListField = EventData.Form.Fields.GetField("Date")
		
		DateListField.Value = Date.Today
	
	
	End Sub
    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
        'TODO add code here to execute when the user presses OK in the form
		
		Dim BankPhone As String
		BankPhone = EventData.Form.Fields.GetField("BankPhone").Value

		Dim BankPhoneFirst3TextField As New TextField
		BankPhoneFirst3TextField = EventData.Form.Fields.GetField("BankPhoneFirst3")
		BankPhoneFirst3TextField.Value = Mid(BankPhone, 2, 3)

		Dim BankPhoneMiddle3TextField As New TextField
		BankPhoneMiddle3TextField = EventData.Form.Fields.GetField("BankPhoneMiddle3")
		BankPhoneMiddle3TextField.Value = Mid(BankPhone, 6, 3)
	
		Dim BankPhoneLast4TextField As New TextField
		BankPhoneLast4TextField = EventData.Form.Fields.GetField("BankPhoneLast4")
		BankPhoneLast4TextField.Value = Mid(BankPhone, 10, 4)
	
		Dim Phone As String
		Phone = EventData.Form.Fields.GetField("Phone").Value

		Dim PhoneFirst3TextField As New TextField
		PhoneFirst3TextField = EventData.Form.Fields.GetField("PhoneFirst3")
		PhoneFirst3TextField.Value = Mid(Phone, 2, 3)

		Dim PhoneMiddle3TextField As New TextField
		PhoneMiddle3TextField = EventData.Form.Fields.GetField("PhoneMiddle3")
		PhoneMiddle3TextField.Value = Mid(Phone, 6, 3)
	
		Dim PhoneLast4TextField As New TextField
		PhoneLast4TextField = EventData.Form.Fields.GetField("PhoneLast4")
		PhoneLast4TextField.Value = Mid(Phone, 10, 4)

		'		Dim SSNTextField As TextField
		'		SSNTextField = eventData.Form.Fields.GetField("SSN")
		'		If (Not (SSNTextField.Value.StartsWith("CC_")))
		'			'Invalid customer code, do not allow scan to continue and display an error message
		'			eventData.Form.IsValid = False
		'			eventData.Form.StatusInfo.MessageType = MessageType.Error
		'			eventData.Form.StatusInfo.Message = "Invalid customer code"
		'		Else
		'			'Customer code looks ok, let the user exit the form
		'		End If


	End Sub
End Module
