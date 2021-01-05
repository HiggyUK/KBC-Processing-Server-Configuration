Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic

Module Script
    Sub Form_OnLoad(ByVal eventData As ClientEventData)
		'TODO: add code here to execute when the form is first shown
		eventData.Form.Fields.GetField("Surname").IsHidden = True
		eventData.Form.Fields.GetField("FirstName").IsHidden = True
		eventData.Form.Fields.GetField("Address").IsHidden = True
		eventData.Form.Fields.GetField("Postcode").IsHidden = True
		eventData.Form.Fields.GetField("GPS").IsHidden = True
		eventData.Form.Fields.GetField("Signature").IsHidden = True
		eventData.Form.Fields.GetField("SignatureName").IsHidden = True
		eventData.Form.Fields.GetField("StatusMessage").IsHidden = False

	End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	
	Sub DeliveryID_OnChange(ByVal eventData As ClientEventData)
	
		Dim txtRead As String = ""
		Dim deliveryID As String = eventData.Form.Fields.GetField("DeliveryID").Value
		Dim txtLine As String()
		Dim Surname As String
		Dim Firstname As String
		Dim Address As String
		Dim Postcode As String
		
		If (File.Exists("C:\Nuance\CompleteDemo\Settings\ProofofDelivery\DeliveryBarcodes.txt")) Then
			Dim objReader As StreamReader = New System.IO.StreamReader("C:\Nuance\CompleteDemo\Settings\ProofofDelivery\DeliveryBarcodes.txt")
			Do 
				txtRead = objReader.ReadLine 	
				txtLine = split(txtRead,"|")
				If txtLine(0) = deliveryID Then
					Surname = txtLine(1)
					Firstname = txtLine(2)
					Address = txtLine(3)
					Postcode = txtLine(4)
				End If
			Loop Until txtRead Is Nothing
			objReader.Close()
			
			eventData.Form.Fields.GetField("Surname").IsHidden = False
			eventData.Form.Fields.GetField("Surname").Value = Surname
			eventData.Form.Fields.GetField("FirstName").IsHidden = False
			eventData.Form.Fields.GetField("FirstName").Value = Firstname
			eventData.Form.Fields.GetField("Address").IsHidden = False
			eventData.Form.Fields.GetField("Address").Value = Address
			eventData.Form.Fields.GetField("Postcode").IsHidden = False
			eventData.Form.Fields.GetField("Postcode").Value = Postcode
			eventData.Form.Fields.GetField("GPS").IsHidden = False
			eventData.Form.Fields.GetField("Signature").IsHidden = False
			eventData.Form.Fields.GetField("SignatureName").IsHidden = False
			eventData.Form.Fields.GetField("StatusMessage").IsHidden = True

			
		End If
		
	End Sub
End Module
