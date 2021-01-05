Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		eventData.Form.Fields.GetField("formNo").value = 1
		setForm(eventData)
		
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
		Dim forms() As String = { "1,2,3,25" , "4,5,6,7,8,9,10,25,26", "11,12,13,14,25,26", "15,16,17,25,26", "18,19,25,26", "20,21,22,23,24,26"}
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
	
	
	Sub CustomerID_OnChange(ByVal eventData As ClientEventData)
		Dim customerName As String = "Not Found"
		Dim address As String = ""
		Dim contact As String = ""
		Dim phone As String =""
		Dim mobile As String = ""
		
		Dim userHomeFolder As String = eventData.User.UserAttributes.Item(1).Data
		
		' Customer Lookup
		
		Dim customerID As String = eventData.Form.Fields.GetField("CustomerID").Value	
		Dim customerTextFile As String = "C:\Nuance\MobileDemo\Settings\FieldService\Customers\" + customerID + ".txt"
		Dim txtRead As String
		Dim txtLine As String()
		
		
		If (File.Exists(customerTextFile)) Then
			Dim objReader As New System.IO.StreamReader(customerTextFile)
			Do While objReader.Peek() <> -1
				txtRead = objReader.ReadLine()
				txtLine = split(txtRead,"=")
				If txtLine(0) = "NAME" Then customerName = txtLine(1)
				If txtLine(0) = "ADDRESS" Then address = txtLine(1)
				If txtLine(0) = "CONTACT" Then contact = txtLine(1)
				If txtLine(0) = "PHONE" Then phone = txtLine(1)
				If txtLine(0) = "MOBILE" Then mobile = txtLine(1)
			Loop
			objReader.Close
			
		End If
		
		eventData.Form.Fields.GetField("CustomerName").Value = customerName
		eventData.Form.Fields.GetField("Address").Value = address
		eventData.Form.Fields.GetField("Contact").Value = contact
		eventData.Form.Fields.GetField("Phone").Value = phone
		eventData.Form.Fields.GetField("Mobile").Value = mobile
		
	End Sub
	Sub MachineID_OnChange(ByVal eventData As ClientEventData)
		Dim make As String = "Not Found"
		Dim serial As String = ""
		Dim model As String = ""
		
		Dim userHomeFolder As String = eventData.User.UserAttributes.Item(1).Data
		
		' Machine Lookup
		
		Dim machineID As String = eventData.Form.Fields.GetField("MachineID").Value	
		Dim machineTextFile As String = "C:\Nuance\MobileDemo\Settings\FieldService\Machines\" + machineID + ".txt"
		Dim txtRead As String
		Dim txtLine As String()
		
		
		If (File.Exists(machineTextFile)) Then
			Dim objReader As New System.IO.StreamReader(machineTextFile)
			Do While objReader.Peek() <> -1
				txtRead = objReader.ReadLine()
				txtLine = split(txtRead,"=")
				If txtLine(0) = "MAKE" Then make = txtLine(1)
				If txtLine(0) = "MODEL" Then model = txtLine(1)
				If txtLine(0) = "SERIAL" Then serial = txtLine(1)
			Loop
			objReader.Close
			
		End If
		
		eventData.Form.Fields.GetField("Make").Value = make + " " + model
		eventData.Form.Fields.GetField("SerialNo").Value = serial
		
	End Sub
End Module
