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

	Sub PatientNumber_OnChange(ByVal eventData As ClientEventData)
		Dim surname As String = "Not Found"
		Dim forename As String = ""
		Dim bloodType As String = ""
		Dim dob As String =""
		
		Dim userHomeFolder As String = eventData.User.UserAttributes.Item(1).Data
		
		' Patient Lookup
		
		Dim patientID As String = eventData.Form.Fields.GetField("PatientNumber").Value	
		Dim patientTextFile As String = "C:\Nuance\MobileDemo\Settings\HealthCare\Patients\" + patientID + ".txt"
		Dim txtRead As String
		Dim txtLine As String()
		Dim patientSurname As String
		Dim patientForename As String
		Dim patientInitials As String
		Dim patientBloodType As String
		Dim patientDOBD As String
		Dim patientDOBM As String
		Dim patientDOBY As String
		
		surname = "Patient Data File Not Found"
		
		If (File.Exists(patientTextFile)) Then
			Dim objReader As New System.IO.StreamReader(patientTextFile)
			Do While objReader.Peek() <> -1
				txtRead = objReader.ReadLine()
				txtLine = split(txtRead,"=")
				If txtLine(0) = "Surname" Then patientSurname = txtLine(1)
				If txtLine(0) = "Forename" Then patientForename = txtLine(1)
				If txtLine(0) = "Initial" Then patientInitials = txtLine(1)
				If txtLine(0) = "BloodType" Then patientBloodType = txtLine(1)
				If txtLine(0) = "DOBD" Then patientDOBD = txtLine(1)
				If txtLine(0) = "DOBM" Then patientDOBM = txtLine(1)
				If txtLine(0) = "DOBY" Then patientDOBY = txtLine(1)
			Loop
			objReader.Close
		Else
			patientTextFile = userHomeFolder + "\Settings\HealthCare\Patients\" + patientID + ".txt"
			If (File.Exists(patientTextFile)) Then
				Dim objReader As New System.IO.StreamReader(patientTextFile)
				Do While objReader.Peek() <> -1
					txtRead = objReader.ReadLine()
					txtLine = split(txtRead,"=")
					If txtLine(0) = "Surname" Then patientSurname = txtLine(1)
					If txtLine(0) = "Forename" Then patientForename = txtLine(1)
					If txtLine(0) = "Initial" Then patientInitials = txtLine(1)
					If txtLine(0) = "BloodType" Then patientBloodType = txtLine(1)
					If txtLine(0) = "DOBD" Then patientDOBD = txtLine(1)
					If txtLine(0) = "DOBM" Then patientDOBM = txtLine(1)
					If txtLine(0) = "DOBY" Then patientDOBY = txtLine(1)
				Loop
				objReader.Close
			End If
			
		End If
		
		eventData.Form.Fields.GetField("Surname").Value = patientSurname
		eventData.Form.Fields.GetField("Forename").Value = patientForename
		eventData.Form.Fields.GetField("DOB").Value = patientDOBM + "-" + patientDOBD + "-" + patientDOBY
		
		

	End Sub
	
	Sub setForm(ByVal eventData As ClientEventData)
	
		Dim totalFields As Integer = eventData.Form.Fields.Count
		Dim forms() As String = { "1,2,3,4,5,49" , "6,7,8,9,49,50", "10,11,12,13,14,49,50", "15,16,17,18,19,49,50", "20,21,22,23,49,50", "24,25,26,27,28,29,30,49,50", "31,32,33,34,35,36,37,49,50", "38,39,40,41,42,49,50","43,44,49,50","45,46,47,48,50"}
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
