Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic


Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		eventData.Form.Fields.GetField("FormNo").Value = 1
		setForm(eventData)
		
		'TODO: add code here to execute when the form is first shown
	End Sub
	
	Sub setForm(ByVal eventData As ClientEventData)
	
		Dim totalFields As Integer = eventData.Form.Fields.Count
		Dim forms() As String = { "1,2,3,20" , "4,5,6,7,8,20,21", "9,10,11,20,21" , "14,15,16,17,18,19,21", "13" }
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
		
	Sub NurseID_OnChange(ByVal eventData As ClientEventData)

		' Nurse Lookup (Defaults first)
		
		Dim nurseName As String = "Nurse Not Found"
		Dim nurseID As String = eventData.Form.Fields.GetField("NurseID").Value	
		Dim nurseTextFile As String = "C:\Nuance\MobileDemo\Settings\HealthCare\Nurses\" + nurseID + ".txt"
		Dim txtRead As String
		Dim txtLine As String()
		Dim nurseSurname As String
		Dim nurseForename As String
		Dim nurseInitials As String
		Dim nurseLevel As String
		Dim userHomeFolder As String = eventData.User.UserAttributes.Item(1).Data
		
		nurseName = "Nurse Data File Not Found"
		
		If (File.Exists(nurseTextFile)) Then
			Dim objReader As New System.IO.StreamReader(nurseTextFile)
			Do While objReader.Peek() <> -1
				txtRead = objReader.ReadLine()
				txtLine = split(txtRead,"=")
				If txtLine(0) = "Surname" Then nurseSurname = txtLine(1)
				If txtLine(0) = "Forename" Then nurseForename = txtLine(1)
				If txtLine(0) = "Initials" Then nurseInitials = txtLine(1)
				If txtLine(0) = "Level" Then nurseLevel = txtLine(1)
			Loop
			objReader.Close
			nurseName = nurseLevel + " " + nurseInitials + " " + nurseSurname
		Else
			nurseTextFile = userHomeFolder + "\Settings\HealthCare\Nurses\" + nurseID + ".txt"
			If (File.Exists(nurseTextFile)) Then
				Dim objReader As New System.IO.StreamReader(nurseTextFile)
				Do While objReader.Peek() <> -1
					txtRead = objReader.ReadLine()
					txtLine = split(txtRead,"=")
					If txtLine(0) = "Surname" Then nurseSurname = txtLine(1)
					If txtLine(0) = "Forename" Then nurseForename = txtLine(1)
					If txtLine(0) = "Initials" Then nurseInitials = txtLine(1)
					If txtLine(0) = "Level" Then nurseLevel = txtLine(1)
				Loop
				objReader.Close
				nurseName = nurseLevel + " " + nurseInitials + " " + nurseSurname
			End If
		End If
	
		eventData.Form.Fields.GetField("NurseName").Value = nurseName
		
				
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
		eventData.Form.Fields.GetField("BloodType").Value = patientBloodType

	
	End Sub
		
	Sub BloodBag_OnChange(ByVal eventData As ClientEventData)
	
		Dim bloodBagType As String = ""
		Dim bloodBagID As String = eventData.Form.Fields.GetField("BloodBag").Value	
		Dim bloodBagTextFile As String = "C:\Nuance\MobileDemo\Settings\HealthCare\BloodBags\" + bloodBagID + ".txt"
		Dim txtRead As String
		Dim txtLine As String()
		Dim userHomeFolder As String = eventData.User.UserAttributes.Item(1).Data
			
		bloodBagType = "Setting Data File Not Found"
		
		If (File.Exists(bloodBagTextFile)) Then
			bloodBagType = "Not Found"
			Dim objReader As New System.IO.StreamReader(bloodBagTextFile)
			Do While objReader.Peek() <> -1
				txtRead = objReader.ReadLine()
				txtLine = split(txtRead,"=")
				If txtLine(0) = "Type" 
					bloodBagType = txtLine(1)
				End If
			Loop
			objReader.Close
		Else
			bloodBagTextFile = userHomeFolder + "\Settings\HealthCare\Patients\" + bloodBagID + ".txt"
			If (File.Exists(bloodBagTextFile)) Then
				bloodBagType = "Not Found"
				Dim objReader As New System.IO.StreamReader(bloodBagTextFile)
				Do While objReader.Peek() <> -1
					txtRead = objReader.ReadLine()
					txtLine = split(txtRead,"=")
					If txtLine(0) = "Type" 
						bloodBagType = txtLine(1)
					End If
				Loop
				objReader.Close
			End If
		End If
		
		eventData.Form.Fields.GetField("BloodBagType").Value = bloodBagType
		
		Dim warningLabel1 As LabelField = eventData.Form.Fields.GetField("WarningLabel1")
		Dim warningLabel2 As LabelField = eventData.Form.Fields.GetField("WarningLabel2")
		Dim warningLabel3 As LabelField = eventData.Form.Fields.GetField("WarningLabel3")
		Dim warningLabel4 As LabelField = eventData.Form.Fields.GetField("WarningLabel4")
		Dim warningLabel5 As LabelField = eventData.Form.Fields.GetField("WarningLabel5")
		Dim warningLabel6 As LabelField = eventData.Form.Fields.GetField("WarningLabel6")

		
		
		
		If eventData.Form.Fields.GetField("BloodType").Value <> eventData.Form.Fields.GetField("BloodBagType").Value Then
			
			warningLabel1.Text = "********************************"
			warningLabel2.Text = "       BLOOD TYPE MISMATCH      "
			warningLabel3.Text = "  PATIENT : " + eventData.Form.Fields.GetField("BloodType").Value 
			warningLabel4.Text = "  BLOOD BAG : " + eventData.Form.Fields.GetField("BloodBagType").Value 
			warningLabel5.Text = "      #### DO NOT USE ####      "
			warningLabel6.Text = "*********************************"
			
			eventData.Form.AllowFormSubmitting=False
			eventData.Form.Fields.GetField("FormNo").Value = 4
			setForm(eventData)
	
				
		Else
			
			eventData.Form.Fields.GetField("FormNo").Value = 5
			setForm(eventData)
			eventData.Form.AllowFormSubmitting=True
		End If
		
	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
		'TODO: add code here to execute when the user presses OK in the form
	End Sub

	Sub Form_OnSubmit(ByVal eventData As ClientEventData)
		If eventData.Form.Fields.GetField("BloodType").Value <> eventData.Form.Fields.GetField("BloodBagType").Value Then
			eventData.Form.Fields.GetField("Match").Value = "No"
		Else
			eventData.Form.Fields.GetField("Match").value = "Yes"
		End If
		'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub Button_OnClick(ByVal eventData As ClientEventData)
	
	
		Dim button As ButtonField = eventData.Form.Fields.GetField(eventData.EventSource)
		Dim formNo As NumericField = eventData.Form.Fields.GetField("FormNo")
		
		If button.Name = "btnNext"
			formNo.value = formNo.Value + 1
			setForm(eventData)
		Else If button.Name = "btnBack"
			formNo.value = formNo.value - 1
			setForm(eventData)
		End If    

		
	End Sub
	
	

		
	
	
End Module
