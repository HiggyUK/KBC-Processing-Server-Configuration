Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic


Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		Dim customPatients  As String = eventData.User.UserAttributes.Item(1).Data + "\Settings\HealthCare\Patients"
		Dim patientList As ListField = eventData.Form.Fields.GetField("PatientID")
		Dim patientFileName As String
		Dim txtRead As String
		Dim txtLine As String()
		Dim patientSurname As String
		Dim patientForename As String
		Dim patientInitials As String
		Dim patientBloodType As String
		Dim patientDOBD As String
		Dim patientDOBM As String
		Dim patientDOBY As String
		patientList.Items.Clear
		
		Dim fileEntries As String() = Directory.GetFiles(customPatients,"*.txt")
		Dim fileName As String
		For Each fileName In fileEntries
			If fileName <> "-DO NOT DELETE-.txt" Then
				fileName = left(fileName,len(fileName)-4)
				patientFileName = right(fileName,len(fileName)-len(customPatients)-1)
				Dim objReader As New System.IO.StreamReader(fileName & ".txt")
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
				Dim listItem As listItem = New ListItem(patientSurname & ", " & patientForename & " (" & patientDOBY & "-" & patientDOBM & "-" & patientDOBD & ") - [" + patientFileName + "]", patientFileName)
				patientList.Items.Add(listItem)
			End If
			
		Next fileName
		
		eventData.Form.Fields.GetField("Surname").IsHidden = True
		eventData.Form.Fields.GetField("Forename").IsHidden = True
		eventData.Form.Fields.GetField("DOB").IsHidden = True
		eventData.Form.Fields.GetField("BloodType").IsHidden = True
		eventData.Form.Fields.GetField("Delete").IsHidden = True

	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
		'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub PatientID_OnChange(ByVal eventData As ClientEventData)
		
		Dim txtRead As String
		Dim txtLine As String()
		Dim patientSurname As String
		Dim patientForename As String
		Dim patientInitials As String
		Dim patientBloodType As String
		Dim patientDOBD As String
		Dim patientDOBM As String
		Dim patientDOBY As String
		Dim patientID As ListField = eventData.Form.Fields.GetField("PatientID")
		Dim fldSurname As BaseField = eventData.Form.Fields.GetField("Surname")
		Dim fldForename As BaseField = eventData.Form.Fields.GetField("Forename")
		Dim fldDOB As DateField = eventData.Form.Fields.GetField("DOB")
		Dim fldBloodType As BaseField = eventData.Form.Fields.GetField("BloodType")
		dim fldDelete as CheckboxField = eventData.Form.Fields.GetField("Delete")
		
		Dim patientRecord As String = eventData.User.UserAttributes.Item(1).Data + "\Settings\HealthCare\Patients\" + patientID.Value + ".txt"
		Dim objReader As New System.IO.StreamReader(patientRecord)
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
		
		fldSurname.value = patientSurname
		fldForename.value = patientForename
		fldDOB.value = DateSerial(patientDOBY,patientDOBM,patientDOBD)
		fldBloodType.value = patientBloodType
		
		fldSurname.IsHidden = False
		fldForename.IsHidden = False
		fldDOB.IsHidden = False
		fldBloodType.IsHidden = False
		fldDelete.IsHidden = False
		
		
	End Sub

	Sub Delete_OnChange(ByVal eventData As ClientEventData)
		
		dim fldDelete as CheckboxField = eventData.Form.Fields.GetField("Delete")
		If fldDelete.Value = 1 Then
			eventData.Form.AllowFormSubmitting = True
		Else
			eventData.Form.AllowFormSubmitting = False
		End If
		
	end sub
		
	
End Module
