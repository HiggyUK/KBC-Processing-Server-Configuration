Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.Diagnostics
Imports System.IO
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

	
		' Document List
		
		Dim DocumentTypeTextFile As String = "C:\Nuance\WebCapture-00001-HeathCare-PatientRecords\Settings\DocumentList.txt"
		Dim docTypes As ListField = eventData.Form.Fields.GetField("DocumentType")
		Dim docRead As String
		
		docTypes.Items.Clear()
			
		If (File.Exists(DocumentTypeTextFile)) Then
			Dim objReader As New System.IO.StreamReader(DocumentTypeTextFile)
			Do While objReader.Peek() <> -1
				docRead = objReader.ReadLine()
				Dim listItem As listItem = New ListItem(docRead,docRead)						
				docTypes.Items.Add(listItem)
			Loop
		Else
			Dim listItem As listItem = New ListItem("No entry found","ERROR")
			docTypes.Items.Add(listItem)
		End If


		' Consultant List
		
		Dim ConsultantTextFile As String = "C:\Nuance\WebCapture-00001-HeathCare-PatientRecords\Settings\ConsultantList.txt"
		Dim consultantList As ListField = eventData.Form.Fields.GetField("Consultant")
		Dim consultantRead As String
		consultantList.Items.Clear()
			
		If (File.Exists(ConsultantTextFile)) Then
			Dim objReader As New System.IO.StreamReader(ConsultantTextFile)
			Do While objReader.Peek() <> -1
				consultantRead = objReader.ReadLine()
				Dim listItem As listItem = New ListItem(consultantRead,consultantRead)						
				consultantList.Items.Add(listItem)
			Loop
		Else
			Dim listItem As listItem = New ListItem("No entry found","ERROR")
			consultantList.Items.Add(listItem)
		End If
		
		eventData.Form.Fields.GetField("Patients").isHidden = true
		eventData.Form.Fields.GetField("SearchAgain").isHidden = true
		eventData.Form.Fields.GetField("PatientNo").isHidden= true
		eventData.Form.Fields.GetField("PatientSurname").isHidden = true
		eventData.Form.Fields.GetField("PatientFirstName").isHidden = true
		eventData.Form.Fields.GetField("PatientDOB").isHidden = true
		eventData.Form.Fields.GetField("DocumentType").isHidden = true
		eventData.Form.Fields.GetField("Consultant").isHidden=true
		eventData.Form.Fields.GetField("DateOfVisit").isHidden=true
			

      'TODO: add code here to execute when the form is first shown
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub
	
	
	Sub Search_OnClick(ByVal eventData As ClientEventData)
	
		' Search based on entries

		Dim patientDir As String = "C:\Nuance\WebCapture-00001-HeathCare-PatientRecords\Output\"
		Dim folder As String
		Dim folderName As String
		Dim patientNoSearch As String
		Dim patientSurnameSearch As String
		Dim patientList As ListField = eventData.Form.Fields.GetField("Patients")
		Dim patientCount As Integer = 0
		Dim patientDetails As Array
	
		patientList.Items.clear()
		patientNoSearch = eventData.Form.Fields.GetField("PatientNoSearch").value
		patientSurnameSearch = eventData.Form.Fields.GetField("PatientSurnameSearch").value
		
		Dim dirEntries As List(Of String) = New List (Of String)(Directory.EnumerateDirectories(patientDir))
			
		For Each folder In dirEntries 
			folderName = folder.Substring(folder.LastIndexOf("\")+1)
			patientDetails = split(folderName,"_")
			If patientNoSearch <> "" And patientSurnameSearch = "" Then
				If instr(patientDetails(0),patientNoSearch) <> 0 Then 
					Dim listItem As listItem = New ListItem(folderName,folderName)
					patientList.Items.Add(listItem)
					patientCount = patientCount+1
				End If
			End If
			If patientNoSearch = "" And patientSurnameSearch <> "" Then
				If instr(patientDetails(1),patientSurnameSearch) <> 0 Then 
					Dim listItem As listItem = New ListItem(folderName,folderName)
					patientList.Items.Add(listItem)
					patientCount = patientCount +1
				End If
			End If		
			If patientNoSearch <> "" And patientSurnameSearch <> "" Then
				If instr(patientDetails(0),patientNoSearch) <> 0 And instr(patientDetails(1), patientSurnameSearch) <> 0 Then 
					Dim listItem As listItem = New ListItem(folderName,folderName)
					patientList.Items.Add(listItem)
					patientCount = patientCount+1
				End If
			End If
		Next
		If patientCount = 0 Then	
			Dim listItem As listItem = New ListItem("No Patients found - please search again","ERROR")
			patientList.Items.Add(listItem)
		End If

		eventData.Form.Fields.GetField("PatientNoSearch").IsHidden = True
		eventData.Form.Fields.GetField("PatientSurnameSearch").IsHidden = True
		eventData.Form.Fields.GetField("Search").IsHidden = True
		eventData.Form.Fields.GetField("Patients").isHidden = False
		eventData.Form.Fields.GetField("SearchAgain").isHidden = False
	
	End Sub
	
	Sub SearchAgain_OnClick(ByVal eventData As ClientEventData)
				
		eventData.Form.Fields.GetField("PatientNoSearch").IsHidden = False
		eventData.Form.Fields.GetField("PatientSurnameSearch").IsHidden = False
		eventData.Form.Fields.GetField("Search").IsHidden = False
		
		eventData.Form.Fields.GetField("Patients").isHidden = True
		eventData.Form.Fields.GetField("SearchAgain").isHidden = True
		
	End Sub
		
	Sub Patients_OnChange(ByVal eventData As ClientEventData)
		
		Dim selectedPatientDetails As Array
		Dim selectedPatient As String
		
		selectedPatient = eventData.Form.Fields.GetField("Patients").Value
		selectedPatientDetails = split(selectedPatient,"_")
		
		eventData.Form.Fields.GetField("PatientNo").value = selectedPatientDetails(0)
		eventData.Form.Fields.GetField("PatientSurname").value = selectedPatientDetails(1)
		eventData.Form.Fields.GetField("PatientFirstName").value = selectedPatientDetails(2)
		eventData.Form.Fields.GetField("PatientDOB").value = selectedPatientDetails(3)
	
		eventData.Form.Fields.GetField("Patients").isHidden = True
		eventData.Form.Fields.GetField("SearchAgain").isHidden = True
		
		eventData.Form.Fields.GetField("PatientNo").isHidden= False
		eventData.Form.Fields.GetField("PatientSurname").isHidden = False
		eventData.Form.Fields.GetField("PatientFirstName").isHidden = False
		eventData.Form.Fields.GetField("PatientDOB").isHidden = False
		eventData.Form.Fields.GetField("DocumentType").isHidden = False
		eventData.Form.Fields.GetField("Consultant").isHidden=False
		eventData.Form.Fields.GetField("DateOfVisit").isHidden=False	

	End Sub
End Module
