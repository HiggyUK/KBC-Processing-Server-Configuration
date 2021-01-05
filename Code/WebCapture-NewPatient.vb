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
		
		Dim DocumentTypeTextFile As String = "C:\Nuance\XeroxEIP-00001-HeathCare-PatientRecords\Settings\DocumentList.txt"
		Dim docTypes As ListField = eventData.Form.Fields.GetField("DocumentType")
		Dim docArray As Array
		Dim docRead as string
		
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
		
		Dim ConsultantTextFile As String = "C:\Nuance\XeroxEIP-00001-HeathCare-PatientRecords\Settings\ConsultantList.txt"
		Dim consultantList As ListField = eventData.Form.Fields.GetField("Consultant")
		Dim consultantArray As Array
		Dim consultantRead as String
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
		
		
		'Get next patient number
		Dim patientDir As String = "C:\Nuance\WebCapture-00001-HeathCare-PatientRecords\Output\"
		Dim nextPatientNumber As Integer = 0
		Dim folder As String
		Dim dirCount As Integer
		Dim patientNumber As String
		Dim nextPatientNumberStr As String
		
		Dim dirEntries as List(of String) = New List (Of String)(Directory.EnumerateDirectories(patientDir))
			
		For Each folder In dirEntries 
			patientNumber = Microsoft.VisualBasic.Left(folder.Substring(folder.LastIndexOf("\")+1),6)
			If CInt(patientNumber) > nextPatientNumber Then
				nextPatientNumber = CInt(patientNumber)
			End If
		Next
		nextPatientNumber = nextPatientNumber + 1
		nextPatientNumberStr = Microsoft.VisualBasic.Right(CStr(1000000+nextPatientNumber),6)
		
		eventData.Form.Fields.GetField("PatientNo").Value = nextPatientNumberStr

      'TODO: add code here to execute when the form is first shown
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
    End Sub
End Module
