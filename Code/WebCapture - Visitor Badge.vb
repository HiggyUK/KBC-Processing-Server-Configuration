Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports System.IO
Imports Microsoft.VisualBasic

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)

		Dim BookedVisits = "C:\Nuance\MobileDemo\Settings\BookedVisits\" & eventData.User.UserName
		Dim vsList As ListField = eventData.Form.Fields.GetField("Visitors")
		
		Dim vsInformation As String()
		Dim vsFileName As String
		Dim vsSurname As String
		Dim vsFirstname As String
		Dim vsVisiting As String
		Dim vsCarParking As String
		Dim vsCarReg As String
		Dim vsDate As String
		Dim vsYear As String
		Dim vsMonth As String
		Dim vsDay As String
		

		vsList.Items.Clear
		
		Dim fileEntries As String() = Directory.GetFiles(BookedVisits,"*.txt")
		Dim fileName As String
		For Each fileName In fileEntries
			If fileName <> "-DO NOT DELETE-.txt" Then
				fileName = left(fileName,len(fileName)-4)
				vsFileName = right(fileName,len(fileName)-len(BookedVisits)-1)
				vsInformation = split(vsFileName,"_")
				vsYear = vsInformation(0)			
				vsMonth = vsInformation(1)
				vsDay = vsInformation(2)
				vsSurname = vsInformation(3)
				vsFirstname = vsInformation(4)
				vsVisiting = vsInformation(5)
			
				Dim listItem As listItem = New ListItem(vsSurname & ", " & vsFirstName & " (" & vsYear & "-" & vsMonth & "-" & vsDay & ")", vsFileName)
				vsList.Items.Add(listItem)
			End If
			
		Next fileName
		
		eventData.Form.Fields.GetField("VisitorDetails").IsHidden = True
		eventData.Form.Fields.GetField("VisitorSurname").IsHidden = True
		eventData.Form.Fields.GetField("VisitorFirstName").IsHidden = True
		eventData.Form.Fields.GetField("VisitorCompany").IsHidden = True
		eventData.Form.Fields.GetField("VisitDetails").IsHidden = True
		eventData.Form.Fields.GetField("VisitDate").IsHidden = True
		eventData.Form.Fields.GetField("VisitParking").IsHidden = True
		eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = True	
	

		
      'TODO: add code here to execute when the form is first shown
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	
	Sub Visitors_OnChange(ByVal eventData As ClientEventData)
	
		Dim vsSelected As String
		Dim vsRead As String = ""
		Dim vsLine As String()
		Dim vsSurname As String
		Dim vsFirstname As String
		Dim vsCompany As String
		Dim vsVisiting As String
		Dim vsCarParking As String
		Dim vsCarReg As String
		Dim vsDate As String
		
		Dim newLine As String = chr(13) & chr(10)

		Dim vsList As ListField = eventData.Form.Fields.GetField("Visitors")
		
		vsSelected = "C:\Nuance\MobileDemo\Settings\BookedVisits\" & eventData.User.UserName & "\" &  vsList.Value & ".txt"
		
		If (File.Exists(vsSelected)) Then
			Dim objReader As StreamReader = New System.IO.StreamReader(vsSelected)
			Do 
				vsRead = objReader.ReadLine 	
				vsLine = split(vsRead,"|")
				If vsLine(0) = "Surname" Then vsSurname = vsLine(1)
				If vsLine(0) = "Firstname" Then vsFirstname = vsLine(1)
				If vsLine(0) = "Host" Then vsVisiting = vsLine(1)
				If vsLine(0) = "DateOfVisit" Then vsDate = vsLine(1)
				If vsLine(0) = "CarParking" Then vsCarParking = vsLine(1)
				If vsLine(0) = "CarReg" Then vsCarReg = vsLine(1)
				If vsLine(0) = "Company" Then vsCompany = vsLine(1)
				
		
			Loop Until vsRead Is Nothing
			objReader.Close()
			
			eventData.Form.Fields.GetField("VisitorDetails").IsHidden = False
			eventData.Form.Fields.GetField("VisitorSurname").IsHidden = False
			eventData.Form.Fields.GetField("VisitorSurname").Value = vsSurname
			eventData.Form.Fields.GetField("VisitorFirstName").IsHidden = False
			eventData.Form.Fields.GetField("VisitorFirstName").Value = vsFirstname
			eventData.Form.Fields.GetField("VisitorCompany").IsHidden = False
			eventData.Form.Fields.GetField("VisitorCompany").Value = vsCompany
			eventData.Form.Fields.GetField("VisitDetails").IsHidden = False
			eventData.Form.Fields.GetField("VisitDate").IsHidden = False
			eventData.Form.Fields.GetField("VisitDate").Value = vsDate
			
			If vsCarParking = "Yes" Then
				eventData.Form.Fields.GetField("VisitParking").IsHidden = False
				eventData.Form.Fields.GetField("VisitParking").Value = vsCarParking
				eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = False
				eventData.Form.Fields.GetField("VisitCarDetails").Value = vsCarReg
			Else
				eventData.Form.Fields.GetField("VisitParking").IsHidden = False
				eventData.Form.Fields.GetField("VisitParking").Value = vsCarParking
				eventData.Form.Fields.GetField("VisitCarDetails").IsHidden = True
				eventData.Form.Fields.GetField("VisitCarDetails").Value = vsCarReg				
			End If
			eventData.Form.Fields.GetField("Visitors").IsHidden = True
		End If
		
	End Sub
End Module
