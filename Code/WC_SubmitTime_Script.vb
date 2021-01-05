Option Strict Off

Imports System
Imports System.Text
Imports Microsoft.VisualBasic


Imports NSi.AutoStore.WebCapture.Workflow

Module Script
	Sub Form_OnLoad(ByVal eventData As ClientEventData)	
		eventData.Form.Fields.GetField("MonStart").IsHidden=True
		eventData.Form.Fields.GetField("MonEnd").IsHidden = True
		eventData.Form.Fields.GetField("MonHours").IsHidden = True
		eventData.Form.Fields.GetField("TueStart").IsHidden=True
		eventData.Form.Fields.GetField("TueEnd").IsHidden = True
		eventData.Form.Fields.GetField("TueHours").IsHidden = True
		eventData.Form.Fields.GetField("WedStart").IsHidden=True
		eventData.Form.Fields.GetField("WedEnd").IsHidden = True
		eventData.Form.Fields.GetField("WedHours").IsHidden = True
		eventData.Form.Fields.GetField("ThuStart").IsHidden=True
		eventData.Form.Fields.GetField("ThuEnd").IsHidden = True
		eventData.Form.Fields.GetField("ThuHours").IsHidden = True
		eventData.Form.Fields.GetField("FriStart").IsHidden=True
		eventData.Form.Fields.GetField("FriEnd").IsHidden = True
		eventData.Form.Fields.GetField("FriHours").IsHidden = True
		'TODO: add code here to execute when the form is first shown
	End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub TimeInvalid_OnChange(ByVal eventData As ClientEventData)
	
	
	
	End Sub
		
	Sub MonStandard_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("MonStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("MonStart").IsHidden=True
			eventData.Form.Fields.GetField("MonEnd").IsHidden = True
			eventData.Form.Fields.GetField("MonHours").IsHidden = True
		Else
			eventData.Form.Fields.GetField("MonStart").IsHidden=False
			eventData.Form.Fields.GetField("MonEnd").IsHidden = False
			eventData.Form.Fields.GetField("MonHours").IsHidden = False

		End If
		
	End Sub
	
	Sub TueStandard_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("TueStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("TueStart").IsHidden=True
			eventData.Form.Fields.GetField("TueEnd").IsHidden = True
			eventData.Form.Fields.GetField("TueHours").IsHidden = True
		Else
			eventData.Form.Fields.GetField("TueStart").IsHidden=False
			eventData.Form.Fields.GetField("TueEnd").IsHidden = False
			eventData.Form.Fields.GetField("TueHours").IsHidden = False
		End If
		
	End Sub
	
	Sub WedStandard_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("WedStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("WedStart").IsHidden=True
			eventData.Form.Fields.GetField("WedEnd").IsHidden = True
			eventData.Form.Fields.GetField("WedHours").IsHidden = True
		Else
			eventData.Form.Fields.GetField("WedStart").IsHidden=False
			eventData.Form.Fields.GetField("WedEnd").IsHidden = False
			eventData.Form.Fields.GetField("WedHours").IsHidden = False
		End If
		
	End Sub

	Sub ThuStandard_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("ThuStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("ThuStart").IsHidden=True
			eventData.Form.Fields.GetField("ThuEnd").IsHidden = True
			eventData.Form.Fields.GetField("ThuHours").IsHidden = True
		Else
			eventData.Form.Fields.GetField("ThuStart").IsHidden=False
			eventData.Form.Fields.GetField("ThuEnd").IsHidden = False
			eventData.Form.Fields.GetField("ThuHours").IsHidden = False
		End If
		
	End Sub

	Sub FriStandard_OnChange(ByVal eventData As ClientEventData)
	
		If eventData.Form.Fields.GetField("FriStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("FriStart").IsHidden=True
			eventData.Form.Fields.GetField("FriEnd").IsHidden = True
			eventData.Form.Fields.GetField("FriHours").IsHidden = True
		Else
			eventData.Form.Fields.GetField("FriStart").IsHidden=False
			eventData.Form.Fields.GetField("FriEnd").IsHidden = False
			eventData.Form.Fields.GetField("FriHours").IsHidden = False
		End If
		
	End Sub

	Sub StandardHoursStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim timeValue As String = eventData.Form.Fields.GetField("StandardHoursStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
		
		If validTime = False Then
			eventData.Form.Fields.GetField("SSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("SSInvalid").IsHidden = True
		End If
		
	End Sub

	
	Sub StandardHoursEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim timeValue As String = eventData.Form.Fields.GetField("StandardHoursEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("StandardHoursStart").Value
		Dim standardHours As BaseField = eventData.Form.Fields.GetField("StandardHoursDaily")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		Dim MonStandard As CheckboxField = eventData.Form.Fields.GetField("MonStandard")
		Dim TueStandard As CheckboxField = eventData.Form.Fields.GetField("TueStandard")
		Dim WedStandard As CheckboxField = eventData.Form.Fields.GetField("WedStandard")
		Dim ThuStandard As CheckboxField = eventData.Form.Fields.GetField("ThuStandard")
		Dim FriStandard As CheckboxField = eventData.Form.Fields.GetField("FriStandard")
		
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue) And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format(dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				standardHours.Value = hoursWorked
				totalHours.Value = UpdateTotalHours(eventData)
			End If
		End If
		
		
		If validTime = False Then
			eventData.Form.Fields.GetField("SEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("SEInvalid").IsHidden = True
		End If
		
		
	End Sub	
	

	Sub MonStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim timeValue As String = eventData.Form.Fields.GetField("MonStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
		
		If validTime = False Then
			eventData.Form.Fields.GetField("MSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("MSInvalid").IsHidden = True
		End If
		
	End Sub
	
	Sub MonEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim timeValue As String = eventData.Form.Fields.GetField("MonEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("MonStart").Value
		Dim monHours As BaseField = eventData.Form.Fields.GetField("MonHours")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue) And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format(dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				monHours.Value = hoursWorked
				totalHours.Value = UpdateTotalHours(eventData)
			End If
		End If
		
		
		If validTime = False Then
			eventData.Form.Fields.GetField("MEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("MEInvalid").IsHidden = True
		End If
		
		
	End Sub	
	
	
	Sub TueStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim timeValue As String = eventData.Form.Fields.GetField("TueStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
		
		
		If validTime = False Then
			eventData.Form.Fields.GetField("TUSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("TUSInvalid").IsHidden = True
		End If
		
	End Sub
	
	Sub TueEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("TueEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("TueStart").Value
		Dim tueHours As BaseField = eventData.Form.Fields.GetField("TueHours")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue) And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format(dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				tueHours.Value = hoursWorked
				totalHours.Value = UpdateTotalHours(eventData)
			End If
		End If
		
		
		
		If validTime = False Then
			eventData.Form.Fields.GetField("TUEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("TUEInvalid").IsHidden = True
		End If
		
		
	End Sub	
	
	Sub WedStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("WedStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If

		
		If validTime = False Then
			eventData.Form.Fields.GetField("WSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("WSInvalid").IsHidden = True
		End If
		
	End Sub
	
	Sub WedEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("WedEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("WedStart").Value
		Dim wedHours As BaseField = eventData.Form.Fields.GetField("WedHours")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue) And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format( dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				wedHours.Value = hoursWorked
				totalHours.Value = UpdateTotalHours(eventData)
			End If
		End If
		
		

		
		If validTime = False Then
			eventData.Form.Fields.GetField("WEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("WEInvalid").IsHidden = True
		End If
		
		
	End Sub	

		
	Sub ThuStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("ThuStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
		

		
		If validTime = False Then
			eventData.Form.Fields.GetField("THSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("THSInvalid").IsHidden = True
		End If
		
	End Sub
	
	Sub ThuEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("ThuEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("ThuStart").Value
		Dim thuHours As BaseField = eventData.Form.Fields.GetField("ThuHours")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue)And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format(dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				thuHours.Value = hoursWorked
				
				
				totalHours.Value = UpdateTotalHours(eventData)

			End If
		End If
		
		
		
		If validTime = False Then
			eventData.Form.Fields.GetField("THEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("THEInvalid").IsHidden = True
		End If
		
	End Sub	

		
	Sub FriStart_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("FriStart").Value		
		Dim validTime As Boolean = True
		
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
		

		
		If validTime = False Then
			eventData.Form.Fields.GetField("FSInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("FSInvalid").IsHidden = True
		End If
		
	End Sub
	
	Sub FriEnd_OnChange(ByVal eventData As ClientEventData)
	
		Dim invalidPopUp As PopupMessageField = eventData.Form.Fields.GetField("TimeInvalid")
		Dim timeValue As String = eventData.Form.Fields.GetField("FriEnd").Value
		Dim startValue As String = eventData.Form.Fields.GetField("FriStart").Value
		Dim friHours As BaseField = eventData.Form.Fields.GetField("FriHours")
		Dim totalHours As BaseField = eventData.Form.Fields.GetField("TotalHours")
		Dim validTime As Boolean = True
				
		'Check Length
		If Len(timeValue) > 5 Then
			validTime = False
		End If
		
		'Check if valid Time
		
		If isDate(timeValue) Then
			validTime = True
		Else
			validTime = False	
		End If
	
		'Check if higher than Start
		If validTime = True Then
			If isDate(startValue) And isDate(timeValue) Then
				Dim startTime As DateTime = startValue
				Dim endTime As DateTime = timeValue
				Dim timeDifference As Integer
				Dim hoursWorked As String
				
				timeDifference = datediff("s",startTime, endTime)
				hoursWorked = Format(dateadd("s",timeDifference,"00:00"),"HH:mm")
	
				friHours.Value = hoursWorked
				
				totalHours.Value = UpdateTotalHours(eventData)
				

			End If
		End If

		
		If validTime = False Then
			eventData.Form.Fields.GetField("FEInvalid").IsHidden = False
		Else
			eventData.Form.Fields.GetField("FEInvalid").IsHidden = True
		End If
		
		
	End Sub	

	Function UpdateTotalHours(ByVal eventData As ClientEventData)
	
		Dim monHours As String
		Dim tueHours As String
		Dim wedHours As String
		Dim thuHours As String
		Dim friHours As String
		
		If eventData.Form.Fields.GetField("MonStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("MonStart").Value = eventData.Form.Fields.GetField("StandardHoursStart").Value
			eventData.Form.Fields.GetField("MonEnd").Value = eventData.Form.Fields.GetField("StandardHoursEnd").Value
			eventData.Form.Fields.GetField("MonHours").Value = eventData.Form.Fields.GetField("StandardHoursDaily").Value
		End If

		If eventData.Form.Fields.GetField("TueStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("TueStart").Value = eventData.Form.Fields.GetField("StandardHoursStart").Value
			eventData.Form.Fields.GetField("TueEnd").Value = eventData.Form.Fields.GetField("StandardHoursEnd").Value
			eventData.Form.Fields.GetField("TueHours").Value = eventData.Form.Fields.GetField("StandardHoursDaily").Value
		End If
		
		If eventData.Form.Fields.GetField("WedStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("WedStart").Value = eventData.Form.Fields.GetField("StandardHoursStart").Value
			eventData.Form.Fields.GetField("WedEnd").Value = eventData.Form.Fields.GetField("StandardHoursEnd").Value
			eventData.Form.Fields.GetField("WedHours").Value = eventData.Form.Fields.GetField("StandardHoursDaily").Value
		End If
		
		If eventData.Form.Fields.GetField("ThuStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("ThuStart").Value = eventData.Form.Fields.GetField("StandardHoursStart").Value
			eventData.Form.Fields.GetField("ThuEnd").Value = eventData.Form.Fields.GetField("StandardHoursEnd").Value
			eventData.Form.Fields.GetField("ThuHours").Value = eventData.Form.Fields.GetField("StandardHoursDaily").Value
		End If
		
		If eventData.Form.Fields.GetField("FriStandard").Value = "Yes" Then
			eventData.Form.Fields.GetField("FriStart").Value = eventData.Form.Fields.GetField("StandardHoursStart").Value
			eventData.Form.Fields.GetField("FriEnd").Value = eventData.Form.Fields.GetField("StandardHoursEnd").Value
			eventData.Form.Fields.GetField("FriHours").Value = eventData.Form.Fields.GetField("StandardHoursDaily").Value
		End If
	
		
		monHours = eventData.Form.Fields.GetField("MonHours").value
		tueHours =	 eventData.Form.Fields.GetField("TueHours").value
		wedHours =  eventData.Form.Fields.GetField("WedHours").value
		thuHours =  eventData.Form.Fields.GetField("ThuHours").value
		friHours =  eventData.Form.Fields.GetField("FriHours").value
	
		Dim tHours As Integer = 0
		Dim tMinutes As Integer = 0
		Dim aHours As Integer = 0
		Dim aMinutes As Integer = 0
		Dim nHours As Integer = 0
		Dim nMinutes As Integer =0 
		
		If len(monHours) >= 5 Then
			aHours = CInt(Microsoft.VisualBasic.Left(monHours,2))
			aMinutes = CInt(Microsoft.VisualBasic.Mid(monHours,4,2))
			tHours = tHours + aHours
			tMinutes = tMinutes + aMinutes
		End If
		If len(tueHours) >= 5 Then
			aHours = CInt(Microsoft.VisualBasic.Left(tueHours,2))
			aMinutes = CInt(Microsoft.VisualBasic.Mid(tueHours,4,2))
			tHours = tHours + aHours
			tMinutes = tMinutes + aMinutes
		End If		
		If len(wedHours) >= 5 Then
			aHours = CInt(Microsoft.VisualBasic.Left(wedHours,2))
			aMinutes = CInt(Microsoft.VisualBasic.Mid(wedHours,4,2))
			tHours = tHours + aHours
			tMinutes = tMinutes + aMinutes
		End If		
		If len(thuHours) >= 5 Then
			aHours = CInt(Microsoft.VisualBasic.Left(thuHours,2))
			aMinutes = CInt(Microsoft.VisualBasic.Mid(thuHours,4,2))
			tHours = tHours + aHours
			tMinutes = tMinutes + aMinutes
		End If	
		If len(friHours) >= 5 Then
			aHours = CInt(Microsoft.VisualBasic.Left(friHours,2))
			aMinutes = CInt(Microsoft.VisualBasic.Mid(friHours,4,2))
			tHours = tHours + aHours
			tMinutes = tMinutes + aMinutes
		End If		
		
		nHours = 100 + tHours 
		nMinutes = tMinutes
		If nMinutes >= 60 Then
			nHours = nHours + 1
			nMinutes = nMinutes - 60
		End If
		nMinutes= 100 + nMinutes
		
		UpdateTotalHours = right(CStr(nHours),2) & ":" & right(CStr(nMinutes),2)
		
		Return UpdateTotalHours
		
	
	End Function
		
End Module
