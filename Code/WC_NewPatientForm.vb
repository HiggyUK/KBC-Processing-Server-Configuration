Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports Microsoft.VisualBasic

Module Script

	Const Form1 As String = "1,2,3,53"
	Const Form2 As String = "4,5,6,7,8,53,54"
	Const Form3 As String = "9,10,11,53,54"
	Const Form4 As String = "12,13,14,53,54"
	Const Form5 As String = "15,16,17,18,19,20,53,54"
	Const Form6 As String = "21,22,23,24,53,54"
	Const Form7 As String = "25,26,27,28,29,30,53,54"
	Const Form8 As String = "31,32,33,34,35,36,53,54"
	Const Form9 As String = "37,38,39,40,53,54"
	Const Form10 As String = "41,42,43,44,45,46,47,48,53,54"
	Const Form11 As String = "49,50,51,52,54"
	Dim FormNo As Integer = 1
	Const MaxField As Integer = 54
	Const MaxFormNo As Integer = 11
		
    Sub Form_OnLoad(ByVal eventData As ClientEventData)
      
		eventData.Form.Fields.Item(1).Display = "Is this your first registration with a GP Practie in the UK?"
		eventData.Form.Fields.Item(2).Display = "Will you be in the area for more than 3 months?"
		eventData.Form.Fields.Item(13).Display = "Community Health Index (CHI) Number"
		eventData.Form.Fields.Item(37).Display = "Date you first came to live in the UK"
		eventData.Form.Fields.Item(38).Display = "If previously in the UK, date of leaving"
		eventData.Form.Fields.Item(39).Display = "Your most recent country of residence"
		eventData.Form.Fields.Item(43).Display = "Are you a reservist?"
		eventData.Form.Fields.Item(45).Display = "Is this your first registration with a GP since leaving?"
		eventData.Form.Fields.Item(46).Display = "Address prior to enlisting"
		eventData.Form.Fields.Item(47).Display="Postcode prior to enlisting"
		eventData.Form.Fields.item(51).Display="Indentification added to this Application"
		SetForm(eventData)
		
    End Sub
	
    Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub
	
	Sub SetForm(ByVal eventData As ClientEventData)
	
		Dim FormItems As String
		Dim DisplayItems As String()
		
	
		Select Case FormNo
			Case 1: FormItems = Form1
			Case 2: FormItems = Form2
			Case 3: FormItems = Form3
			Case 4: FormItems = Form4
			Case 5: FormItems = Form5
			Case 6: FormItems = Form6
			Case 7: FormItems = Form7
			Case 8: FormItems = Form8
			Case 9: FormItems = Form9
			Case 10: FormItems = Form10
			Case 11: FormItems = Form11
			
		End Select
		
		DisplayItems = split(FormItems,",")
	
		' Turn All Fields Off
		For Items As Integer = 0 To MaxField - 1 
			eventData.Form.Fields.Item(Items).IsHidden = True 
		Next
			
		For Each ItemNo As String In DisplayItems
			eventData.Form.Fields.Item(CInt(ItemNo)-1).IsHidden = False
		Next
				
	End Sub	

	Sub Button_OnClick(ByVal eventData As ClientEventData)
		Dim button As ButtonField = eventData.Form.Fields.GetField(eventData.EventSource)

		If button.Name = "Next" Then
			FormNo = FormNo + 1
			If FormNo > MaxFormNo Then
				FormNo = MaxFormNo
			End If
		End If
		If button.Name = "Back" Then
			FormNo = FormNo - 1
			If FormNo < 1 Then
				FormNo = 1
			End If
		End If
		
		SetForm(eventData)

		
		    
	End Sub 

End Module
