Option Strict Off

Imports System
Imports NSi.AutoStore.WebCapture.Workflow
Imports Microsoft.VisualBasic
Imports System.IO

Module Script
    Sub Form_OnLoad(ByVal eventData As ClientEventData)
	
		
		'Load Up Available Vehicle Registrations
		Dim vehicleData  As String = "C:\Nuance\MobileDemo\Data\" + eventData.User.UserName

		Dim vehicleList As ListField = eventData.Form.Fields.GetField("VehicleReg")
		Dim vehicleCount As Integer = 0
		Dim vehicleReg As String
		
		vehicleList.Items.Clear
		
		Dim directoryEntries As String() = Directory.GetDirectories(vehicleData)
		Dim directoryName As String
		For Each directoryName In directoryEntries
			vehicleReg = right(directoryName,len(directoryName)-len(vehicleData)-1)
			Dim listItem As listItem = New ListItem(vehicleReg, vehicleReg)
			vehicleList.Items.Add(listItem)
			vehicleCount = vehicleCount + 1
		Next
		
		If vehicleCount = 0 Then
			Dim listItem As listItem = New ListItem("No Vehicles Found", "No Vehicles Found")
			vehicleList.Items.Add(listItem)
		End If
		
		eventData.Form.Fields.GetField("formNo").Value = 1
		SetForm(eventData)

	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
		'TODO: add code here to execute when the user presses OK in the form
	End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
	End Sub

	Sub SetForm (ByVal eventData As ClientEventData)
	
		Dim totalFields As Integer = eventData.Form.Fields.Count
		Dim forms() As String = { "1,2,3,4" , "5,6,7,8,28,29", "9,10,11,28,29" , "12,13,14,15,16,28,29" , "17,18,19,28,29" , "20,21,22,28,29" , "23,24,25,26,28,29", "27,29" }
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

	Sub VehicleReg_OnChange(ByVal eventData As ClientEventData)
		'Load Up Available Deliveries
		
		
		Dim vehicleList As ListField = eventData.Form.Fields.GetField("VehicleReg")
		Dim vehicleReg As String = vehicleList.Value
		Dim dateList As ListField = eventData.Form.Fields.GetField("DeliveryDate")
		Dim deliveryList As ListField = eventData.Form.Fields.GetField("DeliveryNumber")
		Dim deliveryData  As String = "C:\Nuance\MobileDemo\Data\" + eventData.User.UserName + "\" + vehicleReg
		Dim deliveryDate As String
		Dim dateCount As Integer	
		
		dateList.Items.Clear
		deliveryList.Items.Clear
		
		Dim directoryEntries As String() = Directory.GetDirectories(deliveryData)
		Dim directoryName As String
		For Each directoryName In directoryEntries
			deliveryDate = right(directoryName,len(directoryName)-len(deliveryData)-1)
			Dim listItem As listItem = New ListItem(deliveryDate, deliveryDate)
			dateList.Items.Add(listItem)
			dateCount = dateCount + 1
		Next
		
		If dateCount = 0 Then
			Dim listItem As listItem = New ListItem("No Delivery Dates Found", "No Delivery Found")
			dateList.Items.Add(listItem)
		End If
	End Sub
		
	
	Sub DeliveryDate_OnChange(ByVal eventData As ClientEventData)
		'Load Up Available Deliveries
		
		
		Dim vehicleList As ListField = eventData.Form.Fields.GetField("VehicleReg")
		Dim vehicleReg As String = vehicleList.Value
		Dim dateList As ListField = eventData.Form.Fields.GetField("DeliveryDate")
		Dim deliveryDate As String = dateList.Value
		Dim deliveryList As ListField = eventData.Form.Fields.GetField("DeliveryNumber")
		Dim deliveryData  As String = "C:\Nuance\MobileDemo\Data\" + eventData.User.UserName + "\" + vehicleReg +"\" + deliveryDate
		Dim deliveryFileName As String
		Dim txtRead As String
		Dim txtLine As String()
		Dim OwnerName As String
		Dim OwnerTown As String
		Dim AnimalType As String
		Dim AnimalQty As String
		Dim DeliveryNumber As String
		
		
		'dateList.Items.Clear
		deliveryList.Items.Clear
		
		Dim fileEntries As String() = Directory.GetFiles(deliveryData,"*.txt")
		Dim fileName As String
		For Each fileName In fileEntries
			If fileName <> "-DO NOT DELETE-.txt" Then
				fileName = left(fileName,len(fileName)-4)
				deliveryFileName = right(fileName,len(fileName)-len(deliveryData)-1)
				Dim objReader As New System.IO.StreamReader(fileName & ".txt")
				Do While objReader.Peek() <> -1
					txtRead = objReader.ReadLine()
					txtLine = split(txtRead,"=")
					If txtLine(0) = "OwnerName" Then OwnerName = txtLine(1)
					If txtLine(0) = "OwnerTown" Then OwnerTown = txtLine(1)
					If txtLine(0) = "AnimalType" Then AnimalType = txtLine(1)
					If txtLine(0) = "AnimalQty" Then AnimalQty = txtLine(1)
					If txtLine(0) = "DeliveryNumber" Then DeliveryNumber = txtLine(1)
				Loop
				objReader.Close
				Dim deliverylistItem As ListItem = New ListItem(DeliveryNumber + " - " + OwnerName +","+OwnerTown+" - " + AnimalQty + " x " + AnimalType, fileName +".txt")
				deliveryList.Items.Add(deliverylistItem)

			End If
		Next fileName
	End Sub
	
	
	Sub DeliveryNumber_OnChange(ByVal eventData As ClientEventData)

		
		Dim deliveryList As ListField = eventData.Form.Fields.GetField("DeliveryNumber")
		Dim deliveryFileName As String = deliveryList.Value
		Dim txtRead As String
		Dim txtLine As String()
		Dim OwnerName As String
		Dim OwnerTown As String
		Dim OwnerAddress As String
		Dim AnimalType As String
		Dim AnimalQty As String
		Dim DeliveryNumber As String
		Dim DestinationTown As String
		Dim DestinationAddress As String
		Dim DepartureTown As String
		Dim DepartureAddress As String
		Dim MAFFAuthorisation As String
		Dim HaulageAssuranceNo As String
		Dim Driver As String
		Dim DelNo as String
		
		
		Dim objReader As New System.IO.StreamReader(deliveryFileName)
		
		Do While objReader.Peek() <> -1
			txtRead = objReader.ReadLine()
			txtLine = split(txtRead,"=")
			If txtLine(0) = "OwnerName" Then OwnerName = txtLine(1)
			If txtLine(0) = "OwnerAddress" Then OwnerAddress = txtLine(1)
			If txtLine(0) = "OwnerTown" Then OwnerTown = txtLine(1)
			If txtLine(0) = "AnimalType" Then AnimalType = txtLine(1)
			If txtLine(0) = "AnimalQty" Then AnimalQty = txtLine(1)
			If txtLine(0) = "DestinationTown" Then DestinationTown = txtLine(1)
			If txtLine(0) = "DestinationAddress" Then DestinationAddress = txtLine(1)
			If txtLine(0) = "DepartureTown" Then DepartureTown = txtLine(1)
			If txtLine(0) = "DepartureAddress" Then DepartureAddress = txtLine(1)
			If txtLine(0) = "MAFFAuthorisation" Then MAFFAuthorisation = txtLine(1)
			If txtLine(0) = "Driver" Then Driver = txtLine(1)
			If txtLine(0) = "HaulageAssuranceNo" Then HaulageAssuranceNo = txtLine(1)
			If txtLine(0) = "DelNo" Then DelNo = txtLine(1)
		Loop
		objReader.Close

		eventData.Form.Fields.GetField("OwnerName").Value = OwnerName
		eventData.Form.Fields.GetField("OwnerAddress").Value = OwnerAddress
		eventData.Form.Fields.GetField("OwnerTown").Value = OwnerTown
		eventData.Form.Fields.GetField("AnimalType").Value = AnimalType
		eventData.Form.Fields.GetField("AnimalQty").Value = AnimalQty
		eventData.Form.Fields.GetField("DestinationTown").Value = DestinationTown
		eventData.Form.Fields.GetField("DestinationAddress").Value = DestinationAddress
		eventData.Form.Fields.GetField("DepartureTown").Value = DepartureTown
		eventData.Form.Fields.GetField("DepartureAddress").Value = DepartureAddress
		eventData.Form.Fields.GetField("MAFFAuthorisation").Value = MAFFAuthorisation
		eventData.Form.Fields.GetField("Driver").Value = Driver
		eventData.Form.Fields.GetField("HaulageAssuranceNo").Value=HaulageAssuranceNo
		eventData.Form.Fields.GetField("DelNo").Value = DelNo
		
		eventData.Form.Fields.GetField("formNo").Value = 2
		setForm(eventData)
		
	End Sub
		
End Module
