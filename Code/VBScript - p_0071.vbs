Sub GetCourseDetails_OnLoad
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile("C:\Kofax\MobileDemo\Settings\Education\courses.CSV", 1)
	Do Until (objTextFile.AtEndOfStream)
		the_line = objTextFile.Readline
		arr=Split(the_line, ";")
		dbcourseId = arr(0)
		dbcourseName = arr(1)
		If dbcourseId = courseId Then
			courseName = dbcourseName
		End If
	Loop
	objTextFile.Close

	Set PTopic = KnowledgeContent.GetTopicInterface
	If Not(PTopic Is Nothing) Then
		PTopic.Replace "~USR::CourseName~", courseName
	End If	
End Sub

Sub GetCourseDetails_OnUnload

End Sub
