Public ws, ws_ As Worksheet
Public sridhar As String



Sub pre_q()
	
	Set ws = ThisWorkbook.Sheets("data")
	Set ws_ = ThisWorkbook.Sheets("Dashboard")
	
	If ws_.Cells(2, 4) <> "" And ws_.Cells(3, 4) <> "" Then
		Call gen_data
	Else
		MsgBox "Kindly fill up the 'file extension' & 'location to be saved'"
	End If
	
	
End Sub

'process data
Sub gen_data()
	
	Application.DisplayStatusBar = True
	Application.StatusBar = "Please be patient..."
	
	Dim filepath1, filepath As String
	filepath1 = ws_.Cells(3, 4).Value
	
	Dim x, y As Integer
	x = ws.Range("A1").End(xlDown).Row
	y = ws.Cells(1, Columns.Count).End(xlToLeft).Column
	
	For i = 2 To x
		Application.StatusBar = "Processing file no: " & i & " of " & x
		ws_.Cells(4, 6).Value = "Processing file no: " & i & " of " & x
		
		filepath = filepath1 & ws.Cells(i, 1) & "." & ws_.Cells(2, 4).Value
		Call create_file(filepath, i, y)
		
	Next
	
	ws_.Cells(4, 6).Value = "Processed " & (x * y) & " cells to generate " & (x - 1) & " files :)"
	Application.StatusBar = "sweet :)"
End Sub

'file creation
Function create_file(m, n, o)
	
	Dim fso As Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim oFile As Object
	Set oFile = fso.CreateTextFile(m)
	oFile.WriteLine ws_.Cells(4, 4).Value
	
	For ii = 2 To o
		oFile.WriteLine "<" & ws.Cells(1, ii).Value & ">" & ws.Cells(n, ii).Value & "</" & ws.Cells(1, ii).Value & ">"
	Next
	
	oFile.WriteLine ws_.Cells(5, 4).Value
	oFile.Close
	Set fso = Nothing
	Set oFile = Nothing
	
End Function


