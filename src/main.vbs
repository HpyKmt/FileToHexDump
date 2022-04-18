Option Explicit

Rem - VarType
Const VER_TYPE_NULL = 1
Const VAR_TYPE_BYTES = 8209

Rem - ADODB Stream Type
Const adTypeText   = 2
Const adTypeBinary = 1

Rem - File System Object
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Rem - Convert Long to hex string with zero padding
Function LongToHex(ByVal lngNum)
	LongToHex = Right("00000000" & Hex(lngNum), 8)
End Function

Rem - Write ASCII string to file
Sub WriteTextASCII(ByVal strFp, ByVal strTxt)
	WScript.Echo "WriteTextASCII strFp=" & strFp
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, ForWriting, True)
	Call objTS.Write(strTxt)
	objTS.Close
	Set objFSO = Nothing
End Sub

Rem - Delete text files
Sub DeleteFiles(ByVal argDir)
	WScript.Echo "Delete Files in argDir=" & argDir
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFile: For Each objFile In objFSO.GetFolder(argDir).Files
		WScript.Echo "Deleting: " & objFile.Name
		objFile.Delete
	Next
	Set objFSO = Nothing
	Set objFile = Nothing
End Sub

Class BytesToHexString
	Rem - Output folder path
	Private m_DirOut
	Public Property Let DirOut(ByVal argDirOut)
		m_DirOut = argDirOut
	End Property

	Rem - Column Count
	Private m_ColumnCount
	Public Property Let ColumnCount(ByVal argColumnCount)
		m_ColumnCount = argColumnCount
	End Property

	Rem - Start Position
	Private m_StartPosition
	Private Sub Class_Initialize()
		m_StartPosition = 0
	End Sub
	
	Rem - Convert byte array to hex string and save on file
	Public Sub BytesToFile(ByVal argBytes)
		Dim i, j, k, strBuf, arrBuf, strResult, intUB
		
		Rem - Create column header
		ReDim arrBuf(m_ColumnCount - 1)
		For i = 0 To m_ColumnCount - 1
			arrBuf(i) = Right("00" & Hex(i), 2)
		Next
		strResult = "-------- " & Join(arrBuf, " ") & vbCrlf
		
		Rem - Loop through rows
		intUB = m_ColumnCount - 1
		For i = LBound(argBytes) To UBound(argBytes) Step m_ColumnCount
			Rem - Adjust the upper bound in case byte count is smaller than coulumn length
			If i + (m_ColumnCount - 1) > UBound(argBytes) Then
				intUB = UBound(argBytes) - i
			End If
			Rem - Reset array buffer
			ReDim arrBuf(intUB)
			Rem - Fill the array
			For j = 0 To intUB
				arrBuf(j) = Right("00" & Hex(AscB(MidB(argBytes, i+j+1, 1))), 2)
			Next
			Rem - Update resultant string
			strResult = strResult & LongToHex(m_StartPosition + i) & " " & Join(arrBuf, " ") & vbCrlf			
		Next
		
		Rem - Write the result to file
		Call WriteTextASCII(m_DirOut & "\" & LongToHex(m_StartPosition) & ".txt", strResult)
		
		Rem - Increment start position for the next round
		m_StartPosition = m_StartPosition + UBound(argBytes) + 1
	End Sub

End Class


Sub Main()
	Rem - Arguments
	Dim strFpIn:      strFpIn      = WScript.Arguments(0)
	Dim strDirOut:    strDirOut    = WScript.Arguments(1)
	Dim intRowCount: intRowCount = CInt(WScript.Arguments(2))
	Dim intColumnCount: intColumnCount = CInt(WScript.Arguments(3))
	
	Rem - Clean up output folder.
	DeleteFiles(strDirOut)
	
	Rem - Initialize
	Dim lngBlockSize: lngBlockSize = intRowCount * intColumnCount
	Dim objHex: Set objHex = New BytesToHexString
	objHex.DirOut = strDirOut
	objHex.ColumnCount = intColumnCount

	Rem - Stream
	Dim arrBytes
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = adTypeBinary
		.LoadFromFile(strFpIn)
		Do While True
			Rem - Read Bytes
			arrBytes = .Read(lngBlockSize)
			WScript.Echo "TypeName(arrBytes)=" & TypeName(arrBytes), "VarType(arrBytes)=" & VarType(arrBytes)
			Rem - Process Bytes
			Select Case VarType(arrBytes)
				Case VAR_TYPE_BYTES
					WScript.Echo "UBound(arrBytes)=" & UBound(arrBytes)
					objHex.BytesToFile(arrBytes)
				Case VER_TYPE_NULL
					WScript.Echo "No more data"
					Exit Do
				Case Else
					WScript.Echo "Unexpected VerType", TypeName(arrBytes), VarType(arrBytes)
					WScript.Quit 1
			End Select
		Loop
		.Close
	End With
	Set objStream = Nothing
	
	Set objHex = Nothing
	
End Sub
Call Main
