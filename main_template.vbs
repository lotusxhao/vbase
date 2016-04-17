Option Explicit

Sub Include(file)
	On Error Resume Next

	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing

	If Err.Number <> 0 Then
		If Err.Number = 1041 Then
			Err.Clear
		Else
			WScript.Echo Err.Number & ": " & Err.Description
			WScript.Quit 1
		End If
	End If
End Sub

If WScript.ScriptName = "main_template.vbs" Then
	Include "v_Data"
	Include "v_HTML"
	Include "v_JSON"
	Include "v_Script"
	Include "v_URI"
End If