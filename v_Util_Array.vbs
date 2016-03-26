Function IsArrayAllocated(arrArray)
	IsArrayAllocated = False
	If IsArray(arrArray) Then
		On Error Resume Next
		Dim ub : ub = UBound(arrArray)
		If (Err.Number = 0) And (ub >= 0) Then IsArrayAllocated = True
	End If  
End Function

Sub QuickSort(ByRef arrArray, intLoBound, intHiBound)
	Dim varPivot, _
		intLoSwap, _
		intHiSwap, _
		varTemp

	If intHiBound - intLoBound = 1 Then
		If arrArray(intLoBound) > arrArray(intHiBound) Then
			varTemp = arrArray(intLoBound)
			arrArray(intLoBound) = arrArray(intHiBound)
			arrArray(intHiBound) = varTemp
		End If
	End If

	varPivot = arrArray(CInt((intLoBound + intHiBound) / 2))
	arrArray(CInt((intLoBound + intHiBound) / 2)) = arrArray(intLoBound)
	arrArray(intLoBound) = varPivot
	intLoSwap = intLoBound + 1
	intHiSwap = intHiBound
  
	Do
		While intLoSwap < intHiSwap and arrArray(intLoSwap) <= varPivot
			intLoSwap = intLoSwap + 1
		Wend

		While arrArray(intHiSwap) > varPivot
			intHiSwap = intHiSwap - 1
		Wend

		If intLoSwap < intHiSwap Then
			varTemp = arrArray(intLoSwap)
			arrArray(intLoSwap) = arrArray(intHiSwap)
			arrArray(intHiSwap) = varTemp
		End If
	Loop While intLoSwap < intHiSwap
  
	arrArray(intLoBound) = arrArray(intHiSwap)
	arrArray(intHiSwap) = varPivot
  
	If intLoBound < (intHiSwap - 1) Then Call QuickSort(arrArray, intLoBound, intHiSwap - 1)
	If intHiSwap + 1 < intHiBound Then Call QuickSort(arrArray, intHiSwap + 1, intHiBound)
End Sub

If WScript.ScriptName = "v_Util_Array.vbs" Then

End If