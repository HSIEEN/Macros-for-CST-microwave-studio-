'#Language "WWB-COM"

Option Explicit
Public Const lib_rundef = -1.2345678e27

Sub Main

	MsgBox "Please select two 1D results from a common directory", vbInformation

	Dim n As Integer, points_n As Integer
	Dim selectedItem0 As String, file0 As String, o0 As Object, freqmin0 As Double, freqmax0 As Double
	Dim selectedItem1 As String, file1 As String, o1 As Object, freqmin1 As Double, freqmax1 As Double
	Dim freqGlobalmin As Double, freqGlobalmax As Double


	n = GetNumberOfSelectedTreeItems

	If n = 0 Then
		MsgBox "Nothing picked, please Try again!", vbCritical
		Exit All
	End If

	If n <> 2 Then
		MsgBox "Picked items are invalid, please Try again!", vbCritical
		Exit All
	End If
	selectedItem0 = GetSelectedTreeItem
	selectedItem1 = GetNextSelectedTreeItem

	Dim tmpstr As String

	If StrComp(selectedItem0, selectedItem1) > 0 Then
		tmpstr = selectedItem0
		selectedItem0 = selectedItem1
		selectedItem1 = tmpstr
	End If

	file0 = ResultTree.GetFileFromTreeItem(selectedItem0)
	file1 = ResultTree.GetFileFromTreeItem(selectedItem1)

	On Error Resume Next
	Set o0 = Result1D(file0)
	If Err.Number = 0 Then
		'Set o0 = Mag_Linear2dB(o0)
	Else
		Set o0 = Result1DComplex(file0)
		Set o0 = Mag_Linear2dB(o0)
		Err.Clear
	End If

	On Error Resume Next
	Set o1 = Result1D(file1)
	If Err.Number = 0 Then
		'Set o1 = Mag_Linear2dB(o1)
	Else
		Set o1 = Result1DComplex(file1)
		Set o1 = Mag_Linear2dB(o1)
		Err.Clear
	End If

	With o0
		points_n = .GetN
		freqmin0 = .GetX(0)
		freqmax0 = .GetX(points_n-1)
	End With


	With o1
		points_n = .GetN
		freqmin1 = .GetX(0)
		freqmax1 = .GetX(points_n-1)
	End With

	If freqmin0 >= freqmax1 Or freqmax0 <= freqmin0 Then
		MsgBox "Picked two 1D results have no any intersection part, please check!!", vbCritical
		Exit All
	End If

	If freqmin0 >= freqmin1 Then
		freqGlobalmin = freqmin0
	Else
		freqGlobalmin = freqmin1
	End If

	If freqmax0 <= freqmax1 Then
		freqGlobalmax = freqmax0
	Else
		freqGlobalmax = freqmax1
	End If


	Dim o2 As Object
	Dim o3 As Object

	Set o2 = o0.copy
	Set o3 = o1.copy

	o2.resampleTo(freqGlobalmin,freqGlobalmax,Int((freqGlobalmax-freqGlobalmin)/0.01))
	o3.resampleTo(freqGlobalmin,freqGlobalmax,Int((freqGlobalmax-freqGlobalmin)/0.01))
	o2.Subtract(o3)

	o2.Xlabel("Frequecy/GHz")
	o2.Ylabel("dB")
	o2.save(Right(selectedItem0, Len(selectedItem0)-InStrRev(selectedItem0 ,"\"))+ _
	"_subtract_"+Right(selectedItem1, Len(selectedItem1)-InStrRev(selectedItem1 ,"\"))+".sig")

	Dim treeItem As String
	treeItem = Left(selectedItem0, InStrRev(selectedItem0 ,"\"))+ _
	Right(selectedItem0, Len(selectedItem0)-InStrRev(selectedItem0 ,"\"))+ _
	" - "+Right(selectedItem1, Len(selectedItem1)-InStrRev(selectedItem1 ,"\"))

	o2.AddToTree(treeItem)
	SelectTreeItem(treeItem)
End Sub
Function Mag_Linear2dB(O As Object) As Object
	Dim n As Integer, logFactor As Integer, i As Integer
	With O
		n = .getN
		logFactor = .GetLogarithmicFactor
	End With

	Set O = O.magnitude

	Dim ox As Object
	Set ox = Result1D("")

	For i=0 To n-1
		ox.AppendXY(O.GetX(i), logFactor*CST_Log10(O.GetY(i)))
	Next


	Set Mag_Linear2dB = ox

End Function
Function CST_Log10(value As Double) As Double
	If value <= 0 Then
		CST_Log10 = lib_rundef
	Else
		CST_Log10 = Log(value)/Log(10)
	End If
End Function
