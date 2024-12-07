'#Language "WWB-COM"

Option Explicit

Sub Main

	MsgBox "Please select two 1D results from a common directory", vbInformation

	Dim n As Integer, points_n As Integer
	Dim selectedItem0 As String, file0 As String, o0 As Object, freqmin0 As Double, freqmax0 As Double
	Dim SelectedItem1 As String, file1 As String, o1 As Object, freqmin1 As Double, freqmax1 As Double
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
	SelectedItem1 = GetNextSelectedTreeItem

	file0 = ResultTree.GetFileFromTreeItem(selectedItem0)
	file1 = ResultTree.GetFileFromTreeItem(SelectedItem1)

	Set o0 = Result1D(file0)
	Set o1 = Result1D(file1)

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
	"_subtract_"+Right(SelectedItem1, Len(SelectedItem1)-InStrRev(SelectedItem1 ,"\"))+".sig")
	o2.AddToTree(Left(selectedItem0, InStrRev(selectedItem0 ,"\"))+ _
	Right(selectedItem0, Len(selectedItem0)-InStrRev(selectedItem0 ,"\"))+ _
	" - "+Right(SelectedItem1, Len(SelectedItem1)-InStrRev(SelectedItem1 ,"\")))
End Sub
