' Format 1d plot

'20230614-By Shawn in COROS
'Public FirstChildItem As String
Option Explicit
Sub Main ()
  'Efficiency results parent path


	Dim selectedItem As String
	Dim sselectedItem As String
	Dim curveLabel As String
	Dim index As Integer
	Dim selectedItems(1000) As Variant
	Dim filename As String, sigName As String
	Dim filePath As String
	Dim n As Integer, i As Integer, N_points As Integer, m As Integer

	selectedItem = GetSelectedTreeItem
	'selectedItem = GetNextSelectedTreeItem
	While selectedItem <> ""

		If (InStr(selectedItem,"1D Results") = 0) Then
	        MsgBox("Please select at least a directory containing 1D results before runing this macro.",vbCritical,"Warning")
	        Exit All
	    End If

	    'Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	    Dim isSmith As Boolean
	    Dim test As Boolean
	    isSmith = False

	    If HasChildren(selectedItem) Then

	    	sselectedItem = ResultTree.GetFirstChildName(selectedItem)

			While sselectedItem <> ""
				filename = ResultTree.GetFileFromTreeItem(sselectedItem)
				sigName = Right(filename, Len(filename)-InStrRev(filename,"\"))
	          	Dim magnitude() As Double
	          	Dim X() As Double
	    		Dim Obj_1D As Object
	    		Set Obj_1D = Result1D(filename)
	    		N_points = Obj_1D.GetN
	    		ReDim magnitude(N_points) As Double
	    		ReDim X(N_points) As Double
				For m = 0 To N_points-1
					If Obj_1D.GetY(m)>0 Then
						magnitude(m) = Obj_1D.GetY(m)
					Else
						MsgBox("Data value of curve " +sselectedItem+" is less than zero, continue!",vbInformation,"Information")
						GoTo Got
					End If
					X(m) = Obj_1D.GetX(m)
	    		Next

	    		Dim dB_results As Object

				Set dB_results = Result1D("")

				'Set oPlotMaterialLoss = Result1D("")

				'Plot metal loss

				For m = 1 To N_points-1
					dB_results.AppendXY(X(m), Log(magnitude(m))/Log(10)*10)
				Next
				dB_results.Xlabel("Frequecy/GHz")
				'oPlotMaterialLoss(n).ylabel("Loss in "+ Left(metalList,InStr(metalList,"$")-1)+"/dB" )
				dB_results.Title("Imported data in dB")
				dB_results.Ylabel("dB" )
				dB_results.Save("x_"+sigName)
				dB_results.AddToTree(sselectedItem)
				Got:
				sselectedItem = ResultTree.GetNextItemName(sselectedItem)
			Wend
		End If
		selectedItem = GetNextSelectedTreeItem
    Wend
End Sub

Function HasChildren( Item As String ) As Boolean

	Dim Name As String
	Dim sChild As String

	Name = Item
	sChild = ResultTree.GetFirstChildName ( Name )
	If sChild = "" Then
		HasChildren = False
	Else
		HasChildren = True
	End If

End Function
