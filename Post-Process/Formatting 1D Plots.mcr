' Format 1d plot

'20230614-By Shawn in COROS
'Public FirstChildItem As String
Option Explicit
Public app As String
Sub Main ()
  'Efficiency results parent path


	Dim selectedItem As String
	Dim sselectedItem As String
	Dim curveLabel As String
	Dim index As Integer
	Dim selectedItems(1000) As Variant
	Dim n As Integer, i As Integer

	Dim ItemName As String
	'Dim app As String

	app = Left(GetApplicationName, 2)

	If app = "DS" Then
		selectedItem =DS.GetSelectedTreeItem
	Else
		selectedItem =GetSelectedTreeItem
	End If

	'selectedItem = GetNextSelectedTreeItem
	While selectedItem <> ""
		If app = "DS" Then
			If (InStr(selectedItem,"Tasks\") = 0) Then
		        MsgBox("Please select at least a 1D results before runing this macro.",vbCritical,"Warning")
		        Exit All
		    End If
		Else
			If (InStr(selectedItem,"1D Results") = 0) Then
		        MsgBox("Please select at least a 1D results before runing this macro.",vbCritical,"Warning")
		        Exit All
		    End If
		End If


	    'Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	    Dim isSmith As Boolean
	    isSmith = False

	    If HasChildren(selectedItem) Then
			If app = "DS" Then
				sselectedItem = DSResultTree.GetFirstChildName(selectedItem)
			Else
				sselectedItem = ResultTree.GetFirstChildName(selectedItem)
			End If


			While sselectedItem <> ""
				If HasChildren(sselectedItem)=False Then
				'If (Resulttree.GetResultTypeFromItemName(sselectedItem) = "xysignal" _
				'Or Resulttree.GetResultTypeFromItemName(sselectedItem) = "farfieldpolar" _
				'Or Resulttree.GetResultTypeFromItemName(sselectedItem) = "table") Then
					If app = "DS" Then
						DS.SetPlotStyleForTreeItem(sselectedItem,"linetype=Solid linewidth=4")
					Else
						'SetPlotStyleForTreeItem(sselectedItem,"color=177;1;165 linetype=Dotted linewidth=8")
						curveLabel = Right(sselectedItem,Len(sselectedItem)-InStrRev(sselectedItem,"\"))

					    With Plot1D
					    	index =.GetCurveIndexOfCurveLabel(curveLabel)
					      	If index = -1 Then
					      		.PlotView("magnitudedb")
					      		isSmith = True
					      		index =.GetCurveIndexOfCurveLabel(curveLabel)
					      	End If
							If index = -1 Then
								Exit All
							End If
					      'ReportInformationToWindow("The above curve index is "+CStr(index))
					      'index =.GetCurveIndexOfCurveLabel("S1,1")
					     	.SetLineStyle(index,"Solid",3) ' thick dashed line in while
					     	.SetFont("Tahoma","bold","14")
					     '.SetLineColor(index,255,255,0)  ' yellow
					     	If isSmith = True Then
					     		.PlotView("smith")
					     	End If
					    	 	.Plot ' make changes visible
						End With
					End If
				End If
				If app = "DS" Then
					sselectedItem = DSResultTree.GetNextItemName(sselectedItem)
				Else
					sselectedItem = ResultTree.GetNextItemName(sselectedItem)
				End If
			Wend

		Else
			If app = "DS" Then
				DS.SetPlotStyleForTreeItem(selectedItem,"linetype=Solid linewidth=4")
			Else
				n=GetNumberOfSelectedTreeItems
				For i=0 To n-1
					selectedItems(i)=selectedItem
					selectedItem = GetNextSelectedTreeItem
				Next
				'If n>1 Then
					'SelectTreeItem(Left(selectedItem,InStrRev(selectedItem,"\")-1))
				'End If
				For i=0 To n-1
					SelectTreeItem(selectedItems(i))
					curveLabel = Right(selectedItems(i),Len(selectedItems(i))-InStrRev(selectedItems(i),"\"))
					With Plot1D
					      index =.GetCurveIndexOfCurveLabel(curveLabel)
					      If index = -1 Then
					      	.PlotView("magnitudedb")
					      	isSmith = True
					      	index =.GetCurveIndexOfCurveLabel(curveLabel)
					      End If
					      If index = -1 Then
					      	Exit All
					      End If
					      'selectedItems(n)=selectedItem
						  'n=n+1
					      'ReportInformationToWindow("The above curve index is "+CStr(index))
					      'index =.GetCurveIndexOfCurveLabel("S1,1")
					     .SetLineStyle(index,"Solid",3)
					     .SetFont("Tahoma","bold","14")
					     '.SetLineColor(index,255,255,0)  ' yellow
					     If isSmith = True Then
					     	.PlotView("smith")
					     End If
					     .Plot ' make changes visible
					End With
				Next
			End If
	    End If
		selectedItem = GetNextSelectedTreeItem
    Wend
End Sub

Function HasChildren( Item As String ) As Boolean

	Dim xName As String
	Dim sChild As String

	xName = Item
	If app = "DS" Then
		sChild = DSResultTree.GetFirstChildName ( xName )
	Else
		sChild = ResultTree.GetFirstChildName ( xName )
	End If
	If sChild = "" Then
		HasChildren = False
	Else
		HasChildren = True
	End If

End Function
