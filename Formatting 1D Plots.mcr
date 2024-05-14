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

	selectedItem = GetSelectedTreeItem
	'selectedItem = GetNextSelectedTreeItem
	While selectedItem <> ""

		If (InStr(selectedItem,"1D Results") = 0) Then
	        MsgBox("Please select at least a 1D results before runing this macro.",vbCritical,"Warning")
	        Exit All
	    End If

	    'Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	    Dim isSmith As Boolean
	    isSmith = False

	    If HasChildren(selectedItem) Then

	    	sselectedItem = Resulttree.GetFirstChildName(selectedItem)

			While sselectedItem <> ""
				If HasChildren(sselectedItem)=False Then
				'If (Resulttree.GetResultTypeFromItemName(sselectedItem) = "xysignal" _
				'Or Resulttree.GetResultTypeFromItemName(sselectedItem) = "farfieldpolar" _
				'Or Resulttree.GetResultTypeFromItemName(sselectedItem) = "table") Then

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
				sselectedItem = Resulttree.GetNextItemName(sselectedItem)
			Wend

		Else

			curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))
			SelectTreeItem(Left(selectedItem,InStrRev(selectedItem,"\")-1))

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
			     .SetLineStyle(index,"Solid",3)
			     .SetFont("Tahoma","bold","14")
			     '.SetLineColor(index,255,255,0)  ' yellow
			     If isSmith = True Then
			     	.PlotView("smith")
			     End If
			     .Plot ' make changes visible
			End With

	    End If
		selectedItem = GetNextSelectedTreeItem
    Wend

End Sub

Function HasChildren( Item As String ) As Boolean

	Dim Name As String
	Dim sChild As String

	Name = Item
	sChild = Resulttree.GetFirstChildName ( Name )
	If sChild = "" Then
		HasChildren = False
	Else
		HasChildren = True
	End If

End Function
