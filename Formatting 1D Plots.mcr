' Calculate Power loss in dB

'20220828-By Shawn in COROS
'Public FirstChildItem As String
Option Explicit
Sub Main ()
  'Efficiency results parent path


	Dim selectedItem As String
	Dim sselectedItem As String
	Dim curveLabel As String
	Dim index As Integer


	selectedItem = GetSelectedTreeItem
	'sselectedItem = Resulttree.GetFirstChildName("1D Results\CP directivity\fff")
	While selectedItem <> ""
		If (InStr(selectedItem,"1D Results") = 0) Then
	        MsgBox("Please select at least a 1D results before runing this macro.",vbCritical,"Warning")
	        Exit All
	    End If

	    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long

	    If HasChildren(selectedItem) Then
	    	sselectedItem = Resulttree.GetFirstChildName(selectedItem)
	    	Dim curveIndex As Integer
	    	curveIndex = -1
			While sselectedItem <> ""
				If (Resulttree.GetResultTypeFromItemName(sselectedItem) = _
	    		"xysignal" Or Resulttree.GetResultTypeFromItemName(sselectedItem) = "farfieldpolar") Then
					curveLabel = Right(sselectedItem,Len(sselectedItem)-InStrRev(sselectedItem,"\"))
					curveIndex = curveIndex +1
				   With Plot1D
				      'index =.GetCurveIndexOfCurveLabel(curveLabel)
				      index = curveIndex
				      'ReportInformationToWindow("The above curve index is "+CStr(index))
				      'index =.GetCurveIndexOfCurveLabel("S1,1")
				     .SetLineStyle(index,"Solid",3) ' thick dashed line
				     .SetFont("Tahoma","bold","14")
				     '.SetLineColor(index,255,255,0)  ' yellow
				     .Plot ' make changes visible
					End With
				End If
				sselectedItem = Resulttree.GetNextItemName(sselectedItem)
			Wend

		Else
			curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))

		   With Plot1D
		      index =.GetCurveIndexOfCurveLabel(curveLabel)
		     .SetLineStyle(index,"Solid",3) ' thick dashed line
		     .SetFont("Tahoma","bold","14")
		     '.SetLineColor(index,255,255,0)  ' yellow
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
