' Calculate Power loss in dB

'20220828-By Shawn in COROS
'Public FirstChildItem As String
Option Explicit
Sub Main ()
  'Efficiency results parent path


	Dim selectedItem As String
	Dim curveLabel As String
	Dim index As Integer


	selectedItem = GetSelectedTreeItem
	If (InStr(selectedItem,"1D Results") = 0) Then
        MsgBox("Please select at least a 1D results before runing this macro.",vbCritical,"Warning")
        Exit All
    End If

    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long

    If HasChildren(selectedItem) Then
    	nResults = Resulttree.GetTreeResults(selectedItem,"0D/1D","",paths,types,files,info)

		Dim n As Long

		For n = 0 To nResults-1

			ReportInformationToWindow("path: " + CStr(paths(n)) + vbCrLf  + "type: " + CStr(types(n)) + vbCrLf + "file: " + CStr(files(n)))

			If types(n) = "XYSIGNAL" And Left(paths(n),InStrRev(paths(n),"\")-1) = selectedItem Then
				'curveLabel = Plot1D.GetCurveLabelOfCurveIndex(0)
				curveLabel = Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\"))
			   With Plot1D
			      index =.GetCurveIndexOfCurveLabel(curveLabel)
			      ReportInformationToWindow("The curve index is "+CStr(index))
			      'index =.GetCurveIndexOfCurveLabel("S1,1")
			     .SetLineStyle(index,"Solid",3) ' thick dashed line
			     .SetFont("Tahoma","bold","16")
			     '.SetLineColor(index,255,255,0)  ' yellow
			     .Plot ' make changes visible
				End With

			End If


		Next

	Else
		curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))

	   With Plot1D
	      index =.GetCurveIndexOfCurveLabel(curveLabel)
	     .SetLineStyle(index,"Solid",3) ' thick dashed line
	     .SetFont("Tahoma","bold","16")
	     '.SetLineColor(index,255,255,0)  ' yellow
	     .Plot ' make changes visible
		End With

    End If

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
