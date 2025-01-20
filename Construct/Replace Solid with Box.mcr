' ConstructBox
'Replace the selected solids with rectangular box to reduce mesh numbers and simulation time

Sub Main ()
	Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
	Dim sn As Integer, i As Integer, m As Integer
	Dim sname As String, SelectedItem As String, tmpname As String, compname As String, solidname As String
	Dim smaterial As String
	Dim historycontent As String, historyname As String

	'MsgBox("Please Select solids you want to replace with box",vbInformation,"Attention")
	On Error GoTo Message

	sn = GetNumberOfSelectedTreeItems
    SelectedItem = GetSelectedTreeItem
	'WCS.ActivateWCS("global")
	While SelectedItem <> ""
		tmpname =Replace(Right(SelectedItem,Len(SelectedItem)-InStr(SelectedItem,"\")),"\","/")
		i = InStrRev(tmpname,"/")
		If i = 0 Then
			MsgBox("No solids are selected!!",vbCritical,"Warning")
			Exit All
		End If
		group_sname = GetQualifiedNameFromTreeName(SelectedItem)
		sname = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
		'tmpname1 = Left(tmpname,i-1)
		'tmpname2 =
        'tmpname2 = Replace(Right(tmpname,Len(tmpname)-i+1),"/",":")
        'ssname = Replace(tmpname,"/",":",1,1)
        'sname = Left(tmpname,i-1)+Replace(Right(tmpname,Len(tmpname)-i+1),"/",":")
        smaterial = Solid.GetMaterialNameForShape(sname)
        compname = Left(sname,InStr(sname,":")-1)
        solidname = Right(sname,Len(sname)-InStr(sname,":"))+"_box"
		Solid.GetLooseBoundingBoxOfShape(sname,xmin,xmax,ymin,ymax,zmin,zmax)

		historycontent = ""
		historycontent = historycontent + "WCS.ActivateWCS("""+"global"+""")"+ vbLf
		historycontent = historycontent + "With Brick" + vbLf
		historycontent = historycontent + "   .Reset" + vbLf
		historycontent = historycontent + "   .Name("""+ solidname + """)" + vbLf
		historycontent = historycontent + "   .Component("""+ compname +""")" + vbLf
		historycontent = historycontent + "   .Material(""" + smaterial + """)" + vbLf
		historycontent = historycontent + "   .Xrange(" + CStr(xmin) + "," + CStr(xmax) + ")" + vbLf
		historycontent = historycontent + "   .Yrange(" + CStr(ymin) + "," + CStr(ymax) + ")" + vbLf
		historycontent = historycontent + "   .Zrange(" + CStr(zmin) + "," + CStr(zmax) + ")" + vbLf
		historycontent = historycontent + "   .Create" + vbLf
		historycontent = historycontent + "End With" + vbLf
        historycontent = historycontent + "WCS.ActivateWCS("""+"local"+""")"+ vbLf
		historyname = "Create solid box: " + solidname
		AddToHistory(historyname,historycontent)

	    historycontent = ""
	    historycontent = historycontent + "Solid.ChangeMaterial """ + sname + """,""Vacuum"""+vbLf
	    'Solid.ChangeMaterial "4 FPCs/LPM013M461B_FPC:LPM013M461B_FPC", "Vacuum"
	    'historycontent = historycontent + "Group.AddItem ""solid$" + sname + """,""Excluded from Simulation""" + vbLf
	    historycontent = historycontent + "Group.AddItem """ + group_sname + """,""Excluded from Simulation""" + vbLf
	    'Group.AddItem "solid$component1:xieh06F0DAF46v8c_1_box", "Excluded from Simulation"
	    historyname = "Exclude" + " " + sname + " " + "from Simulation"
		AddToHistory(historyname,historycontent)

		SelectedItem = GetNextSelectedTreeItem

	Wend

	'Mesh.Update

	'WCS.ActivateWCS("local")
	Exit All

	Message:
           MsgBox("No solids are selected!!",vbCritical,"Warning")
		   Exit All



End Sub
