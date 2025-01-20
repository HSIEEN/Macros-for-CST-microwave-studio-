' ConstructCylinder
'Replace the selected solids with cylinder to reduce mesh numbers and simulation time

Sub Main ()
	Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
	Dim sn As Integer, i As Integer, m As Integer
	Dim sname As String, SelectedItem As String, tmpname As String, tmpname1 As String, tmpname2 As String, compname As String, solidname As String
	Dim smaterial As String
	Dim historycontent As String, historyname As String
	Dim Radius As Double,xcenter As Double, ycenter As Double, zcenter As Double
	Dim deltaxy As Double, deltaxz As Double, deltayz As Double
	Dim Axis As String

	'MsgBox("Please Select solids you want to replace with box",vbInformation,"Attention")
	On Error GoTo Message

	sn = GetNumberOfSelectedTreeItems
    SelectedItem = GetSelectedTreeItem
	WCS.ActivateWCS("global")
	While SelectedItem <> ""
		tmpname =Replace(Right(SelectedItem,Len(SelectedItem)-InStr(SelectedItem,"\")),"\","/")
		i = InStrRev(tmpname,"/")
		If i = 0 Then
			MsgBox("No solids are selected!!",vbCritical,"Warning")
			Exit All
		End If

		tmpname1 = Left(tmpname,i-1)
		tmpname2 = Right(tmpname,Len(tmpname)-i+1)
        tmpname2 = Replace(tmpname2,"/",":")
        sname = tmpname1 + tmpname2
        smaterial = Solid.GetMaterialNameForShape(sname)
        compname = Left(sname,InStr(sname,":")-1)
        solidname = Right(sname,Len(sname)-InStr(sname,":"))+"_Cylinder"
		Solid.GetLooseBoundingBoxOfShape(sname,xmin,xmax,ymin,ymax,zmin,zmax)

		deltaxy = Abs(Abs(xmax-xmin)-Abs(ymax-ymin))
		deltaxz = Abs(Abs(xmax-xmin)-Abs(zmax-zmin))
		deltayz = Abs(Abs(ymax-ymin)-Abs(zmax-zmin))
		If deltaxy < deltayz And deltaxy < deltaxz Then
			Axis = "z"
			Radius = Abs(ymax-ymin)/2
			xcenter = (xmax+xmin)/2
			ycenter = (ymax+ymin)/2
		ElseIf deltaxz < deltayz And deltaxz < deltaxy Then
			Axis = "y"
			Radius = Abs(xmax-xmin)/2
			xcenter = (xmax+xmin)/2
			zcenter = (zmax+zmin)/2

		ElseIf deltayz < deltaxy And deltayz < deltaxz Then
			Axis = "x"
			Radius = Abs(zmax-zmin)/2
			ycenter = (ymax+ymin)/2
			zcenter = (zmax+zmin)/2

		End If


		historycontent = ""
		historycontent = historycontent + "WCS.ActivateWCS("""+"global"+""")"+ vbLf
		historycontent = historycontent + "With Cylinder" + vbLf
		historycontent = historycontent + ".Reset" + vbLf
		historycontent = historycontent + ".Name("""+ solidname + """)" + vbLf
		historycontent = historycontent + ".Component("""+ compname +""")" + vbLf
		historycontent = historycontent + ".Material(""" + smaterial + """)" + vbLf
		historycontent = historycontent + ".OuterRadius(" + CStr(Radius) + ")" + vbLf
		historycontent = historycontent + ".InnerRadius(" + "0" + ")" + vbLf
		historycontent = historycontent + ".Axis(""" + Axis + """)" + vbLf
		Select Case Axis
		Case "x"
			historycontent = historycontent + ".Xrange(" + CStr(xmin)+"," + CStr(xmax)+ ")" + vbLf
			historycontent = historycontent + ".Ycenter(" + CStr(ycenter)+ ")" + vbLf
			historycontent = historycontent + ".Zcenter(" + CStr(zcenter)+")" + vbLf
		Case "y"
			historycontent = historycontent + ".Yrange(" + CStr(ymin)+"," + CStr(ymax) +")" + vbLf
			historycontent = historycontent + ".Xcenter(" + CStr(xcenter)+ ")" + vbLf
			historycontent = historycontent + ".Zcenter(" + CStr(zcenter)+")" + vbLf
		Case "z"
			historycontent = historycontent + ".Zrange(" + CStr(zmin)+"," + CStr(zmax) +")" + vbLf
			historycontent = historycontent + ".Xcenter(" + CStr(xcenter)+ ")" + vbLf
			historycontent = historycontent + ".Ycenter(" + CStr(ycenter)+")" + vbLf
		End Select
		historycontent = historycontent + ".Segments(0)" + vbLf
		historycontent = historycontent + ".Create" + vbLf
		historycontent = historycontent + "End With" + vbLf
        historycontent = historycontent + "WCS.ActivateWCS("""+"local"+""")"+ vbLf
		historyname = "Create solid box: " + solidname
		AddToHistory(historyname,historycontent)

	    historycontent = ""
	    historycontent = historycontent + "Solid.ChangeMaterial """ + sname + """,""Vacuum"""+vbLf
	    'Solid.ChangeMaterial "4 FPCs/LPM013M461B_FPC:LPM013M461B_FPC", "Vacuum"
	    historycontent = historycontent + "Group.AddItem ""solid$" + sname + """,""Excluded from Simulation""" + vbLf
	    'Group.AddItem "solid$component1:xieh06F0DAF46v8c_1_box", "Excluded from Simulation"
	    historyname = "Exclude" + " " + sname + " " + "from Simulation"
		AddToHistory(historyname,historycontent)

		SelectedItem = GetNextSelectedTreeItem

	Wend

	WCS.ActivateWCS("local")
	Exit All

	Message:
           MsgBox("No solids are selected!!",vbCritical,"Warning")
		   Exit All



End Sub
