'#Language "WWB-COM"
'#include "vba_globals_all.lib"

Option Explicit

Sub Main


    If MsgBox("Please select a component which only contains a solid.",vbYesNo,"Notice") <> vbYes Then
    	Exit Sub
    End If

	Dim sn As Integer
	Dim sName As String,fullName As String, sComponent As String
	Dim i As Integer
	Dim wStep As Double
	Dim Axis As Integer
	Dim xStep As Double, yStep As Double, zStep As Double
	Dim anStep As Double
	Dim sSolid As String
	Dim isWCS As Boolean
	Dim sCommand As String

	isWCS=IIf(WCS.IsWCSActive()="local",True,False)
	If isWCS=True Then
		WCS.Store("originalWCS")
	End If
	'Dim fullName As String
	'fullName = Solid.GetNameOfShapeFromIndex(0)
	'fullName = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")


	sComponent = GetSelectedTreeItem
	If sComponent = "" Then
		MsgBox "No components is selected!!", vbCritical, "Error"
		Exit All
	ElseIf HasChildren(sComponent) = False Then
		MsgBox "The selected item is not a component Or No shapes is contained in the selected component!!", vbCritical, "Error"
		Exit All
	ElseIf Resulttree.GetNextItemName(Resulttree.GetFirstChildName(sComponent)) <> "" Then
		MsgBox "There are at leat two shapes in the selected component!!",vbCritical,"Error"
		Exit All
	ElseIf HasChildren(Resulttree.GetFirstChildName(sComponent)) = True Then
		MsgBox "At least one component is contained in the selected componnent!!",vbCritical,"Error"
		Exit All
	End If

	'sSolid = Resulttree.GetFirstChildName(sComponent)
	'ReDim Preserve solidArray(n)

	Begin Dialog UserDialog 410,175,"Along z-axis" ' %GRID:10,7,1,1
		Text 20,7,270,14,"Pleas input slice dimensions:",.Text1
		Text 20,28,90,14,"Along x-axis:",.Text2
		Text 20,49,90,14,"Along y-axis:",.Text3
		Text 20,70,90,14,"Along z-axis:",.Text4
		TextBox 110,28,50,14,.xStep
		TextBox 110,49,50,14,.yStep
		TextBox 110,70,50,14,.zStep
		OKButton 30,140,90,21
		CancelButton 140,140,90,21
		Text 20,91,280,14,"Or input slice angle for annular structures:",.Text5
		Text 20,112,60,14,"Angle:",.Text6
		TextBox 80,112,50,14,.anStep
	End Dialog
	Dim dlg As UserDialog
	'Dialog dlg
	dlg.xStep = "0.0"
	dlg.yStep = "0.0"
	dlg.zStep = "0.0"
	dlg.anStep = "0.0"

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If


	xStep = Evaluate(dlg.xStep)
    yStep = Evaluate(dlg.yStep)
    zStep = Evaluate(dlg.zStep)
    anStep = Evaluate(dlg.anStep)

    'Default slice step or initial slice step
    wStep = 6
    'Parse input parameters
    If xStep = 0 And yStep = 0 And zStep = 0 And anStep = 0 Then
    	MsgBox("Invalid parameters, please re-run this macro and input valid parameters!!",vbCritical,"Error")
    	Exit All
    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep <> 0) Then
    	MsgBox("Too many parameters, please re-run this macro and input valid parameters!!",vbCritical,"Error")
    	Exit All
    'Slice by dimension steps
    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep = 0) Then
    	AutoSliceBySteps(sComponent,xStep,yStep,zStep,wStep)
    'Slice by angles
    ElseIf (anStep <> 0) And (xStep = 0  And yStep = 0  And zStep = 0) Then
    	AutoSliceByAngle(sComponent,anStep)
    End If
    'WCS.ActivateWCS("global")
    'IIf(isWCS, WCS.ActivateWCS("local"),WCS.ActivateWCS("global"))
     If isWCS = True Then
	 	'WCS.ActivateWCS("local")
		sCommand="WCS.ActivateWCS ""local"""
		AddToHistory "Activate local coordinates system", sCommand
		WCS.Restore("originalWCS")
	 Else
	 	sCommand="WCS.ActivateWCS ""global"""
		AddToHistory "Activate global coordinates system", sCommand

	 End If


	
End Sub

Function AutoSliceAlongAxis(fName As String, sStep As Double, sAxis As Integer, xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double)
	Dim sName As String, CompName As String
	Dim xcut As Double, ycut As Double,zcut As Double
	Dim Steps As Integer
	Dim n As Integer
	Dim sCommand As String
	Dim commandName As String
	sCommand = ""
    'WCS.ActivateWCS("local")
	sName = Right(fName,Len(fName)-InStr(fName,":"))
	'sName = Right(fName,Len(fName)-InStrRev(fName,"\"))
	CompName = Left(fName,InStr(fName,":")-1)

	Select Case sAxis
	Case 0
		'commandName = "slice shape"+fName+" by dimensions along x axis with step of" + Cstr(sStep)
		Steps =  CInt((xMax-xMin)/sStep)
		xcut = xMin
		'WCS.SetNormal(1,0,0)
		'sCommand = sCommand + "WCS.SetNormal ""1"", ""0"", ""0""" + vbLf
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				xcut = xcut + sStep
				'WCS.SetOrigin(xcut,ymin,zmin)
				sCommand = sCommand + "WCS.SetOrigin """+CStr(xcut)+""","""+CStr(yMin)+""","""+Cstr(zMin)+""""+vbLf
				'Solid.SliceShape(sName,CompName)
				sCommand = sCommand + "Solid.SliceShape """ + sName+""","""+CompName+""""+vbLf
			Next
		Else
			sCommand = ""
		End If

	Case 1
		'commandName = "slice shape"+fName+" by dimensions along y axis with step of" + Cstr(sStep)
		Steps =  CInt((yMax-yMin)/sStep)
		ycut = yMin
		'WCS.SetNormal(0,1,0)
		'sCommand = sCommand + "WCS.SetNormal ""0"", ""1"", ""0""" + vbLf
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				ycut = ycut + sStep
				'WCS.SetOrigin(xmin,ycut,zmin)
				sCommand = sCommand + "WCS.SetOrigin """+CStr(xMin)+""","""+CStr(ycut)+""","""+Cstr(zMin)+""""+vbLf
				Solid.SliceShape(sName,CompName)
				sCommand = sCommand + "Solid.SliceShape """ + sName+""","""+CompName+""""+vbLf
			Next
		Else
			sCommand = ""
		End If

	Case 2
		'commandName = "slice shape"+fName+" by dimensions along z axis with step of" + Cstr(sStep)
		Steps =  CInt((zMax-zMin)/sStep)
		zcut = zMin
		'WCS.SetNormal(0,0,1)
		'sCommand = sCommand + "WCS.SetNormal ""0"", ""0"", ""1""" + vbLf
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				zcut = zcut + sStep
				'WCS.SetOrigin(xmin,ymin,zcut)
				sCommand = sCommand + "WCS.SetOrigin """+CStr(xMin)+""","""+CStr(yMin)+""","""+Cstr(zcut)+""""+vbLf
				'Solid.SliceShape(sName,CompName)
				sCommand = sCommand + "Solid.SliceShape """ + sName+""","""+CompName+""""+vbLf
			Next
		Else
			sCommand = ""
		End If

	Case Else
		Exit Function
	End Select
	'AddToHistory(commandName,sCommand)
	AutoSliceAlongAxis = sCommand
End Function
Function AutoSliceBySteps(sComponent As String, xStep As Double,yStep As Double,zStep As Double,wStep As Double)
	Dim fullName As String
	Dim path As String
	Dim sn As Integer, i As Integer, n As Integer
	Dim sCommand As String, commandName As String, tCommand As String
	Dim isSliced As Boolean
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	Dim gxMin As Double, gxMax As Double, gyMin As Double, gyMax As Double, gzMin As Double, gzMax As Double
	Dim solidsInComponent() As String, group_sname As String
	gxMin = 1e200
	gxMax = -1e200
	gyMin = 1e200
	gyMax = -1e200
	gzMin = 1e200
	gzMax = -1e200
	'When xStep is not less than wStep, slice once with step of wStep
	solidsInComponent=getChildSolidsInComponent(sComponent, n)
	'sn = Solid.GetNumberOfShapes
	For i = 0 To n-1 STEP 1
		'fullName = Solid.GetNameOfShapeFromIndex(i)
		fullName=solidsInComponent(i)
		group_sname = GetQualifiedNameFromTreeName(fullName)
		fullName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
		'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
		'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
		Solid.GetLooseBoundingBoxOfShape(fullName,xMin,xMax,yMin,yMax,zMin,zMax)
		If xMin < gxMin Then
			gxMin = xMin
		End If
		If xMax > gxMax Then
			gxMax = xMax
		End If
		If yMin < gyMin Then
			gyMin = yMin
		End If
		If yMax > gyMax Then
			gyMax = yMax
		End If
		If zMin < gzMin Then
			gzMin = zMin
		End If
		If zMax > gzMax Then
			gzMax = zMax
		End If
		'End If
	Next

	isSliced = False
	sCommand = ""
	If WCS.IsWCSActive() = "global" Then
		commandName = "Set local WCS"
		sCommand = sCommand + "WCS.ActivateWCS ""local""" + vbLf
		AddToHistory(commandName,sCommand)
		sCommand = ""
	End If

	'commandName = "Slice Shape in " +sComponent+" with xStep of "+Cstr(xStep)+", ystep of "+CStr(yStep)+" and zStep of "+ CStr(zStep)
	'When xStep is not less than wStep, slice once with step of xStep
	If xStep <> 0 Then
		commandName = "Slice Shape in " +sComponent+" with xStep of "+Cstr(xStep)
		sCommand = sCommand + "WCS.SetNormal ""1"", ""0"", ""0""" + vbLf
		'sn =  Solid.GetNumberOfShapes
		solidsInComponent=getChildSolidsInComponent(sComponent, n)
	    If n > 0 Then
	    	For i = 0 To n-1 STEP 1
				'fullName = Solid.GetNameOfShapeFromIndex(i)
				fullName=solidsInComponent(i)

				group_sname = GetQualifiedNameFromTreeName(fullName)
				fullName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
				tCommand = AutoSliceAlongAxis(fullName,xStep,0,gxMin, gxMax, gyMin, gyMax, gzMin, gzMax)
				If tCommand <>""Then
					sCommand = sCommand + tCommand
					isSliced = True
				End If
				'End If

	    	Next i
	    End If
	    If isSliced = True Then
			AddToHistory(commandName,sCommand)
		    sCommand = ""
		    isSliced = False
	    End If
	End If

	'When yStep is not less than wStep, slice once with step of yStep
	If yStep <> 0 Then
		sCommand = sCommand + "WCS.SetNormal ""0"", ""1"", ""0""" + vbLf
		'sn =  Solid.GetNumberOfShapes
		solidsInComponent=getChildSolidsInComponent(sComponent, n)
		commandName = "Slice Shape in " +sComponent+" with yStep of "+Cstr(yStep)
	    If n > 0 Then
	    	For i = 0 To n-1 STEP 1
				'fullName = Solid.GetNameOfShapeFromIndex(i)
				fullName=solidsInComponent(i)

				group_sname = GetQualifiedNameFromTreeName(fullName)
				fullName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
				tCommand = AutoSliceAlongAxis(fullName,yStep,1,gxMin, gxMax, gyMin, gyMax, gzMin, gzMax)
				If tCommand <>""Then
					sCommand = sCommand + tCommand
					isSliced = True
				End If
				'End If
	    	Next i
	    End If
	    If isSliced = True Then
			AddToHistory(commandName,sCommand)
		    sCommand = ""
		    isSliced = False
	    End If
	End If

	'When zStep is not less than wStep, slice once with step of zStep
	If zStep <> 0 Then
		sCommand = sCommand + "WCS.SetNormal ""0"", ""0"", ""1""" + vbLf
		'sn =  Solid.GetNumberOfShapes
		solidsInComponent=getChildSolidsInComponent(sComponent, n)
		commandName = "Slice Shape in " +sComponent+" with zStep of "+Cstr(zStep)
	    If n > 0 Then
	    	For i = 0 To n-1 STEP 1
				'fullName = Solid.GetNameOfShapeFromIndex(i)
				fullName=solidsInComponent(i)

				group_sname = GetQualifiedNameFromTreeName(fullName)
				fullName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
				tCommand = AutoSliceAlongAxis(fullName,zStep,2,gxMin, gxMax, gyMin, gyMax, gzMin, gzMax)
				If tCommand <>""Then
					sCommand = sCommand + tCommand
					isSliced = True
				End If
				'End If
	    	Next i
	    End If
	    If isSliced = True Then
			AddToHistory(commandName,sCommand)
		    sCommand = ""
		    isSliced = False
	    End If
	End If

End Function

Function AutoSliceByAngle(sComponent As String, anStep As Double)

	Dim sName As String, CompName As String, fName As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	Dim deltaxy As Double,deltayz As Double,deltaxz As Double
	Dim Axis As String
	Dim Angle As Double
	Dim xcenter As Double,ycenter As Double,zcenter As Double
	Dim sn As Integer, i As Integer, n As Integer
	Dim path As String
	Dim sCommand As String, commandName As String
	Dim group_sname As String
	Dim solidsInComponent() As String
	Dim maximalSolid As String
	Dim minimalSolid As String
	sCommand = ""

	'sn = Solid.GetNumberOfShapes
	fName = Resulttree.GetFirstChildName(sComponent)
	group_sname = GetQualifiedNameFromTreeName(fName)
	fName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
	'WCS.ActivateWCS("local")
	sCommand = sCommand + "WCS.ActivateWCS ""local""" + vbLf
	sName = Right(fName,Len(fName)-InStr(fName,":"))
	CompName = Left(fName,InStr(fName,":")-1)
	Solid.GetLooseBoundingBoxOfShape(fName,xMin,xMax,yMin,yMax,zMin,zMax)

	deltaxy = Abs(Abs(xMax-xMin)-Abs(yMax-yMin))
	deltaxz = Abs(Abs(xMax-xMin)-Abs(zMax-zMin))
	deltayz = Abs(Abs(yMax-yMin)-Abs(zMax-zMin))

	If deltaxy < deltayz And deltaxy < deltaxz Then
		Axis = "z"
		xcenter = (xMax+xMin)/2
		ycenter = (yMax+yMin)/2
		'WCS.SetNormal(0,0,1)
		sCommand = sCommand + "WCS.SetNormal ""0"", ""0"", ""1""" + vbLf
		'WCS.SetOrigin(xcenter,ycenter,zmin)
		sCommand = sCommand + "WCS.SetOrigin """+CStr(xcenter)+""","""+CStr(ycenter)+""","""+Cstr(zMin)+""""+vbLf

	ElseIf deltaxz < deltayz And deltaxz < deltaxy Then
		Axis = "y"
		xcenter = (xMax+xMin)/2
		zcenter = (zMax+zMin)/2
		'WCS.SetNormal(0,1,0)
		sCommand = sCommand + "WCS.SetNormal ""0"", ""1"", ""0""" + vbLf
		'WCS.SetOrigin(xcenter,ymin,zcenter)
		sCommand = sCommand + "WCS.SetOrigin """+CStr(xcenter)+""","""+CStr(yMin)+""","""+Cstr(zcenter)+""""+vbLf
	ElseIf deltayz < deltaxy And deltayz < deltaxz Then
		Axis = "x"
		ycenter = (yMax+yMin)/2
		zcenter = (zMax+zMin)/2
		'WCS.SetNormal(1,0,0)
		sCommand = sCommand + "WCS.SetNormal ""1"", ""0"", ""0""" + vbLf
		'WCS.SetOrigin(xmin,ycenter,zcenter)
		sCommand = sCommand + "WCS.SetOrigin """+CStr(xMin)+""","""+CStr(ycenter)+""","""+Cstr(zcenter)+""""+vbLf
	End If
	'Initialize the total rotated angle, this angle should be less than 180 degree
	Angle = 0
	'WCS.RotateWCS("u",90)
	sCommand = sCommand + "WCS.RotateWCS ""u"", ""90.0"""+vbLf
	While Angle < 180
		'sn = Solid.GetNumberOfShapes
		solidsInComponent=getChildSolidsInComponent(sComponent, n)
		getSolidsWithExtremalNameLength(solidsInComponent, n, minimalSolid, maximalSolid)
		'For i = 0 To n-1 STEP 1
			'fName = Solid.GetNameOfShapeFromIndex(i)
			'fName=solidsInComponent(i)
			fName=minimalSolid
			'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
			'path = Replace(Left(fName,InStr(fName,":")-1),"/","\")
			'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
			sName = Right(fName,Len(fName)-InStrRev(fName,"\"))
			'CompName = Left(fName,InStr(fName,":")-1)
			'Solid.SliceShape(sName,CompName)
			sCommand = sCommand + "Solid.SliceShape """ + sName+""","""+CompName+""""+vbLf
			If StrComp(maximalSolid,minimalSolid) <> 0 Then
				fName=maximalSolid
				'path = Replace(Left(fullName,InStr(fullName,":")-1),"/","\")
				'path = Replace(Left(fName,InStr(fName,":")-1),"/","\")
				'If Right(path,Len(path)-InStrRev(path,"\")) = Right(sComponent,Len(sComponent)-InStrRev(sComponent,"\")) Then
				sName = Right(fName,Len(fName)-InStrRev(fName,"\"))
				'CompName = Left(fName,InStr(fName,":")-1)
				'Solid.SliceShape(sName,CompName)
				sCommand = sCommand + "Solid.SliceShape """ + sName+""","""+CompName+""""+vbLf
			End If

		'Next

		'WCS.RotateWCS("v",anStep)
		commandName = "Slice Shape in " +sComponent+" with angle of "+ Cstr(Angle)
		AddToHistory(commandName,sCommand)
		Angle = Angle + anStep
		'Wait 0.1
		sCommand = ""
		sCommand = sCommand + "WCS.RotateWCS ""v"",""" + CStr(anStep) + """"+vbLf

	Wend

End Function
Function HasChildren( Item As String ) As Boolean

	Dim xName As String
	Dim sChild As String

	xName = Item
	sChild = Resulttree.GetFirstChildName ( xName )
	If sChild = "" Then
		HasChildren = False
	Else
		HasChildren = True
	End If

End Function

Function getChildSolidsInComponent(sComponent As String, n As Integer)
	Dim selectedTreeItems(1000) As String
	Dim tmpItem As String
	'Dim n As Integer
	n=0
	selectedTreeItems(n) = Resulttree.GetFirstChildName(sComponent)
	While selectedTreeItems(n) <> ""
		'selectTreeItem(selectedTreeItems(n))
		tmpItem = selectedTreeItems(n)
		n = n+1
		selectedTreeItems(n) = Resulttree.GetNextItemName(tmpItem)
	Wend
	'ReDim Preserve selectedTreeItems(n)
	getChildSolidsInComponent=selectedTreeItems
End Function
Function getSolidsWithExtremalNameLength(solids() As String, n As Integer, solidWithMinima As String, solidWithMaxima As String)
	Dim i As Integer
	'Dim n As Integer
	solidWithMinima=solids(0)
	solidWithMaxima=solids(0)
	For i=1 To n-1 STEP 1
		If Len(solids(i))>=Len(solidWithMaxima) Then
			solidWithMaxima=solids(i)
		End If
		If Len(solids(i))<=Len(solidWithMinima) Then
			solidWithMinima=solids(i)
		End If
	Next
End Function
