'#Language "WWB-COM"

Option Explicit

Sub Main

    If MsgBox("Please copy the target solid to a new project and then run this macro!!!",vbYesNo,"Notice") <> vbYes Then
    	Exit Sub
    End If


	Begin Dialog UserDialog 410,175,"Slice Parameters"' %GRID:10,7,1,1
		Text 20,7,270,14,"Pleas input slice dimensions:",.Text1
		Text 20,28,90,14,"xStep:",.Text2
		Text 20,49,90,14,"yStep:",.Text3
		Text 20,70,90,14,"zStep:",.Text4
		TextBox 70,28,50,14,.xStep
		TextBox 70,49,50,14,.yStep
		TextBox 70,70,50,14,.zStep
		OKButton 30,140,90,21
		CancelButton 140,140,90,21
		Text 20,91,280,14,"Or input slice angle for annular structures:",.Text5
		Text 20,112,90,14,"anStep:",.Text6
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

	Dim sn As Integer
	Dim sName As String,fullName As String, ComponentName As String
	Dim i As Integer
	Dim wStep As Double
	Dim Axis As Integer
	Dim xStep As Double, yStep As Double, zStep As Double
	Dim anStep As Double

	xStep = Evaluate(dlg.xStep)
    yStep = Evaluate(dlg.yStep)
    zStep = Evaluate(dlg.zStep)
    anStep = Evaluate(dlg.anStep)
    'Parse input parameters
    If xStep = 0 And yStep = 0 And zStep = 0 And anStep = 0 Then
    	MsgBox("Invalid parameters, please re-run this macro and input valid parameters!!",vbCritical,"Error")
    	Exit All
    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep <> 0) Then
    	MsgBox("Too many parameters, please re-run this macro and input valid parameters!!",vbCritical,"Error")
    	Exit All
    'Slice by dimension steps
    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep = 0) Then
    	AutoSliceBySteps(xStep,yStep,zStep)
    'Slice by angles
    ElseIf (anStep <> 0) And (xStep = 0  And yStep = 0  And zStep = 0) Then
    	AutoSliceByAngle(anStep)
    End If
    WCS.ActivateWCS("global")


	
End Sub

Function AutoSliceAlongAxis(fName As String, sStep As Double, sAxis As Integer, xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double)
	Dim sName As String, CompName As String
	Dim xcut As Double, ycut As Double,zcut As Double
	Dim Steps As Integer
	Dim n As Integer
    WCS.ActivateWCS("local")
	sName = Right(fName,Len(fName)-InStr(fName,":"))
	CompName = Left(fName,InStr(fName,":")-1)
	Select Case sAxis
	Case 0
		Steps =  CInt((xMax-xMin)/sStep)
		xcut = xMin
		WCS.SetNormal(1,0,0)
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				xcut = xcut + sStep
				WCS.SetOrigin(xcut,yMin,zMin)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case 1
		Steps =  CInt((yMax-yMin)/sStep)
		ycut = yMin
		WCS.SetNormal(0,1,0)
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				ycut = ycut + sStep
				WCS.SetOrigin(xMin,ycut,zMin)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case 2
		Steps =  CInt((zMax-zMin)/sStep)
		zcut = zMin
		WCS.SetNormal(0,0,1)
		If Steps > 1 Then
			For n = 1 To Steps STEP 1
				zcut = zcut + sStep
				WCS.SetOrigin(xMin,yMin,zcut)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case Else
		Exit Function
	End Select
End Function
Function AutoSliceBySteps(xStep As Double,yStep As Double,zStep As Double)
	Dim fullName As String
	Dim sn As Integer, i As Integer
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	Dim gxMin As Double, gxMax As Double, gyMin As Double, gyMax As Double, gzMin As Double, gzMax As Double
	gxMin = 1e200
	gxMax = -1e200
	gyMin = 1e200
	gyMax = -1e200
	gzMin = 1e200
	gzMax = -1e200
	'When xStep is not less than wStep, slice once with step of wStep
	sn = Solid.GetNumberOfShapes
	For i = 0 To sn-1 STEP 1
		fullName = Solid.GetNameOfShapeFromIndex(i)
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
	Next
	If xStep <> 0 Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,xStep,0,xMin, gxMax, gyMin, gyMax, gzMin, gzMax)
	    	Next i
	    End If
	End If
	'When yStep is not less than wStep, slice once with step of wStep
	If yStep <> 0 Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,xStep,1,xMin, gxMax, gyMin, gyMax, gzMin, gzMax)
	    	Next i
	    End If
	End If
	'When zStep is not less than wStep, slice once with step of wStep
	If zStep <> 0 Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,xStep,2,xMin, gxMax, gyMin, gyMax, gzMin, gzMax)
	    	Next i
	    End If
	End If

End Function

Function AutoSliceByAngle(anStep As Double)

	Dim sName As String, CompName As String, fName As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	Dim deltaxy As Double,deltayz As Double,deltaxz As Double
	Dim Axis As String
	Dim Angle As Double
	Dim xcenter As Double,ycenter As Double,zcenter As Double
	Dim sn As Integer, i As Integer

	sn = Solid.GetNumberOfShapes
	fName = Solid.GetNameOfShapeFromIndex(0)
	WCS.ActivateWCS("local")
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
		WCS.SetNormal(0,0,1)
		WCS.SetOrigin(xcenter,ycenter,zMin)

	ElseIf deltaxz < deltayz And deltaxz < deltaxy Then
		Axis = "y"
		xcenter = (xMax+xMin)/2
		zcenter = (zMax+zMin)/2
		WCS.SetNormal(0,1,0)
		WCS.SetOrigin(xcenter,yMin,zcenter)
	ElseIf deltayz < deltaxy And deltayz < deltaxz Then
		Axis = "x"
		ycenter = (yMax+yMin)/2
		zcenter = (zMax+zMin)/2
		WCS.SetNormal(1,0,0)
		WCS.SetOrigin(xMin,ycenter,zcenter)
	End If
	'Initialize the total rotated angle, this angle should be less than 180 degree
	Angle = 0
	WCS.RotateWCS("u",90)
	While Angle <= 180
		sn = Solid.GetNumberOfShapes
		For i = 0 To sn-1 STEP 1
			fName = Solid.GetNameOfShapeFromIndex(i)
			sName = Right(fName,Len(fName)-InStr(fName,":"))
			CompName = Left(fName,InStr(fName,":")-1)
			Solid.SliceShape(sName,CompName)
		Next
		Angle = Angle + anStep
		WCS.RotateWCS("v",anStep)

	Wend

End Function
