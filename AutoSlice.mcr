'#Language "WWB-COM"

Option Explicit

Sub Main

    MsgBox("请将要切片的物体先拷贝到新工程里然后运行此脚本，切勿在原工程里运行此脚本！！！",vbInformation,"Notice")
	Begin Dialog UserDialog 410,175,"Step input:",.DialogFunction ' %GRID:10,7,1,1
		Text 20,7,270,14,"请输入切片尺寸（适用任意结构）：",.Text1
		Text 20,28,90,14,"xStep:",.Text2
		Text 20,49,90,14,"yStep:",.Text3
		Text 20,70,90,14,"zStep:",.Text4
		TextBox 70,28,50,14,.xStep
		TextBox 70,49,50,14,.yStep
		TextBox 70,70,50,14,.zStep
		OKButton 30,140,90,21
		CancelButton 140,140,90,21
		Text 20,91,280,14,"或者输入切片角度（适用环形结构）：",.Text5
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


	
End Sub

Function AutoSliceAlongAxis(fName As String, sStep As Double, sAxis As Integer)
	Dim sName As String, CompName As String
	Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
	Dim xcut As Double, ycut As Double,zcut As Double
	Dim Steps As Integer
	Dim n As Integer
    WCS.ActivateWCS("local")
	sName = Right(fName,Len(fName)-InStr(fName,":"))
	CompName = Left(fName,InStr(fName,":")-1)
	Solid.GetLooseBoundingBoxOfShape(fName,xmin,xmax,ymin,ymax,zmin,zmax)

	Select Case sAxis
	Case 0
		Steps =  CInt((xmax-xmin)/sStep)
		xcut = xmin
		WCS.SetNormal(1,0,0)
		If Steps > 0 Then
			For n = 1 To Steps STEP 1
				xcut = xcut + sStep
				WCS.SetOrigin(xcut,ymin,zmin)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case 1
		Steps =  CInt((ymax-ymin)/sStep)
		ycut = ymin
		WCS.SetNormal(0,1,0)
		If Steps > 0 Then
			For n = 1 To Steps STEP 1
				ycut = ycut + sStep
				WCS.SetOrigin(xmin,ycut,zmin)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case 2
		Steps =  CInt((zmax-zmin)/sStep)
		zcut = zmin
		WCS.SetNormal(0,0,1)
		If Steps > 0 Then
			For n = 1 To Steps STEP 1
				zcut = zcut + sStep
				WCS.SetOrigin(xmin,ymin,zcut)
				Solid.SliceShape(sName,CompName)
			Next
		End If

	Case Else
		Exit Function
	End Select
End Function

Function AutoSliceByAngle(anStep As Double)

	Dim sName As String, CompName As String, fName As String
	Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
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
	Solid.GetLooseBoundingBoxOfShape(fName,xmin,xmax,ymin,ymax,zmin,zmax)

	deltaxy = Abs(Abs(xmax-xmin)-Abs(ymax-ymin))
	deltaxz = Abs(Abs(xmax-xmin)-Abs(zmax-zmin))
	deltayz = Abs(Abs(ymax-ymin)-Abs(zmax-zmin))

	If deltaxy < deltayz And deltaxy < deltaxz Then
		Axis = "z"
		xcenter = (xmax+xmin)/2
		ycenter = (ymax+ymin)/2
		WCS.SetNormal(0,0,1)
		WCS.SetOrigin(xcenter,ycenter,zmin)

	ElseIf deltaxz < deltayz And deltaxz < deltaxy Then
		Axis = "y"
		xcenter = (xmax+xmin)/2
		zcenter = (zmax+zmin)/2
		WCS.SetNormal(0,1,0)
		WCS.SetOrigin(xcenter,ymin,zcenter)
	ElseIf deltayz < deltaxy And deltayz < deltaxz Then
		Axis = "x"
		ycenter = (ymax+ymin)/2
		zcenter = (zmax+zmin)/2
		WCS.SetNormal(1,0,0)
		WCS.SetOrigin(xmin,ycenter,zcenter)
	End If
	'Initialize the total rotated angle, this angle should be less than 180 degree
	Angle = 0
	WCS.RotateWCS("u",90)
	While Angle < 180
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

Function AutoSliceBySteps(xStep As Double,yStep As Double,zStep As Double,wStep As Double)
	Dim fullName As String
	Dim sn As Integer, i As Integer
	'When xStep is not less than wStep, slice once with step of wStep
	If xStep <> 0 And xStep >= wStep Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,0)
	    	Next i
	    End If
	'When xStep is less than wStep, slice twice
    ElseIf xStep <> 0 And xStep < wStep Then
    	'First slice is done with step of wStep
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,0)
	    	Next i
	    End If
	    'second slice is done with step of xStep
		sn =  Solid.GetNumberOfShapes
		If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,xStep,0)
	    	Next i
	    End If
	End If
	'When yStep is not less than wStep, slice once with step of wStep
	If yStep <> 0 And yStep >= wStep Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,1)
	    	Next i
	    End If
	'When yStep is less than wStep, slice twice
    ElseIf yStep <> 0 And yStep < wStep Then
    	'First slice is done with step of wStep
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,1)
	    	Next i
	    End If
	    'second slice is done with step of yStep
		sn =  Solid.GetNumberOfShapes
		If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,yStep,1)
	    	Next i
	    End If
	End If
	'When zStep is not less than wStep, slice once with step of wStep
	If zStep <> 0 And zStep >= wStep Then
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,2)
	    	Next i
	    End If
	'When zStep is less than wStep, slice twice
    ElseIf zStep <> 0 And zStep < wStep Then
    	'First slice is done with step of wStep
		sn =  Solid.GetNumberOfShapes
	    If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,wStep,2)
	    	Next i
	    End If
	    'second slice is done with step of zStep
		sn =  Solid.GetNumberOfShapes
		If sn > 0 Then
	    	For i = 0 To sn-1 STEP 1
				fullName = Solid.GetNameOfShapeFromIndex(i)
			    AutoSliceAlongAxis(fullName,zStep,2)
	    	Next i
	    End If
	End If

End Function


Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean

	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box

		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
			DialogFunction = False

			Dim sn As Integer
			Dim sName As String,fullName As String, ComponentName As String
			Dim i As Integer
			Dim wStep As Double
			Dim Axis As Integer
			Dim xStep As Double, yStep As Double, zStep As Double
			Dim anStep As Double

			xStep = Evaluate(DlgText("xStep"))
		    yStep = Evaluate(DlgText("yStep"))
		    zStep = Evaluate(DlgText("zStep"))
		    anStep = Evaluate(DlgText("anStep"))
		    'Default slice step or initial slice step
		    wStep = 8
		    'Parse input parameters
		    If xStep = 0 And yStep = 0 And zStep = 0 And anStep =0 Then
		    	MsgBox("未输入有效参数，请重新运行脚本并输入有效参数！！",vbCritical,"错误")
		    	Exit All
		    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep <> 0) Then
		    	MsgBox("输入参数有冲突，请重新运行脚本并选择切片尺寸或者切片角度参数之一进行输入！！",vbCritical,"错误")
		    	Exit All
		    'Slice by dimension steps
		    ElseIf (xStep <> 0 Or yStep <> 0 Or zStep <> 0) And (anStep = 0) Then
		    	AutoSliceBySteps(xStep,yStep,zStep,wStep)
		    'Slice by angles
		    ElseIf (anStep <> 0) And (xStep = 0  And yStep = 0  And zStep = 0) Then
		    	AutoSliceByAngle(anStep)
		    End If

		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function
