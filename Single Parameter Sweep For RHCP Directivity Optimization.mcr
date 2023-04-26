'sweep single parameter to meet the target
'2023-01-10 By Shawn
'#include "vba_globals_all.lib"

Sub Main ()
	Dim parameterArray(1000) As String, IndexSelection() As Integer
	Dim ii As Integer, jj As Integer, jnow As Integer, sArray() As String
	Dim sAllSelectedParaNames As String
	Dim sPara As String
	For ii = 0 To GetNumberOfParameters-1
		sPara = GetParameterName(ii)
		parameterArray(ii) = sPara
	Next ii

	Begin Dialog UserDialog 730,350,"Single parameter sweep for RHCP directivity" ' %GRID:10,7,1,1
		GroupBox 500,7,230,56,"Select a parameter:",.GroupBox1
		GroupBox 500,63,230,63,"Sweep settings:",.GroupBox2
		OKButton 510,322,90,21
		CancelButton 630,322,90,21
		Text 510,84,80,14,"Sweep from",.Text1
		Text 660,84,20,14,"to",.Text2
		Text 520,105,100,14,"with step size:",.Text3
		TextBox 600,84,50,14,.xMin
		TextBox 680,84,40,14,.xMax
		TextBox 630,105,50,14,.stepSize
		GroupBox 500,133,230,56,"Frequency settings:",.GroupBox3
		Text 530,147,80,14,"Frequency1:",.Text5
		Text 530,168,80,14,"Frequency2:",.Text7
		TextBox 620,147,40,14,.f1
		TextBox 620,168,40,14,.f2
		GroupBox 500,196,230,63,"Cut angle settings:",.GroupBox4
		OptionGroup .Group1
			OptionButton 550,217,40,14,"¦È",.OptionButton1
			OptionButton 620,217,40,14,"¦Õ",.OptionButton2
		Text 520,238,90,14,"with angle of",.Text6
		TextBox 620,238,60,14,.Angle
		GroupBox 500,259,230,56,"Calculate point settings:",.GroupBox5
		Text 530,287,30,14,"¦È0",.Text8
		TextBox 560,287,50,14,.theta0
		Text 640,287,30,14,"¦Õ0",.Text11
		TextBox 670,287,50,14,.phi0
		Text 600,217,20,14,"or",.Text4
		DropListBox 520,28,200,21,parameterArray(),.parameterIndex
		Picture 0,7,500,343,GetInstallPath + "\Library\Macros\Coros\Simulation\Single parameter sweep instructions For RHCP optimization.bmp",0,.Picture1
	End Dialog
	Dim dlg As UserDialog
	dlg.xMin = "0"
	dlg.xMax = "0"
	'dlg.Mono = "0"
	dlg.f1 = "0"
	dlg.f2 = "0"
	dlg.theta0 = "0"
	dlg.phi0 = "0"
	dlg.Group1 = 0
	dlg.Angle = "90"
	'dlg.f3 = "0"
	'dlg.Q1 = "0"
	'dlg.Q2 = "0"
	'dlg.Q3 = "0"
	'dlg.AR1 = "0"
	'dlg.AR2 = "0"
	'dlg.theta1 = "0"
	'dlg.theta2 = "0"
	'dlg.phi1 = "0"
	'dlg.phi2 = "0"
	dlg.stepSize = "0"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If


	Dim parameter As String
	Dim xMin As Double, xMax As Double, theta0 As Double, phi0 As Double, directivity As Double
	Dim xSim As Double
	Dim stepWidth As Double

	parameter = parameterArray(dlg.parameterIndex)
	If DoesParameterExist(parameter) = False Then
		MsgBox("The input paramter does not exist!!",vbCritical,"Error")
		Exit All
    End If
    xMin = Evaluate(dlg.xMin)
    xMax = Evaluate(dlg.xMax)
    theta0 = Evaluate(dlg.theta0)
    phi0 = Evaluate(dlg.phi0)
    'Mono = Evaluate(dlg.Mono)
    'xSim = xMin
    f1 = Evaluate(dlg.f1)
    f2 = Evaluate(dlg.f2)

	'Check validity of frequency
	Dim monitorNumber As Integer, i As Integer
	Dim flag1 As Boolean, flag2 As Boolean

	flag1 = False
	flag2 = False
	monitorNumber = Monitor.GetNumberOfMonitors()
	For i = 0 To monitorNumber-1 STEP 1
		If Monitor.GetMonitorTypeFromIndex(i)="Farfield" Then
			If Monitor.GetMonitorFrequencyFromIndex(i) = f1 Then
				flag1 = True
			End If
			If Monitor.GetMonitorFrequencyFromIndex(i) = f2 Then
				flag2 = True
			End If
		End If
	Next
	If (f1 <> 0  And flag1 = False) Or (f2 <> 0 And flag2 = False) Then
		MsgBox("The input frequencies do not exist!!",vbCritical,"Error")
		Exit All
	End If

    stepSize = Evaluate(dlg.stepSize)

	Dim n As Integer
	Dim rotateAngle As Double
	'rotateAngle = xMin
	n = 0
	rotateAngle = xMin

	Dim prjPath As String
	Dim dataFile As String
	Dim groupValue As Integer, cutAngle As Double
	Dim fieldComponent As String
	groupValue = dlg.Group1
	cutAngle = Evaluate(dlg.Angle)

	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\sweeplog.txt"
   	Open dataFile For Output As #2
   	Print #2, "##########Sweep of " + parameter + " is in progress###########."
   	Print #2, " "
	'run with specified value of xMin
	FarfieldPlot.Reset
	'FarfieldPlot.SetAntennaType("directional_linear")
	'FarfieldPlot.Plot
	FarfieldPlot.SetAntennaType("directional_circular")
	'FarfieldPlot.Plot
	FarfieldPlot.SetPlotMode("directivity")
	fieldComponent = "ludwig3 circular right abs"
    While rotateAngle <= xMax And rotateAngle < xMin+360
		Print #2, "% On step "+CStr(n+1)+": "+parameter+"="+Cstr(rotateAngle)+"."
		'run with specified value of xSim
		runWithParameter(parameter,rotateAngle)
		Print #2, "% Step "+CStr(n+1)+" simulation is done."
		If f1 <> 0 Then
			directivity = Copy1DFarfieldResult(groupValue, cutAngle, theta0, phi0, rotateAngle, f1)
			Print #2, "% Directivity at frequency "+ CStr(f1)+ "Ghz is "+ CStr(Round(directivity, 2)) + "dB."
		End If

		If f2 <> 0 Then
			directivity = Copy1DFarfieldResult(groupValue, cutAngle, theta0, phi0, rotateAngle, f2)
			Print #2, "% Directivity at frequency "+ CStr(f2)+ "Ghz is "+ CStr(Round(directivity, 2)) + "dB."
		End If
		Print #2, " "
		n = n+1
		If stepSize <> 0 Then
			rotateAngle = xMin + n*stepSize
		Else
			rotateAngle = xMax+1
		End If

	Wend
	Close #2
	MsgBox("Maximum rotate angle reached, exit the sweep progress",vbInformation, "Attention")

End Sub
Sub runWithParameter(para As String, value As Double)
 	StoreParameter(para,value)
	Rebuild
	Solver.MeshAdaption(False)
	Solver.SteadyStateLimit(-40)
	Solver.Start
End Sub

Function Copy1DFarfieldResult(groupValue As Integer, cutAngle As Double, theta0 As Double, phi0 As Double, rotateAngle As Double, freq As Double)
	'parameters: groupValue denotes the cutting plane,0->theta, 1->phi; cutAngle denotes the angle in the plane specified by groupvalue; theta0 and phi0
	'denote the angle where the directivity is estimated; rotateAngle is the rotate angle of the ham; freq is the operation frequency we take care
		Dim SelectedItem As String, PortStr As String, FrequencyStr As String

		PortStr = "1"
		FrequencyStr = CStr(freq)

		SelectTreeItem("Farfields\farfield (f="+FrequencyStr+") [" +PortStr+"]")

		FarfieldPlot.Reset
		FarfieldPlot.SetAntennaType("directional_circular")
		FarfieldPlot.SetPlotMode("directivity")
		FarfieldPlot.PlotType("polar")
		FarfieldPlot.SelectComponent("Axial Ratio")
		FarfieldPlot.Plot

		If groupValue = 0 Then

			FarfieldPlot.Vary("angle2")
			FarfieldPlot.Theta(cutAngle)

		Else

			FarfieldPlot.Vary("angle1")
			FarfieldPlot.Phi(cutAngle)

		End If

		'FarfieldPlot.Plot
		Dim DirName As String

		DirName = "CP directivity\Theta0="+CStr(theta0)+", Phi0="+CStr(phi0)+" and Rotate_angle="+CStr(rotateAngle)+ " @"
		Dim ChildItem As String
			If Resulttree.DoesTreeItemExist("1D Results\"+DirName+FrequencyStr+"GHz") Then
				ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName+FrequencyStr+"GHz")
				While ChildItem <> ""
					With Resulttree
						.Name ChildItem
						.Delete
					End With
					ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName+FrequencyStr+"GHz")
				Wend

			End If
		FarfieldPlot.SelectComponent("Abs")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Abs")
		FarfieldPlot.SelectComponent("Right")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Right")
		FarfieldPlot.SelectComponent("Left")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Left")

        'CurrentItem = FirstChildItem
        Copy1DFarfieldResult = getDirectivity(groupValue, cutAngle, theta0, phi0, rotateAngle, freq)
        SelectTreeItem("1D Results\"+DirName+FrequencyStr+"GHz")
End Function

Function getDirectivity(groupValue As Integer, cutAngle As Double, theta0 As Double, phi0 As Double, rotateAngle As Double, frequency As Double)
	Dim calTheta As Double, calPhi As Double, totAngle As Double
	Dim farfieldComponent As String
	Dim directivity As Double

	farfieldComponent = "ludwig3 circular right abs"
	totAngle = theta0 + rotateAngle
	diffAngle = theta0 - rotateAngle
	If groupValue = 1 And phi0 = 0 Then
		If totAngle < 180 Then
			calTheta = totAngle
			calPhi = 0
		ElseIf totAngle > 180 And totAngle < 360 Then
			calTheta = 360-totAngle
			calPhi = 180
		ElseIf totAngle > 360 Then
			calTheta = totAngle - 360
			calPhi = 0
		End If
	ElseIf groupValue = 1 And phi0 = 180 Then
		If diffAngle > 0 Then
			calTheta = diffAngle
			calPhi = 180
		ElseIf diffAngle < 0 And diffAngle > -180 Then
			calTheta = -diffAngle
			calPhi = 0
		ElseIf diffAngle > -360 And diffAngle < -180 Then
			calTheta = -diffAngle - 180
			calPhi = 180
		End If

	End If
	'SelectTreeItem
	directivity = FarfieldPlot.CalculatePoint(calTheta, calPhi, farfieldComponent, "")
	getDirectivity = directivity
End Function




