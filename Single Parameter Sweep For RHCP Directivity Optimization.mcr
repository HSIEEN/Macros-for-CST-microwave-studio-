'sweep single parameter to meet the target
'2023-01-10 By Shawn
'#include "vba_globals_all.lib"
Sub Main ()
	Dim parameterArray(1000) As String
	Dim ii As Integer
	Dim sAllSelectedParaNames As String
	For ii = 0 To GetNumberOfParameters-1
		parameterArray(ii) = GetParameterName(ii)
	Next ii

	Begin Dialog UserDialog 810,434,"Single parameter sweep for RHCP directivity comparison" ' %GRID:10,7,1,1
		GroupBox 600,7,210,56,"Select a parameter:",.GroupBox1
		GroupBox 600,63,210,63,"Parameter sweep settings:",.GroupBox2
		OKButton 660,315,90,21
		CancelButton 660,364,90,21
		Text 620,84,40,14,"From",.Text1
		Text 720,84,20,14,"to",.Text2
		Text 630,105,100,14,"with step size:",.Text3
		TextBox 670,84,40,14,.xMin
		TextBox 750,84,40,14,.xMax
		TextBox 740,105,50,14,.stepSize
		GroupBox 600,126,210,63,"Frequency settings:",.GroupBox3
		Text 630,147,80,14,"Frequency1:",.Text5
		Text 630,168,80,14,"Frequency2:",.Text7
		TextBox 720,147,40,14,.f1
		TextBox 720,168,40,14,.f2
		GroupBox 600,189,210,91,"Cut angle settings in 1D plot:",.GroupBox4
		OptionGroup .Group1
			OptionButton 660,210,40,14,"θ",.OptionButton1
			OptionButton 660,238,40,14,"φ",.OptionButton2
		Text 630,259,90,14,"with angle of",.Text6
		TextBox 720,259,40,14,.Angle
		DropListBox 620,28,180,21,parameterArray(),.parameterIndex
		Picture 0,7,600,399,GetInstallPath + "\Library\Macros\Coros\Simulation\Single parameter sweep instructions For RHCP optimization.bmp",0,.Picture1
	End Dialog
	Dim dlg As UserDialog
	dlg.xMin = "0"
	dlg.xMax = "0"
	'dlg.Mono = "0"
	dlg.f1 = "0"
	dlg.f2 = "0"

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

	Dim projectPath As String
	Dim dataFile As String
	Dim groupValue As Integer, cutAngle As Double
	Dim fieldComponent As String
	groupValue = dlg.Group1
	cutAngle = Evaluate(dlg.Angle)

	projectPath = GetProjectPath("Project")
   	dataFile = projectPath + "\sweeplog_"+Replace(CStr(Time),":","_")+".txt"
   	Open dataFile For Output As #2
   	Print #2, "##########Sweep of " + parameter + " begins at " + CStr(Now) +"'###########."
   	Print #2, " "

	FarfieldPlot.SetPlotMode("directivity")
    While rotateAngle <= xMax And rotateAngle < xMin+360
		Print #2, "%-%-% On step "+CStr(n+1)+": "+parameter+"="+Cstr(rotateAngle)+"."
		Print #2, "%-%-% Start time: " + CStr(Now) +"."
		'run with specified value of xSim
		runWithParameter(parameter,rotateAngle)
		Print #2, "% Step "+CStr(n+1)+" simulation is done."
		If f1 <> 0 Then
			directivity = Copy1DFarfieldResult(groupValue, cutAngle, rotateAngle, f1)
			Print #2, "%-%-% RHCP Directivity at frequency "+ CStr(f1)+ "Ghz is "+ CStr(Round(directivity, 2)) + "dBi."
		End If

		If f2 <> 0 Then
			directivity = Copy1DFarfieldResult(groupValue, cutAngle, rotateAngle, f2)
			Print #2, "%-%-% RHCP Directivity at frequency "+ CStr(f2)+ "Ghz is "+ CStr(Round(directivity, 2)) + "dBi."
		End If
		Print #2, "%-%-% Finish time: " + CStr(Now) +"."
		Print #2, " "
		n = n+1
		If stepSize = 0 Then
			Exit While
		End If
		rotateAngle = xMin + n*stepSize
	Wend
	Print #2, "##########Sweep of " + parameter + " ends at " + CStr(Now) +"'###########."
	Close #2
	MsgBox("Maximum rotate angle reached, the sweep progress finished.",vbInformation, "Attention")

End Sub
Sub runWithParameter(para As String, value As Double)
 	StoreParameter(para,value)
	Rebuild
	Solver.MeshAdaption(False)
	Solver.SteadyStateLimit(-40)
	Solver.Start
End Sub

Function Copy1DFarfieldResult(groupValue As Integer, cutAngle As Double, rotateAngle As Double, freq As Double)
	'parameters: groupValue denotes the cutting plane,0->theta, 1->phi; cutAngle denotes the angle in the plane specified by groupvalue; theta0 and phi0
	'denote the angle where the directivity is estimated; rotateAngle is the rotate angle of the ham; freq is the operation frequency we take care
		Dim SelectedItem As String, PortStr As String, FrequencyStr As String

		PortStr = "1"
		FrequencyStr = CStr(freq)

		SelectTreeItem("Farfields\farfield (f="+FrequencyStr+") [" +PortStr+"]")

		FarfieldPlot.Reset

		'FarfieldPlot.Plot
		FarfieldPlot.SelectComponent("Abs")
		FarfieldPlot.PlotType("polar")
		If groupValue = 0 Then

			FarfieldPlot.Vary("angle2")
			FarfieldPlot.Theta(cutAngle)

		Else

			FarfieldPlot.Vary("angle1")
			FarfieldPlot.Phi(cutAngle)

		End If

		FarfieldPlot.SetAxesType("currentwcs")
		FarfieldPlot.SetAntennaType("unknown")
		FarfieldPlot.SetPlotMode("Directivity")
		'FarfieldPlot.SetAntennaType("directional_linear")
		'FarfieldPlot.SetAntennaType("directional_circular")
		FarfieldPlot.SetCoordinateSystemType("ludwig3")
		FarfieldPlot.SetAutomaticCoordinateSystem("True")
		FarfieldPlot.SetPolarizationType("Circular")

		FarfieldPlot.StoreSettings

		FarfieldPlot.Plot

		'FarfieldPlot.Plot
		Dim DirName As String

		DirName = "CP directivity\Rotation angle="+CStr(rotateAngle)+ "@"+FrequencyStr+"GHz"
		Dim ChildItem As String
			If Resulttree.DoesTreeItemExist("1D Results\"+DirName) Then
				ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName)
				While ChildItem <> ""
					With Resulttree
						.Name ChildItem
						.Delete
					End With
					ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName)
				Wend

			End If
		FarfieldPlot.SelectComponent("Abs")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Abs")
		FarfieldPlot.SelectComponent("Right")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Right")
		FarfieldPlot.SelectComponent("Left")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Left")

		saveCircularDirectivity(rotateAngle, freq)
        'CurrentItem = FirstChildItem
        Copy1DFarfieldResult = getDirectivity()
        SelectTreeItem("1D Results\"+DirName)
End Function

Function getDirectivity()

	Dim farfieldComponent As String
	Dim directivity As Double

	farfieldComponent = "ludwig3 circular right abs"
	'SelectTreeItem
	directivity = FarfieldPlot.CalculatePoint(0, 0, farfieldComponent, "")
	getDirectivity = directivity
End Function
Sub saveCircularDirectivity(rotateAngle As Double,frequency As Double)

    Dim SelectedItem As String

    Dim n As Integer

    Dim FrequencyStr As String
    Dim PortStr As String


	PortStr = "1"
	FrequencyStr = CStr(frequency)

	SelectTreeItem("Farfields\farfield (f="+FrequencyStr+") [" +PortStr+"]")
    '==============Upper Hemisphere rhcp directivity and rhcp directivity estimation===============

    Dim  upperHemisphereRHCPdirectivity() As Double, upperHemisphereLHCPdirectivity() As Double

    Dim rhcpDirectivity() As Double, lhcpDirectivity() As Double

    Dim Theta As Double, Phi As Double

    Dim position_theta() As Double, position_phi() As Double
    Dim Columns As String
    Dim dataFile As String
    Dim projectPath As String

    For Phi=0 To 360 STEP 30

         For Theta = 0 To 180 STEP 15

             FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", frequency)

         Next Theta

    Next Phi

    FarfieldPlot.CalculateList("")

    'UHPower = FarfieldPlot.GetList("Spherical  abs")

    upperHemisphereRHCPdirectivity = FarfieldPlot.GetList("Spherical circular right abs")

    upperHemisphereLHCPdirectivity = FarfieldPlot.GetList("Spherical circular left abs")

    position_theta = FarfieldPlot.GetList("Point_T")

    position_phi = FarfieldPlot.GetList("Point_P")

    ReDim rhcpDirectivity(13,13)
    ReDim lhcpDirectivity(13,13)
    'EIRP of an isotropic antenna
    For n = 0 To UBound(upperHemisphereRHCPdirectivity)

    	'linear to dB, Log(UHRHCPEffi)/Log(10)*10
         rhcpDirectivity(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = upperHemisphereRHCPdirectivity(n)
         '10*CST_Log10(upperHemisphereRHCPdirectivity(n)) 'Log(upperHemisphereRHCPdirectivity(n)/AVGPower)/Log(10)*10
         lhcpDirectivity(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = upperHemisphereLHCPdirectivity(n)
         '10*CST_Log10(upperHemisphereLHCPdirectivity(n))'Log(upperHemisphereLHCPdirectivity(n)/AVGPower)/Log(10)*10
    Next n
	 '==============================write directivity data============================
	projectPath = GetProjectPath("Project")
	dataFile = projectPath+"\Circularly polarized directivity_frquency="+FrequencyStr+"GHz_Port="+PortStr+"_"+Replace(CStr(Time),":","_")+".xlsx"
	Columns = "BCDEFGHIJKLMN"

	NoticeInformation = "The directivity data is under（"+projectPath+"\）"
    ReportInformationToWindow(NoticeInformation)

    Set O = CreateObject("Excel.Application")
	If Dir(dataFile) = "" Then
	    Dim wBook As Object
	    Set wBook = O.Workbooks.Add
	    With wBook
	        .Title = "Title"
	        .Subject = "Subject"
	        .SaveAs Filename:= dataFile
		End With
	Else
	    Set wBook = O.Workbooks.Open(dataFile)
	End If


	'Add a sheet and rename it

	wBook.Sheets.Add.Name = "rotate angle="+CStr(rotateAngle)
	Set wSheet = wBook.Sheets("rotate angle="+CStr(rotateAngle))

	'write rhcp directivity
	wSheet.Range("A1").value = "Polarization"
	wSheet.Range("B1").value = "RHCP"
	wSheet.Range("C1").value = "Frequency"
	wSheet.Range("D1").value = FrequencyStr+"GHz"
	wSheet.Range("E1").value = "Port"
	wSheet.Range("F1").value = PortStr
	wSheet.Range("A2").value = "Phi\Theta"

	For n = 0 To Len(Columns)-1
		wSheet.Range(Mid(Columns,n+1,1)+"2").value = n*15
		wSheet.Range("A"+Cstr(n+3)) = n*30
	Next

	Dim i As Integer, j As Integer

	For i  = 0 To 12
		For j = 0 To 12
			wSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).value = Round(rhcpDirectivity(i,j),2)
		Next
	Next
	'write lhcp directivity
	wSheet.Range("A18").value = "Polarization"
	wSheet.Range("B18").value = "LHCP"
	wSheet.Range("C18").value = "Frequency"
	wSheet.Range("D18").value = FrequencyStr+"GHz"
	wSheet.Range("E18").value = "Port"
	wSheet.Range("F18").value = PortStr
	wSheet.Range("A19").value = "Phi\Theta"

	For n = 0 To Len(Columns)-1
		wSheet.Range(Mid(Columns,n+1,1)+"19").value = n*15
		wSheet.Range("A"+Cstr(n+20)) = n*30
	Next

	For i  = 0 To 12
		For j = 0 To 12
			wSheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).value = Round(lhcpDirectivity(i,j),2)
		Next
	Next

	Dim sheet As Object
	For Each sheet In wBook.Sheets
	    If sheet.Name Like "Sheet*" Then
	        sheet.Delete
	    End If
	Next
	'process sheet data, axial ratio, coloring, scoring and so on
	processDirectivityData(wSheet, Columns)
	wBook.Save
	O.ActiveWorkbook.Close
	O.quit

End Sub

Sub processDirectivityData(sheet As Object, Columns As String)

	Dim i As Integer
	Dim j As Integer

	sheet.Range("A35").value = "Polarization"
	sheet.Range("B35").value = "AR"
	sheet.Range("C35").value = "Frequency"
	sheet.Range("D35").value = sheet.Range("D1").Value
	sheet.Range("E35").value = "Port"
	sheet.Range("F35").value = sheet.Range("F1").Value
	sheet.Range("A36").value = "Phi\Theta"
	For i = 0 To Len(Columns)-1
		sheet.Range(Mid(Columns,i+1,1)+"36").value = i*15
		sheet.Range("A"+Cstr(i+37)) = i*30
	Next

	'hide lhcp directivity data
	sheet.Rows("17:33").Hidden = True
	Dim Dvalue As Double, deltaDirectivity As Double, axialRatio As Double
	'coloring and resizing cells
	sheet.Columns("A").ColumnWidth = 18

	sheet.Rows("1").RowHeight = 25
	sheet.Rows("35").RowHeight = 25
	'sheet.Range("A1:Z100").HorizontalAlignment = xlCenter


	For i  = 0 To 12
		For j = 0 To 12

			'=======================Axial ratio estimating and coloring============================
			deltaDirectivity = sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value - sheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).Value
			axialRatio = Sgn(deltaDirectivity)*20*CST_Log10((10^(deltaDirectivity/20)+1)/(Abs(10^(deltaDirectivity/20)-1)+0.01))
			sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).value = Round(axialRatio,2)
			If axialRatio >= 0 And axialRatio < 3 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(0, 130, 0)
			ElseIf axialRatio < 6 And axialRatio >= 3 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(0, 180, 0)
			ElseIf axialRatio < 10 And axialRatio >= 6 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(145, 218, 0)
			ElseIf axialRatio < 18 And axialRatio >= 10 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(216, 254, 154)
			ElseIf axialRatio >= 18 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 255, 0)
			ElseIf axialRatio < -14 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 200, 0)
			ElseIf axialRatio < -6 And axialRatio >= -14 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 0, 0)
			ElseIf axialRatio < 0 And axialRatio >= -6  Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(150, 0, 0)
			End If
			'color = sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color
			'========================Coloring rhcp directvity data===================================
			'reference total efficiency -8dB
			Dvalue = sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value
			If Dvalue >= 2 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 130, 0)
			ElseIf Dvalue < 2 And Dvalue >= 0 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 180, 0)
			ElseIf Dvalue < 0 And Dvalue >= -2 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(145, 218, 0)
			ElseIf Dvalue < -2 And Dvalue >= -4 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(216, 254, 154)
			ElseIf Dvalue < -4 And Dvalue >= -6 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 255, 0)
			ElseIf Dvalue < -6 And Dvalue >= -8 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 200, 0)
			ElseIf Dvalue < -8 And Dvalue >= -10 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 0, 0)
			ElseIf Dvalue < -10  Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(150, 0, 0)
			End If

		Next
	Next
	'======================================UH/Tot===============================
	'sheet.Range("A51") = "UHPower ratio"
	'sheet.Range("B51") = Round(10*CST_Log10(getUpperHemisphereRatio(sheet, columns)),2)
	'sheet.Range("C51") = "dB"
	writeAverageDirectivity(sheet, Columns)

End Sub
Sub writeAverageDirectivity(sheet As Object, Columns As String)

	sheet.Columns("P").ColumnWidth = 15
	sheet.Range("P2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("P3:P7").Interior.Color = RGB(0, 176, 240)
	sheet.Range("Q2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("Q3:Q7").Interior.Color = RGB(255, 217, 102)
	sheet.Columns("Q").ColumnWidth = 25
	sheet.Range("P2:Q7").Font.Bold = True

	sheet.Range("P2").Value = "Within theta"
	sheet.Range("Q2").Value = "Weighted Average direcitivy"

	sheet.Range("P3").Value = "30"
	sheet.Range("P4").Value = "45"
	sheet.Range("P5").Value = "60"
	sheet.Range("P6").Value = "90"
	sheet.Range("P7").Value = "120"

	sheet.Range("Q3").Value = Round(10*CST_Log10(((10^(sheet.Range("B3").Value/10)+10^(sheet.Range("B4").Value/10)+10^(sheet.Range("B5").Value/10)+10^(sheet.Range("B6").Value/10)+10^(sheet.Range("B7").Value/10)+10^(sheet.Range("B8").Value/10)+10^(sheet.Range("B9").Value/10)+10^(sheet.Range("B10").Value/10)+10^(sheet.Range("B11").Value/10)+10^(sheet.Range("B12").Value/10)+10^(sheet.Range("B13").Value/10)+10^(sheet.Range("B14").Value/10))*(1-Cos(pi/24))*pi/6+(10^(sheet.Range("C3").Value/10)+10^(sheet.Range("C4").Value/10)+10^(sheet.Range("C5").Value/10)+10^(sheet.Range("C6").Value/10)+10^(sheet.Range("C7").Value/10)+10^(sheet.Range("C8").Value/10)+10^(sheet.Range("C9").Value/10)+10^(sheet.Range("C10").Value/10)+10^(sheet.Range("C11").Value/10)+10^(sheet.Range("C12").Value/10)+10^(sheet.Range("C13").Value/10)+10^(sheet.Range("C14").Value/10))*(Cos(pi/24)-Cos(pi/8))*pi/6+(10^(sheet.Range("D3").Value/10)+10^(sheet.Range("D4").Value/10)+10^(sheet.Range("D5").Value/10)+10^(sheet.Range("D6").Value/10)+10^(sheet.Range("D7").Value/10)+10^(sheet.Range("D8").Value/10)+10^(sheet.Range("D9").Value/10)+10^(sheet.Range("D10").Value/10)+10^(sheet.Range("D11").Value/10)+10^(sheet.Range("D12").Value/10)+10^(sheet.Range("D13").Value/10)+10^(sheet.Range("D14").Value/10))*(Cos(pi/8)-Cos(pi/6))*pi/6)/(2*pi*(1-Cos(pi/6)))),2)

	sheet.Range("Q4").Value = Round(10*CST_Log10(((10^(sheet.Range("Q3").Value/10)*(2*pi*(1-Cos(pi/6)))+(10^(sheet.Range("D3").Value/10)+10^(sheet.Range("D4").Value/10)+10^(sheet.Range("D5").Value/10)+10^(sheet.Range("D6").Value/10)+10^(sheet.Range("D7").Value/10)+10^(sheet.Range("D8").Value/10)+10^(sheet.Range("D9").Value/10)+10^(sheet.Range("D10").Value/10)+10^(sheet.Range("D11").Value/10)+10^(sheet.Range("D12").Value/10)+10^(sheet.Range("D13").Value/10)+10^(sheet.Range("D14").Value/10))*(Cos(pi/6)-Cos(5*pi/24))*pi/6+(10^(sheet.Range("E3").Value/10)+10^(sheet.Range("E4").Value/10)+10^(sheet.Range("E5").Value/10)+10^(sheet.Range("E6").Value/10)+10^(sheet.Range("E7").Value/10)+10^(sheet.Range("E8").Value/10)+10^(sheet.Range("E9").Value/10)+10^(sheet.Range("E10").Value/10)+10^(sheet.Range("E11").Value/10)+10^(sheet.Range("E12").Value/10)+10^(sheet.Range("E13").Value/10)+10^(sheet.Range("E14").Value/10))*(Cos(5*pi/24)-Cos(pi/4))*pi/6)/(2*pi*(1-Cos(pi/4))))),2)

	sheet.Range("Q5").Value = Round(10*CST_Log10(((10^(sheet.Range("Q4").Value/10)*(2*pi*(1-Cos(pi/4)))+(10^(sheet.Range("E3").Value/10)+10^(sheet.Range("E4").Value/10)+10^(sheet.Range("E5").Value/10)+10^(sheet.Range("E6").Value/10)+10^(sheet.Range("E7").Value/10)+10^(sheet.Range("E8").Value/10)+10^(sheet.Range("E9").Value/10)+10^(sheet.Range("E10").Value/10)+10^(sheet.Range("E11").Value/10)+10^(sheet.Range("E12").Value/10)+10^(sheet.Range("E13").Value/10)+10^(sheet.Range("E14").Value/10))*(Cos(pi/4)-Cos(7*pi/24))*pi/6+((10^(sheet.Range("F3").Value/10)+10^(sheet.Range("F4").Value/10)+10^(sheet.Range("F5").Value/10)+10^(sheet.Range("F6").Value/10)+10^(sheet.Range("F7").Value/10)+10^(sheet.Range("F8").Value/10)+10^(sheet.Range("F9").Value/10)+10^(sheet.Range("F10").Value/10)+10^(sheet.Range("F11").Value/10)+10^(sheet.Range("F12").Value/10)+10^(sheet.Range("F13").Value/10)+10^(sheet.Range("F14").Value/10))*(Cos(7*pi/24)-Cos(pi/3))*pi/6))/(2*pi*(1-Cos(pi/3))))),2)

	sheet.Range("Q6").Value = Round(10*CST_Log10(((10^(sheet.Range("Q5").Value/10)*2*pi*(1-Cos(pi/3))+(10^(sheet.Range("F3").Value/10)+10^(sheet.Range("F4").Value/10)+10^(sheet.Range("F5").Value/10)+10^(sheet.Range("F6").Value/10)+10^(sheet.Range("F7").Value/10)+10^(sheet.Range("F8").Value/10)+10^(sheet.Range("F9").Value/10)+10^(sheet.Range("F10").Value/10)+10^(sheet.Range("F11").Value/10)+10^(sheet.Range("F12").Value/10)+10^(sheet.Range("F13").Value/10)+10^(sheet.Range("F14").Value/10))*(Cos(8*pi/24)-Cos(9*pi/24))*pi/6+(10^(sheet.Range("G3").Value/10)+10^(sheet.Range("G4").Value/10)+10^(sheet.Range("G5").Value/10)+10^(sheet.Range("G6").Value/10)+10^(sheet.Range("G7").Value/10)+10^(sheet.Range("G8").Value/10)+10^(sheet.Range("G9").Value/10)+10^(sheet.Range("G10").Value/10)+10^(sheet.Range("G11").Value/10)+10^(sheet.Range("G12").Value/10)+10^(sheet.Range("G13").Value/10)+10^(sheet.Range("G14").Value/10))*(Cos(9*pi/24)-Cos(11*pi/24))*pi/6+(10^(sheet.Range("H3").Value/10)+10^(sheet.Range("H4").Value/10)+10^(sheet.Range("H5").Value/10)+10^(sheet.Range("H6").Value/10)+10^(sheet.Range("H7").Value/10)+10^(sheet.Range("H8").Value/10)+10^(sheet.Range("H9").Value/10)+10^(sheet.Range("H10").Value/10)+10^(sheet.Range("H11").Value/10)+10^(sheet.Range("H12").Value/10)+10^(sheet.Range("H13").Value/10)+10^(sheet.Range("H14").Value/10))*(Cos(11*pi/24)-Cos(pi/2))*pi/6))/(2*pi*(1-Cos(pi/2)))),2)

	sheet.Range("Q7").Value = Round(10*CST_Log10((10^(sheet.Range("Q6").Value/10)*2*pi*(1-Cos(pi/2))+(10^(sheet.Range("H3").Value/10)+10^(sheet.Range("H4").Value/10)+10^(sheet.Range("H5").Value/10)+10^(sheet.Range("H6").Value/10)+10^(sheet.Range("H7").Value/10)+10^(sheet.Range("H8").Value/10)+10^(sheet.Range("H9").Value/10)+10^(sheet.Range("H10").Value/10)+10^(sheet.Range("H11").Value/10)+10^(sheet.Range("H12").Value/10)+10^(sheet.Range("H13").Value/10)+10^(sheet.Range("H14").Value/10))*(Cos(pi/2)-Cos(13*pi/24))*pi/6+(10^(sheet.Range("I3").Value/10)+10^(sheet.Range("I4").Value/10)+10^(sheet.Range("I5").Value/10)+10^(sheet.Range("I6").Value/10)+10^(sheet.Range("I7").Value/10)+10^(sheet.Range("I8").Value/10)+10^(sheet.Range("I9").Value/10)+10^(sheet.Range("I10").Value/10)+10^(sheet.Range("I11").Value/10)+10^(sheet.Range("I12").Value/10)+10^(sheet.Range("I13").Value/10)+10^(sheet.Range("I14").Value/10))*(Cos(13*pi/24)-Cos(15*pi/24))*pi/6+(10^(sheet.Range("J3").Value/10)+10^(sheet.Range("J4").Value/10)+10^(sheet.Range("J5").Value/10)+10^(sheet.Range("J6").Value/10)+10^(sheet.Range("J7").Value/10)+10^(sheet.Range("J8").Value/10)+10^(sheet.Range("J9").Value/10)+10^(sheet.Range("J10").Value/10)+10^(sheet.Range("J11").Value/10)+10^(sheet.Range("J12").Value/10)+10^(sheet.Range("J13").Value/10)+10^(sheet.Range("J14").Value/10))*(Cos(15*pi/24)-Cos(16*pi/24))*pi/6)/(2*pi*(1-Cos(2*pi/3)))),2)


End Sub




