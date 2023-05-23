'sweep single parameter to meet the target
'2023-01-10 By Shawn
'#include "vba_globals_all.lib"
Public startTime, parameter

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
	dlg.xMax = "360"
	'dlg.Mono = "0"
	dlg.f1 = "0"
	dlg.f2 = "0"

	dlg.Group1 = 1
	dlg.Angle = "0"
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
	dlg.stepSize = "10"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If


	'Dim parameter As String
	Dim xMin As Double, xMax As Double, theta0 As Double, phi0 As Double, directivity As Double
	Dim xSim As Double
	Dim stepWidth As Double

	If dlg.parameterIndex = -1 Then
		MsgBox("No parameter is selected!!",vbCritical,"Error")
		Exit All
	End If
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
	startTime = Replace(CStr(Time),":","_")
   	dataFile = projectPath + "\Sweep log_"+Replace(CStr(Time),":","_")+".txt"
   	Open dataFile For Output As #2
   	Print #2, "##########Sweep of " + parameter + " begins at " + CStr(Now) +"'###########."
   	Print #2, " "

	FarfieldPlot.SetPlotMode("directivity")
    While rotateAngle <= xMax And rotateAngle < xMin+360
		Print #2, "%-%-% On step "+CStr(n+1)+": "+parameter+"="+Cstr(rotateAngle)+"."
		Print #2, "%-%-% Start time: " + CStr(Now) +"."
		'run with specified value of xSim
		runWithParameter(parameter,rotateAngle)
		Print #2, "%-%-%  Step "+CStr(n+1)+" simulation is done."
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
	MsgBox("Maximum rotation angle reached, the sweep progress finished.",vbInformation, "Attention")

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
		Dim selectedItem As String, portStr As String, frequencyStr As String

		portStr = "1"
		frequencyStr = CStr(freq)

		SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")

		FarfieldPlot.Reset

		'FarfieldPlot.Plot
		'FarfieldPlot.SelectComponent("Abs")
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

		DirName = "CP directivity\"+parameter+"="+CStr(rotateAngle)+ "@"+frequencyStr+"GHz"
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
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+frequencyStr+")["+portStr+"]_Abs")
		FarfieldPlot.SelectComponent("Right")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+frequencyStr+")["+portStr+"]_Right")
		FarfieldPlot.SelectComponent("Left")
		'FarfieldPlot.PlotType("polar")
		FarfieldPlot.Plot
		FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+frequencyStr+")["+portStr+"]_Left")

		saveCircularDirectivity(rotateAngle, freq)
        'CurrentItem = FirstChildItem
        Copy1DFarfieldResult = getDirectivity()
        SelectTreeItem("1D Results\"+DirName)

		Dim curveLabel As String
		Dim index As Integer
		'Dim SelectedItem As String

		selectedItem = Resulttree.GetFirstChildName("1D Results\"+DirName)
		While selectedItem <> ""
			'SelectTreeItem(selectedItem)
			curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))

		   With Plot1D
		      index =.GetCurveIndexOfCurveLabel(curveLabel)
		     .SetLineStyle(index,"Solid",3) ' thick dashed line
		     .SetFont("Tahoma","bold","16")
		     '.SetLineColor(index,255,255,0)  ' yellow
		     .Plot ' make changes visible
			End With

			selectedItem = Resulttree.GetNextItemName(selectedItem)
		Wend
        'Plot1D.SetFont("Tahoma","bold","16")
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

    Dim selectedItem As String
    Dim n As Integer
    Dim frequencyStr As String
    Dim portStr As String

	portStr = "1"
	frequencyStr = CStr(frequency)

	SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")
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
	dataFile = projectPath+"\Circularly polarized directivity_frequency="+frequencyStr+"GHz Port="+portStr+"_"+startTime+".xlsx"
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

	wBook.Sheets.Add.Name = parameter+"="+CStr(rotateAngle)
	Set wSheet = wBook.Sheets(parameter+"="+CStr(rotateAngle))

	'write rhcp directivity
	wSheet.Range("A1").value = "Polarization"
	wSheet.Range("B1").value = "RHCP"
	wSheet.Range("C1").value = "Frequency"
	wSheet.Range("D1").value = frequencyStr+"GHz"
	wSheet.Range("E1").value = "Port"
	wSheet.Range("F1").value = portStr
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
	wSheet.Range("D18").value = frequencyStr+"GHz"
	wSheet.Range("E18").value = "Port"
	wSheet.Range("F18").value = portStr
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
	sheet.Columns("A").ColumnWidth = 14

	sheet.Columns("C").ColumnWidth = 10

	sheet.Rows("1").RowHeight = 25
	sheet.Rows("35").RowHeight = 25
	sheet.Range("A1:Z100").HorizontalAlignment =  -4108
	sheet.Range("P2:Q8").Borders.LineStyle = 1


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
	writeAverageDirectivityAndRating(sheet, Columns)

End Sub
Sub writeAverageDirectivityAndRating(sheet As Object, Columns As String)
	'Formating cells

	sheet.Columns("P").ColumnWidth = 15
	sheet.Range("P2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("P3:P7").Interior.Color = RGB(0, 176, 240)
	sheet.Range("P8").Interior.Color = RGB(221, 235, 247)
	sheet.Range("Q2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("Q3:Q7").Interior.Color = RGB(255, 217, 102)
	sheet.Range("Q8").Interior.Color = RGB(255, 217, 102)
	sheet.Columns("Q").ColumnWidth = 33
	sheet.Range("A1:Q100").Font.Bold = True
	sheet.Range("Q3:Q8").Font.Color = RGB(255, 0, 0)

	sheet.Range("P2").Value = "Within theta"
	sheet.Range("Q2").Value = "Weighted Average directivity"
	sheet.Range("P8").Value = "RHCPD rating"

	sheet.Range("P3").Value = "30"
	sheet.Range("P4").Value = "45"
	sheet.Range("P5").Value = "60"
	sheet.Range("P6").Value = "90"
	sheet.Range("P7").Value = "120"

	
	sheet.Range("Q3").Formula = "=ROUND(10*LOG10(((10^(B3/10)+10^(B4/10)+10^(B5/10)+10^(B6/10)"+ _
	"+10^(B7/10)+10^(B8/10)+10^(B9/10)+10^(B10/10)+10^(B11/10)+10^(B12/10)+10^(B13/10)+"+ _
	"10^(B14/10))*(1-COS(PI()/24))*PI()/6+(10^(C3/10)+10^(C4/10)+10^(C5/10)+10^(C6/10)+"+ _
	"10^(C7/10)+10^(C8/10)+10^(C9/10)+10^(C10/10)+10^(C11/10)+10^(C12/10)+10^(C13/10)+"+ _
	"10^(C14/10))*(COS(PI()/24)-COS(PI()/8))*PI()/6+(10^(D3/10)+10^(D4/10)+10^(D5/10)+"+ _
	"10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+10^(D10/10)+10^(D11/10)+10^(D12/10)+"+ _
	"10^(D13/10)+10^(D14/10))*(COS(PI()/8)-COS(PI()/6))*PI()/6)/(2*PI()*(1-COS(PI()/6)))),2)"

	sheet.Range("Q4").Formula = "=ROUND(10*LOG10(((10^(Q3/10)*(2*PI()*(1-COS(PI()/6)))+"+ _
	"(10^(D3/10)+10^(D4/10)+10^(D5/10)+10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+"+ _
	"10^(D10/10)+10^(D11/10)+10^(D12/10)+10^(D13/10)+10^(D14/10))*(COS(PI()/6)-"+ _
	"COS(5*PI()/24))*PI()/6+(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+"+ _
	"10^(E8/10)+10^(E9/10)+10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))"+ _
	"*(COS(5*PI()/24)-COS(PI()/4))*PI()/6)/(2*PI()*(1-COS(PI()/4))))),2)"

	sheet.Range("Q5").Formula = "=ROUND(10*LOG10(((10^(Q4/10)*(2*PI()*(1-COS(PI()/4)))+"+ _
	"(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+10^(E8/10)+10^(E9/10)+"+ _
	"10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))*(COS(PI()/4)-"+ _
	"COS(7*PI()/24))*PI()/6+((10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+"+ _
	"10^(F7/10)+10^(F8/10)+10^(F9/10)+10^(F10/10)+10^(F11/10)+10^(F12/10)+"+ _
	"10^(F13/10)+10^(F14/10))*(COS(7*PI()/24)-COS(PI()/3))*PI()/6))/(2*PI()*(1-COS(PI()/3))))),2)"

	sheet.Range("Q6").Formula = "=ROUND(10*LOG10(((10^(Q5/10)*2*PI()*(1-COS(PI()/3))+"+ _
	"(10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+10^(F7/10)+10^(F8/10)+10^(F9/10)+"+ _
	"10^(F10/10)+10^(F11/10)+10^(F12/10)+10^(F13/10)+10^(F14/10))*(COS(8*PI()/24)-"+ _
	"COS(9*PI()/24))*PI()/6+(10^(G3/10)+10^(G4/10)+10^(G5/10)+10^(G6/10)+10^(G7/10)+"+ _
	"10^(G8/10)+10^(G9/10)+10^(G10/10)+10^(G11/10)+10^(G12/10)+10^(G13/10)+10^(G14/10))"+ _
	"*(COS(9*PI()/24)-COS(11*PI()/24))*PI()/6+(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)"+ _
	"+10^(H7/10)+10^(H8/10)+10^(H9/10)+10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+"+ _
	"10^(H14/10))*(COS(11*PI()/24)-COS(PI()/2))*PI()/6))/(2*PI()*(1-COS(PI()/2)))),2)"

	sheet.Range("Q7").Formula = "=ROUND(10*LOG10((10^(Q6/10)*2*PI()*(1-COS(PI()/2))+"+ _
	"(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)+10^(H7/10)+10^(H8/10)+10^(H9/10)+"+ _
	"10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+10^(H14/10))*(COS(PI()/2)-"+ _
	"COS(13*PI()/24))*PI()/6+(10^(I3/10)+10^(I4/10)+10^(I5/10)+10^(I6/10)+10^(I7/10)"+ _
	"+10^(I8/10)+10^(I9/10)+10^(I10/10)+10^(I11/10)+10^(I12/10)+10^(I13/10)+"+ _
	"10^(I14/10))*(COS(13*PI()/24)-COS(15*PI()/24))*PI()/6+(10^(J3/10)+10^(J4/10)+"+ _
	"10^(J5/10)+10^(J6/10)+10^(J7/10)+10^(J8/10)+10^(J9/10)+10^(J10/10)+10^(J11/10)+"+ _
	"10^(J12/10)+10^(J13/10)+10^(J14/10))*(COS(15*PI()/24)-COS(16*PI()/24))*PI()/6)/(2*PI()*(1-COS(2*PI()/3)))),2)"

	sheet.Range("Q8").Formula = _
	"=ROUND((117-0.5*(1.75*(1.5*SUMPRODUCT((B3:E15<=-9)*(B3:E15>-100))+1.25*SUMPRODUCT((B3:E15<=-8)*(B3:E15>-9))"+ _
	"+SUMPRODUCT((B3:E15<=-7)*(B3:E15>-8))+0.75*SUMPRODUCT((B3:E15<=-6)*(B3:E15>-7))+0.5*SUMPRODUCT((B3:E15<=-5)"+ _
	"*(B3:E15>-6))+0.25*SUMPRODUCT((B3:E15<=-4)*(B3:E15>-5)))+1.5*SUMPRODUCT((F3:H15<=-9)*(F3:H15>-100))+"+ _
	"1.25*SUMPRODUCT((F3:H15<=-8)*(F3:H15>-9))+SUMPRODUCT((F3:H15<=-7)*(F3:H15>-8))+0.75*SUMPRODUCT((F3:H15<=-6)"+ _
	"*(F3:H15>-7))+0.5*SUMPRODUCT((F3:H15<=-5)*(F3:H15>-6))+0.25*SUMPRODUCT((F3:H15<=-4)*(F3:H15>-5))+0.5*"+ _
	"(1.5*SUMPRODUCT((I3:J15<=-9)*(I3:J15>-100))+1.25*SUMPRODUCT((I3:J15<=-8)*(I3:J15>-9))+SUMPRODUCT((I3:J15<=-7)*"+ _
	"(I3:J15>-8))+0.75*SUMPRODUCT((I3:J15<=-6)*(I3:J15>-7))+0.5*SUMPRODUCT((I3:J15<=-5)*(I3:J15>-6))"+ _
	"+0.25*SUMPRODUCT((I3:J15<=-4)*(I3:J15>-5)))))/117*100,2)"
End Sub




