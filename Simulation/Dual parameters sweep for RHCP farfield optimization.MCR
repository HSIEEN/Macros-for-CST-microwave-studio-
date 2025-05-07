'sweep single parameter to meet the target
'2023-01-10 By Shawn
'#include "vba_globals_all.lib"
Option Explicit
Public startTime As String, parameter1 As String, parameter2 As String, portStr As String, componentNames() As String
Public cutPlaneValue As Integer, farfieldComponentValue As Integer, cutAngle As Double
Public TE As Double, RE As Double

Sub Main ()
	Dim parameterArray(1000) As String
	Dim portArray(100) As String
	Dim FarfieldFreq() As Double
	Dim HasFarfieldMonitor As Boolean
    Dim TempStr As String
	Dim ii As Integer, i As Integer, m As Integer

	ReDim componentNames(2)
	componentNames(0) = "directivity"
	componentNames(1) = "gain"
	componentNames(2) = "realized gain"

	'Dim sAllSelectedParaNames As String
	'============Collect parameters and store them to an array================
	For ii = 0 To GetNumberOfParameters-1
		parameterArray(ii) = GetParameterName(ii)
	Next ii

	'===========Collect ports and store them to an array====================
	For ii = 0 To Port.StartPortNumberIteration-1 STEP 1
		portArray(ii) = Port.GetNextPortNumber
	Next

	'===========Collect farfield monitors and store them to an array=============
	ii = Monitor.GetNumberOfMonitors
    ReDim FarfieldFreq(ii-1)
    If ii < 1 Then
         MsgBox("No monitors found!",vbCritical,"Warning")
    	Exit Sub
    End If

	HasFarfieldMonitor  = False

    m = 0

    For i = 0 To ii-1 STEP 1
    	TempStr = Monitor.GetMonitorTypeFromIndex(i)

    	If InStr(Monitor.GetMonitorTypeFromIndex(i),"Farfield") <> 0 Then
    		'Dim MidValue0 As Double
    		'MidValue0 =  Monitor.GetMonitorFrequencyFromIndex(i)
    		FarfieldFreq(m) = Monitor.GetMonitorFrequencyFromIndex(i)
    		m = m+1

    		If HasFarfieldMonitor = False Then
    			HasFarfieldMonitor = True
    		End If

    	End If

    Next i
    'No farfield monitors, Exit

    If HasFarfieldMonitor = False Then
    	 MsgBox("No Farfield monitors found!",vbCritical,"Warning")
    	 Exit Sub
    End If
	'=============================================================================
	Dim InformationStr As String
    InformationStr = "Input values should be within "+CStr(FarfieldFreq(0))+"GHz to "+CStr(FarfieldFreq(m-1))+"GHz:"

	Begin Dialog UserDialog 1110,441,"Dual parameters sweep for farfield optimization" ' %GRID:10,7,1,1

		GroupBox 640,7,460,84,"Select two parameters:",.GroupBox1
		Text 710,35,80,14,"Parameter1:",.Text4
		DropListBox 810,28,170,21,parameterArray(),.parameterIndex1
		Text 710,56,80,14,"Parameter2:",.Text8
		DropListBox 810,56,170,21,parameterArray(),.parameterIndex2

		GroupBox 640,98,470,84,"Parameter sweep settings:",.GroupBox2
		'----------------------------------------------------------
		Text 670,119,80,14,"Parameter1:",.Text12
		Text 760,119,40,14,"from",.Text1
		TextBox 810,119,40,14,.xMin
		Text 860,119,20,14,"to",.Text2
		TextBox 890,119,40,14,.xMax
		Text 950,119,70,14,"step size:",.Text3
		TextBox 1030,119,50,14,.xStep
		'---------------------------------------------------------
		Text 670,147,80,14,"Parameter2:",.Text13
		Text 760,147,40,14,"from",.Text9
		TextBox 810,147,40,14,.yMin
		Text 860,147,20,14,"to",.Text10
		TextBox 890,147,40,14,.yMax
		Text 950,147,70,14,"step size:",.Text11
		TextBox 1030,147,50,14,.yStep

		GroupBox 640,189,470,42,"Select a port:",.GroupBox5
		DropListBox 830,203,80,21,portArray(),.portIndex

		GroupBox 640,245,470,56,"Frequency settings:",.GroupBox3
		Text 660,259,440,14,InformationStr,.Text15
		Text 730,280,40,14,"From:",.Text5
		TextBox 780,280,40,14,.fMin
		Text 830,280,30,14,"to:",.Text7
		TextBox 860,280,40,14,.fMax
		Text 920,280,40,14,"step:",.Text14
		TextBox 960,280,40,14,.fStep

		GroupBox 640,308,470,49,"Set plot mode:",.GroupBox6
		OptionGroup .Group2
			OptionButton 690,329,100,14,"Directivity",.OptionButton3
			OptionButton 820,329,70,14,"Gain",.OptionButton4
			OptionButton 920,329,120,14,"Realized Gain",.OptionButton5

		GroupBox 640,364,470,49,"Cut angle settings in 1D plot:",.GroupBox4
		OptionGroup .Group1
			OptionButton 700,385,40,14,"¦È",.OptionButton1
			OptionButton 770,385,40,14,"¦Õ",.OptionButton2
		Text 830,385,90,14,"with angle of",.Text6
		TextBox 940,385,40,14,.Angle

		Picture 0,7,630,434,GetInstallPath + "\Library\Macros\Coros\Simulation\Single parameter sweep instructions For RHCP optimization.bmp",0,.Picture1

		OKButton 750,420,90,21
		CancelButton 890,420,90,21

	End Dialog

	Dim dlg As UserDialog

	dlg.xMin = "0"
	dlg.xMax = "360"
	dlg.xStep = "10"

	dlg.yMin = "0"
	dlg.yMax = "360"
	dlg.yStep = "10"

	'dlg.Mono = "0"
	dlg.fMin = CStr(FarfieldFreq(0))
	dlg.fMax = CStr(FarfieldFreq(m-1))
	dlg.fStep = CStr(Round(FarfieldFreq(1)-FarfieldFreq(0),2))
	dlg.Group1 = 1
	dlg.Angle = "270"

	dlg.Group2 = 0

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

	'Dim parameter As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, theta0 As Double, phi0 As Double, directivity As Double
	Dim xSim As Double, ySim As Double
	Dim xStep As Double, yStep As Double
	Dim fmin As Double, fMax As Double, fStep As Double

	If dlg.parameterIndex1 = -1 Or dlg.parameterIndex2 = -1 Then
		MsgBox("No parameter is selected!!",vbCritical,"Error")
		Exit All
	End If

	If dlg.parameterIndex1 = dlg.parameterIndex2 Then
		MsgBox "Selected two parameters are with the same name, please check and re-run the macro again!", vbCritical, "Invalid inputs"
		Exit All
	End If

	parameter1 = parameterArray(dlg.parameterIndex1)
	parameter2 = parameterArray(dlg.parameterIndex2)
	portStr = portArray(dlg.portIndex)

	If DoesParameterExist(parameter1) = False Or DoesParameterExist(parameter1) = False  Then
		MsgBox("The input paramters does not exist!!",vbCritical,"Error")
		Exit All
    End If


    fmin = Evaluate(dlg.fMin)
    fMax = Evaluate(dlg.fMax)
	fStep = Evaluate(dlg.fStep)
	'============Check validity of frequency================
	'Dim monitorNumber As Integer
	'Dim monitorFrequency As Double
	'Dim flag1 As Boolean, flag2 As Boolean

	'flag1 = False
	'flag2 = False

	'monitorNumber = Monitor.GetNumberOfMonitors()
	'For i = 0 To monitorNumber-1 STEP 1

	'	If Monitor.GetMonitorTypeFromIndex(i)="Farfield" Then
	'		monitorFrequency = Monitor.GetMonitorFrequencyFromIndex(i)
	'		If Abs(monitorFrequency-f1)<1e-5 Then
	'			flag1 = True
	'		End If
	'		If Abs(monitorFrequency-f2)<1e-5 Then
	'			flag2 = True
	'		End If
	'	End If
	'Next

	'If (f1 <> 0  And flag1 = False) Or (f2 <> 0 And flag2 = False) Then
	'	MsgBox("The input frequencies do not exist in farfield monitors!!",vbCritical,"Error")
	'	Exit All
	'End If
	If fmin - FarfieldFreq(0)<-1e-2 Or fMax - FarfieldFreq(m-1)>1e-2 Then
		MsgBox("Input frequency is out of Range",1,"Warning")
		Exit All
	End If

	Dim n As Integer

	n = 0
	m = 0
	'rotateAngle = xMin
	xMin = Evaluate(dlg.xMin)
    xMax = Evaluate(dlg.xMax)
    yMin = Evaluate(dlg.yMin)
    yMax = Evaluate(dlg.yMax)
    xStep = Evaluate(dlg.xStep)
    yStep = Evaluate(dlg.yStep)
	xSim = xMin
	ySim = yMin

	Dim projectPath As String
	Dim logFile As String

	cutPlaneValue = dlg.Group1
	cutAngle = Evaluate(dlg.Angle)

	'farfieldComponentValue = 0 for directivity, 1 for gain, 2 for realized gain
	farfieldComponentValue = dlg.Group2

	projectPath = GetProjectPath("Project")
	startTime = Replace(CStr(Time),":","_")

	logFile = projectPath + "\"+componentNames(farfieldComponentValue)+" sweep log_"+startTime+".txt"


   	Open logFile For Output As #2
   	Print #2, "##########Parameters sweep begins at " + CStr(Now) +"'###########."
   	Print #2, " "

	FarfieldPlot.SetPlotMode(componentNames(farfieldComponentValue))

	If FarfieldPlot.IsScaleLinear = True Then
		FarfieldPlot.SetScaleLinear(False)
   	End If

   	Dim simulation_loops As Integer
   	simulation_loops = 0

    While xSim <= xMax
    	m = 0
    	ySim = yMin
    	While ySim <= yMax
			Print #2, "%-%-% At step "+CStr(simulation_loops+1)+": "+parameter1+"="+CStr(xSim)+" and "+parameter2+"="+CStr(ySim)
			Print #2, "%-%-% Start time: " + CStr(Now) +"."
			'run simulation with specified value of xSim
			runWithParameter(parameter1,xSim, parameter2, ySim)
			Print #2, "%-%-%  Step "+CStr(simulation_loops+1)+" simulation is done."
			For i = 0 To UBound(FarfieldFreq) STEP 1

		    	If FarfieldFreq(i) >= fmin And FarfieldFreq(i) <= fMax Then
					'Check if the frquency is valid
					Dim quotient As Double
					quotient = (FarfieldFreq(i) - fmin)/fStep
					If Abs(quotient - Int(quotient))<1e-6 Then
						directivity = Copy1DFarfieldResult(xSim, ySim, FarfieldFreq(i))
						Print #2, "%-%-% RHCP "+componentNames(farfieldComponentValue)+" at frequency "+ CStr(FarfieldFreq(i))+ "Ghz is "+ CStr(Round(directivity, 2)) + "dBi."
				    End If
			    End If
		    Next i

			Print #2, "%-%-% Finish time: " + CStr(Now) +"."
			Print #2, " "
			simulation_loops = simulation_loops + 1
			m = m+1
			If yStep = 0 Then
				Exit While
			End If
			ySim = yMin + m*yStep
		Wend
		n = n+1
		If xStep = 0 Then
			Exit While
		End If
		xSim = xMin + n*xStep
	Wend

	Print #2, "##########Parameters sweep ends at " + CStr(Now) +"'###########."
	Close #2

	MsgBox("Maximum value reached, the sweep progress finished.",vbInformation, "Attention")

End Sub
Sub runWithParameter(para1 As String, value1 As Double, para2 As String, value2 As Double)

 	StoreParameter(para1,value1)
 	StoreParameter(para2,value2)
	Rebuild
	Solver.MeshAdaption(False)
	Solver.SteadyStateLimit(-40)
	Solver.Start

End Sub

Function Copy1DFarfieldResult(xSim As Double, ySim As Double, freq As Double)
	'parameters: cutPlaneValue denotes the cutting plane,0->theta, 1->phi; cutAngle denotes the angle in the plane specified by cutPlaneValue; theta0 and phi0
	'denote the angle where the directivity is estimated; rotateAngle is the rotate angle of the ham; freq is the operation frequency we take care
		Dim selectedItem As String, frequencyStr As String

		'portStr = "1"
		frequencyStr = CStr(freq)

		SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")

		FarfieldPlot.Reset
		FarfieldPlot.PlotType("polar")

		If FarfieldPlot.IsScaleLinear = True Then
			FarfieldPlot.SetScaleLinear(False)
	   	End If

		If cutPlaneValue = 0 Then

			FarfieldPlot.Vary("angle2")
			FarfieldPlot.Theta(cutAngle)

		Else

			FarfieldPlot.Vary("angle1")
			FarfieldPlot.Phi(cutAngle)

		End If

		FarfieldPlot.SetAxesType("currentwcs")
		FarfieldPlot.SetAntennaType("unknown")

		FarfieldPlot.SetPlotMode(componentNames(farfieldComponentValue))

		FarfieldPlot.SetCoordinateSystemType("ludwig3")
		FarfieldPlot.SetAutomaticCoordinateSystem("True")
		FarfieldPlot.SetPolarizationType("Circular")
		FarfieldPlot.StoreSettings
		FarfieldPlot.Plot

		'FarfieldPlot.Plot
		Dim DirName As String
		DirName = "CP "+componentNames(farfieldComponentValue)+"@port="+portStr+"\"+parameter1+"="+CStr(xSim)+" "+parameter2+"="+CStr(ySim)+ "@"+frequencyStr+"GHz"

		Dim ChildItem As String
		If ResultTree.DoesTreeItemExist("1D Results\"+DirName) Then
			ChildItem = ResultTree.GetFirstChildName("1D Results\"+DirName)
			While ChildItem <> ""
				With ResultTree
					.Name ChildItem
					.Delete
				End With
				ChildItem = ResultTree.GetFirstChildName("1D Results\"+DirName)
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
		If FarfieldPlot.GetSystemTotalEfficiency > -100 Then

			TE = FarfieldPlot.GetSystemTotalEfficiency
			RE = FarfieldPlot.GetSystemRadiationEfficiency
		Else
			TE = FarfieldPlot.GetTotalEfficiency
			RE = FarfieldPlot.GetRadiationEfficiency
		End If
		savefarfieldComponent(xSim, ySim, freq)
        'CurrentItem = FirstChildItem
        Copy1DFarfieldResult = getfarfieldComponent()
        SelectTreeItem("1D Results\"+DirName)

		Dim curveLabel As String
		Dim index As Integer
		'Dim SelectedItem As String

		selectedItem = ResultTree.GetFirstChildName("1D Results\"+DirName)
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

			selectedItem = ResultTree.GetNextItemName(selectedItem)
		Wend

        'Plot1D.SetFont("Tahoma","bold","16")
End Function

Function getfarfieldComponent()

	Dim farfieldComponent As String
	Dim directivity As Double

	farfieldComponent = "ludwig3 circular right abs"
	directivity = FarfieldPlot.CalculatePoint(0, 0, farfieldComponent, "")
	getfarfieldComponent = directivity

End Function
Sub savefarfieldComponent(xSim As Double, ySim As Double, frequency As Double)

    Dim selectedItem As String
    Dim n As Integer
    Dim frequencyStr As String

	'portStr = "1"
	frequencyStr = CStr(frequency)

	SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")
    '==============Upper Hemisphere rhcp directivity and rhcp directivity estimation===============

    Dim  upperHemisphereRHCPdirectivity() As Double, upperHemisphereLHCPdirectivity() As Double
    Dim rhcpDirectivity() As Double, lhcpDirectivity() As Double
    Dim Theta As Double, Phi As Double
    Dim position_theta() As Double, position_phi() As Double
    Dim Columns As String
    Dim xlsxFile As String
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

    'missmatchLoss = Round(FarfieldPlot.GetTotalEfficiency - FarfieldPlot.GetRadiationEfficiency,2)
	 '==============================write directivity data============================
	projectPath = GetProjectPath("Project")

	xlsxFile = projectPath+"\Circularly polarized "+componentNames(farfieldComponentValue)+"_frequency="+frequencyStr+"GHz Port="+portStr+ _
   	" "+parameter1+"="+CStr(xSim)+" "+parameter2+"="+CStr(ySim)+ " @" +Replace(CStr(Time),":","-")+".xlsx"


	Columns = "BCDEFGHIJKLMN"

	Dim NoticeInformation As String
	Dim O As Object
	NoticeInformation = "The "+componentNames(farfieldComponentValue)+" data is under£¨"+projectPath+"\£©"

    ReportInformationToWindow(NoticeInformation)

    Set O = CreateObject("Excel.Application")
	If Dir(xlsxFile) = "" Then
	    Dim wBook As Object
	    Set wBook = O.Workbooks.Add
	    With wBook
	        .Title = "Title"
	        .Subject = "Subject"
	        .SaveAs Filename:= xlsxFile
		End With
	Else
	    Set wBook = O.Workbooks.Open(xlsxFile)
	End If


	'Add a sheet and rename it
	Dim wSheet As Object
	Dim sheetName As String
	sheetName= parameter1+"="+CStr(xSim)+" "+parameter2+"="+CStr(ySim)
	If Len(sheetName) >= 30 Then
		MsgBox "Name of a sheet is too long, please make change!!", vbCritical
		Exit All
	End If
	wBook.Sheets.Add.Name = parameter1+"="+CStr(xSim)+" "+parameter2+"="+CStr(ySim)
	Set wSheet = wBook.Sheets(parameter1+"="+CStr(xSim)+" "+parameter2+"="+CStr(ySim))

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

	'formatting cells

	sheet.Columns("A").ColumnWidth = 14
	sheet.Columns("C").ColumnWidth = 10
	sheet.Rows("1").RowHeight = 25
	sheet.Rows("35").RowHeight = 25
	sheet.Range("A1:Z100").HorizontalAlignment =  -4108 	'Center
	sheet.Range("P2:Q8").Borders.LineStyle = 1			 'Continous border

	Dim Dvalue As Double, deltaDirectivity As Double, axialRatio As Double
	For i  = 0 To 12
		For j = 0 To 12

			'=======================Axial ratio estimating and coloring============================
			deltaDirectivity = sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value - sheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).Value

			axialRatio = Sgn(deltaDirectivity-0.0001)*20*CST_Log10((10^(deltaDirectivity/20)+1)/(Abs(10^(deltaDirectivity/20)-1)+0.0001))

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

			Dvalue = sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value

			Select Case farfieldComponentValue
		   	Case 0	'reference total efficiency -7dB when the directivity is selected as the farfield component
				'Dvalue = Dvalue
				Dvalue = Dvalue+TE
			Case 1
				Dvalue = Dvalue+TE-RE
			Case 2
				Dvalue = Dvalue
		   	End Select


			If Dvalue >= -6 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 130, 0)
			ElseIf Dvalue < -6 And Dvalue >= -8 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 180, 0)
			ElseIf Dvalue <-8 And Dvalue >= -10 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(145, 218, 0)
			ElseIf Dvalue < -10 And Dvalue >= -12 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(216, 254, 154)
			ElseIf Dvalue < -12 And Dvalue >= -14 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 255, 0)
			ElseIf Dvalue < -14 And Dvalue >= -16 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 200, 0)
			ElseIf Dvalue < -16 And Dvalue >= -20 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 0, 0)
			ElseIf Dvalue < -20 Then
				sheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(150, 0, 0)
			End If

		Next
	Next
	'======================================UH/Tot===============================
	'sheet.Range("A51") = "UHPower ratio"
	'sheet.Range("B51") = Round(10*CST_Log10(getUpperHemisphereRatio(sheet, columns)),2)
	'sheet.Range("C51") = "dB"
	writeAverageFarfieldComponentAndRating(sheet, Columns)

End Sub
Sub writeAverageFarfieldComponentAndRating(sheet As Object, Columns As String)

	'Formatting cells
	With sheet
	    .Columns("P").ColumnWidth = 15
	    .Columns("Q").ColumnWidth = 33
	    .Range("P2,P8,Q2").Interior.Color = RGB(221, 235, 247)
	    .Range("P3:P7").Interior.Color = RGB(0, 176, 240)
	    .Range("Q3:Q8").Interior.Color = RGB(255, 217, 102)
	    .Range("A1:Q100").Font.Bold = True
	    .Range("Q3:Q8").Font.Color = RGB(255, 0, 0)
	End With

	With sheet
	    .Range("P2").Value = "Within theta"
	    .Range("P3").Value = "30"
	    .Range("P4").Value = "45"
	    .Range("P5").Value = "60"
	    .Range("P6").Value = "90"
	    .Range("P7").Value = "120"
	    .Range("P8").Value = "RHCPD rating"
	    .Range("Q2").Value = "Weighted Average directivity"
	End With

	sheet.Range("Q3").Formula = _
	"=ROUND(10*LOG10(((10^(B3/10)+10^(B4/10)+10^(B5/10)+10^(B6/10)"+ _
	"+10^(B7/10)+10^(B8/10)+10^(B9/10)+10^(B10/10)+10^(B11/10)+10^(B12/10)+10^(B13/10)+"+ _
	"10^(B14/10))*(1-COS(PI()/24))*PI()/6+(10^(C3/10)+10^(C4/10)+10^(C5/10)+10^(C6/10)+"+ _
	"10^(C7/10)+10^(C8/10)+10^(C9/10)+10^(C10/10)+10^(C11/10)+10^(C12/10)+10^(C13/10)+"+ _
	"10^(C14/10))*(COS(PI()/24)-COS(PI()/8))*PI()/6+(10^(D3/10)+10^(D4/10)+10^(D5/10)+"+ _
	"10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+10^(D10/10)+10^(D11/10)+10^(D12/10)+"+ _
	"10^(D13/10)+10^(D14/10))*(COS(PI()/8)-COS(PI()/6))*PI()/6)/(2*PI()*(1-COS(PI()/6)))),2)"

	sheet.Range("Q4").Formula = _
	"=ROUND(10*LOG10(((10^(Q3/10)*(2*PI()*(1-COS(PI()/6)))+"+ _
	"(10^(D3/10)+10^(D4/10)+10^(D5/10)+10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+"+ _
	"10^(D10/10)+10^(D11/10)+10^(D12/10)+10^(D13/10)+10^(D14/10))*(COS(PI()/6)-"+ _
	"COS(5*PI()/24))*PI()/6+(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+"+ _
	"10^(E8/10)+10^(E9/10)+10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))"+ _
	"*(COS(5*PI()/24)-COS(PI()/4))*PI()/6)/(2*PI()*(1-COS(PI()/4))))),2)"

	sheet.Range("Q5").Formula = _
	"=ROUND(10*LOG10(((10^(Q4/10)*(2*PI()*(1-COS(PI()/4)))+"+ _
	"(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+10^(E8/10)+10^(E9/10)+"+ _
	"10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))*(COS(PI()/4)-"+ _
	"COS(7*PI()/24))*PI()/6+((10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+"+ _
	"10^(F7/10)+10^(F8/10)+10^(F9/10)+10^(F10/10)+10^(F11/10)+10^(F12/10)+"+ _
	"10^(F13/10)+10^(F14/10))*(COS(7*PI()/24)-COS(PI()/3))*PI()/6))/(2*PI()*(1-COS(PI()/3))))),2)"

	sheet.Range("Q6").Formula = _
	"=ROUND(10*LOG10(((10^(Q5/10)*2*PI()*(1-COS(PI()/3))+"+ _
	"(10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+10^(F7/10)+10^(F8/10)+10^(F9/10)+"+ _
	"10^(F10/10)+10^(F11/10)+10^(F12/10)+10^(F13/10)+10^(F14/10))*(COS(8*PI()/24)-"+ _
	"COS(9*PI()/24))*PI()/6+(10^(G3/10)+10^(G4/10)+10^(G5/10)+10^(G6/10)+10^(G7/10)+"+ _
	"10^(G8/10)+10^(G9/10)+10^(G10/10)+10^(G11/10)+10^(G12/10)+10^(G13/10)+10^(G14/10))"+ _
	"*(COS(9*PI()/24)-COS(11*PI()/24))*PI()/6+(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)"+ _
	"+10^(H7/10)+10^(H8/10)+10^(H9/10)+10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+"+ _
	"10^(H14/10))*(COS(11*PI()/24)-COS(PI()/2))*PI()/6))/(2*PI()*(1-COS(PI()/2)))),2)"

	sheet.Range("Q7").Formula = _
	"=ROUND(10*LOG10((10^(Q6/10)*2*PI()*(1-COS(PI()/2))+"+ _
	"(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)+10^(H7/10)+10^(H8/10)+10^(H9/10)+"+ _
	"10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+10^(H14/10))*(COS(PI()/2)-"+ _
	"COS(13*PI()/24))*PI()/6+(10^(I3/10)+10^(I4/10)+10^(I5/10)+10^(I6/10)+10^(I7/10)"+ _
	"+10^(I8/10)+10^(I9/10)+10^(I10/10)+10^(I11/10)+10^(I12/10)+10^(I13/10)+"+ _
	"10^(I14/10))*(COS(13*PI()/24)-COS(15*PI()/24))*PI()/6+(10^(J3/10)+10^(J4/10)+"+ _
	"10^(J5/10)+10^(J6/10)+10^(J7/10)+10^(J8/10)+10^(J9/10)+10^(J10/10)+10^(J11/10)+"+ _
	"10^(J12/10)+10^(J13/10)+10^(J14/10))*(COS(15*PI()/24)-COS(16*PI()/24))*PI()/6)/(2*PI()*(1-COS(2*PI()/3)))),2)"

	Select Case farfieldComponentValue
  	Case 0	'reference total efficiency -10dB when the directivity is selected as the farfield component
		sheet.Range("Q8").Formula = _
		"=ROUND((88-0.5*(1.5*(1.75*SUMPRODUCT((B4:F15<=(-17-"+CStr(RE)+"))*(B4:F15>-100))+"+ _
		"1.5*SUMPRODUCT((B4:F15<=(-16-"+CStr(RE)+"))*(B4:F15>(-17-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((B4:F15<=(-15-"+CStr(RE)+"))*(B4:F15>(-16-"+CStr(RE)+")))+"+ _
		"0.75*SUMPRODUCT((B4:F15<=(-14-"+CStr(RE)+"))*(B4:F15>(-15-"+CStr(RE)+")))+"+ _
		"0.5*SUMPRODUCT((B4:F15<=(-13-"+CStr(RE)+"))*(B4:F15>(-14-"+CStr(RE)+")))+"+ _
		"0.25*SUMPRODUCT((B4:F15<=(-11-"+CStr(RE)+"))*(B4:F15>(-13-"+CStr(RE)+"))))+"+ _
		"1.3*(1.75*(SUMPRODUCT((G4:H6<=(-17-"+CStr(RE)+"))*(G4:H6>-100))+"+ _
		"SUMPRODUCT((G12:H15>-100)*(G12:H15<=(-17-"+CStr(RE)+"))))+"+ _
		"1.5*(SUMPRODUCT((G4:H6<=(-16-"+CStr(RE)+"))*(G4:H6>(-17-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-16-"+CStr(RE)+"))*(G12:H15>(-17-"+CStr(RE)+"))))+"+ _
		"1*(SUMPRODUCT((G4:H6<=(-15-"+CStr(RE)+"))*(G4:H6>(-16-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-15-"+CStr(RE)+"))*(G12:H15>(-16-"+CStr(RE)+"))))+"+ _
		"0.75*(SUMPRODUCT((G4:H6<=(-14-"+CStr(RE)+"))*(G4:H6>(-15-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((G12:H15=(-14-"+CStr(RE)+"))*(G12:H15>(-15-"+CStr(RE)+"))))+"+ _
		"0.5*(SUMPRODUCT((G4:H6<=(-13-"+CStr(RE)+"))*(G4:H6>(-14-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-13-"+CStr(RE)+"))*(G12:H15>(-14-"+CStr(RE)+"))))+"+ _
		"0.25*(SUMPRODUCT((G4:H6<=(-11-"+CStr(RE)+"))*(G4:H6>(-13-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-11-"+CStr(RE)+"))*(G12:H15>(-13-"+CStr(RE)+")))))+"+ _
		"1.1*(1.75*(SUMPRODUCT((I4:J6<=(-17-"+CStr(RE)+"))*(I4:J6>-100))+"+ _
		"SUMPRODUCT((I12:J15<=(-17-"+CStr(RE)+"))*(I12:J15>-100)))+"+ _
		"1.5*(SUMPRODUCT((I4:J6<=(-16-"+CStr(RE)+"))*(I4:J6>(-17-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-16+"+CStr(RE)+"))*(I12:J15>(-17-"+CStr(RE)+"))))+"+ _
		"(SUMPRODUCT((I4:J6<=(-15-"+CStr(RE)+"))*(I4:J6>(-16-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-15-"+CStr(RE)+"))*(I12:J15>(-16-"+CStr(RE)+"))))+"+ _
		"0.75*(SUMPRODUCT((I4:J6<=(-14-"+CStr(RE)+"))*(I4:J6>(-15-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-14-"+CStr(RE)+"))*(I12:J15>(-15-"+CStr(RE)+"))))+"+ _
		"0.5*(SUMPRODUCT((I4:J6<=(-13-"+CStr(RE)+"))*(I4:J6>(-14-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-13-"+CStr(RE)+"))*(I12:J15>(-14-"+CStr(RE)+"))))+"+ _
		"0.25*(SUMPRODUCT((I4:J6<=(-11-"+CStr(RE)+"))*(I4:J6>(-13-"+CStr(RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-11-"+CStr(RE)+"))*(I12:J15>(-13-"+CStr(RE)+")))))))/88*100,2)"
	Case 1
		sheet.Range("Q8").Formula = _
		"=ROUND((88-0.5*(1.5*(1.75*SUMPRODUCT((B4:F15<=-17)*(B4:F15>-100))+"+ _
		"1.5*SUMPRODUCT((B4:F15<=-16)*(B4:F15>-17))+SUMPRODUCT((B4:F15<=-15)*(B4:F15>-16))+"+ _
		"0.75*SUMPRODUCT((B4:F15<=-14)*(B4:F15>-15))+0.5*SUMPRODUCT((B4:F15<=-13)*(B4:F15>-14))+"+ _
		"0.25*SUMPRODUCT((B4:F15<=-11)*(B4:F15>-13)))+1.3*(1.75*(SUMPRODUCT((G4:H6<=-17)*(G4:H6>-100))+"+ _
		"SUMPRODUCT((G12:H15>-100)*(G12:H15<=-17)))+1.5*(SUMPRODUCT((G4:H6<=-16)*(G4:H6>-17))+"+ _
		"SUMPRODUCT((G12:H15<=-16)*(G12:H15>-17)))+1*(SUMPRODUCT((G4:H6<=-15)*(G4:H6>-16))+"+ _
		"SUMPRODUCT((G12:H15<=-15)*(G12:H15>-16)))+0.75*(SUMPRODUCT((G4:H6<=-14)*(G4:H6>-15))+"+ _
		"SUMPRODUCT((G12:H15=-14)*(G12:H15>-15)))+0.5*(SUMPRODUCT((G4:H6<=-13)*(G4:H6>-14))+"+ _
		"SUMPRODUCT((G12:H15<=-13)*(G12:H15>-14)))+0.25*(SUMPRODUCT((G4:H6<=-11)*(G4:H6>-13))+"+ _
		"SUMPRODUCT((G12:H15<=-11)*(G12:H15>-13))))+1.1*(1.75*(SUMPRODUCT((I4:J6<=-17)*(I4:J6>-100))+"+ _
		"SUMPRODUCT((I12:J15<=-17)*(I12:J15>-100)))+1.5*(SUMPRODUCT((I4:J6<=-16)*(I4:J6>-17))+"+ _
		"SUMPRODUCT((I12:J15<=-16)*(I12:J15>-17)))+(SUMPRODUCT((I4:J6<=-15)*(I4:J6>-16))+"+ _
		"SUMPRODUCT((I12:J15<=-15)*(I12:J15>-16)))+0.75*(SUMPRODUCT((I4:J6<=-14)*(I4:J6>-15))+"+ _
		"SUMPRODUCT((I12:J15<=-14)*(I12:J15>-15)))+0.5*(SUMPRODUCT((I4:J6<=-13)*(I4:J6>-14))+"+ _
		"SUMPRODUCT((I12:J15<=-13)*(I12:J15>-14)))+0.25*(SUMPRODUCT((I4:J6<=-11)*(I4:J6>-13))+"+ _
		"SUMPRODUCT((I12:J15<=-11)*(I12:J15>-13))))))/88*100,2)"

	Case 2
		sheet.Range("Q8").Formula = _
		"=ROUND((88-0.5*(1.5*(1.75*SUMPRODUCT((B4:F15<=(-17+"+CStr(TE-RE)+"))*(B4:F15>-100))+"+ _
		"1.5*SUMPRODUCT((B4:F15<=(-16+"+CStr(TE-RE)+"))*(B4:F15>(-17+"+CStr(TE-RE)+")))+SUMPRODUCT((B4:F15<=(-15+"+CStr(TE-RE)+"))*(B4:F15>(-16+"+CStr(TE-RE)+")))+"+ _
		"0.75*SUMPRODUCT((B4:F15<=(-14+"+CStr(TE-RE)+"))*(B4:F15>(-15+"+CStr(TE-RE)+")))+0.5*SUMPRODUCT((B4:F15<=(-13+"+CStr(TE-RE)+"))*(B4:F15>(-14+"+CStr(TE-RE)+")))+"+ _
		"0.25*SUMPRODUCT((B4:F15<=(-11+"+CStr(TE-RE)+"))*(B4:F15>(-13+"+CStr(TE-RE)+"))))+1.3*(1.75*(SUMPRODUCT((G4:H6<=(-17+"+CStr(TE-RE)+"))*(G4:H6>-100))+"+ _
		"SUMPRODUCT((G12:H15>-100)*(G12:H15<=(-17+"+CStr(TE-RE)+"))))+1.5*(SUMPRODUCT((G4:H6<=(-16+"+CStr(TE-RE)+"))*(G4:H6>(-17+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-16+"+CStr(TE-RE)+"))*(G12:H15>(-17+"+CStr(TE-RE)+"))))+1*(SUMPRODUCT((G4:H6<=(-15+"+CStr(TE-RE)+"))*(G4:H6>(-16+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-15+"+CStr(TE-RE)+"))*(G12:H15>(-16+"+CStr(TE-RE)+"))))+0.75*(SUMPRODUCT((G4:H6<=(-14+"+CStr(TE-RE)+"))*(G4:H6>(-15+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((G12:H15=(-14+"+CStr(TE-RE)+"))*(G12:H15>(-15+"+CStr(TE-RE)+"))))+0.5*(SUMPRODUCT((G4:H6<=(-13+"+CStr(TE-RE)+"))*(G4:H6>(-14+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-13+"+CStr(TE-RE)+"))*(G12:H15>(-14+"+CStr(TE-RE)+"))))+0.25*(SUMPRODUCT((G4:H6<=(-11+"+CStr(TE-RE)+"))*(G4:H6>(-13+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((G12:H15<=(-11+"+CStr(TE-RE)+"))*(G12:H15>(-13+"+CStr(TE-RE)+")))))+1.1*(1.75*(SUMPRODUCT((I4:J6<=(-17+"+CStr(TE-RE)+"))*(I4:J6>-100))+"+ _
		"SUMPRODUCT((I12:J15<=(-17+"+CStr(TE-RE)+"))*(I12:J15>-100)))+1.5*(SUMPRODUCT((I4:J6<=(-16+"+CStr(TE-RE)+"))*(I4:J6>(-17+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-16+"+CStr(TE-RE)+"))*(I12:J15>(-17+"+CStr(TE-RE)+"))))+(SUMPRODUCT((I4:J6<=(-15+"+CStr(TE-RE)+"))*(I4:J6>(-16+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-15+"+CStr(TE-RE)+"))*(I12:J15>(-16+"+CStr(TE-RE)+"))))+0.75*(SUMPRODUCT((I4:J6<=(-14+"+CStr(TE-RE)+"))*(I4:J6>(-15+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-14+"+CStr(TE-RE)+"))*(I12:J15>(-15+"+CStr(TE-RE)+"))))+0.5*(SUMPRODUCT((I4:J6<=(-13+"+CStr(TE-RE)+"))*(I4:J6>(-14+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-13+"+CStr(TE-RE)+"))*(I12:J15>(-14+"+CStr(TE-RE)+"))))+0.25*(SUMPRODUCT((I4:J6<=(-11+"+CStr(TE-RE)+"))*(I4:J6>(-13+"+CStr(TE-RE)+")))+"+ _
		"SUMPRODUCT((I12:J15<=(-11+"+CStr(TE-RE)+"))*(I12:J15>(-13+"+CStr(TE-RE)+")))))))/88*100,2)"
   	End Select
End Sub



