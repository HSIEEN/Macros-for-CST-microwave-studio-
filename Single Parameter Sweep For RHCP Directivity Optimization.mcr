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
   	dataFile = projectPath + "\sweeplog.txt"
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

		DirName = "CP directivity\rotateAngle="+CStr(rotateAngle)+ "@"+FrequencyStr+"GHz"
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
         rhcpDirectivity(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = upperHemisphereRHCPdirectivity(n)'10*CST_Log10(upperHemisphereRHCPdirectivity(n)) 'Log(upperHemisphereRHCPdirectivity(n)/AVGPower)/Log(10)*10
         lhcpDirectivity(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = upperHemisphereLHCPdirectivity(n)'10*CST_Log10(upperHemisphereLHCPdirectivity(n))'Log(upperHemisphereLHCPdirectivity(n)/AVGPower)/Log(10)*10
    Next n

	projectPath = GetProjectPath("Project")
	dataFile = projectPath+"\Circularly polarized directivity @"+FrequencyStr+"GHz and Port "+PortStr+".xlsx"
	Columns = "BCDEFGHIJKLMN"

	NoticeInformation = "The directivity data is under（"+projectPath+"\）"
    ReportInformationToWindow(NoticeInformation)
	Dim IsFileExist As String
    Set O = CreateObject("Excel.Application")
	IsFileExist = Dir(dataFile)
	If IsFileExist = "" Then
		Set wBook  = O.Workbooks.Add
		With wBook
			.Title = "Title"
			.Subject = "Subject"
			.SaveAs Filename:= dataFile
		End With
	Else
		Set wBook = O.Workbooks.Open(dataFile)
	End If


	'Add a sheet and rename it

	wBook.Sheets.Add.Name = CStr(rotateAngle)
	Set wSheet = wBook.Sheets(CStr(rotateAngle))

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
	'estimate axial ratio
	'coloring
	'scores?

	wBook.Save
	O.ActiveWorkbook.Close
	O.quit

End Sub



