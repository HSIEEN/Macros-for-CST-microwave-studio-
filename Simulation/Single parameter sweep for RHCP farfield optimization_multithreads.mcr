'sweep single parameter to meet the target
'2023-01-10 By Shawn
'#include "vba_globals_all.lib"

Option Explicit
Public startTime As String, parameter As String, portStr As String, componentNames() As String
Public cutPlaneValue As Integer, farfieldComponentValue As Integer, cutAngle As Double
Public TE As Double, RE As Double

Sub Main ()
	MsgBox "Please set solver to store results in cache and apply it!", vbExclamation
	Dim parameterArray(1000) As String
	Dim portArray(100) As String
	Dim FarfieldFreq() As Double
	Dim HasFarfieldMonitor As Boolean
    Dim TempStr As String
	Dim ii As Integer, m As Integer, i As Integer

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

    	If StrComp(TempStr, "Farfield") = 0 Then
    		'Dim MidValue0 As Double
    		'MidValue0 =  Monitor.GetMonitorFrequencyFromIndex(i)
    		FarfieldFreq(m) = Monitor.GetMonitorFrequencyFromIndex(i)
    		m = m+1

    		If HasFarfieldMonitor = False Then
    			HasFarfieldMonitor = True
    		End If

    	End If

    Next i
    If HasFarfieldMonitor = False Then
    	 MsgBox("No Farfield monitors found!",vbCritical,"Warning")
    	 Exit Sub
    End If

	Begin Dialog UserDialog 990,427,"Single parameter sweep for farfield optimization" ' %GRID:10,7,1,1
		GroupBox 640,7,340,42,"Select a parameter:",.GroupBox1
		DropListBox 740,21,170,21,parameterArray(),.parameterIndex

		GroupBox 640,56,340,42,"Parameter sweep settings:",.GroupBox2
		Text 650,77,40,14,"From",.Text1
		TextBox 700,77,40,14,.xMin
		Text 750,77,20,14,"to",.Text2
		TextBox 780,77,40,14,.xMax
		Text 840,77,80,14,"stepwidth:",.Text3
		TextBox 920,77,50,14,.xStep

		GroupBox 640,140,340,105,"Frequency settings:",.GroupBox3
		TextBox 680,175,260,14,.sampleList
		Text 670,196,30,14,"Or",.Text10
		Text 670,217,40,14,"From:",.Text5
		Text 670,154,280,14,"Samples: use semicolon as a separator",.Text9
		TextBox 710,217,40,14,.fMin
		Text 760,217,30,14,"to:",.Text7
		TextBox 790,217,40,14,.fMax
		Text 840,217,70,14,"stepwidth:",.Text8
		TextBox 920,217,40,14,.fStep

		GroupBox 640,252,340,42,"Multithreads settings:",.GroupBox7
		Text 660,273,210,14,"Number of parallel simulations:",.Text4
		TextBox 870,273,40,14,.paraThreads

		GroupBox 640,350,340,42,"Cut angle settings in 1D plot:",.GroupBox4
		OptionGroup .Group1
			OptionButton 670,371,40,14,"¦È",.OptionButton1
			OptionButton 740,371,40,14,"¦Õ",.OptionButton2
		Text 800,371,90,14,"with angle of",.Text6
		TextBox 900,371,40,14,.Angle

		Picture 0,7,630,420,GetInstallPath + "\Library\Macros\Coros\Simulation\Single parameter sweep instructions For RHCP optimization.bmp",0,.Picture1
		GroupBox 640,98,340,42,"Select a port:",.GroupBox5
		DropListBox 780,105,80,21,portArray(),.portIndex

		GroupBox 640,301,340,42,"Set plot mode:",.GroupBox6
		OptionGroup .Group2
			OptionButton 650,315,100,14,"Directivity",.OptionButton3
			OptionButton 770,315,70,14,"Gain",.OptionButton4
			OptionButton 840,315,120,14,"Realized Gain",.OptionButton5

		OKButton 680,399,90,21
		CancelButton 830,399,90,21


	End Dialog

	Dim dlg As UserDialog
	dlg.xMin = "0"
	dlg.xMax = "100"
	dlg.xStep = "1"
	'dlg.Mono = "0"
	dlg.sampleList = "0;0"
	dlg.fMin = CStr(FarfieldFreq(0))
	dlg.fMax = CStr(FarfieldFreq(m-1))
	dlg.fStep = CStr(Round(FarfieldFreq(1)-FarfieldFreq(0),2))
	dlg.paraThreads = "4"
	dlg.Group1 = 1
	dlg.Angle = "270"

	dlg.Group2 = 0

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

	'Dim parameter As String
	Dim xMin As Double, xMax As Double, theta0 As Double, phi0 As Double, directivity As Double
	Dim xSim As Double
	Dim xStep As Double
	Dim fMin As Double, fMax As Double, fStep As Double, fSamples As String
	Dim threads As Integer

	If dlg.parameterIndex = -1 Then
		MsgBox("No parameter is selected!!",vbCritical,"Error")
		Exit All
	End If

	parameter = parameterArray(dlg.parameterIndex)
	portStr = portArray(dlg.portIndex)

	If DoesParameterExist(parameter) = False Then
		MsgBox("The input paramter does not exist!!",vbCritical,"Error")
		Exit All
    End If

    xMin = Evaluate(dlg.xMin)
    xMax = Evaluate(dlg.xMax)
	xStep = Evaluate(dlg.xStep)

    fMin = Evaluate(dlg.fMin)
    fMax = Evaluate(dlg.fMax)
    fStep = Evaluate(dlg.fStep)
    fSamples = dlg.sampleList

	'============first check the validity of frequency================
	If fMin - FarfieldFreq(0)<-1e-2 Or fMax - FarfieldFreq(m-1)>1e-2 Then
		MsgBox("Input frequency is out of Range",1,"Warning")
		Exit All
	End If
	'============Normalize the frequency samples======
	Dim freqSamples() As Double, freqStrArray() As String

	If fSamples <> "0;0" Then
		freqStrArray = Split(fSamples, ";")
		ReDim freqSamples(UBound(freqStrArray))
		For i=0 To UBound(freqStrArray)
			freqSamples(i) = CDbl(freqStrArray(i))
		Next
	Else
		ReDim freqSamples(Round((fMax-fMin)/fStep))
		For i=0 To Round((fMax-fMin)/fStep)
			freqSamples(i) = fMin+i*fStep
		Next
	End If
	'==============Double check the validity of frequency samples======
	Dim index As Integer
	Dim freqFinalSamples() As Double, nn As Integer
	ReDim freqFinalSamples(UBound(FarfieldFreq))
	nn = 0
	For i = 0 To UBound(FarfieldFreq) STEP 1
		For index=0 To UBound(freqSamples)
			If Abs(FarfieldFreq(i)-freqSamples(index))<1e-10 Then
				freqFinalSamples(nn) = freqSamples(index)
				nn = nn+1
				Exit For
			End If
		Next
    Next i
    ReDim Preserve freqFinalSamples(nn-1)


	Dim n As Integer
	'Dim rotateAngle As Double
	Dim projectPath As String
	Dim logFile As String
	'Dim cutPlaneValue As Integer, cutAngle As Double
	'Dim fieldComponentValue As Integer
	'Dim farfieldComponentValue As String

	cutPlaneValue = dlg.Group1
	cutAngle = Evaluate(dlg.Angle)

	'farfieldComponentValue = 0 for directivity, 1 for gain, 2 for realized gain
	farfieldComponentValue = dlg.Group2

	projectPath = GetProjectPath("Project")
	startTime = Replace(CStr(Time),":","_")

	logFile = projectPath + "\"+componentNames(farfieldComponentValue)+" sweep log_"+startTime+".txt"

   	Open logFile For Output As #2
   	Print #2, "##########Sweep of " + parameter + " begins at " + CStr(Now) +"'###########."
   	Print #2, " "

	FarfieldPlot.SetPlotMode(componentNames(farfieldComponentValue))

	If FarfieldPlot.IsScaleLinear = True Then
		FarfieldPlot.SetScaleLinear(False)
   	End If

   	threads = Evaluate(dlg.paraThreads)

	Dim sequenceName As String
	'ReDim x(threads)
	n = 0
	xSim = xMin

	sequenceName = "Single_para_sweep " & parameter
	'Solver.RemoteCalculation "True"
	Solver.StoreTDResultsInCache  "True"
	AddToHistory("Store results in cache", _
	"Solver.StoreTDResultsInCache ""True""")
	'Setting parameter sweep--------------------------------------------
    While xSim <= xMax
		ParameterSweep.DeleteAllSequences
		ParameterSweep.AddSequence(sequenceName)
		MaxNumberOfDistributedComputingParameters(threads)
		ParameterSweep.UseDistributedComputing(True)
		If xSim + (threads-1)*xStep <= xMax Then
			ParameterSweep.AddParameter_Samples(sequenceName, parameter, xSim, xSim + (threads-1)*xStep, threads, False)
			xSim = xSim + threads*xStep
		Else
			If Abs(xSim-xMax)>1e-10 Then
				ParameterSweep.AddParameter_StepWidth(sequenceName, parameter, xSim, xMax, xStep)
			Else
				ParameterSweep.AddParameter_Samples(sequenceName, parameter, xSim, xSim, 1, False)
			'	runWithParameter(parameter,xSim)
			End If
			xSim = xMax + xStep
		End If
		ReportInformationToWindow "###Parameter sweep with DC computing is in progress......"
		ParameterSweep.Start()
		retrieveSimulationResultsFromCache(freqFinalSamples)
	Wend
End Sub
'Sub runWithParameter(para As String, value As Double)
'
' 	StoreParameter(para,value)
'	Rebuild
'	Solver.MeshAdaption(False)
'	Solver.SteadyStateLimit(-40)
'	Solver.Start
'	ReportInformationToWindow "###Finish simulation loop with " +para+"="+Cstr(Round(value,3))
'End Sub

Sub retrieveSimulationResultsFromCache(f() As Double)
	Dim app As Object
	Dim mws As Object
	Dim pValue As Double

	Dim prjPath As String, cachePath As String
	Dim subDir As String, subDirs() As String
	Dim n As Integer, i As Integer, nDir As Integer
	ReDim subDirs(1000)

	prjPath = GetProjectPath("Result")
	cachePath = prjPath & "Cache\"
	subDir = Dir(cachePath, vbDirectory)
	nDir=0 '
	Do While subDir <> ""
        ' exclude . and ..
        If subDir <> "." And subDir <> ".." Then
            ' determine whether it is a dir
            If (GetAttr(cachePath & subDir) And vbDirectory) = vbDirectory Then
                ' print the name of the subdir
				subDirs(nDir) = cachePath & subDir
                'Debug.Print cachePath & subDir
                nDir = nDir+1
            End If
        End If
        ' get a next subdir
        subDir = Dir
    Loop
    If nDir >= 1 Then
    	ReDim Preserve subDirs(nDir-1)
    Else
    	MsgBox "No Cache results found, please check",vbCritical
    	Exit All
	End If
	Dim fileName As String
	fileName = BaseName(getProjectPath("Project")) & ".cst"
	Set app = CreateObject ("CSTStudio.Application")

	For i=0 To nDir-1
		Set mws = app.OpenFile(subDirs(i) & "\" & fileName)
		'mws.SelectTreeItem("1D Results\S-Parameters") 'just for test
    	pValue = mws.RestoreParameter(parameter)
    	'For n=0 To UBound(f)
		Copy1DFarfieldResult(mws, pValue, f)

    	'Next
	Next

	'delete all subdirectories under .\Result\Cache\
	Dim fso As Object
    Dim folderPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    For i=0 To nDir-1
		folderPath = subDirs(i)
		If fso.FolderExists(folderPath) Then
        	On Error Resume Next
        	fso.DeleteFolder(folderPath, True)

        	' Check if deleting succesfully
        	If Err.Number <> 0 Then
           		MsgBox "Deletion failed: " & Err.Description, vbExclamation
       	 	Else
            	ReportInformationToWindow "Cache result delete successfully!"
        	End If
        	On Error GoTo 0
    	Else
        	MsgBox "Directory does not exist!", vbExclamation
    	End If
    Next
    Set fso = Nothing ' release the object
	DeleteResults
End Sub


Function Copy1DFarfieldResult(ByRef mws As Object, xSim As Double, freq() As Double)

	'parameters: cutPlaneValue denotes the cutting plane,0->theta, 1->phi; cutAngle denotes the angle in the plane specified by cutPlaneValue; theta0 and phi0
	'denote the angle where the directivity is estimated; xSim is the rotate angle of the ham; freq is the operation frequency we care about
		Dim selectedItem As String, frequencyStr As String
		Dim nn As Integer, i As Integer
		For nn=0 To UBound(freq)
			'portStr = "1"
			frequencyStr = CStr(freq(nn))
			'ReportInformationToWindow "###Copying farfield results to 1D when parameter="+CStr(Round(xSim, 3))+" and frequency="+frequencyStr+"GHz"
			'mws.SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")

			mws.FarfieldPlot.Reset
			mws.FarfieldPlot.PlotType("polar")

			If mws.FarfieldPlot.IsScaleLinear = True Then
				mws.FarfieldPlot.SetScaleLinear(False)
		   	End If

			'If cutPlaneValue = 0 Then

			'	mws.FarfieldPlot.Vary("angle2")
			'	mws.FarfieldPlot.Theta(cutAngle)

			'Else

			'	mws.FarfieldPlot.Vary("angle1")
			'	mws.FarfieldPlot.Phi(cutAngle)

			'End If

			mws.FarfieldPlot.SetAxesType("currentwcs")
			mws.FarfieldPlot.SetAntennaType("unknown")

			mws.FarfieldPlot.SetPlotMode(componentNames(farfieldComponentValue))

			mws.FarfieldPlot.SetCoordinateSystemType("ludwig3")
			mws.FarfieldPlot.SetAutomaticCoordinateSystem("True")
			mws.FarfieldPlot.SetPolarizationType("Circular")
			mws.FarfieldPlot.StoreSettings
			mws.FarfieldPlot.Plot

			'FarfieldPlot.Plot
			'Dim dirName As String
			'dirName = "CP "+componentNames(farfieldComponentValue)+"@port="+portStr+"\"+parameter+"="+CStr(xSim)+ "@"+frequencyStr+"GHz"


			'Dim ChildItem As String
			'If mws.ResultTree.DoesTreeItemExist("1D Results\"+dirName) Then
			'	ChildItem = mws.ResultTree.GetFirstChildName("1D Results\"+dirName)
			'	While ChildItem <> ""
			'		With mws.ResultTree
			'			.Name ChildItem
			'			.Delete
			'		End With
			'		ChildItem = mws.ResultTree.GetFirstChildName("1D Results\"+dirName)
			'	Wend

			'End If

			'Dim resFile As String, resItem As String, resMainFile As String
			'Dim temp As Object
			'Dim circularComponents() As String
			'ReDim circularComponents(2)
			'circularComponents(0)="Abs"
			'circularComponents(1)="Right"
			'circularComponents(2)="Left"

			'For i=0 To UBound(circularComponents)
			'	mws.FarfieldPlot.SelectComponent(circularComponents(i))
			'	mws.FarfieldPlot.Plot
			'	mws.FarfieldPlot.CopyFarfieldTo1DResults(dirName, "farfield (f="+frequencyStr+")["+portStr+"]_"+circularComponents(i))
			'	resItem = "1D Results\"+dirName+"\"+"farfield (f="+frequencyStr+")["+portStr+"]_"+circularComponents(i)
			'	resFile = mws.Resulttree.GetFileFromTreeItem(resItem)
			'	resMainFile = Left(resFile, InStr(resFile, "\Result")-1) & Right(resFile, Len(resFile)-InStrRev(resFile, "\Result")+1)
			'	Set temp = Result1D(resFile)
			'	temp.save(resMainFile)
			'	temp.AddToTree(resItem)

			'Next
			'Set temp = Nothing
			'==================================================================================
			'---------------------------------------------------------------------------------------------------

			If mws.FarfieldPlot.GetSystemTotalEfficiency > -100 Then

				TE = mws.FarfieldPlot.GetSystemTotalEfficiency
				RE = mws.FarfieldPlot.GetSystemRadiationEfficiency
			Else
				TE = mws.FarfieldPlot.GetTotalEfficiency
				RE = mws.FarfieldPlot.GetRadiationEfficiency
			End If

			savefarfieldComponent(mws, xSim, freq(nn))
	        'CurrentItem = FirstChildItem
	        Copy1DFarfieldResult = getfarfieldComponent(mws)
	    Next
	    mws.save()
		mws.quit()
		Wait 5

		'Next
        'Plot1D.SetFont("Tahoma","bold","16")
End Function

Function getfarfieldComponent(ByRef mws As Object)

	Dim farfieldComponent As String
	Dim directivity As Double

	farfieldComponent = "ludwig3 circular right abs"
	directivity = mws.FarfieldPlot.CalculatePoint(0, 0, farfieldComponent, "")
	getfarfieldComponent = directivity

End Function
Sub savefarfieldComponent(ByRef mws As Object, xSim As Double,frequency As Double)
	ReportInformationToWindow "###Saving farfield components to xlsx file when parameter="+CStr(Round(xSim, 3))+" and frequency="+CStr(Round(frequency, 3)) +"GHz......"
    Dim selectedItem As String
    Dim n As Integer
    Dim frequencyStr As String

	'portStr = "1"
	frequencyStr = CStr(frequency)

	mws.SelectTreeItem("Farfields\farfield (f="+frequencyStr+") [" +portStr+"]")
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

             mws.FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", frequency)

         Next Theta

    Next Phi

    mws.FarfieldPlot.CalculateList("")

    'UHPower = FarfieldPlot.GetList("Spherical  abs")

    upperHemisphereRHCPdirectivity = mws.FarfieldPlot.GetList("Spherical circular right abs")

    upperHemisphereLHCPdirectivity = mws.FarfieldPlot.GetList("Spherical circular left abs")

    position_theta = mws.FarfieldPlot.GetList("Point_T")

    position_phi = mws.FarfieldPlot.GetList("Point_P")

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
	 'put the xlsx file into the main project directory
	projectPath = GetProjectPath("Project")

	xlsxFile = projectPath & "\CP " & componentNames(farfieldComponentValue) & " f=" _
	& frequencyStr & " P=" & portStr & " " & parameter & "=" & CStr(xSim) & " @" & Replace(CStr(Time), ":", "-") & ".xlsx"

	Dim NoticeInformation As String
	NoticeInformation = "The " & componentNames(farfieldComponentValue) & " data for frequency="+frequencyStr+"GHz is under£¨" & projectPath & "\£©"
	ReportInformationToWindow(NoticeInformation)

	Dim O As Object
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

	Columns = "BCDEFGHIJKLMN"
	'Add a sheet and rename it

	wBook.Sheets.Add.Name = "CP farfield data"
	wBook.Sheets.Add.Name = "S-para and Efficiencies"

	'wBook.Save
	'O.ActiveWorkbook.Close
	'O.quit

	Dim farfieldSheet As Object
	Set farfieldSheet = wBook.Sheets("CP farfield data")

	Dim sSheet As Object
	Set sSheet = wBook.Sheets("S-para and Efficiencies")

	'write rhcp directivity
	farfieldSheet.Range("A1").value = "Polarization"
	farfieldSheet.Range("B1").value = "RHCP"
	farfieldSheet.Range("C1").value = "Frequency"
	farfieldSheet.Range("D1").value = frequencyStr+"GHz"
	farfieldSheet.Range("E1").value = "Port"
	farfieldSheet.Range("F1").value = portStr
	farfieldSheet.Range("A2").value = "Phi\Theta"

	For n = 0 To Len(Columns)-1
		farfieldSheet.Range(Mid(Columns,n+1,1)+"2").value = n*15
		farfieldSheet.Range("A"+Cstr(n+3)) = n*30
	Next

	Dim i As Integer, j As Integer

	For i  = 0 To 12
		For j = 0 To 12
			farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).value = Round(rhcpDirectivity(i,j),2)
		Next
	Next
	'write lhcp directivity
	farfieldSheet.Range("A18").value = "Polarization"
	farfieldSheet.Range("B18").value = "LHCP"
	farfieldSheet.Range("C18").value = "Frequency"
	farfieldSheet.Range("D18").value = frequencyStr+"GHz"
	farfieldSheet.Range("E18").value = "Port"
	farfieldSheet.Range("F18").value = portStr
	farfieldSheet.Range("A19").value = "Phi\Theta"

	For n = 0 To Len(Columns)-1
		farfieldSheet.Range(Mid(Columns,n+1,1)+"19").value = n*15
		farfieldSheet.Range("A"+Cstr(n+20)) = n*30
	Next

	For i  = 0 To 12
		For j = 0 To 12
			farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).value = Round(lhcpDirectivity(i,j),2)
		Next
	Next

	Dim sheet As Object

	For Each sheet In wBook.Sheets
	    If sheet.Name Like "Sheet*" Then
	        sheet.Delete
	    End If
	Next
	'process sheet data, axial ratio, coloring, scoring and so on
	processDirectivityData(farfieldSheet, Columns)
	process1DResults(sSheet)
	Dim xlsxFileNewName As String, pos As Integer
	pos = InStr(xlsxFile, "@")
	xlsxFileNewName = Left(xlsxFile, pos-1) & "Sc=" & CStr(Round(farfieldSheet.Range("Q8").value)) & Right(xlsxFile, Len(xlsxFile)-pos+1)&".xlsx"
	wBook.SaveAS(xlsxFileNewName)
	O.ActiveWorkbook.Close
	O.quit
	Kill xlsxFile
	ReportInformationToWindow "###Finish saving farfield components to xlsx file when parameter="+CStr(Round(xSim, 3))+" and frequency="+CStr(Round(frequency, 3)) +"GHz"
End Sub

Sub processDirectivityData(sheet As Object, Columns As String)
	ReportInformationToWindow "%%%Processing farfield data in the xlsx file......"
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
	ReportInformationToWindow "%%%Processing is done......"
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

Sub process1DResults(sSheet As Object)

	Dim fileName As String, nPoints As Integer, n As Integer
	sSheet.Range("A1").Value = "Freq_" & portStr
	sSheet.Range("B1").Value = "S11/dB"
	sSheet.Range("C1").Value = "Rad_eff"
	sSheet.Range("D1").Value = "Tot_eff"

	'System efficiencies or efficiencies?
	fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\System Rad. Efficiency [" & portStr & "]")
	If fileName = "" Then
		fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Rad. Efficiency [" & portStr & "]")
	End If
	Dim realPart As Object
	With Result1DComplex(fileName) 'load data
		Set realPart = .real()
		nPoints = .GetN 'get number of points

		For n = 0 To nPoints-1

		'read all points, index of first point is zero.

			sSheet.Range("A" & CStr(n+2)).Value = .GetX(n)

			sSheet.Range("C" & CStr(n+2)).Value = Round(10*CST_log10(realPart.GetY(n)),2)

		Next n

	End With

	fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\System Tot. Efficiency [" & portStr & "]")
	If fileName = "" Then
		fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Tot. Efficiency [" & portStr & "]")
	End If
	'Dim reaPart As Object
	With Result1DComplex(fileName) 'load data
		Set realPart = .real()
		'nPoints = .GetN 'get number of points

		For n = 0 To nPoints-1

		'read all points, index of first point is zero.

			'sSheet.Range("A" & CStr(n+2)).Value = .GetX(n)

			sSheet.Range("D" & CStr(n+2)).Value = Round(10*CST_log10(realPart.GetY(n)),2)

		Next n

	End With

	fileName = ResultTree.GetFileFromTreeItem("1D Results\S-Parameters\S" & portStr & ","& portStr)
	Dim m As Integer, x As Double, y As Double
	With Result1DComplex(fileName) 'load data
		Set realPart = .real()
		'nPoints = .GetN 'get number of points

		For n = 0 To nPoints-1
			m = realPart.GetClosestIndexFromX(sSheet.Range("A" & CStr(n+2)).Value)
		'read all points, index of first point is zero.

			'C.Value = .GetX(n)

			'sSheet.Range("B" & CStr(n+2)).Value = 10*CST_log10(realPart.GetY(m))
			'sSheet.Range("B" & CStr(n+2)).Value = realPart.GetY(m)
			x = .GetYRe(m)
			y = .GetYIm(m)
			sSheet.Range("B" & CStr(n+2)).Value = Round(20*CST_Log10(Sqr((x^2+y^2))),2)

		Next n

	End With
	With sSheet
	    '.Columns("P").ColumnWidth = 15
	    '.Columns("Q").ColumnWidth = 33
	    .Range("A1").Interior.Color = RGB(221, 235, 247)
	    .Range("B1").Interior.Color = RGB(0, 176, 240)
	    .Range("C1:D1").Interior.Color = RGB(255, 217, 102)
	    .Range("A1:D100").Font.Bold = True
	    '.Range("Q3:Q8").Font.Color = RGB(255, 0, 0)
	End With
End Sub

