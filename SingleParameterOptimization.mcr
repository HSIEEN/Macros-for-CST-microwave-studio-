'sweep single parameter to meet the target
'2023-01-10 By Shawn
Const macrofile = "C:\Users\Simulation-1\AppData\Roaming\Dassault Systemes\CST STUDIO SUITE\Library\Macros\Coros\PostProcess\StoreResonantFrequencyAndQ.mcr"
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

	Begin Dialog UserDialog 340,287,"Single parameter optimization" ' %GRID:10,7,1,1
		TextBox 180,231,40,14,.theta2
		TextBox 260,210,40,14,.phi1
		TextBox 260,231,40,14,.phi2
		Text 20,14,220,14,"Please select a parameter:",.Text1
		ListBox 20,35,240,56,parameterArray(),.parameterList
		Text 20,91,240,14,"Please set the sweep settings:",.Text21
		Text 20,105,90,14,"Range from",.Text2
		TextBox 100,105,40,14,.xMin
		TextBox 200,126,40,14,.nSim
		Text 150,105,20,14,"to",.Text4
		TextBox 170,105,40,14,.xMax
		Text 20,147,180,14,"Setting goals:",.Text6
		Text 20,161,260,14,"========Resonance frequency=======",.Text3
		Text 30,175,30,14,"f1:",.Text9
		TextBox 60,175,40,14,.f1
		Text 110,175,30,14,"f2:",.Text10
		TextBox 140,175,40,14,.f2
		Text 190,175,40,14,"f3:",.Text11
		TextBox 220,175,40,14,.f3
		Text 20,196,320,14,"===============Axial Ratio===============",.Text8
		Text 30,210,50,14,"AR1",.Text15
		TextBox 70,210,40,14,.AR1
		Text 30,231,60,14,"AR2",.Text16
		TextBox 70,231,40,14,.AR2
		TextBox 180,210,40,14,.theta1
		Text 110,210,70,14,"dB at ¦È1=",.Text17
		Text 230,210,30,14,"¦Õ1=",.Text19
		Text 230,231,30,14,"¦Õ2=",.Text20
		Text 110,231,70,14,"dB at ¦È2=",.Text18
		OKButton 70,259,90,21
		CancelButton 180,259,90,21
		Text 20,126,180,14,"Max.Number of evaluations:",.Text23
	End Dialog
	Dim dlg As UserDialog
	dlg.xMin = "0"
	dlg.xMax = "0"
	'dlg.Mono = "0"
	dlg.f1 = "0"
	dlg.f2 = "0"
	dlg.f3 = "0"
	'dlg.Q1 = "0"
	'dlg.Q2 = "0"
	'dlg.Q3 = "0"
	dlg.AR1 = "0"
	dlg.AR2 = "0"
	dlg.theta1 = "0"
	dlg.theta2 = "0"
	dlg.phi1 = "0"
	dlg.phi2 = "0"
	dlg.nSim = "10"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If


	Dim parameter As String
	Dim xMin As Double, xMax As Double, Mono As Boolean
	Dim xSim As Double
	Dim fo As Double, fcur As Double, fmin As Double, fmax As Double
	Dim sfmin As Double, sfmax As Double
	Dim nSim As Integer

	sfmin = Solver.GetFmin
	sfmax = Solver.GetFmax

	parameter = parameterArray(dlg.parameterList)
	If DoesParameterExist(parameter) = False Then
		MsgBox("Input paramter does not exist!!",vbCritical,"Error")
		Exit All
    End If
    xMin = Evaluate(dlg.xMin)
    xMax = Evaluate(dlg.xMax)
    'Mono = Evaluate(dlg.Mono)
    'xSim = xMin
    fo = Evaluate(dlg.f1)
    If fo < sfmin Or fo > sfmax Then
    	MsgBox("Target frequency is out of sover's frequency range!!",vbCritical,"Error")
    	Exit All
    End If
    nSim = Evaluate(dlg.nSim)

    Dim prjPath As String
	Dim dataFile As String
	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\optlog.txt"
   	Open dataFile For Output As #2
   	Print #2, "Optimizing of " + parameter + " is in progress..."
   	Print #2, "% Calculating the target function when "+parameter+"="+Cstr(xMin)
	'run with specified value of xMin
	runWithParameter(parameter,xMin)
	'store resonant frequency and Q value
	MacroRun(macrofile)
	'get fmin
	fmin = GetFrequencyValue()(0)
	Print #2, "% Value of target function is "+ Cstr(fmin)

	'run with specified value of xMax
	Print #2, "% Calculating the target function when "+parameter+"="+Cstr(xMax)
	runWithParameter(parameter,xMax)
	'store resonant frequency and Q value
	MacroRun(macrofile)
	'get fmax
	fmax = GetFrequencyValue()(0)
	Print #2, "% Value of target function is "+ Cstr(fmax)
	Print #2, "% Running iteration..."
	If fo < fmin Or fo > fmax Then
			MsgBox("Target frequency is may Not in search range !!",vbCritical,"Error")
			Exit All
	End If
	Dim n As Integer
	n = 1
    Do
    	Print #2, "% Iteration turn "+ Cstr(n) + ":"
		xSim = xMin+(fo-fmin)/(fmax-fmin)*(xMax-xMin)
		Print #2, "% Attempting to calculate the target function when "+parameter+"="+Cstr(Round(xSim,2))



		Debug.Print CStr(xSim)
		Debug.Print CStr(n)

    	Dim SimNotice As String
    	SimNotice = "Time domain simulation is ongoing with "+parameter+"="+Cstr(xSim)+"..."
    	'MsgBox(SimNotice,vbOkOnly,"Information")


		'run with specified value of xSim
		runWithParameter(parameter,xSim)
		'store resonant frequency and Q value
		MacroRun(macrofile)
		'get fcur
		fcur = GetFrequencyValue()(0)
		Print #2, "% Value of target function is "+ Cstr(fcur)

		Dim isMet As Boolean
		isMet = False
		If Abs(fcur-fo)/fo < 0.01 Then
			Print #2, "% Target is Met when " + parameter + "=" + CStr(Round(xSim,2)) + " and the target function value is "+ Cstr(fcur)
			isMet = True
			Exit Do
		ElseIf fcur<fo Then
			Print #2, "% Target is not met yet... "

			fmin = fcur
			xMin = xSim
		ElseIf fcur>fo Then
			Print #2, "% Target is not met yet... "

			fmax = fcur
			xMax = xSim
		End If

		n = n+1
    Loop Until nm > nSim
    If isMet = True Then
    	ReportInformationToWindow "Target is Met when " + parameter + "=" + CStr(Round(xSim,3)) + " and the resonant frequency is "+ Cstr(fcur)
    Else
    	ReportInformationToWindow "Sweep has been done and the target is not met."
    	Print #2, "% Max Iterations exceeded. Optimization did NOT converge" +vbNewLine+ "The current "+ parameter + "=" + CStr(Round(xSim,3)) + " and the target function value is  "+ Cstr(fcur)
    End If

    Debug.Clear
	Close #2

End Sub
Sub runWithParameter(para As String, value As Double)
 	StoreParameter(para,value)
	Rebuild
	Solver.MeshAdaption(False)
	Solver.SteadyStateLimit(-40)
	Solver.Start
End Sub
Function GetFrequencyValue() As Variant
    Dim prjPath As String
	Dim dataFile As String
	Dim freq() As Double
	Dim resStr As String
	Dim n As Integer
	Dim lineRead As String
	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\freq_Q.txt"
   	n = 1
	Open dataFile For Input As #1
	While Not EOF(1)
		Line Input #1, lineRead
		While Not EOF(1) And Left(lineRead,1)<>"F"
			Line Input #1, lineRead
		Wend

		If Left(lineRead,1) = "F" Then
			ReDim Preserve freq(n)
			resStr = Split(lineRead, "=")(1)
			freq(n-1) = CDbl(resStr) 'realized resonance frequency

		End If
		n = n+1
	Wend
	GetFrequencyValue = freq
	Close #1
End Function





