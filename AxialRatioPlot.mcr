' Axial ratio Vs Frequency plot, Farfied monitors with strictly increasing frequencies are necessary
'frequency step is not available by now
'Axial ratio is in decible
'2022-05-25
'Option Explicit

Public FarfieldFreq() As Double, m As Integer
'#include "vba_globals_all.lib"

Sub Main ()
    'Get frequencies of farfield monitors
    'Dim FarfieldFreq() As Double
    Dim NumOfMonitor As Integer, i As Integer
    Dim HasFarfiledMonitor As Boolean
    Dim TempStr As String

    NumOfMonitor = Monitor.GetNumberOfMonitors
    ReDim FarfieldFreq(NumOfMonitor-1)

    If NumOfMonitor < 1 Then
         MsgBox("No monitors found!",vbCritical,"Warning")
    	Exit Sub
    End If

    HasFarfieldMonitor  = False

    m = 0

    For i = 0 To NumOfMonitor-1 STEP 1
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

    Dim InformationStr As String
    InformationStr = "请输入要计算的频率范围（"+CStr(FarfieldFreq(0))+"GHz到"+CStr(FarfieldFreq(m-1))+"GHz之间）:"
	Begin Dialog UserDialog 410,231,"轴比-频率作图",.DialogFunction ' %GRID:10,7,1,1
		Text 20,7,390,14,InformationStr,.Text1
		TextBox 40,35,35,14,.Fmin
		TextBox 130,35,35,14,.Fmax
		OKButton 60,203,90,21
		CancelButton 190,203,90,21
		Text 20,84,40,14,"步长:",.Text2
		TextBox 60,84,50,14,.Fstep
		Text 20,112,220,14,"请输入要计算的远场方向：",.Text3
		Text 20,133,30,14,"θ1=",.Text4
		Text 20,154,30,14,"θ2=",.Text11
		TextBox 60,133,40,14,.Theta1
		TextBox 60,154,40,14,.Theta2
		Text 110,133,30,14,"φ1=",.Text5
		Text 110,154,30,14,"φ2=",.Text12
		TextBox 150,133,40,14,.Phi1
		TextBox 150,154,40,14,.Phi2
		Text 20,35,20,14,"从",.Text6
		Text 120,84,40,14,"GHz",.Text7
		Text 80,35,50,14,"GHz 到",.Text8
		Text 170,35,30,14,"GHz",.Text9
		Text 20,63,110,14,"请选择端口：",.Text10
		TextBox 110,63,40,14,.PortNum
	End Dialog
	Dim dlg As UserDialog
	dlg.Fmin = CStr(FarfieldFreq(0))
	dlg.Fmax = CStr(FarfieldFreq(m-1))
	dlg.Fstep = CStr(FarfieldFreq(1)-FarfieldFreq(0))
	dlg.Theta1 = "45"
	dlg.Phi1 = "0"
	dlg.Theta2 = "45"
	dlg.Phi2 = "90"
	dlg.PortNum = "1"


	If Dialog(dlg,1) = 0 Then
		Exit All
	End If

End Sub

Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Dim parameterFile As String
   	Dim prjPath As String

	Select Case Action
	Case 1 ' Dialog box initialization
		prjPath = GetProjectPath("Project")
   		parameterFile = prjPath + "\dialog_parameter.txt"
   		'parameterFile = file = "D:\Simulation\SXW\Research\Basics\a.txt"
		ReStoreAllDialogSettings_LIB(parameterFile)
	Case 2 ' Value changing or button pressed
		'DialogFunction = True ' Prevent button press from closing the dialog box

		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
			'DialogFunction = True
			prjPath = GetProjectPath("Project")
   			parameterFile = prjPath + "\dialog_parameter.txt"
   			'parameterFile = "D:\Simulation\SXW\Research\Basics\a.txt"
			StoreAllDialogSettings_LIB(parameterFile)
			Dim FreqMin As Double, FreqMax As Double, FreqStep As Double
		    Dim Theta1 As Integer,Theta2 As Integer
		    Dim Phi1 As Integer, Phi2 As Integer
		    Dim PortNum As String

		    FreqMin = Evaluate(DlgText("Fmin"))
		    FreqMax = Evaluate(DlgText("Fmax"))
		    FreqStep = Evaluate(DlgText("Fstep"))

		    Theta1 = Evaluate(DlgText("Theta1"))
		    Phi1 = Evaluate(DlgText("Phi1"))
		    Theta2 = Evaluate(DlgText("Theta2"))
		    Phi2 = Evaluate(DlgText("Phi2"))
		    PortNum = DlgText("PortNum")




		    If FreqMin < FarfieldFreq(0) Or FreqMax > FarfieldFreq(m-1) Then
		    	MsgBox("输入的频率超出可计算范围，请重新输入",1,"警告")
		    	Exit All
		    End If


		    Dim AxialRatio1() As Double, AxialRatio2() As Double

		    ReDim AxialRatio1(m-1)
		    ReDim AxialRatio2(m-1)


		    Dim o As Object

		    'For test

		   ' m = 1

		    'FarfieldPlot.Reset

		    FarfieldPlot.SetScaleLinear(False)

		    Dim Nstep As Integer
		    Dim PlotFreq() As Double
		    ReDim PlotFreq(m-1) As Double

		    Nstep = 0

		    For i = 0 To m-1 STEP 1

		    	If FarfieldFreq(i) >= FreqMin And FarfieldFreq(i) <= FreqMax Then
		    		Dim FarfieldName As String
		    		FarfieldName = "farfield (f="+ CStr(FarfieldFreq(i))+") ["+PortNum+"]"
			    	If SelectTreeItem("Farfields\"+ FarfieldName) <> False Then
			    		'Dim MidValue As Double

			    		'MidValue = FarfieldFreq(i)

			    		'FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", FarfieldFreq(i))
			    		AxialRatio1(Nstep) = FarfieldPlot.CalculatePoint(Theta1,Phi1,"Spherical  circular axialratio",FarfieldName)
			    		AxialRatio2(Nstep) = FarfieldPlot.CalculatePoint(Theta2,Phi2,"Spherical  circular axialratio",FarfieldName)
			    		PlotFreq(Nstep) = FarfieldFreq(i)
			    		Nstep = Nstep+1
			    	Else
			    		MsgBox("The selected port does not exist!",1,"warning")
		            	Exit All

			    	End If
			    	End If

		    Next i



		    '-------------------------For test

		    '------------------------


		    Set o1 = Result1D("")
		    Set o2 = Result1D("")


		    For i = 0 To Nstep-1 STEP 1

		    	o1.AppendXY(PlotFreq(i),AxialRatio1(i))
		    	o2.AppendXY(PlotFreq(i),AxialRatio2(i))

		    Next i

		    o1.ylabel("AR/dB")

		    o2.ylabel("AR/dB")

		    o1.xlabel("Frequency/GHz")

		    o2.xlabel("Frequency/GHz")

			o1.Save("AxialRatio@Port="+PortNum+"_Theta="+CStr(Theta1)+"_Phi="+CStr(Phi1)+".sig")

			o2.Save("AxialRatio@Port="+PortNum+"_Theta="+CStr(Theta2)+"_Phi="+CStr(Phi2)+".sig")

			'o.AddToTree("1D Results\AxialRatio\AR_Freq("+CStr(FreqMin)+"GHz_to_"+CStr(FreqMax)+"GHz")_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg")
			'Dim ResultItem As String
			'ResultItem = "1D Results\AxialRatio\AR_Freq_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg"
			Dim sName1 As String, sName2 As String
			sName1 = "1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta1)+"deg_φ="+CStr(Phi1)+"deg"
			sName2 = "1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta2)+"deg_φ="+CStr(Phi2)+"deg"
			AddPlotToTree_LIB(o1, sName1, True)
			AddPlotToTree_LIB(o2, sName2, True)
			'o1.AddToTree("1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta1)+"deg_φ="+CStr(Phi1)+"deg")
			'o2.AddToTree("1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta2)+"deg_φ="+CStr(Phi2)+"deg")

			'SelectTreeItem("1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta1)+"deg_φ="+CStr(Phi1)+"deg")

		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function

