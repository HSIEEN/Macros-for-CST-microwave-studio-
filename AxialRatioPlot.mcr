' Axial ratio Vs Frequency plot, Farfied monitors with strictly increasing frequencies are necessary
'frequency step is not available by now
'Axial ratio is in decible
'2022-05-25
'Option Explicit

Public FarfieldFreq() As Double, m As Integer

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
	Begin Dialog UserDialog 410,182,"轴比-频率作图",.DialogFunction ' %GRID:10,7,1,1
		Text 20,7,390,14,InformationStr,.Text1
		TextBox 50,28,40,21,.Fmin
		TextBox 170,28,40,21,.Fmax
		OKButton 60,154,90,21
		CancelButton 190,154,90,21
		Text 20,84,40,14,"步长",.Text3
		TextBox 60,77,50,21,.Fstep
		Text 20,105,220,14,"请输入要计算的远场方向：",.Text4
		Text 20,133,20,14,"θ=",.Text5
		TextBox 50,126,40,21,.Theta
		Text 100,133,20,14,"φ=",.Text6
		TextBox 130,126,40,21,.Phi
		Text 20,35,20,14,"从",.Text2
		Text 120,84,40,14,"GHz",.Text8
		Text 100,35,50,14,"GHz 到",.Text9
		Text 220,35,30,14,"GHz",.Text10
		Text 20,56,110,14,"请选择端口：",.Text7
		TextBox 120,56,30,21,.PortNum
	End Dialog
	Dim dlg As UserDialog

	dlg.Fmin = CStr(FarfieldFreq(0))
	dlg.Fmax = CStr(FarfieldFreq(m-1))
	dlg.Fstep = CStr(FarfieldFreq(1)-FarfieldFreq(0))
	dlg.Theta = "60"
	dlg.Phi = "0"
	dlg.PortNum = "1"


	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

End Sub

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

			Dim FreqMin As Double, FreqMax As Double, FreqStep As Double
		    Dim Theta As Integer
		    Dim Phi As Integer
		    Dim PortNum As String

		    FreqMin = Evaluate(DlgText("Fmin"))
		    FreqMax = Evaluate(DlgText("Fmax"))
		    FreqStep = Evaluate(DlgText("Fstep"))

		    Theta = Evaluate(DlgText("Theta"))
		    Phi = Evaluate(DlgText("Phi"))
		    PortNum = DlgText("PortNum")



		    If FreqMin < FarfieldFreq(0) Or FreqMax > FarfieldFreq(m-1) Then
		    	MsgBox("输入的频率超出可计算范围，请重新输入",1,"警告")
		    	Exit All
		    End If

		    Dim AxialRatio() As Double

		    ReDim AxialRatio(m-1)


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
			    	SelectTreeItem("Farfields\"+ FarfieldName)
			    	'Dim MidValue As Double

			    	'MidValue = FarfieldFreq(i)

			    	'FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", FarfieldFreq(i))
			    	AxialRatio(Nstep) = FarfieldPlot.CalculatePoint(Theta,Phi,"Spherical  circular axialratio",FarfieldName)
			    	PlotFreq(Nstep) = FarfieldFreq(i)
			    	Nstep = Nstep+1
			    	End If

		    Next i



		    '-------------------------For test

		    '------------------------


		    Set o = Result1D("")


		    For i = 0 To Nstep-1 STEP 1

		    	o.AppendXY(PlotFreq(i),AxialRatio(i))

		    Next i

		    o.ylabel("AR/dB")

		    o.xlabel("Frequency/GHz")

			o.Save("AxialRatio@Port="+PortNum+"_Theta="+CStr(Theta)+"_Phi="+CStr(Phi)+".sig")

			'o.AddToTree("1D Results\AxialRatio\AR_Freq("+CStr(FreqMin)+"GHz_to_"+CStr(FreqMax)+"GHz")_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg")
			'Dim ResultItem As String
			'ResultItem = "1D Results\AxialRatio\AR_Freq_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg"
			o.AddToTree("1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg")

			SelectTreeItem("1D Results\AxialRatio\AR_Port["+PortNum+"]_θ="+CStr(Theta)+"deg_φ="+CStr(Phi)+"deg")

		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function

