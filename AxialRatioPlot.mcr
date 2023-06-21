' Axial ratio Vs Frequency plot, Farfied monitors with strictly increasing frequencies are necessary
'frequency step is not available by now
'Axial ratio is in decible
'2022-05-25
Option Explicit

Public FarfieldFreq() As Double, m As Integer
Public portArray(100) As String
Public portNum As String
'#include "vba_globals_all.lib"

Sub Main ()
    'Get frequencies of farfield monitors
    'Dim FarfieldFreq() As Double
    Dim NumOfMonitor As Integer, i As Integer
    Dim HasFarfieldMonitor As Boolean
    Dim TempStr As String
    Dim farfieldPath As String

	farfieldPath = "Farfields\"
    'Dim portArray(100) As String
    Dim ii As Integer
    Dim childItem As String
    Dim element As Variant
    Dim Port As String

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

    childItem = Resulttree.GetFirstChildName(farfieldPath)
	 ii = 1
	childItem = Resulttree.GetNextItemName(childItem)
    While childItem <> ""
		Port = Mid(childItem,InStr(childItem,"[")+1,InStr(childItem,"]")-InStr(childItem,"[")-1)
   		If ii > 1 Then
   			For Each element In portArray
   				If element = Port Then
   					Exit While
   				End If
   			Next
   		End If
   		portArray(ii-1) = Port
   		ii = ii+1
   		childItem = Resulttree.GetNextItemName(childItem)

    Wend

    Dim InformationStr As String
    InformationStr = "Input valueS should be within "+CStr(FarfieldFreq(0))+"GHz to "+CStr(FarfieldFreq(m-1))+"GHz:"
	Begin Dialog UserDialog 420,210,"Axial Raio_Frequency Plot",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 0,0,420,70,"Frequency settings:",.GroupBox1
		Text 10,21,390,14,InformationStr,.Text1
		TextBox 50,42,30,14,.Fmin
		TextBox 150,42,30,14,.Fmax
		OKButton 80,189,90,21
		CancelButton 190,189,90,21
		GroupBox 0,112,420,70,"Farfield Cuts settings",.GroupBox3
		TextBox 310,42,50,14,.Fstep
		Text 60,133,30,14,"θ1=",.Text4
		Text 60,161,30,14,"θ2=",.Text11
		TextBox 100,133,40,14,.Theta1
		TextBox 100,161,40,14,.Theta2
		Text 180,133,30,14,"φ1=",.Text5
		Text 180,161,30,14,"φ2=",.Text12
		TextBox 220,133,40,14,.Phi1
		TextBox 220,161,40,14,.Phi2
		Text 10,42,40,14,"From",.Text6
		Text 370,42,30,14,"GHz",.Text7
		GroupBox 0,70,420,42,"Port settings:",.GroupBox2
		Text 90,42,50,14,"GHz  to",.Text8
		Text 190,42,120,14,"GHz with step of ",.Text9
		DropListBox 150,84,90,14,portArray(),.port
	End Dialog
	Dim dlg As UserDialog

	dlg.Fmin = CStr(FarfieldFreq(0))
	dlg.Fmax = CStr(FarfieldFreq(m-1))
	dlg.Fstep = CStr(FarfieldFreq(1)-FarfieldFreq(0))
	dlg.Theta1 = "45"
	dlg.Phi1 = "0"
	dlg.Theta2 = "45"
	dlg.Phi2 = "90"
	'dlg.PortNum = "1"


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
		    Dim Theta1 As Integer,Theta2 As Integer
		    Dim Phi1 As Integer, Phi2 As Integer
		    Dim i As Integer

		    FreqMin = Evaluate(DlgText("Fmin"))
		    FreqMax = Evaluate(DlgText("Fmax"))
		    FreqStep = Evaluate(DlgText("Fstep"))

		    Theta1 = Evaluate(DlgText("Theta1"))
		    Phi1 = Evaluate(DlgText("Phi1"))
		    Theta2 = Evaluate(DlgText("Theta2"))
		    Phi2 = Evaluate(DlgText("Phi2"))

			portNum = DlgText("port")


		    If FreqMin < FarfieldFreq(0) Or FreqMax > FarfieldFreq(m-1) Then
		    	MsgBox("Input frequency is out of Range",1,"Warning")
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
		    		FarfieldName = "farfield (f="+ CStr(FarfieldFreq(i))+") ["+portNum+"]"
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
		    Dim o1 As Object, o2 As Object


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

			o1.Save("AxialRatio@Port="+portNum+"_Theta="+CStr(Theta1)+"_Phi="+CStr(Phi1)+".sig")

			o2.Save("AxialRatio@Port="+portNum+"_Theta="+CStr(Theta2)+"_Phi="+CStr(Phi2)+".sig")

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

