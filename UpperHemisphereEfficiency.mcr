' Upper hemisphere efficiencies with RHCP and LHCP characteristics are supported
'2022-05-05 by Shawn Shi
'Option Explicit
'#include "vba_globals_all.lib"

Sub Main ()
    'Get current farfield plot mode
    Dim CurrentPlotMode As String

     If MsgBox("Please make sure the axis W in current WCS coordinated system point to the Zenith",vbYesNo,"Information") <> vbYes Then
		Exit Sub
     End If
    'MsgBoxTimer("Please make sure the axis W in current WCS coordinated system point to the Zenith",1,"Attention:",64)
    'MessageBoxTimeout(0, "Hello World", "Tips", vbOkCancel, 0, 10000)

    CurrentPlotMode = FarfieldPlot.GetPlotMode

    'Farfield plot settings
    FarfieldPlot.SetPlotMode("pfield")

    FarfieldPlot.Distance(1)

    FarfieldPlot.SetScaleLinear("True")

    FarfieldPlot.SetAxesType("currentwcs")

    FarfieldPlot.StoreSettings

    FarfieldPlot.Plot

    'FarField Calculation

    Dim SelectedItem As String

    Dim n As Integer

    Dim Frequency As Double, FrequencyStr As String

    Dim PortStr As String

    SelectedItem = GetSelectedTreeItem

    If (InStr(SelectedItem,"farfield (") = 0) Then

        MsgBox("Please select a farfield result before runing this macro.",vbCritical,"Warning")
        'MsgBoxTimer("Please select a farfield result before runing this macro.",1,"Warning:",16)

        Exit All

    Else
    	'Get Port number in string

    	PortStr = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)

        'Get the frequency of the selected item

        FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)

        Frequency = CDbl(FrequencyStr)

        FarfieldPlot.Reset
        '==============Upper Hemisphere TRP Calculation===============

        Dim UHPower() As Double, UHRHCPPower() As Double, UHLHCPPower() As Double

        Dim UHTRP As Double, UHRHCPTRP As Double, UHLHCPTRP As Double

         Dim Theta As Double, Phi As Double

        Dim position_theta() As Double, position_phi() As Double

        For Phi=0 To 360 STEP 5

             For Theta = 0 To 90 STEP 5

                 FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", Frequency)

             Next Theta

        Next Phi

        FarfieldPlot.CalculateList("")

        UHPower = FarfieldPlot.GetList("Spherical  abs")

        UHRHCPPower = FarfieldPlot.GetList("Spherical circular right abs")

        UHLHCPPower = FarfieldPlot.GetList("Spherical circular left abs")

        position_theta = FarfieldPlot.GetList("Point_T")

        position_phi = FarfieldPlot.GetList("Point_P")

        UHTRP = 0

        UHRHCPTRP = 0

        UHLHCPTRP = 0

        For n = 0 To UBound(UHPower)

            'Theta and phi step is 5deg in default

             If position_theta(n) = 0 Then

                 UHTRP = UHTRP + UHPower(n)*(1-cosD(2.5))*pi/36

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*(1-cosD(2.5))*pi/36


                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*(1-cosD(2.5))*pi/36


             ElseIf position_theta(n) = 90 Then

                 UHTRP = UHTRP + UHPower(n)*(CosD(87.5)-CosD(90))*pi/36

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*(CosD(87.5)-CosD(90))*pi/36

                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*(CosD(87.5)-CosD(90))*pi/36


            ElseIf (position_theta(n) <> 0 And position_theta(n) <> 90) Then

                 UHTRP = UHTRP + UHPower(n)*(CosD(position_theta(n)-2.5)-CosD(position_theta(n)+2.5))*pi/36

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*(CosD(position_theta(n)-2.5)-CosD(position_theta(n)+2.5))*pi/36

                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*(CosD(position_theta(n)-2.5)-CosD(position_theta(n)+2.5))*pi/36

            'Total = Total + Power_am(n)*pi/36*(sinD(position_theta(n))+sinD(position_theta(n-1)))/2*pi/36

            End If

        Next n

        Dim TRP As Double,StimPower As Double,AcceptPower As Double, RadEffi As Double, TotEffi As Double, SysRadEffi As Double, SysTotEffi As Double

        TRP = FarfieldPlot.GetTRP

        RadEffi = FarfieldPlot.GetRadiationEfficiency

        TotEffi = FarfieldPlot.GetTotalEfficiency

        SysRadEffi =  FarfieldPlot.GetSystemRadiationEfficiency

        SysTotEffi = FarfieldPlot.GetSystemTotalEfficiency
        'If system total/radiation efficiency is available, which means combined results may be available, use system total/radiation efficiency to calculate.

        '!!Attention: when aperture tuning happends, the calculated results are not accurate since the radiation pattern differs

        If SysTotEffi > -100 Then
        	StimPower = TRP/SysTotEffi
        	AcceptPower = TRP/SysRadEffi
        Else
        	StimPower = TRP/TotEffi
        	AcceptPower = TRP/RadEffi

        End If


        Dim UHTotEffi As Double, dBTotal As Double

        Dim UHRHCPTotEffi As Double, dBRight As Double

        Dim UHLHCPTotEffi As Double,dBLeft As Double

        'efficiency calculation, if combined results found, the efficiencies are in system type, instead of 3d simulation

        UHTotEffi = UHTRP/StimPower

        dBTotal = 10*CST_Log10(UHTotEffi)'Log(UHTotEffi)/Log(10)*10

        UHRHCPEffi = UHRHCPTRP/StimPower

        dBRight = 10*CST_Log10(UHRHCPEffi)'Log(UHRHCPEffi)/Log(10)*10

        UHLHCPEffi = UHLHCPTRP/StimPower

        dBLeft = 10*CST_Log10(UHLHCPEffi)'Log(UHLHCPEffi)/Log(10)*10


        'Print information to the message window

        ReportInformationToWindow( _
        "上半球总效率f="+FrequencyStr+"GHz@Port"+PortStr+": "+Left(Cstr(UHTotEffi*100),InStr(Cstr(UHTotEffi*100),".")+2)+"% ("+Left(Cstr(dBTotal),InStr(Cstr(dBTotal),".")+2)+ "dB)"+vbCrLf+ _
        "上半球右旋效率f="+FrequencyStr+"GHz@Port"+PortStr+": "+Left(Cstr(UHRHCPEffi*100),InStr(Cstr(UHRHCPEffi*100),".")+2)+"% ("+Left(Cstr(dBRight),InStr(Cstr(dBRight),".")+2)+ "dB)"+ vbCrLf+ _
        "上半球左旋效率f="+FrequencyStr+"GHz@Port"+PortStr+": "+Left(Cstr(UHLHCPEffi*100),InStr(Cstr(UHLHCPEffi*100),".")+2)+"% ("+Left(Cstr(dBLeft),InStr(Cstr(dBLeft),".")+2)+ "dB)")

    End If

    'resume last plot mode

    FarfieldPlot.Reset

    FarfieldPlot.SetScaleLinear("False")

    FarfieldPlot.SetPlotMode(CurrentPlotMode)

    FarfieldPlot.StoreSettings


End Sub

