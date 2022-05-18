' LowerHemisphereEfficiency 

Sub Main ()
	 'Get current farfield plot mode
    Dim CurrentPlotMode As String

    CurrentPlotMode = FarfieldPlot.GetPlotMode

    'Farfield plot settings
    FarfieldPlot.SetPlotMode("pfield")

    FarfieldPlot.Distance(1)

    FarfieldPlot.SetScaleLinear("True")

    FarfieldPlot.StoreSettings

    FarfieldPlot.Plot

    'FarField Calculation

    Dim SelectedItem As String

    Dim n As Integer

    Dim Frequency As Double, FrequencyStr As String

    SelectedItem = GetSelectedTreeItem

    If InStr(SelectedItem,"farfield") = 0 Then

        MsgBox("Please select a farfield result before runing this macro.",vbCritical,"Warning")

        Exit All

    Else

        'Get the frequency of calculation

        FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)

        Frequency = CDbl(FrequencyStr)

        FarfieldPlot.Reset
        '==============Upper Hemisphere TRP Calculation===============

        Dim UHPower() As Double, UHRHCPPower() As Double, UHLHCPPower() As Double

        Dim UHTRP As Double, UHRHCPTRP As Double, UHLHCPTRP As Double

         Dim Theta As Double, Phi As Double

        Dim position_theta() As Double, position_phi() As Double

        For Phi=0 To 360 STEP 5

             For Theta = 90 To 180 STEP 5

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

            'Theta step is 5 in default

             If position_theta(n) = 180 Then

                 UHTRP = UHTRP + UHPower(n)*pi/72*sinD(2.5)*pi/72

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*pi/72*sinD(2.5)*pi/72

                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*pi/72*sinD(2.5)*pi/72

             ElseIf position_theta(n) = 90 Then

                 UHTRP = UHTRP + UHPower(n)*pi/72*sinD(position_theta(n))*pi/36

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*pi/72*sinD(position_theta(n))*pi/36

                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*pi/72*sinD(position_theta(n))*pi/36


            ElseIf (position_theta(n) <> 180 And position_theta(n) <> 90) Then

                 UHTRP = UHTRP + UHPower(n)*pi/36*sinD(position_theta(n))*pi/36

                 UHRHCPTRP = UHRHCPTRP + UHRHCPPower(n)*pi/36*sinD(position_theta(n))*pi/36

                 UHLHCPTRP = UHLHCPTRP + UHLHCPPower(n)*pi/36*sinD(position_theta(n))*pi/36

            'Total = Total + Power_am(n)*pi/36*(sinD(position_theta(n))+sinD(position_theta(n-1)))/2*pi/36

            End If

        Next n

        Dim TRP As Double,StimPower As Double,AcceptPower As Double, RadEffi As Double, TotEffi As Double

        TRP = FarfieldPlot.GetTRP

        RadEffi = FarfieldPlot.GetRadiationEfficiency

        TotEffi = FarfieldPlot.GetTotalEfficiency

        StimPower = TRP/TotEffi

        AcceptPower = TRP/RadEffi

        Dim UHTotEffi As Double, dBTotal As Double

        Dim UHRHCPTotEffi As Double, dBRight As Double

        Dim UHLHCPTotEffi As Double,dBLeft As Double

        'efficiency calculation

        UHTotEffi = UHTRP/StimPower

        dBTotal = Log(UHTotEffi)/Log(10)*10

        UHRHCPEffi = UHRHCPTRP/StimPower

        dBRight = Log(UHRHCPEffi)/Log(10)*10

        UHLHCPEffi = UHLHCPTRP/StimPower

        dBLeft = Log(UHLHCPEffi)/Log(10)*10


        'Print information to the message window

        ReportInformationToWindow( _
        "下半球总效率@"+FrequencyStr+"GHz: "+Left(Cstr(UHTotEffi*100),InStr(Cstr(UHTotEffi*100),".")+2)+"% ("+Left(Cstr(dBTotal),InStr(Cstr(dBTotal),".")+2)+ "dB)"+vbCrLf+ _
        "下半球右旋效率@"+FrequencyStr+"GHz: "+Left(Cstr(UHRHCPEffi*100),InStr(Cstr(UHRHCPEffi*100),".")+2)+"% ("+Left(Cstr(dBRight),InStr(Cstr(dBRight),".")+2)+ "dB)"+ vbCrLf+ _
        "下半球左旋效率@"+FrequencyStr+"GHz: "+Left(Cstr(UHLHCPEffi*100),InStr(Cstr(UHLHCPEffi*100),".")+2)+"% ("+Left(Cstr(dBLeft),InStr(Cstr(dBLeft),".")+2)+ "dB)")

    End If

    'resume last plot mode

    FarfieldPlot.Reset

    FarfieldPlot.SetScaleLinear("False")

    FarfieldPlot.SetPlotMode(CurrentPlotMode)

    FarfieldPlot.StoreSettings


End Sub
