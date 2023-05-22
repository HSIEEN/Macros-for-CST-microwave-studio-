' Upper hemisphere efficiencies with RHCP and LHCP characteristics are supported
'2022-05-05 by Shawn Shi
'Option Explicit
'#include "vba_globals_all.lib"

Sub Main ()
    'Get current farfield plot mode
    'Dim CurrentPlotMode As String

    If MsgBox("Please make sure the axis W in current WCS coordinated system point to the Zenith",vbYesNo,"Information") <> vbYes Then
    	Exit Sub
    End If

	'FarfieldPlot.SelectComponent("Abs")
	FarfieldPlot.PlotType("polar")
	FarfieldPlot.SetAxesType("currentwcs")
	FarfieldPlot.SetAntennaType("unknown")
	FarfieldPlot.SetPlotMode("Directivity")
	'FarfieldPlot.SetAntennaType("directional_linear")
	'FarfieldPlot.SetAntennaType("directional_circular")
	FarfieldPlot.SetCoordinateSystemType("ludwig3")
	FarfieldPlot.SetAutomaticCoordinateSystem("True")
	FarfieldPlot.SetPolarizationType("Circular")

    'FarField Calculation

    Dim SelectedItem As String

    Dim n As Integer

    Dim Frequency As Double, FrequencyStr As String
    Dim PortStr As String

    SelectedItem = GetSelectedTreeItem

    If (InStr(SelectedItem,"farfield (") = 0) Then

        MsgBox("Please select a farfield result before runing this macro.",vbCritical,"Warning")

        Exit All

    Else

        'Get the frequency and port of the selected item

        FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)
        PortStr  = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)
        Frequency = CDbl(FrequencyStr)

        FarfieldPlot.Reset
        '==============Upper Hemisphere RHCP and LHCP gain abstraction===============

        Dim upperHemisphereRHCPdirectivity() As Double, upperHemisphereLHCPdirectivity() As Double

        Dim Theta As Double, Phi As Double

        Dim position_theta() As Double, position_phi() As Double
        Dim AVGPower As Double
        Dim Columns As String

        For Phi=0 To 360 STEP 30

             For Theta = 0 To 180 STEP 15

                 FarfieldPlot.AddListEvaluationPoint(Theta, Phi, 0, "spherical", "frequency", Frequency)

             Next Theta

        Next Phi

        FarfieldPlot.CalculateList("")

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
		dataFile = projectPath+"\Circularly polarized directivity_frequency="+FrequencyStr+"GHz Port="+PortStr+".xlsx"
		Columns = "BCDEFGHIJKLMN"

		NoticeInformation = "The directivity data is under£¨"+projectPath+"\£©"
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

		wBook.Sheets.Add.Name = "Frequency="+FrequencyStr+"GHz"
		Set wSheet = wBook.Sheets("Frequency="+FrequencyStr+"GHz")

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

    End If

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
	sheet.Columns("A").ColumnWidth = 18

	sheet.Columns("C").ColumnWidth = 10

	sheet.Rows("1").RowHeight = 25
	sheet.Rows("35").RowHeight = 25
	'sheet.Range("A1:Z100").HorizontalAlignment = xlCenter


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
	writeAverageDirectivity(sheet, Columns)

End Sub
Sub writeAverageDirectivity(sheet As Object, Columns As String)
	'Formating cells

	sheet.Columns("P").ColumnWidth = 15
	sheet.Range("P2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("P3:P7").Interior.Color = RGB(0, 176, 240)
	sheet.Range("P8").Interior.Color = RGB(221, 235, 247)
	sheet.Range("Q2").Interior.Color = RGB(221, 235, 247)
	sheet.Range("Q3:Q7").Interior.Color = RGB(255, 217, 102)
	sheet.Range("Q8").Interior.Color = RGB(255, 217, 102)
	sheet.Columns("Q").ColumnWidth = 30
	sheet.Range("A1:Q100").Font.Bold = True
	sheet.Range("Q3:Q8").Font.Color = RGB(255, 0, 0)

	sheet.Range("P2").Value = "Within theta"
	sheet.Range("Q2").Value = "Weighted Average direcitivy"
	sheet.Range("P8").Value = "RHCPD rating"

	sheet.Range("P3").Value = "30"
	sheet.Range("P4").Value = "45"
	sheet.Range("P5").Value = "60"
	sheet.Range("P6").Value = "90"
	sheet.Range("P7").Value = "120"

	sheet.Range("Q3").Value = Round(10*CST_Log10(((10^(sheet.Range("B3").Value/10)+ _
	10^(sheet.Range("B4").Value/10)+10^(sheet.Range("B5").Value/10)+ _
	10^(sheet.Range("B6").Value/10)+10^(sheet.Range("B7").Value/10)+ _
	10^(sheet.Range("B8").Value/10)+10^(sheet.Range("B9").Value/10)+ _
	10^(sheet.Range("B10").Value/10)+10^(sheet.Range("B11").Value/10)+ _
	10^(sheet.Range("B12").Value/10)+10^(sheet.Range("B13").Value/10)+ _
	10^(sheet.Range("B14").Value/10))*(1-Cos(pi/24))*pi/6+(10^(sheet.Range("C3").Value/10)+ _
	10^(sheet.Range("C4").Value/10)+10^(sheet.Range("C5").Value/10)+ _
	10^(sheet.Range("C6").Value/10)+10^(sheet.Range("C7").Value/10)+ _
	10^(sheet.Range("C8").Value/10)+10^(sheet.Range("C9").Value/10)+ _
	10^(sheet.Range("C10").Value/10)+10^(sheet.Range("C11").Value/10)+ _
	10^(sheet.Range("C12").Value/10)+10^(sheet.Range("C13").Value/10)+ _
	10^(sheet.Range("C14").Value/10))*(Cos(pi/24)-Cos(pi/8))*pi/6+ _
	(10^(sheet.Range("D3").Value/10)+10^(sheet.Range("D4").Value/10)+ _
	10^(sheet.Range("D5").Value/10)+10^(sheet.Range("D6").Value/10)+ _
	10^(sheet.Range("D7").Value/10)+10^(sheet.Range("D8").Value/10)+ _
	10^(sheet.Range("D9").Value/10)+10^(sheet.Range("D10").Value/10)+ _
	10^(sheet.Range("D11").Value/10)+10^(sheet.Range("D12").Value/10)+ _
	10^(sheet.Range("D13").Value/10)+ _
	10^(sheet.Range("D14").Value/10))*(Cos(pi/8)-Cos(pi/6))*pi/6)/(2*pi*(1-Cos(pi/6)))),2)

	sheet.Range("Q4").Value = Round(10*CST_Log10(((10^(sheet.Range("Q3").Value/10)*(2*pi*(1-Cos(pi/6)))+ _
	(10^(sheet.Range("D3").Value/10)+10^(sheet.Range("D4").Value/10)+ _
	10^(sheet.Range("D5").Value/10)+10^(sheet.Range("D6").Value/10)+ _
	10^(sheet.Range("D7").Value/10)+10^(sheet.Range("D8").Value/10)+ _
	10^(sheet.Range("D9").Value/10)+10^(sheet.Range("D10").Value/10)+ _
	10^(sheet.Range("D11").Value/10)+10^(sheet.Range("D12").Value/10)+ _
	10^(sheet.Range("D13").Value/10)+ _
	10^(sheet.Range("D14").Value/10))*(Cos(pi/6)-Cos(5*pi/24))*pi/6+ _
	(10^(sheet.Range("E3").Value/10)+10^(sheet.Range("E4").Value/10)+ _
	10^(sheet.Range("E5").Value/10)+10^(sheet.Range("E6").Value/10)+ _
	10^(sheet.Range("E7").Value/10)+10^(sheet.Range("E8").Value/10)+ _
	10^(sheet.Range("E9").Value/10)+10^(sheet.Range("E10").Value/10)+ _
	10^(sheet.Range("E11").Value/10)+10^(sheet.Range("E12").Value/10)+ _
	10^(sheet.Range("E13").Value/10)+ _
	10^(sheet.Range("E14").Value/10))*(Cos(5*pi/24)-Cos(pi/4))*pi/6)/(2*pi*(1-Cos(pi/4))))),2)

	sheet.Range("Q5").Value = Round(10*CST_Log10(((10^(sheet.Range("Q4").Value/10)*(2*pi*(1-Cos(pi/4)))+ _
	(10^(sheet.Range("E3").Value/10)+10^(sheet.Range("E4").Value/10)+ _
	10^(sheet.Range("E5").Value/10)+10^(sheet.Range("E6").Value/10)+ _
	10^(sheet.Range("E7").Value/10)+10^(sheet.Range("E8").Value/10)+ _
	10^(sheet.Range("E9").Value/10)+10^(sheet.Range("E10").Value/10)+ _
	10^(sheet.Range("E11").Value/10)+10^(sheet.Range("E12").Value/10)+ _
	10^(sheet.Range("E13").Value/10)+ _
	10^(sheet.Range("E14").Value/10))*(Cos(pi/4)-Cos(7*pi/24))*pi/6+ _
	((10^(sheet.Range("F3").Value/10)+10^(sheet.Range("F4").Value/10)+ _
	10^(sheet.Range("F5").Value/10)+10^(sheet.Range("F6").Value/10)+ _
	10^(sheet.Range("F7").Value/10)+10^(sheet.Range("F8").Value/10)+ _
	10^(sheet.Range("F9").Value/10)+10^(sheet.Range("F10").Value/10)+ _
	10^(sheet.Range("F11").Value/10)+10^(sheet.Range("F12").Value/10)+ _
	10^(sheet.Range("F13").Value/10)+ _
	10^(sheet.Range("F14").Value/10))*(Cos(7*pi/24)-Cos(pi/3))*pi/6))/(2*pi*(1-Cos(pi/3))))),2)

	sheet.Range("Q6").Value = Round(10*CST_Log10(((10^(sheet.Range("Q5").Value/10)*2*pi*(1-Cos(pi/3))+ _
	(10^(sheet.Range("F3").Value/10)+10^(sheet.Range("F4").Value/10)+ _
	10^(sheet.Range("F5").Value/10)+10^(sheet.Range("F6").Value/10)+ _
	10^(sheet.Range("F7").Value/10)+10^(sheet.Range("F8").Value/10)+ _
	10^(sheet.Range("F9").Value/10)+10^(sheet.Range("F10").Value/10)+ _
	10^(sheet.Range("F11").Value/10)+10^(sheet.Range("F12").Value/10)+ _
	10^(sheet.Range("F13").Value/10)+10^(sheet.Range("F14").Value/10))*(Cos(8*pi/24)-Cos(9*pi/24))*pi/6+ _
	(10^(sheet.Range("G3").Value/10)+10^(sheet.Range("G4").Value/10)+ _
	10^(sheet.Range("G5").Value/10)+10^(sheet.Range("G6").Value/10)+ _
	10^(sheet.Range("G7").Value/10)+10^(sheet.Range("G8").Value/10)+ _
	10^(sheet.Range("G9").Value/10)+10^(sheet.Range("G10").Value/10)+ _
	10^(sheet.Range("G11").Value/10)+10^(sheet.Range("G12").Value/10)+ _
	10^(sheet.Range("G13").Value/10)+ _
	10^(sheet.Range("G14").Value/10))*(Cos(9*pi/24)-Cos(11*pi/24))*pi/6+ _
	(10^(sheet.Range("H3").Value/10)+10^(sheet.Range("H4").Value/10)+ _
	10^(sheet.Range("H5").Value/10)+10^(sheet.Range("H6").Value/10)+ _
	10^(sheet.Range("H7").Value/10)+10^(sheet.Range("H8").Value/10)+ _
	10^(sheet.Range("H9").Value/10)+10^(sheet.Range("H10").Value/10)+ _
	10^(sheet.Range("H11").Value/10)+10^(sheet.Range("H12").Value/10)+ _
	10^(sheet.Range("H13").Value/10)+ _
	10^(sheet.Range("H14").Value/10))*(Cos(11*pi/24)-Cos(pi/2))*pi/6))/(2*pi*(1-Cos(pi/2)))),2)

	sheet.Range("Q7").Value = Round(10*CST_Log10((10^(sheet.Range("Q6").Value/10)*2*pi*(1-Cos(pi/2))+ _
	(10^(sheet.Range("H3").Value/10)+10^(sheet.Range("H4").Value/10)+ _
	10^(sheet.Range("H5").Value/10)+10^(sheet.Range("H6").Value/10)+ _
	10^(sheet.Range("H7").Value/10)+10^(sheet.Range("H8").Value/10)+ _
	10^(sheet.Range("H9").Value/10)+10^(sheet.Range("H10").Value/10)+ _
	10^(sheet.Range("H11").Value/10)+10^(sheet.Range("H12").Value/10)+ _
	10^(sheet.Range("H13").Value/10)+ _
	10^(sheet.Range("H14").Value/10))*(Cos(pi/2)-Cos(13*pi/24))*pi/6+ _
	(10^(sheet.Range("I3").Value/10)+10^(sheet.Range("I4").Value/10)+ _
	10^(sheet.Range("I5").Value/10)+10^(sheet.Range("I6").Value/10)+ _
	10^(sheet.Range("I7").Value/10)+10^(sheet.Range("I8").Value/10)+ _
	10^(sheet.Range("I9").Value/10)+10^(sheet.Range("I10").Value/10)+ _
	10^(sheet.Range("I11").Value/10)+10^(sheet.Range("I12").Value/10)+ _
	10^(sheet.Range("I13").Value/10)+ _
	10^(sheet.Range("I14").Value/10))*(Cos(13*pi/24)-Cos(15*pi/24))*pi/6+ _
	(10^(sheet.Range("J3").Value/10)+10^(sheet.Range("J4").Value/10)+ _
	10^(sheet.Range("J5").Value/10)+10^(sheet.Range("J6").Value/10)+ _
	10^(sheet.Range("J7").Value/10)+10^(sheet.Range("J8").Value/10)+ _
	10^(sheet.Range("J9").Value/10)+10^(sheet.Range("J10").Value/10)+ _
	10^(sheet.Range("J11").Value/10)+10^(sheet.Range("J12").Value/10)+ _
	10^(sheet.Range("J13").Value/10)+10^(sheet.Range("J14").Value/10))*(Cos(15*pi/24)- _
	Cos(16*pi/24))*pi/6)/(2*pi*(1-Cos(2*pi/3)))),2)

	sheet.Range("Q8").Formula = "=ROUND((117-0.5*(1.75*(1.5*SUMPRODUCT((B3:E15<=-9)*(B3:E15>-100))+1.25*SUMPRODUCT((B3:E15<=-8)*(B3:E15>-9))+SUMPRODUCT((B3:E15<=-7)*(B3:E15>-8))+0.75*SUMPRODUCT((B3:E15<=-6)*(B3:E15>-7))+0.5*SUMPRODUCT((B3:E15<=-5)*(B3:E15>-6))+0.25*SUMPRODUCT((B3:E15<=-4)*(B3:E15>-5)))+1.5*SUMPRODUCT((F3:H15<=-9)*(F3:H15>-100))+1.25*SUMPRODUCT((F3:H15<=-8)*(F3:H15>-9))+SUMPRODUCT((F3:H15<=-7)*(F3:H15>-8))+0.75*SUMPRODUCT((F3:H15<=-6)*(F3:H15>-7))+0.5*SUMPRODUCT((F3:H15<=-5)*(F3:H15>-6))+0.25*SUMPRODUCT((F3:H15<=-4)*(F3:H15>-5))+0.5*(1.5*SUMPRODUCT((I3:J15<=-9)*(I3:J15>-100))+1.25*SUMPRODUCT((I3:J15<=-8)*(I3:J15>-9))+SUMPRODUCT((I3:J15<=-7)*(I3:J15>-8))+0.75*SUMPRODUCT((I3:J15<=-6)*(I3:J15>-7))+0.5*SUMPRODUCT((I3:J15<=-5)*(I3:J15>-6))+0.25*SUMPRODUCT((I3:J15<=-4)*(I3:J15>-5)))))/117*100,2)"
End Sub
