' Upper hemisphere efficiencies with RHCP and LHCP characteristics are supported
'2022-05-05 by Shawn Shi
Option Explicit
'#include "vba_globals_all.lib"
Public farfieldMode As Integer, TE As Double, RE As Double, PortStr As String, FrequencyStr As String, Frequency As Double
Public projectPath As String, dataFile As String

Sub Main ()
    'Get current farfield plot mode

    If MsgBox("Local WCS shall be used for farfield calculation, please make sure the current WCS has been properly aligned.",vbYesNo,"Information") <> vbYes Then
    	Exit Sub
    End If


	Begin Dialog UserDialog 320,70,"Save CP farfield componets",.DialogFunction ' %GRID:10,7,1,1
		OKButton 70,49,90,21
		CancelButton 170,49,90,21
		GroupBox 0,7,320,42,"Set plot mode:",.GroupBox6
		OptionGroup .Group1
			OptionButton 10,28,100,14,"Directivity",.OptionButton3
			OptionButton 120,28,60,14,"Gain",.OptionButton4
			OptionButton 190,28,120,14,"Realized Gain",.OptionButton5
	End Dialog
	Dim dlg As UserDialog
	dlg.Group1 = 0

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

End Sub
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	'Dim parameterFile As String
   	'Dim prjPath As String

   	'prjPath = GetProjectPath("Project")
   	'parameterFile = prjPath + "\dialog_parameter.txt"

	Select Case Action%
	Case 1 ' Dialog box initialization
		'ReStoreAllDialogSettings_LIB(parameterFile)
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
			Dim plotModes As Variant
			Dim NoticeInformation As String
			Dim SelectedItem As String
		    Dim n As Integer
		    Dim Frequency As Double, FrequencyStr As String
		    Dim Columns As String

			plotModes = Array("directivity", "gain", "realized gain")
			'farfieldMode = dlg.Group1
			farfieldMode = DlgValue("Group1")
			'FarfieldPlot.SelectComponent("Abs")
			FarfieldPlot.PlotType("polar")
			FarfieldPlot.SetAxesType("currentwcs")
			FarfieldPlot.SetAntennaType("unknown")

			'Select Case farfieldMode
		   	'Case 0
			'	FarfieldPlot.SetPlotMode("directivity")
			'Case 1
			'	FarfieldPlot.SetPlotMode("gain")
			'Case 2
			'	FarfieldPlot.SetPlotMode("realized gain")
		   	'End Select
			FarfieldPlot.SetPlotMode(plotModes(farfieldMode))

		   	If FarfieldPlot.IsScaleLinear = True Then
				FarfieldPlot.SetScaleLinear(False)
		   	End If

			'FarfieldPlot.SetPlotMode("Directivity")
			'FarfieldPlot.SetAntennaType("directional_linear")
			'FarfieldPlot.SetAntennaType("directional_circular")
			FarfieldPlot.SetCoordinateSystemType("ludwig3")
			FarfieldPlot.SetAutomaticCoordinateSystem("True")
			FarfieldPlot.SetPolarizationType("Circular")

		    'FarField Calculation
		    'Dim PortStr As String

		    SelectedItem = GetSelectedTreeItem

		    If (InStr(SelectedItem,"farfield (") = 0) Then

		        MsgBox("Please select a farfield result before runing this macro.",vbCritical,"Warning")

		        Exit All

		    Else

		        'Get the frequency and port of the selected item

		        FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)
		        PortStr  = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)
		        PortStr = Replace(PortStr,"[","")
		        PortStr = Replace(PortStr,"]","")
		        Frequency = CDbl(FrequencyStr)

				 '==============================write directivity data============================
				projectPath = GetProjectPath("Project")+"\Result"
				'Select Case farfieldMode
			   	'Case 0
				'	dataFile = projectPath+"\Circularly polarized directivity_frequency="+FrequencyStr+"GHz Port="+PortStr+"_"+Replace(CStr(Time),":","_")+".xlsx"
				'Case 1
				'	dataFile = projectPath+"\Circularly polarized gain_frequency="+FrequencyStr+"GHz Port="+PortStr+"_"+Replace(CStr(Time),":","_")+".xlsx"
				'Case 2
				'	dataFile = projectPath+"\Circularly polarized realized gain_frequency="+FrequencyStr+"GHz Port="+PortStr+"_"+Replace(CStr(Time),":","_")+".xlsx"
			   	'End Select
			   	dataFile = projectPath+"\Circularly polarized "+plotModes(farfieldMode)+"_frequency="+FrequencyStr+"GHz Port="+PortStr+"_"+Replace(CStr(Time),":","_")+".xlsx"
			   	dataFile = Replace(dataFile, "[", "")
			   	dataFile = Replace(dataFile, "]", "")
				'dataFile = projectPath+"\Circularly polarized directivity_frequency="+FrequencyStr+"GHz Port="+PortStr+".xlsx"
				Columns = "BCDEFGHIJKLMN"

				'NoticeInformation = "The directivity data is under£¨"+projectPath+"\£©"
				ReportInformationToWindow "Data export in progress............."
				'RE = Round(FarfieldPlot.GetTotalEfficiency - FarfieldPlot.GetRadiationEfficiency,2)
				If FarfieldPlot.GetSystemTotalEfficiency > -100 Then

					TE = FarfieldPlot.GetSystemTotalEfficiency
					RE = FarfieldPlot.GetSystemRadiationEfficiency
				Else
					TE = FarfieldPlot.GetTotalEfficiency
					RE = FarfieldPlot.GetRadiationEfficiency
				End If

				Dim O As Object

			    Set O = CreateObject("Excel.Application")
				If Dir(dataFile) = "" Then
				    Dim wBook As Object
				    Set wBook = O.Workbooks.Add
				    With wBook
				        .Title = "Title"
				        .Subject = "Subject"
				        .SaveAs fileName:= dataFile
					End With
				Else
				    Set wBook = O.Workbooks.Open(dataFile)
				End If


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

				Dim sheet As Object
				For Each sheet In wBook.Sheets
				    If sheet.Name Like "Sheet*" Then
				        sheet.Delete
				    End If
				Next
				'process sheet data, axial ratio, coloring, scoring and so on
				ReportInformationToWindow "Data estimation in progress............."
				processDirectivityData(farfieldSheet, Columns)
				process1DResults(sSheet)
				wBook.Save
				O.ActiveWorkbook.Close
				O.quit

			   	NoticeInformation = "The "&plotModes(farfieldMode)&" data is under£¨"+projectPath+"\£©"
			   	ReportInformationToWindow "Progress ends successfully."
			   	ReportInformationToWindow(NoticeInformation)
			   	MsgBox "Farfied data are exported successfully.",vbInformation,"Done"

			End If
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


Sub processDirectivityData(farfieldSheet As Object, Columns As String)
	FarfieldPlot.Reset
    '==============Upper Hemisphere RHCP and LHCP directivity abstraction===============

    Dim upperHemisphereRHCPdirectivity() As Double, upperHemisphereLHCPdirectivity() As Double

    Dim Theta As Double, Phi As Double

    Dim position_theta() As Double, position_phi() As Double
    Dim AVGPower As Double
    Dim n As Integer

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
	'write rhcp directivity
	farfieldSheet.Range("A1").value = "Polarization"
	farfieldSheet.Range("B1").value = "RHCP"
	farfieldSheet.Range("C1").value = "Frequency"
	farfieldSheet.Range("D1").value = FrequencyStr+"GHz"
	farfieldSheet.Range("E1").value = "Port"
	farfieldSheet.Range("F1").value = PortStr
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
	farfieldSheet.Range("D18").value = FrequencyStr+"GHz"
	farfieldSheet.Range("E18").value = "Port"
	farfieldSheet.Range("F18").value = PortStr
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


    Dim headers As Variant
    headers = Array("Polarization", "AR", "Frequency", farfieldSheet.Range("D1").Value, "Port", farfieldSheet.Range("F1").Value)
    Dim headerCols As Variant
    headerCols = Array("A", "B", "C", "D", "E", "F")

    For i = LBound(headers) To UBound(headers)
        farfieldSheet.Range(headerCols(i) & "35").Value = headers(i)
    Next i

	farfieldSheet.Range("A36").value = "Phi\Theta"

	For i = 0 To Len(Columns)-1
		farfieldSheet.Range(Mid(Columns,i+1,1)+"36").value = i*15
		farfieldSheet.Range("A"+Cstr(i+37)) = i*30
	Next

	'hide lhcp directivity data
	farfieldSheet.Rows("17:33").Hidden = True

	'formatting cells

    With farfieldSheet
       .Columns("A").ColumnWidth = 14
       .Columns("C").ColumnWidth = 10
       .Rows("1").RowHeight = 25
       .Rows("35").RowHeight = 25
       .Range("A1:Z100").HorizontalAlignment = -4108
       .Range("P2:Q8").Borders.LineStyle = 1
    End With

	Dim Dvalue As Double, deltaDirectivity As Double, axialRatio As Double
	For i  = 0 To 12
		For j = 0 To 12

			'=======================Axial ratio estimating and coloring============================
			deltaDirectivity = farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value - farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).Value

			axialRatio = Sgn(deltaDirectivity-0.0001)*20*CST_Log10((10^(deltaDirectivity/20)+1)/(Abs(10^(deltaDirectivity/20)-1)+0.0001))

			farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).value = Round(axialRatio,2)

			If axialRatio >= 0 And axialRatio < 3 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(0, 130, 0)
			ElseIf axialRatio < 6 And axialRatio >= 3 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(0, 180, 0)
			ElseIf axialRatio < 10 And axialRatio >= 6 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(145, 218, 0)
			ElseIf axialRatio < 18 And axialRatio >= 10 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(216, 254, 154)
			ElseIf axialRatio >= 18 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 255, 0)
			ElseIf axialRatio < -14 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 200, 0)
			ElseIf axialRatio < -6 And axialRatio >= -14 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(255, 0, 0)
			ElseIf axialRatio < 0 And axialRatio >= -6  Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+37)).Interior.Color = RGB(150, 0, 0)
			End If
			'color = farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color
			'========================Coloring rhcp directvity data===================================

			Dvalue = farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Value

			Select Case farfieldMode
		   	Case 0	'reference total efficiency -7dB when the directivity is selected as the farfield component
				'Dvalue = Dvalue
				Dvalue = Dvalue+TE
			Case 1
				Dvalue = Dvalue+TE-RE
			Case 2
				Dvalue = Dvalue
		   	End Select


			If Dvalue >= -6 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 130, 0)
			ElseIf Dvalue < -6 And Dvalue >= -8 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(0, 180, 0)
			ElseIf Dvalue <-8 And Dvalue >= -10 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(145, 218, 0)
			ElseIf Dvalue < -10 And Dvalue >= -12 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(216, 254, 154)
			ElseIf Dvalue < -12 And Dvalue >= -14 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 255, 0)
			ElseIf Dvalue < -14 And Dvalue >= -16 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 200, 0)
			ElseIf Dvalue < -16 And Dvalue >= -20 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(255, 0, 0)
			ElseIf Dvalue < -20 Then
				farfieldSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).Interior.Color = RGB(150, 0, 0)
			End If


		Next
	Next
	'======================================UH/Tot===============================
	'farfieldSheet.Range("A51") = "UHPower ratio"
	'farfieldSheet.Range("B51") = Round(10*CST_Log10(getUpperHemisphereRatio(farfieldSheet, columns)),2)
	'farfieldSheet.Range("C51") = "dB"
	writeAverageDirectivityAndRating(farfieldSheet, Columns)

End Sub
Sub writeAverageDirectivityAndRating(farfieldSheet As Object, Columns As String)

	'Formatting cells

	With farfieldSheet
	    .Columns("P").ColumnWidth = 15
	    .Columns("Q").ColumnWidth = 33
	    .Range("P2,P8,Q2").Interior.Color = RGB(221, 235, 247)
	    .Range("P3:P7").Interior.Color = RGB(0, 176, 240)
	    .Range("Q3:Q8").Interior.Color = RGB(255, 217, 102)
	    .Range("A1:Q100").Font.Bold = True
	    .Range("Q3:Q8").Font.Color = RGB(255, 0, 0)
	End With

	With farfieldSheet
	    .Range("P2").Value = "Within theta"
	    .Range("P3").Value = "30"
	    .Range("P4").Value = "45"
	    .Range("P5").Value = "60"
	    .Range("P6").Value = "90"
	    .Range("P7").Value = "120"
	    .Range("P8").Value = "RHCPD rating"
	    .Range("Q2").Value = "Weighted Average directivity"
	End With

	farfieldSheet.Range("Q3").Formula = _
	"=ROUND(10*LOG10(((10^(B3/10)+10^(B4/10)+10^(B5/10)+10^(B6/10)"+ _
	"+10^(B7/10)+10^(B8/10)+10^(B9/10)+10^(B10/10)+10^(B11/10)+10^(B12/10)+10^(B13/10)+"+ _
	"10^(B14/10))*(1-COS(PI()/24))*PI()/6+(10^(C3/10)+10^(C4/10)+10^(C5/10)+10^(C6/10)+"+ _
	"10^(C7/10)+10^(C8/10)+10^(C9/10)+10^(C10/10)+10^(C11/10)+10^(C12/10)+10^(C13/10)+"+ _
	"10^(C14/10))*(COS(PI()/24)-COS(PI()/8))*PI()/6+(10^(D3/10)+10^(D4/10)+10^(D5/10)+"+ _
	"10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+10^(D10/10)+10^(D11/10)+10^(D12/10)+"+ _
	"10^(D13/10)+10^(D14/10))*(COS(PI()/8)-COS(PI()/6))*PI()/6)/(2*PI()*(1-COS(PI()/6)))),2)"

	farfieldSheet.Range("Q4").Formula = _
	"=ROUND(10*LOG10(((10^(Q3/10)*(2*PI()*(1-COS(PI()/6)))+"+ _
	"(10^(D3/10)+10^(D4/10)+10^(D5/10)+10^(D6/10)+10^(D7/10)+10^(D8/10)+10^(D9/10)+"+ _
	"10^(D10/10)+10^(D11/10)+10^(D12/10)+10^(D13/10)+10^(D14/10))*(COS(PI()/6)-"+ _
	"COS(5*PI()/24))*PI()/6+(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+"+ _
	"10^(E8/10)+10^(E9/10)+10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))"+ _
	"*(COS(5*PI()/24)-COS(PI()/4))*PI()/6)/(2*PI()*(1-COS(PI()/4))))),2)"

	farfieldSheet.Range("Q5").Formula = _
	"=ROUND(10*LOG10(((10^(Q4/10)*(2*PI()*(1-COS(PI()/4)))+"+ _
	"(10^(E3/10)+10^(E4/10)+10^(E5/10)+10^(E6/10)+10^(E7/10)+10^(E8/10)+10^(E9/10)+"+ _
	"10^(E10/10)+10^(E11/10)+10^(E12/10)+10^(E13/10)+10^(E14/10))*(COS(PI()/4)-"+ _
	"COS(7*PI()/24))*PI()/6+((10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+"+ _
	"10^(F7/10)+10^(F8/10)+10^(F9/10)+10^(F10/10)+10^(F11/10)+10^(F12/10)+"+ _
	"10^(F13/10)+10^(F14/10))*(COS(7*PI()/24)-COS(PI()/3))*PI()/6))/(2*PI()*(1-COS(PI()/3))))),2)"

	farfieldSheet.Range("Q6").Formula = _
	"=ROUND(10*LOG10(((10^(Q5/10)*2*PI()*(1-COS(PI()/3))+"+ _
	"(10^(F3/10)+10^(F4/10)+10^(F5/10)+10^(F6/10)+10^(F7/10)+10^(F8/10)+10^(F9/10)+"+ _
	"10^(F10/10)+10^(F11/10)+10^(F12/10)+10^(F13/10)+10^(F14/10))*(COS(8*PI()/24)-"+ _
	"COS(9*PI()/24))*PI()/6+(10^(G3/10)+10^(G4/10)+10^(G5/10)+10^(G6/10)+10^(G7/10)+"+ _
	"10^(G8/10)+10^(G9/10)+10^(G10/10)+10^(G11/10)+10^(G12/10)+10^(G13/10)+10^(G14/10))"+ _
	"*(COS(9*PI()/24)-COS(11*PI()/24))*PI()/6+(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)"+ _
	"+10^(H7/10)+10^(H8/10)+10^(H9/10)+10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+"+ _
	"10^(H14/10))*(COS(11*PI()/24)-COS(PI()/2))*PI()/6))/(2*PI()*(1-COS(PI()/2)))),2)"

	farfieldSheet.Range("Q7").Formula = _
	"=ROUND(10*LOG10((10^(Q6/10)*2*PI()*(1-COS(PI()/2))+"+ _
	"(10^(H3/10)+10^(H4/10)+10^(H5/10)+10^(H6/10)+10^(H7/10)+10^(H8/10)+10^(H9/10)+"+ _
	"10^(H10/10)+10^(H11/10)+10^(H12/10)+10^(H13/10)+10^(H14/10))*(COS(PI()/2)-"+ _
	"COS(13*PI()/24))*PI()/6+(10^(I3/10)+10^(I4/10)+10^(I5/10)+10^(I6/10)+10^(I7/10)"+ _
	"+10^(I8/10)+10^(I9/10)+10^(I10/10)+10^(I11/10)+10^(I12/10)+10^(I13/10)+"+ _
	"10^(I14/10))*(COS(13*PI()/24)-COS(15*PI()/24))*PI()/6+(10^(J3/10)+10^(J4/10)+"+ _
	"10^(J5/10)+10^(J6/10)+10^(J7/10)+10^(J8/10)+10^(J9/10)+10^(J10/10)+10^(J11/10)+"+ _
	"10^(J12/10)+10^(J13/10)+10^(J14/10))*(COS(15*PI()/24)-COS(16*PI()/24))*PI()/6)/(2*PI()*(1-COS(2*PI()/3)))),2)"

	Select Case farfieldMode
  	Case 0	'reference total efficiency -10dB when the directivity is selected as the farfield component
		farfieldSheet.Range("Q8").Formula = _
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
		farfieldSheet.Range("Q8").Formula = _
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
		farfieldSheet.Range("Q8").Formula = _
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
	sSheet.Range("A1").Value = "Freq_" & PortStr
	sSheet.Range("B1").Value = "S11/dB"
	sSheet.Range("C1").Value = "Rad_eff"
	sSheet.Range("D1").Value = "Tot_eff"

	'System efficiencies or efficiencies?
	fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\System Rad. Efficiency [" & PortStr & "]")
	If fileName = "" Then
		fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Rad. Efficiency [" & PortStr & "]")
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

	fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\System Tot. Efficiency [" & PortStr & "]")
	If fileName = "" Then
		fileName = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Tot. Efficiency [" & PortStr & "]")
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

	fileName = ResultTree.GetFileFromTreeItem("1D Results\S-Parameters\S" & PortStr & ","& PortStr)
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
	ExecuteVBACodeToPlot(sSheet)
End Sub

Sub ExecuteVBACodeToPlot(ws As Object)
    Dim vbaCode As String
    Dim moduleObj As Object, wBook As Object
    Dim sheetName As String

    Set wBook = ws.Parent
    sheetName = ws.Name

    vbaCode = "Sub CreateChartFromCode()" & vbCrLf & _
                "    Dim ws As Worksheet" & vbCrLf & _
                "    Set ws = ThisWorkbook.Sheets(""" & sheetName & """)" & vbCrLf & _
                "    ws.Columns(""A:D"").Select" & vbCrLf & _
                "    ws.Shapes.AddChart.Select" & vbCrLf & _
                "    ActiveChart.ChartType = xlXYScatterSmoothNoMarkers" & vbCrLf & _
                "    ActiveChart.SetSourceData Source:=ws.Range(""A:D"")" & vbCrLf & _
                "    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).IncrementLeft -50" & vbCrLf & _
                "    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).IncrementTop -50" & vbCrLf & _
                "    ActiveChart.Axes(xlCategory).Select" & vbCrLf & _
     			"    Dim minVal As Double, maxVal As Double" & vbCrLf & _
                "    minVal = WorksheetFunction.Min(ws.Range(""A:A""))" & vbCrLf & _
                "    maxVal = WorksheetFunction.Max(ws.Range(""A:A""))" & vbCrLf & _
                "    ActiveChart.Axes(xlCategory).MinimumScale = minVal" & vbCrLf & _
                "    ActiveChart.Axes(xlCategory).MaximumScale = maxVal" & vbCrLf & _
                "    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).ScaleWidth 2.044791776, msoFalse, msoScaleFromTopLeft" & vbCrLf & _
                "    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).ScaleHeight 1.9982640712, msoFalse, msoScaleFromTopLeft" & vbCrLf & _
                "End Sub"

    On Error Resume Next
    Set moduleObj = wBook.VBProject.VBComponents.Add(1)
    If Err.Number <> 0 Then
        MsgBox "Unable to insert code module: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    moduleObj.CodeModule.AddFromString vbaCode
    Dim Application As Object
    Set Application = ws.Application
    On Error Resume Next
    Application.Run "CreateChartFromCode"
    If Err.Number <> 0 Then
        MsgBox "Executing module failed!: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
    On Error Resume Next
    wBook.VBProject.VBComponents.Remove moduleObj
    On Error GoTo 0
End Sub

