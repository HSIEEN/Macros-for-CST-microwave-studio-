' Upper hemisphere efficiencies with RHCP and LHCP characteristics are supported
'2022-05-05 by Shawn Shi
Option Explicit
'#include "vba_globals_all.lib"
Public farfieldMode As Integer, TE As Double, RE As Double

Sub Main ()
    'Get current farfield plot mode

    If MsgBox("Local WCS shall be used for farfield calculation, please make sure the current WCS has been properly aligned.",vbYesNo,"Information") <> vbYes Then
    	Exit Sub
    End If


	Begin Dialog UserDialog 220,105,"Save CP farfield componets", .DialogFunction ' %GRID:10,7,1,1
		OKButton 20,84,90,21
		CancelButton 120,84,90,21
		GroupBox 10,7,210,70,"Set plot mode:",.GroupBox6
		OptionGroup .Group1
			OptionButton 20,28,100,14,"Directivity",.OptionButton3
			OptionButton 130,28,70,14,"Gain",.OptionButton4
			OptionButton 20,49,120,14,"Realized Gain",.OptionButton5
	End Dialog
	Dim dlg As UserDialog
	dlg.Group1 = 0

	If Dialog(dlg,-2) = 0 Then
		Exit All
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

			Select Case farfieldMode
		   	Case 0	'reference total efficiency -7dB when the directivity is selected as the farfield component
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
	writeAverageDirectivityAndRating(sheet, Columns)

End Sub
Sub writeAverageDirectivityAndRating(sheet As Object, Columns As String)

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

	Select Case farfieldMode
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
		        PortStr = Replace(PortStr,"[","")
		        PortStr = Replace(PortStr,"]","")
		        Frequency = CDbl(FrequencyStr)

		        FarfieldPlot.Reset
		        '==============Upper Hemisphere RHCP and LHCP directivity abstraction===============

		        Dim upperHemisphereRHCPdirectivity() As Double, upperHemisphereLHCPdirectivity() As Double

		        Dim Theta As Double, Phi As Double

		        Dim position_theta() As Double, position_phi() As Double
		        Dim AVGPower As Double
		        Dim Columns As String
		        Dim projectPath As String, dataFile As String
		        Dim NoticeInformation As String

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

				'NoticeInformation = "The directivity data is under（"+projectPath+"\）"
				ReportInformationToWindow "Data export in progress............."
				'RE = Round(FarfieldPlot.GetTotalEfficiency - FarfieldPlot.GetRadiationEfficiency,2)
				TE = FarfieldPlot.GetTotalEfficiency
				RE = FarfieldPlot.GetRadiationEfficiency
				Dim O As Object

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

				Dim wSheet As Object
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
				ReportInformationToWindow "Data estimation in progress............."
				processDirectivityData(wSheet, Columns)
				wBook.Save
				O.ActiveWorkbook.Close
				O.quit

			   	NoticeInformation = "The realized "&plotModes(farfieldMode)&" data is under（"+projectPath+"\）"
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




