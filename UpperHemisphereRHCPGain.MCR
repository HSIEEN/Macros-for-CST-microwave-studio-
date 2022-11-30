' Upper hemisphere efficiencies with RHCP and LHCP characteristics are supported
'2022-05-05 by Shawn Shi
'Option Explicit

Sub Main ()
    'Get current farfield plot mode
    Dim CurrentPlotMode As String

    MsgBox("Please make sure the axis W in current WCS coordinated system point to the Zenith",vbInformation,"Information")

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

        Exit All

    Else

        'Get the frequency of the selected item

        FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)
        PortStr  = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)
        Frequency = CDbl(FrequencyStr)

        FarfieldPlot.Reset
        '==============Upper Hemisphere TRP Calculation===============

        Dim UHPower() As Double, UHRHCPPower() As Double, UHLHCPPower() As Double

        Dim RHCPGain() As Double, LHCPGain() As Double

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

        UHPower = FarfieldPlot.GetList("Spherical  abs")

        UHRHCPPower = FarfieldPlot.GetList("Spherical circular right abs")

        UHLHCPPower = FarfieldPlot.GetList("Spherical circular left abs")

        position_theta = FarfieldPlot.GetList("Point_T")

        position_phi = FarfieldPlot.GetList("Point_P")

        ReDim RHCPGain(13,13)
        ReDim LHCPGain(13,13)
        AVGPower = 0.5/(4*pi)

        For n = 0 To UBound(UHPower)

             RHCPGain(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = Log(UHRHCPPower(n)/AVGPower)/Log(10)*10
             LHCPGain(CInt(position_phi(n)/30),CInt(position_theta(n)/15)) = Log(UHLHCPPower(n)/AVGPower)/Log(10)*10

            'linear to dB, Log(UHRHCPEffi)/Log(10)*10

        Next n


		PrjPath = GetProjectPath("Project")
		DataFile = PrjPath+"\CircularlyPolarizedGain.xlsx"
		Columns = "BCDEFGHIJKLMN"

		NoticeInformation = "生成的增益数据文件放在工程目录下（"+PrjPath+"\）"
	    MsgBox(NoticeInformation,vbInformation,"注意")

		Dim IsFileExist As String
	    Set O = CreateObject("Excel.Application")
		IsFileExist = Dir(DataFile)
		If IsFileExist = "" Then
			Set wBook  = O.Workbooks.Add
			With wBook
				.Title = "Title"
				.Subject = "Subject"
				.SaveAs Filename:= DataFile
			End With
		Else
			Set wBook = O.Workbooks.Open(DataFile)
		End If
		Set wSheet = wBook.Sheets(1)
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
				wSheet.Range(Mid(Columns,j+1,1) + CStr(i+3)).value = Round(RHCPGain(i,j),2)
			Next
		Next

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
				wSheet.Range(Mid(Columns,j+1,1) + CStr(i+20)).value = Round(LHCPGain(i,j),2)
			Next
		Next

		wBook.Save
		O.ActiveWorkbook.Close
		O.quit

    End If

    'resume last plot mode

    FarfieldPlot.Reset

    FarfieldPlot.SetScaleLinear("False")

    FarfieldPlot.SetPlotMode(CurrentPlotMode)

    FarfieldPlot.StoreSettings


End Sub
