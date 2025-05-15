'#Language "WWB-COM"
'
Option Explicit
Public Const lib_rundef = -1.2345678e27

Sub Main


	Dim time_suffix As String
	Dim time_str As String
	Dim date_str As String

	'time_str = Cstr(Time)
	time_str = Format(Time, "hhmmss")

	'time_str = Replace(time_str, ":", "")

	'date_str = CStr(Month(Date))+CStr(Day(Date))
	date_str = Format(Date, "mmdd")

	time_suffix = date_str + time_str

	'formating simulated results
	'If there is only one sp result, one re result and one te result, get the mag format of them and move it to comparison directory
	'else, only format them without moving
	Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
	'nResults = ResultTree.GetTreeResults("1D Results\Efficiencies", "folder 0D/1D recursive","","","","",)
	nResults = ResultTree.GetTreeResults("1D Results\Efficiencies", "0D/1D recursive","",paths,types,files,info)

	If nResults=2 Then
		'Format sp, re and te results and move them to "SP", "RE", "TE" 1D results directories
		Dim sSP As Object, sre As Object, ste As Object
		Dim spfile As String, refile As String, tefile As String

		spfile = ResultTree.GetFileFromTreeItem("1D Results\S-Parameters\S1,1")
		refile = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Rad. Efficiency [1]")
		tefile = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Tot. Efficiency [1]")
		Set sSP = Result1DComplex(spfile)
		Set sre = Result1DComplex(refile)
		Set ste = Result1DComplex(tefile)

		Set sSP = Mag_Linear2dB(sSP)
		Set sre = Mag_Linear2dB(sre)
		Set ste = Mag_Linear2dB(ste)

		sSP.DeleteAt("never")
		sSP.Type("dB")
		sSP.Xlabel("Frequecy/GHz")
		sSP.Ylabel("dB")
		sSP.Save("simulated_SP_mag_"+time_suffix+".sig")
		sSP.AddToTree("1D Results\SP\Simulated_SP_"+time_suffix)


		sre.DeleteAt("never")
		sre.Type("dB")
		sre.Xlabel("Frequecy/GHz")
		sre.Ylabel("dB")
		sre.Save("simulated_RE_mag_"+time_suffix+".sig")
		sre.AddToTree("1D Results\RE\Simulated_RE_"+time_suffix)

		ste.DeleteAt("never")
		ste.Type("dB")
		ste.Xlabel("Frequecy/GHz")
		ste.Ylabel("dB")
		ste.Save("simulated_TE_mag_"+time_suffix+".sig")
		ste.AddToTree("1D Results\TE\Simulated_TE_"+time_suffix)

	ElseIf nResults>2 Then

		MsgBox "Please select one S-Parameter From 1D results/S-Parameters ,it must be Sn,n", vbInformation
		Dim selectedItem0 As String , S1 As Variant ,  pos As Variant , Length As Variant, S2 As Variant
		Wait

		selectedItem0 = GetSelectedTreeItem
		'MsgBox selectedItem0
		pos = InStr( selectedItem0, ",")
		Length = Len(selectedItem0)- pos
		S1 = Right(selectedItem0, Length)
		'MsgBox S1
		S2 = Mid(selectedItem0,26,Length)
		'MsgBox S2
		If S1 <> S2 Then
			MsgBox "The selected parameter is incorrect. Please select again."
				GoTo theEnd
		End If



		Dim sSP1 As Object, sre1 As Object, ste1 As Object
		Dim spfile1 As String, refile1 As String, tefile1 As String

		spfile1 = ResultTree.GetFileFromTreeItem("1D Results\S-Parameters\S"+S1+","+S1)
		refile1 = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Rad. Efficiency ["+S1+"]")
		tefile1 = ResultTree.GetFileFromTreeItem("1D Results\Efficiencies\Tot. Efficiency ["+S1+"]")
		Set sSP1 = Result1DComplex(spfile1)
		Set sre1 = Result1DComplex(refile1)
		Set ste1 = Result1DComplex(tefile1)

		Set sSP1 = Mag_Linear2dB(sSP1)
		Set sre1 = Mag_Linear2dB(sre1)
		Set ste1 = Mag_Linear2dB(ste1)

		sSP1.DeleteAt("never")
		sSP1.Type("dB")
		sSP1.Xlabel("Frequecy/GHz")
		sSP1.Ylabel("dB")
		sSP1.Save("Port"+S1+"_simulated_SP_mag_"+time_suffix+".sig")
		sSP1.AddToTree("1D Results\Port"+S1+"_SP\Simulated_SP_"+time_suffix)


		sre1.DeleteAt("never")
		sre1.Type("dB")
		sre1.Xlabel("Frequecy/GHz")
		sre1.Ylabel("dB")
		sre1.Save("simulated_RE_mag_"+time_suffix+".sig")
		sre1.AddToTree("1D Results\Port"+S1+"_RE\Simulated_RE_"+time_suffix)

		ste1.DeleteAt("never")
		ste1.Type("dB")
		ste1.Xlabel("Frequecy/GHz")
		ste1.Ylabel("dB")
		ste1.Save("simulated_TE_mag_"+time_suffix+".sig")
		ste1.AddToTree("1D Results\Port"+S1+"_TE\Simulated_TE_"+time_suffix)








    	'Dim childItem As String
    	'Dim os As Object
    	'Dim sfile As String
    	'Dim suffix_0 As String, suffix_1 As String, suffix As String

    	'childItem = ResultTree.GetFirstChildName("1D Results\S-Parameters\")

		'While childItem <>"" And InStr(childItem, "Simulated") = 0
			'suffix_0 = Right(childItem, Len(childItem)-InStr(childItem,","))
			'suffix_1 = Mid(childItem,InStr(childItem, ",")-1, 1)
			'If StrComp(suffix_1, suffix_0) = 0 Then
				'sfile = ResultTree.GetFileFromTreeItem(childItem)
				'Set os = Result1DComplex(sfile)
				'Set os = Mag_Linear2dB(os)
				'os.Xlabel("Frequecy/GHz")
				'os.Ylabel("dB")
				'os.Save("simulated_sp_"+suffix+".sig")
				'os.AddToTree("1D Results\S-Parameters\Simulated_SP_"+suffix)
			'End If
			'childItem = ResultTree.GetNextItemName(childItem)
			'suffix = Right(childItem, Len(childItem)-InStr(childItem,","))
		'Wend

		'Dim childItem As String
    	'Dim os As Object
    	'Dim sfile As String
    	'Dim suffix As String
    	'Dim prefix As String

    	'childItem = ResultTree.GetFirstChildName("1D Results\Efficiencies\")

		'suffix = Right(childItem, 1)
		'While childItem <>"" And InStr(childItem, "Simulated") = 0
			'prefix = Mid$(childItem$,InStrRev(childItem,"\")+1,3)
			'suffix = Mid$(childItem$,InStr(childItem,"[")+1,InStr(childItem,"]")-InStr(childItem,"[")-1)
			'sfile = ResultTree.GetFileFromTreeItem(childItem)
			'Set os = Result1DComplex(sfile)
			'Set os = Mag_Linear2dB(os)
			'os.Xlabel("Frequecy/GHz")
			'os.Ylabel("dB")
			'os.Save("simulated_Effi_"+prefix+"_"+suffix+".sig")
			'os.AddToTree("1D Results\Efficiencies\Simulated_Eff_"+prefix+"_"+suffix)
			'childItem = ResultTree.GetNextItemName(childItem)
			'prefix = Mid$(childItem$,InStrRev(childItem,"\")+1,3)
			'suffix = Mid$(childItem$,InStr(childItem,"[")+1,InStr(childItem,"]")-InStr(childItem,"[")-1)
		'Wend
	ElseIf nResults<2 Then
		'Nothing to do

	End If

	'------------------------------------------------------"

	'Select a Excel file containing target data

	Dim efilename As String
	efilename  = GetFilePath("", "xls;xlsx", "", "Browse excel-file exported from GTS test system", 0)

	If efilename = "" Then
		GoTo theEnd
	End If

	'Open the Excel file and begin reading data
	Dim O As Object
	Set O = CreateObject("Excel.Application")

	Dim wBook As Object
	Set wBook = O.Workbooks.Open(efilename)

	Dim efilename_suffix As String



	'efilename_suffix =


	Dim wSheet As Object
	Set wSheet = wBook.Sheets("Sheet1")

	'Check the row numbers of the data
	Dim rowNum As Integer, i As Integer
	Dim cellName As String, cellValue As String
	Dim mfreq() As Double, mre() As Double
	Dim msp() As Double, mte() As Double
	rowNum=0
	For i=5 To 1000
		cellName = Replace("A" & Str(i)," ", "")
		cellValue = wSheet.Range(cellName).Value
		If Not cellValue = "" Then
			rowNum = rowNum + 1
		Else
			Exit For
		End If
	Next


	'Store data to varibles
	ReDim Preserve mfreq(rowNum-1) As Double
	ReDim Preserve mte(rowNum-1) As Double
	ReDim Preserve msp(rowNum-1) As Double
	ReDim Preserve mre(rowNum-1) As Double
	'ReDim Preserve mfreq(rowNum) As Double

	For i = 0 To rowNum-1
		mfreq(i)= wSheet.Range(Replace("B" & Str(i+5)," ", "")).Value/1000
		mte(i)= wSheet.Range(Replace("C" & Str(i+5)," ", "")).Value
		msp(i)= -wSheet.Range(Replace("Z" & Str(i+5)," ", "")).Value
		mre(i)= wSheet.Range(Replace("AC" & Str(i+5)," ", "")).Value
	Next

	wBook.Save
	O.ActiveWorkbook.Close
	O.quit

	'plot data

	Dim osp As Object
	Dim ore As Object
	Dim ote As Object

	Set osp = Result1D("")
	Set ore = Result1D("")
	Set ote = Result1D("")

	For i = 0 To rowNum-1
		osp.AppendXY(mfreq(i), msp(i))
		ore.AppendXY(mfreq(i), mre(i))
		ote.AppendXY(mfreq(i), mte(i))
	Next
	osp.Xlabel("Frequecy/GHz")
	ore.Xlabel("Frequecy/GHz")
	ote.Xlabel("Frequecy/GHz")

	osp.Title("Measured SP"+"/dB")
	ore.Title("Measured RE"+"/dB")
	ote.Title("Measured TE"+"/dB")

	osp.Ylabel("dB" )
	ore.Ylabel("dB" )
	ote.Ylabel("dB" )

	osp.DeleteAt("never")
	ore.DeleteAt("never")
	ote.DeleteAt("never")


	osp.Type("dB")
	ore.Type("dB")
	ote.Type("dB")

If nResults=2 Then


	osp.Save("measured_SP_From_GTS_"+time_suffix+".sig")
	ore.Save("measured_RE_From_GTS_"+time_suffix+".sig")
	ote.Save("measured_TE_From_GTS_"+time_suffix+".sig")

	osp.AddToTree("1D Results\SP\Measured_SP_"+time_suffix)
	ore.AddToTree("1D Results\RE\Measured_RE_"+time_suffix)
	ote.AddToTree("1D Results\TE\Measured_TE_"+time_suffix)

ElseIf nResults>2 Then

	osp.Save("measured_SP_From_GTS_"+time_suffix+".sig")
	ore.Save("measured_RE_From_GTS_"+time_suffix+".sig")
	ote.Save("measured_TE_From_GTS_"+time_suffix+".sig")

	osp.AddToTree("1D Results\Port"+S1+"_SP\Measured_SP_"+time_suffix)
	ore.AddToTree("1D Results\Port"+S1+"_RE\Measured_RE_"+time_suffix)
	ote.AddToTree("1D Results\Port"+S1+"_TE\Measured_TE_"+time_suffix)

	ElseIf nResults<2 Then
		'Nothing to do

	End If



	theEnd:
	MsgBox "Operation finished", vbInformation, "Done"


End Sub

Function Mag_Linear2dB(O As Object) As Object
	Dim n As Integer, logFactor As Integer, i As Integer
	With O
		n = .getN
		logFactor = .GetLogarithmicFactor
	End With

	Set O = O.magnitude

	Dim ox As Object
	Set ox = Result1D("")

	For i=0 To n-1
		ox.AppendXY(O.GetX(i), logFactor*CST_Log10(O.GetY(i)))
	Next


	Set Mag_Linear2dB = ox

End Function
Function CST_Log10(value As Double) As Double
	If value <= 0 Then
		CST_Log10 = lib_rundef
	Else
		CST_Log10 = Log(value)/Log(10)
	End If
End Function
