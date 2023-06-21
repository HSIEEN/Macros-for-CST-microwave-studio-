' Calculate the average efficiency according to the user input
'Option Explicit
'20220828-By Shawn in COROS

Public effiPath As String
Public childItem As String
'Public FirstChildItem As String

Sub Main ()
  'Efficiency results parent path
	Dim ii As Integer
	Dim portArray(100) As String
   effiPath = "1D Results\Efficiencies"
   childItem = Resulttree.GetFirstChildName(effiPath)
   If childItem = "" Then
   	 MsgBox("No Efficiency results found!",vbCritical,"Warning")
   	 Exit All
   End If

    'effiPath = "1D Results\Efficiencies"
    'childItem = Resulttree.GetFirstChildName(effiPath)
	ii = 1
    While childItem <> ""
   	 portArray(ii-1) = Mid(childItem,InStr(childItem,"[")+1,InStr(childItem,"]")-InStr(childItem,"[")-1)
   	 ii = ii+1
   	 childItem = Resulttree.GetNextItemName(childItem)
    Wend
    'InformationStr = "eg L1 L5 W2 W5"
	Begin Dialog UserDialog 250,140,"Average Efficiency ",.DialogFunction ' %GRID:10,7,1,1
		GroupBox 10,42,220,63,"Band selection:",.GroupBox2
		GroupBox 10,0,220,42,"Port selection:",.GroupBox1
		TextBox 80,77,80,14,.Band
		OKButton 40,112,90,21
		CancelButton 140,112,90,21
		DropListBox 80,14,90,14,portArray(),.portNum
		Text 30,56,130,14,"eg: L1 L5 W2 W5",.Text1
	End Dialog
	Dim dlg As UserDialog

	dlg.Band = "L1"
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

			Dim Bands As String
		    Dim PortNum As String
		    Dim CurrentItem As String
		    Bands = DlgText("Band")
		    PortNum = DlgText("PortNum")

            CurrentItem = Resulttree.GetFirstChildName(effiPath)
            'Dim TempStr As String, m As Integer, n As Integer
            'm = InStr(CurrentItem,"[")
            'n = InStr(CurrentItem,"]")
            'TempStr = Mid$(CurrentItem,InStr(CurrentItem,"[")+1,InStr(CurrentItem,"]")-InStr(CurrentItem,"[")-1)
            While CurrentItem <> ""

            	If  Mid$(CurrentItem,InStr(CurrentItem,"[")+1,InStr(CurrentItem,"]")-InStr(CurrentItem,"[")-1) =  PortNum Then
            		Dim EffiType As String, FileName As String
            		Dim nPoints As Long, n As Integer, Num As Integer, Ysum As Double, Avg As Double, dBAvg As Double, X As Double, Y As Double
            		Dim O As Object
            		EffiType = Mid(CurrentItem,Len(effiPath)+2,InStr(CurrentItem,"[")-Len(effiPath)-2)
            		FileName = Resulttree.GetFileFromTreeItem(CurrentItem)
            		Set O = Result1DComplex(FileName)
            		nPoints = O.GetN
            		Ysum = 0
            		Num = 0
            		'Bands
            		If InStr(Bands,"L1")<> 0 Then
            			For n = 0 To nPoints-1
            				X = O.GetX(n)
            				If X < 1.61 And X > 1.56 Then
            					Num = Num+1
            					Y = O.GetYRe(n)
            					Ysum = Ysum + Y
            				End If

            			Next
            			Avg = Ysum/Num
            			dBAvg = Log(Avg)/Log(10)*10
            			ReportInformationToWindow( _
       					 EffiType+"@Port"+PortNum+" Over Band GNSS L1 is: "+Left(Cstr(Avg*100),InStr(Cstr(Avg*100),".")+2)+"% ("+Left(Cstr(dBAvg),InStr(Cstr(dBAvg),".")+2)+ "dB)")
            		End If
            		If InStr(Bands,"L5")<> 0 Then
            			For n = 0 To nPoints-1
            				X = O.GetX(n)
            				If X < 1.21 And X > 1.16 Then
            					Num = Num+1
            					Y = O.GetYRe(n)
            					Ysum = Ysum + Y
            				End If

            			Next
            			Avg = Ysum/Num
            			dBAvg = Log(Avg)/Log(10)*10
            			ReportInformationToWindow( _
       					 EffiType+"@Port"+PortNum+" Over Band GNSS L5 is: "+Left(Cstr(Avg*100),InStr(Cstr(Avg*100),".")+2)+"% ("+Left(Cstr(dBAvg),InStr(Cstr(dBAvg),".")+2)+ "dB)")
            		End If
            		If InStr(Bands,"W2")<> 0 Then
            			For n = 0 To nPoints-1
            				X = O.GetX(n)
            				If X < 2.5 And X > 2.4 Then
            					Num = Num+1
            					Y = O.GetYRe(n)
            					Ysum = Ysum + Y
            				End If

            			Next
            			Avg = Ysum/Num
            			dBAvg = Log(Avg)/Log(10)*10
            			ReportInformationToWindow( _
       					 EffiType+"@Port"+PortNum+" Over Band Wifi 2.4GHz is: "+Left(Cstr(Avg*100),InStr(Cstr(Avg*100),".")+2)+"% ("+Left(Cstr(dBAvg),InStr(Cstr(dBAvg),".")+2)+ "dB)")
            		End If
            		If InStr(Bands,"W5")<> 0 Then
            			For n = 0 To nPoints-1
            				X = O.GetX(n)
            				If X < 5.85 And X > 5.15 Then
            					Num = Num+1
            					Y = O.GetYRe(n)
            					Ysum = Ysum + Y
            				End If

            			Next
            			Avg = Ysum/Num
            			dBAvg = Log(Avg)/Log(10)*10
            			ReportInformationToWindow( _
       					 EffiType+"@Port"+PortNum+" Over Band Wifi 5GHz is: "+Left(Cstr(Avg*100),InStr(Cstr(Avg*100),".")+2)+"% ("+Left(Cstr(dBAvg),InStr(Cstr(dBAvg),".")+2)+ "dB)")
            		End If

            	End If
            	CurrentItem = Resulttree.GetNextItemName(CurrentItem)
            Wend


		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function

