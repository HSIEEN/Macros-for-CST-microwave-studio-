' Calculate the average efficiency according to the user input
'Option Explicit
'20220828-By Shawn in COROS

Public EffiPath As String
Public FirstChildItem As String

Sub Main ()
  'Efficiency results parent path

   EffiPath = "1D Results\Efficiencies"
   FirstChildItem = Resulttree.GetFirstChildName(EffiPath)
   If FirstChildItem = "" Then
   	 MsgBox("No Efficiency results found!",vbCritical,"Warning")
   	 Exit All
   End If

    Dim InformationStr As String
    InformationStr = "请输入要计算的频段（eg L1 L5 W2 W5）："
	Begin Dialog UserDialog 410,98,"平均效率求算",.DialogFunction ' %GRID:10,7,1,1
		Text 10,28,340,14,InformationStr,.Text1
		TextBox 20,49,90,14,.Band
		OKButton 20,77,90,21
		CancelButton 120,77,90,21
		Text 10,7,110,14,"请选择端口：",.Text10
		TextBox 100,7,40,14,.PortNum
	End Dialog
	Dim dlg As UserDialog

	dlg.Band = "L1"
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

			Dim Bands As String
		    Dim PortNum As String
		    Dim CurrentItem As String
		    Bands = DlgText("Band")
		    PortNum = DlgText("PortNum")

            CurrentItem = FirstChildItem
            'Dim TempStr As String, m As Integer, n As Integer
            'm = InStr(CurrentItem,"[")
            'n = InStr(CurrentItem,"]")
            'TempStr = Mid$(CurrentItem,InStr(CurrentItem,"[")+1,InStr(CurrentItem,"]")-InStr(CurrentItem,"[")-1)
            While CurrentItem <> ""

            	If  Mid$(CurrentItem,InStr(CurrentItem,"[")+1,InStr(CurrentItem,"]")-InStr(CurrentItem,"[")-1) =  PortNum Then
            		Dim EffiType As String, FileName As String
            		Dim nPoints As Long, n As Integer, Num As Integer, Ysum As Double, Avg As Double, dBAvg As Double, X As Double, Y As Double
            		Dim O As Object
            		EffiType = Mid(CurrentItem,Len(EffiPath)+2,InStr(CurrentItem,"[")-Len(EffiPath)-2)
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

