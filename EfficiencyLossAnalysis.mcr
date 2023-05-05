' Calculate Power loss in dB
'Option Explicit
'20220828-By Shawn in COROS

Public PowerPath As String
'Public FirstChildItem As String

Sub Main ()
  'Efficiency results parent path

   PowerPath = "1D Results\Power"

    Dim InformationStr As String
    'InformationStr = "请输入要计算的频段（eg L1 L5 W2 W5）："
	Begin Dialog UserDialog 410,63,"效率损失分析",.DialogFunction ' %GRID:10,7,1,1
		OKButton 10,35,90,21
		CancelButton 110,35,90,21
		Text 10,7,110,14,"请选择端口：",.Text10
		TextBox 100,7,40,14,.PortNum
	End Dialog
	Dim dlg As UserDialog

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
		    Dim PortNum As String
		    'Dim CurrentItem As String
		    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
		    PortNum = DlgText("PortNum")

		    PowerPath = PowerPath + "\Excitation [" + PortNum + "]"
            nResults = Resulttree.GetTreeResults(PowerPath,"0D/1D recursive","",paths,types,files,info)

            'data record
            Dim n As Long
            'metal material list
            Dim MetalList As String
             'dielectric  material list
            Dim DielectricList As String
            'Number of metal materials
            Dim MetalNum As Integer
            'Number of dielectric materials
            Dim DielectricNum As Integer
            'metal loss data
            Dim MetalLoss(100,1000) As Double
            'dielectric loss data
            Dim DielectricLoss(100,1000) As Double
            MetalList = ""
            DielectricList = ""
            MetalNum = 0
            DielectricNum = 0

            Dim filename As String
            'Dim DSfilename As String
            'Dim Path As String
            'Path ="1D Results\S-Parameters\S"+PortNum+","+PortNum
			filename = Resulttree.GetFileFromTreeItem("1D Results\S-Parameters\S"+PortNum+","+PortNum)
			'DSfilename = DSResultTree.GetFileFromTreeItem("Tasks\SPara1\S-Parameters\S")

            For n = 0 To nResults-1
            	Dim m As Long
				'Record power Stimulated
        		If InStr(paths(n),"Power Stimulated") <> 0 Then
            		'Dim EffiType As String, FileName As String
            		Dim nPoints As Long, losPoints As Long,radPoints As Long, X() As Double, Psti() As Double
            		'power stimulated object
            		Dim oPStim As Object

            		Set oPStim = Result1DComplex(files(n))
            		nPoints = oPStim.GetN
            		ReDim X(nPoints) As Double
            		ReDim Psti(nPoints) As Double
        			For m = 0 To nPoints-1
        				X(m) = oPStim.GetX(m)
        				Psti(m) = oPStim.GetYRe(m)
            		Next
            	'get Coupling power
            	ElseIf InStr(paths(n),"Power Outgoing all Ports") <> 0 Then
            		'Dim EffiType As String, FileName As String
                    Dim Pcoulp() As Double
            		Dim oPCoup As Object
            		Set oPCoup = Result1DComplex(files(n))
            		nPoints = oPCoup.GetN
            		ReDim Pcoulp(nPoints) As Double
        			For m = 0 To nPoints-1
        				'Pcoulp(m) = oPCoup.GetYRe(m) - Ypref(m)
        				Pcoulp(m) = oPCoup.GetYRe(m)
            		Next
            	'Record Accepted power
            	ElseIf  (filename <> "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted") _
            	Or (filename = "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted (DS)") Then
	            		'Dim Xacp() As Double
	                    Dim Ypacp() As Double
	                    Dim Xacp() As Double
	            		Dim oPAcpt As Object
	            		Dim acpPoints As Long
	            		Set oPAcpt = Result1DComplex(files(n))
	            		acpPoints = oPAcpt.GetN
	            		ReDim Ypacp(acpPoints) As Double
	            		ReDim Xacp(acpPoints)
	        			For m = 0 To acpPoints-1
	        				Ypacp(m) = oPAcpt.GetYRe(m)
	        				'Xrad(m) = Oprad.GetX(m)
	            		Next


            	'Record metal loss
        		ElseIf InStr(paths(n),"Loss in Metals") <> 0 Then
            		'Dim EffiType As String, FileName As String
                    Dim Ypmtl() As Double
                    Dim Xmtl() As Double
            		Dim oPMtl As Object
            		Set oPMtl = Result1DComplex(files(n))
            		losPoints = oPMtl.GetN
            		ReDim Ypmtl(losPoints) As Double
            		ReDim Xmtl(losPoints) As Double
        			For m = 0 To losPoints-1
        				Ypmtl(m) = oPMtl.GetYRe(m)
        				Xmtl(m) = oPMtl.GetX(m)
            		Next

            	'Record dielectric loss
        		ElseIf InStr(paths(n),"Loss in Dielectrics") <> 0 Then
            		'Dim EffiType As String, FileName As String
                    Dim Ypdll() As Double
            		Dim oPDll As Object
            		Set oPDll = Result1DComplex(files(n))
            		losPoints = oPDll.GetN
            		ReDim Ypdll(losPoints) As Double
        			For m = 0 To losPoints-1
        				Ypdll(m) = oPDll.GetYRe(m)
            		Next
            	ElseIf InStr(paths(n),"Metal loss") <> 0 Then
            		MetalList = MetalList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-14)+"$"
                    Dim Ymtloss() As Double
                    Dim oMtLoss As Object
                    Set oMtLoss = Result1DComplex(files(n))
                    Dim MetPoints As Long
                    MetPoints = oMtLoss.GetN
                    For m = 0 To MetPoints-1
                    	MetalLoss(MetalNum,m) = oMtLoss.GetYRe(m)
                    Next
            		MetalNum = MetalNum+1

            	ElseIf InStr(paths(n),"Volume loss") <> 0 Then
            		DielectricList = DielectricList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-15)+"$"
                    Dim Ydlloss() As Double
                    Dim oDlLoss As Object
                    Set oDlLoss = Result1DComplex(files(n))
                    Dim DiePoints As Long
                    DiePoints = oDlLoss.GetN
                    For m = 0 To DiePoints-1
                    	DielectricLoss(DielectricNum,m) = oDlLoss.GetYRe(m)
                    Next
            		DielectricNum = DielectricNum+1
            	End If
            Next
             'Calculate reflected power at feeding port
            If filename <> "" Then
				Dim oPRef As Object
				Set oPRef = Result1DComplex(filename)
				Dim YRe() As Double, YIm() As Double, Pref() As Double
				ReDim YRe(nPoints) As Double
				ReDim YIm(nPoints) As Double
				ReDim Pref(nPoints) As Double
				'ReDim Yref(nPoints) As Double
				For n = 0 To nPoints-1
					YRe(n) = oPRef.GetYRe(n)
					YIm(n) = oPRef.GetYIm(n)
					Pref(n) = (YRe(n)^2+YIm(n)^2)*Psti(n)

				Next
			'Else 'This port is postprocessing port, no sparameters
			End If


    		'Plot
    		Dim oPlotRefLoss As Object
    		Dim oPlotCouLoss As Object
    		Dim oPlotMetLoss As Object
    		Dim oPlotDieLoss As Object
    		Dim oPlotMaterialLoss() As Object
    		Set oPlotRefLoss = Result1D("")
    		Set oPlotCouLoss = Result1D("")
    		Set oPlotMetLoss = Result1D("")
    		Set oPlotDieLoss = Result1D("")
    		'Set oPlotMaterialLoss = Result1D("")

    		'Plot metal loss
    		ReDim oPlotMaterialLoss(MetalNum) As Object
    		For n = 0 To MetalNum-1
                Set oPlotMaterialLoss(n) = Result1D("")
    			For m = 1 To nPoints-1
    				Dim i As Integer
    				For i = 0 To MetPoints-1
    					If Xmtl(i) <= X(m) And Xmtl(i)> (X(m-1)+X(m))/2 Then
                			oPlotMaterialLoss(n).AppendXY(X(m),Log((Ypacp(m)-MetalLoss(n,i))/Ypacp(m))/Log(10)*10)
                		ElseIf Xmtl(i) >= X(m-1) And Xmtl(i)< (X(m-1)+X(m))/2 Then
                            oPlotMaterialLoss(n).AppendXY(X(m-1),Log((Ypacp(m-1)-MetalLoss(n,i))/Ypacp(m-1))/Log(10)*10)
						End If

    				Next
    			Next
				oPlotMaterialLoss(n).xlabel("Frequecy/GHz")
				oPlotMaterialLoss(n).ylabel("Loss in "+ Left(MetalList,InStr(MetalList,"$")-1)+"/dB" )
				oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(MetalList,InStr(MetalList,"$")-1)+ "@Port="+PortNum+".sig")
				oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation Efficiency Metallic Loss\Loss in "+Left(MetalList,InStr(MetalList,"$")-1))
				MetalList = Right(MetalList,Len(MetalList)-InStr(MetalList,"$"))
    		Next

    		'Plot dielectric loss
    		ReDim oPlotMaterialLoss(DielectricNum) As Object
    		For n = 0 To DielectricNum-1
                Set oPlotMaterialLoss(n) = Result1D("")
    			For m = 1 To nPoints-1

    				For i = 0 To MetPoints-1
    					If Xmtl(i) <= X(m) And Xmtl(i)> X(m-1) Then
                			oPlotMaterialLoss(n).AppendXY(X(m),Log((Ypacp(m)-DielectricLoss(n,i))/Ypacp(m))/Log(10)*10)
                		ElseIf Xmtl(i) >= X(m-1) And Xmtl(i)< (X(m-1)+X(m))/2 Then
                            oPlotMaterialLoss(n).AppendXY(X(m-1),Log((Ypacp(m-1)-DielectricLoss(n,i))/Ypacp(m-1))/Log(10)*10)
						End If

    				Next
    			Next
				oPlotMaterialLoss(n).xlabel("Frequecy/GHz")
				oPlotMaterialLoss(n).ylabel("Loss in "+ Left(DielectricList,InStr(DielectricList,"$")-1)+"/dB" )
				oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(DielectricList,InStr(DielectricList,"$")-1)+ "@Port="+PortNum+".sig")
				oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation Efficiency Dielectric Loss\Loss in "+Left(DielectricList,InStr(DielectricList,"$")-1))
				DielectricList = Right(DielectricList,Len(DielectricList)-InStr(DielectricList,"$"))
    		Next

            If filename <> "" Then
	            'Plot reflection and coupling loss as well as total metal and dielectric loss
	    		For n = 0 To nPoints-1
	    			If (Psti(n)-Pref(n))<=0 Then
	    				Psti(n) = Pref(n) + 1e-3

	    			End If
					oPlotRefLoss.AppendXY(X(n),Log((Psti(n)-Pref(n))/Psti(n))/Log(10)*10)
					oPlotCouLoss.AppendXY(X(n),Log((Psti(n)-Pcoulp(n)+Pref(n))/Psti(n))/Log(10)*10)
	                If X(n) >= Xmtl(0) Then
	                	For m = 0 To losPoints-1
	                		If Xmtl(m) <= X(n) And Xmtl(m)> (X(n-1)+X(n))/2 Then
	                			oPlotMetLoss.AppendXY(X(n),Log((Psti(n)-Ypmtl(m))/Psti(n))/Log(10)*10)
								oPlotDieLoss.AppendXY(X(n),Log((Psti(n)-Ypdll(m))/Psti(n))/Log(10)*10)
							ElseIf Xmtl(m) >= X(n-1) And Xmtl(m)< (X(n-1)+X(n))/2 Then
								oPlotMetLoss.AppendXY(X(n-1),Log((Psti(n-1)-Ypmtl(m))/Psti(n-1))/Log(10)*10)
								oPlotDieLoss.AppendXY(X(n-1),Log((Psti(n-1)-Ypdll(m))/Psti(n-1))/Log(10)*10)
							End If
	               		 Next
	               	End If
	    		Next
			    oPlotRefLoss.ylabel("Total efficiency Loss due to reflection/dB")
			    oPlotCouLoss.ylabel("Total efficiency Loss due to coupling/dB")
			    oPlotMetLoss.ylabel("Total efficiency Loss due to Metal Loss/dB")
			    oPlotDieLoss.ylabel("Total efficiency Loss due to Dielectric Loss/dB")

			    oPlotRefLoss.xlabel("Frequency/GHz")
			    oPlotCouLoss.xlabel("Frequency/GHz")
			    oPlotMetLoss.xlabel("Frequency/GHz")
			    oPlotDieLoss.xlabel("Frequency/GHz")

				oPlotRefLoss.Save("TotalEfficiencyLossDueToReflection @Port="+PortNum+".sig")
				oPlotCouLoss.Save("TotoalEfficiencyLossDueToCoupling @Port="+PortNum+".sig")
				oPlotMetLoss.Save("TotoalEfficiencyLossDueToMetalLoss @Port="+PortNum+".sig")
				oPlotDieLoss.Save("TotoalEfficiencyLossDueToDielectricLoss @Port="+PortNum+".sig")

				oPlotRefLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Reflection")
				oPlotCouLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Coupling")
				oPlotMetLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Metal Loss")
				oPlotDieLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Dielectric Loss")
			End If
            'Change Plot Styles
			Dim SelectedItem As String
			Dim CurveLabel As String
			Dim index As Integer
			SelectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation Efficiency Dielectric Loss")
			While SelectedItem <> ""
				SelectTreeItem(SelectedItem)
                CurveLabel = Right(SelectedItem,Len(SelectedItem)-InStrRev(SelectedItem,"\"))

			    index =Plot1D.GetCurveIndexOfCurveLabel(CurveLabel)

			     Plot1D.SetLineStyle(index,"Solid",2) ' thick dashed line

			     '.SetLineColor(index,255,255,0)  ' yellow

			     Plot1D.Plot ' make changes visible


				SelectedItem = Resulttree.GetNextItemName(SelectedItem)
			Wend
			SelectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation Efficiency Metallic Loss")
			While SelectedItem <> ""
				SelectTreeItem(SelectedItem)
                CurveLabel = Right(SelectedItem,Len(SelectedItem)-InStrRev(SelectedItem,"\"))
                With Plot1D

				      index =.GetCurveIndexOfCurveLabel(CurveLabel)

				     .SetLineStyle(index,"Solid",2) ' thick dashed line

				     '.SetLineColor(index,255,255,0)  ' yellow

				     .Plot ' make changes visible

				End With

				SelectedItem = Resulttree.GetNextItemName(SelectedItem)
			Wend


		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function

