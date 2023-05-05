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
	Begin Dialog UserDialog 410,63,"效率损失分析" ' %GRID:10,7,1,1
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
    Dim PortNum As String
    'Dim CurrentItem As String
    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
    PortNum = dlg.portNum

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
    		Dim oPstim As Object

    		Set oPstim = Result1DComplex(files(n))
    		nPoints = oPstim.GetN
    		ReDim X(nPoints) As Double
    		ReDim Psti(nPoints) As Double
			For m = 0 To nPoints-1
				X(m) = oPstim.GetX(m)
				Psti(m) = oPstim.GetYRe(m)
    		Next
    	'get Coupling power
    	ElseIf InStr(paths(n),"Power Outgoing all Ports") <> 0 Then
    		'Dim EffiType As String, FileName As String
            Dim Pcoup() As Double
    		Dim oPcoup As Object
    		Set oPcoup = Result1DComplex(files(n))
    		nPoints = oPcoup.GetN
    		ReDim Pcoup(nPoints) As Double
			For m = 0 To nPoints-1
				'Pcoup(m) = oPcoup.GetYRe(m) - Ypref(m)
				Pcoup(m) = oPcoup.GetYRe(m)
    		Next
    	'Record Accepted power
    	ElseIf  (filename <> "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted") _
    	Or (filename = "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted (DS)") Then
        		'Dim Xacp() As Double
                Dim Pacc() As Double
                Dim Xacp() As Double
        		Dim oPacc As Object
        		Dim acpPoints As Long
        		Set oPacc = Result1DComplex(files(n))
        		acpPoints = oPacc.GetN
        		ReDim Pacc(acpPoints) As Double
        		ReDim Xacp(acpPoints)
    			For m = 0 To acpPoints-1
    				Pacc(m) = oPacc.GetYRe(m)
    				'Xrad(m) = Oprad.GetX(m)
        		Next


    	'Record metal total loss
		ElseIf InStr(paths(n),"Loss in Metals") <> 0 Then
    		'Dim EffiType As String, FileName As String
            Dim lossOfMetal() As Double
            Dim Xmtl() As Double
    		Dim oPmetalLoss As Object
    		Set oPmetalLoss = Result1DComplex(files(n))
    		losPoints = oPmetalLoss.GetN
    		ReDim lossOfMetal(losPoints) As Double
    		ReDim Xmtl(losPoints) As Double
			For m = 0 To losPoints-1
				lossOfMetal(m) = oPmetalLoss.GetYRe(m)
				Xmtl(m) = oPmetalLoss.GetX(m)
    		Next

    	'Record dielectric total loss
		ElseIf InStr(paths(n),"Loss in Dielectrics") <> 0 Then
    		'Dim EffiType As String, FileName As String
            Dim lossOfDielectric() As Double
    		Dim oPdielectricLoss As Object
    		Set oPdielectricLoss = Result1DComplex(files(n))
    		losPoints = oPdielectricLoss.GetN
    		ReDim lossOfDielectric(losPoints) As Double
			For m = 0 To losPoints-1
				lossOfDielectric(m) = oPdielectricLoss.GetYRe(m)
    		Next
    	'Record loss of per metal
    	ElseIf InStr(paths(n),"Metal loss") <> 0 Then
    		MetalList = MetalList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-14)+"$"
            'Dim Ymtloss() As Double
            Dim oMetalLoss As Object
            Set oMetalLoss = Result1DComplex(files(n))
            Dim MetPoints As Long
            MetPoints = oMetalLoss.GetN
            For m = 0 To MetPoints-1
            	MetalLoss(MetalNum,m) = oMetalLoss.GetYRe(m)
            Next
    		MetalNum = MetalNum+1
		'Record loss of per dielectric
    	ElseIf InStr(paths(n),"Volume loss") <> 0 Then
    		DielectricList = DielectricList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-15)+"$"
            'Dim Ydlloss() As Double
            Dim oDielectricLoss As Object
            Set oDielectricLoss = Result1DComplex(files(n))
            Dim DiePoints As Long
            DiePoints = oDielectricLoss.GetN
            For m = 0 To DiePoints-1
            	DielectricLoss(DielectricNum,m) = oDielectricLoss.GetYRe(m)
            Next
    		DielectricNum = DielectricNum+1
    	End If
    Next
     'Calculate reflected power at feeding port
    If filename <> "" Then
		Dim oPref As Object
		Set oPref = Result1DComplex(filename)
		Dim YRe() As Double, YIm() As Double, Pref() As Double
		ReDim YRe(nPoints) As Double
		ReDim YIm(nPoints) As Double
		ReDim Pref(nPoints) As Double
		'ReDim Yref(nPoints) As Double
		For n = 0 To nPoints-1
			YRe(n) = oPref.GetYRe(n)
			YIm(n) = oPref.GetYIm(n)
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
        			oPlotMaterialLoss(n).AppendXY(X(m),Log((Pacc(m)-MetalLoss(n,i))/Pacc(m))/Log(10)*10)
        		ElseIf Xmtl(i) >= X(m-1) And Xmtl(i)< (X(m-1)+X(m))/2 Then
                    oPlotMaterialLoss(n).AppendXY(X(m-1),Log((Pacc(m-1)-MetalLoss(n,i))/Pacc(m-1))/Log(10)*10)
				End If

			Next
		Next
		oPlotMaterialLoss(n).xlabel("Frequecy/GHz")
		oPlotMaterialLoss(n).ylabel("Loss in "+ Left(MetalList,InStr(MetalList,"$")-1)+"/dB" )
		oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(MetalList,InStr(MetalList,"$")-1)+ "@Port="+PortNum+".sig")
		oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation efficiency loss due to metal loss\Loss in "+Left(MetalList,InStr(MetalList,"$")-1))
		MetalList = Right(MetalList,Len(MetalList)-InStr(MetalList,"$"))
	Next

	'Plot dielectric loss
	ReDim oPlotMaterialLoss(DielectricNum) As Object
	For n = 0 To DielectricNum-1
        Set oPlotMaterialLoss(n) = Result1D("")
		For m = 1 To nPoints-1

			For i = 0 To MetPoints-1
				If Xmtl(i) <= X(m) And Xmtl(i)> X(m-1) Then
        			oPlotMaterialLoss(n).AppendXY(X(m),Log((Pacc(m)-DielectricLoss(n,i))/Pacc(m))/Log(10)*10)
        		ElseIf Xmtl(i) >= X(m-1) And Xmtl(i)< (X(m-1)+X(m))/2 Then
                    oPlotMaterialLoss(n).AppendXY(X(m-1),Log((Pacc(m-1)-DielectricLoss(n,i))/Pacc(m-1))/Log(10)*10)
				End If

			Next
		Next
		oPlotMaterialLoss(n).xlabel("Frequecy/GHz")
		oPlotMaterialLoss(n).ylabel("Loss in "+ Left(DielectricList,InStr(DielectricList,"$")-1)+"/dB" )
		oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(DielectricList,InStr(DielectricList,"$")-1)+ "@Port="+PortNum+".sig")
		oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation efficiency loss due to dielectric loss\Loss in "+Left(DielectricList,InStr(DielectricList,"$")-1))
		DielectricList = Right(DielectricList,Len(DielectricList)-InStr(DielectricList,"$"))
	Next

    If filename <> "" Then
        'Plot reflection and coupling loss as well as total metal and dielectric loss
		For n = 0 To nPoints-1
			If (Psti(n)-Pref(n))<=0 Then
				Psti(n) = Pref(n) + 1e-3

			End If
			oPlotRefLoss.AppendXY(X(n),Log((Psti(n)-Pref(n))/Psti(n))/Log(10)*10)
			oPlotCouLoss.AppendXY(X(n),Log((Psti(n)-Pcoup(n)+Pref(n))/Psti(n))/Log(10)*10)
            If X(n) >= Xmtl(0) Then
            	For m = 0 To losPoints-1
            		If Xmtl(m) <= X(n) And Xmtl(m)> (X(n-1)+X(n))/2 Then
            			oPlotMetLoss.AppendXY(X(n),Log((Psti(n)-lossOfMetal(m))/Psti(n))/Log(10)*10)
						oPlotDieLoss.AppendXY(X(n),Log((Psti(n)-lossOfDielectric(m))/Psti(n))/Log(10)*10)
					ElseIf Xmtl(m) >= X(n-1) And Xmtl(m)< (X(n-1)+X(n))/2 Then
						oPlotMetLoss.AppendXY(X(n-1),Log((Psti(n-1)-lossOfMetal(m))/Psti(n-1))/Log(10)*10)
						oPlotDieLoss.AppendXY(X(n-1),Log((Psti(n-1)-lossOfDielectric(m))/Psti(n-1))/Log(10)*10)
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
	SelectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation efficiency loss due to dielectric loss")
	While SelectedItem <> ""
		SelectTreeItem(SelectedItem)
        CurveLabel = Right(SelectedItem,Len(SelectedItem)-InStrRev(SelectedItem,"\"))

	    index =Plot1D.GetCurveIndexOfCurveLabel(CurveLabel)

	     Plot1D.SetLineStyle(index,"Solid",2) ' thick dashed line

	     '.SetLineColor(index,255,255,0)  ' yellow

	     Plot1D.Plot ' make changes visible


		SelectedItem = Resulttree.GetNextItemName(SelectedItem)
	Wend
	SelectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation efficiency loss due to metal loss")
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

End Sub


