' Calculate Power loss in dB
'Option Explicit
'20220828-By Shawn in COROS

Public PowerPath As String
'Public FirstChildItem As String

Sub Main ()
  'Efficiency results parent path

   PowerPath = "1D Results\Power"
   Dim portArray(100) As String
   Dim ii As Integer
   Dim childItem As String

   childItem = Resulttree.GetFirstChildName(PowerPath)
	ii = 1
   While childItem <> ""
   	portArray(ii-1) = Mid(childItem,InStr(childItem,"[")+1,InStr(childItem,"]")-InStr(childItem,"[")-1)
   	ii = ii+1
   	childItem = Resulttree.GetNextItemName(childItem)
   Wend

   'For ii = 1 To Port.StartPortNumberIteration
   		'portArray(ii-1) = CStr(Port.GetNextPortNumber)
   'Next

    Dim informationStr As String
    'informationStr = "请输入要计算的频段（eg L1 L5 W2 W5）："
	Begin Dialog UserDialog 180,91,"效率损失分析" ' %GRID:10,7,1,1
		OKButton 10,63,80,21
		CancelButton 110,63,70,21
		GroupBox 10,7,160,49,"Please select a port",.GroupBox1
		DropListBox 50,28,70,14,portArray(),.portNumber
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If
    Dim portNumber As String
    'Dim CurrentItem As String
    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
    portNumber = portArray(dlg.portNumber)

    PowerPath = PowerPath + "\Excitation [" + portNumber + "]"
    nResults = Resulttree.GetTreeResults(PowerPath,"0D/1D recursive","",paths,types,files,info)

    'data record
    Dim n As Long
    'metal material list
    Dim metalList As String
     'dielectric  material list
    Dim dielectricList As String
    'Number of metal materials
    Dim metalNumber As Integer
    'Number of dielectric materials
    Dim dielectricNumber As Integer
    'metal loss data
    Dim metalLoss(100,1000) As Double
    'dielectric loss data
    Dim dielectricLoss(100,1000) As Double
    metalList = ""
    dielectricList = ""
    metalNumber = 0
    dielectricNumber = 0

    Dim filename As String
    'Dim DSfilename As String
    'Dim Path As String
    'Path ="1D Results\S-Parameters\S"+portNumber+","+portNumber
	filename = Resulttree.GetFileFromTreeItem("1D Results\S-Parameters\S"+portNumber+","+portNumber)
	'DSfilename = DSResultTree.GetFileFromTreeItem("Tasks\SPara1\S-Parameters\S")

    For n = 0 To nResults-1
    	Dim m As Long
		'Record power Stimulated
		If InStr(paths(n),"Power Stimulated") <> 0 Then
    		'Dim EffiType As String, FileName As String
    		Dim nPoints As Long, losPoints As Long,radPoints As Long, X() As Double, pStimulate() As Double
    		'power stimulated object
    		Dim opStimulatem As Object

    		Set opStimulatem = Result1DComplex(files(n))
    		nPoints = opStimulatem.GetN
    		ReDim X(nPoints) As Double
    		ReDim pStimulate(nPoints) As Double
			For m = 0 To nPoints-1
				X(m) = opStimulatem.GetX(m)
				pStimulate(m) = opStimulatem.GetYRe(m)
    		Next
    	'get Coupling power
    	ElseIf InStr(paths(n),"Power Outgoing all Ports") <> 0 Then
    		'Dim EffiType As String, FileName As String
            Dim pCoupling() As Double
    		Dim oCoupling As Object
    		Set oCoupling = Result1DComplex(files(n))
    		nPoints = oCoupling.GetN
    		ReDim pCoupling(nPoints) As Double
			For m = 0 To nPoints-1
				'pCoupling(m) = oCoupling.GetYRe(m) - Ypref(m)
				pCoupling(m) = oCoupling.GetYRe(m)
    		Next
    	'Record Accepted power
    	ElseIf  (filename <> "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted") _
    	Or (filename = "" And Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")) = "Power Accepted (DS)") Then
        		'Dim Xacp() As Double
                Dim pAccept() As Double
                Dim xAccept() As Double
        		Dim opAccept As Object
        		Dim acpPoints As Long
        		Set opAccept = Result1DComplex(files(n))
        		acpPoints = opAccept.GetN
        		ReDim pAccept(acpPoints) As Double
        		ReDim xAccept(acpPoints)
    			For m = 0 To acpPoints-1
    				pAccept(m) = opAccept.GetYRe(m)
    				'Xrad(m) = Oprad.GetX(m)
        		Next


    	'Record metal total loss
		ElseIf InStr(paths(n),"Loss in Metals") <> 0 Then
    		'Dim EffiType As String, FileName As String
            Dim lossOfMetal() As Double
            Dim xMetal() As Double
    		Dim oPmetalLoss As Object
    		Set oPmetalLoss = Result1DComplex(files(n))
    		losPoints = oPmetalLoss.GetN
    		ReDim lossOfMetal(losPoints) As Double
    		ReDim xMetal(losPoints) As Double
			For m = 0 To losPoints-1
				lossOfMetal(m) = oPmetalLoss.GetYRe(m)
				xMetal(m) = oPmetalLoss.GetX(m)
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
    		metalList = metalList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-14)+"$"
            'Dim Ymtloss() As Double
            Dim ometalLoss As Object
            Set ometalLoss = Result1DComplex(files(n))
            Dim MetPoints As Long
            MetPoints = ometalLoss.GetN
            For m = 0 To MetPoints-1
            	metalLoss(metalNumber,m) = ometalLoss.GetYRe(m)
            Next
    		metalNumber = metalNumber+1
		'Record loss of per dielectric
    	ElseIf InStr(paths(n),"Volume loss") <> 0 Then
    		dielectricList = dielectricList + Right(paths(n),Len(paths(n))-InStrRev(paths(n),"\")-15)+"$"
            'Dim Ydlloss() As Double
            Dim odielectricLoss As Object
            Set odielectricLoss = Result1DComplex(files(n))
            Dim dielectricPoints As Long
            dielectricPoints = odielectricLoss.GetN
            For m = 0 To dielectricPoints-1
            	dielectricLoss(dielectricNumber,m) = odielectricLoss.GetYRe(m)
            Next
    		dielectricNumber = dielectricNumber+1
    	End If
    Next
     'Calculate reflected power at feeding port
    If filename <> "" Then
		Dim opReflct As Object
		Set opReflct = Result1DComplex(filename)
		Dim YRe() As Double, YIm() As Double, pReflct() As Double
		ReDim YRe(nPoints) As Double
		ReDim YIm(nPoints) As Double
		ReDim pReflct(nPoints) As Double
		'ReDim Yref(nPoints) As Double
		For n = 0 To nPoints-1
			YRe(n) = opReflct.GetYRe(n)
			YIm(n) = opReflct.GetYIm(n)
			pReflct(n) = (YRe(n)^2+YIm(n)^2)*pStimulate(n)

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
	ReDim oPlotMaterialLoss(metalNumber) As Object
	For n = 0 To metalNumber-1
        Set oPlotMaterialLoss(n) = Result1D("")
		For m = 1 To nPoints-1
			Dim i As Integer
			For i = 0 To MetPoints-1
				If xMetal(i) <= X(m) And xMetal(i)> (X(m-1)+X(m))/2 Then
        			oPlotMaterialLoss(n).AppendXY(X(m),Log((pAccept(m)-metalLoss(n,i))/pAccept(m))/Log(10)*10)
        			'oPlotMaterialLoss(n).AppendXY(X(m),Log(metalLoss(n,i)/pAccept(m))/Log(10)*10)
        			Exit For
        		ElseIf xMetal(i) >= X(m-1) And xMetal(i)< (X(m-1)+X(m))/2 Then
                    oPlotMaterialLoss(n).AppendXY(X(m-1),Log((pAccept(m-1)-metalLoss(n,i))/pAccept(m-1))/Log(10)*10)
                    'oPlotMaterialLoss(n).AppendXY(X(m-1),Log(metalLoss(n,i)/pAccept(m-1))/Log(10)*10)
                    Exit For
				End If

			Next
		Next
		oPlotMaterialLoss(n).Xlabel("Frequecy/GHz")
		'oPlotMaterialLoss(n).ylabel("Loss in "+ Left(metalList,InStr(metalList,"$")-1)+"/dB" )
		oPlotMaterialLoss(n).Title("Loss in "+ Left(metalList,InStr(metalList,"$")-1)+"/dB")
		oPlotMaterialLoss(n).Ylabel("dB" )
		oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(metalList,InStr(metalList,"$")-1)+ "@Port="+portNumber+".sig")
		oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation efficiency loss due to metal loss\Loss in "+Left(metalList,InStr(metalList,"$")-1))
		metalList = Right(metalList,Len(metalList)-InStr(metalList,"$"))
	Next

	'Plot dielectric loss
	ReDim oPlotMaterialLoss(dielectricNumber) As Object
	For n = 0 To dielectricNumber-1
        Set oPlotMaterialLoss(n) = Result1D("")
		For m = 1 To nPoints-1

			For i = 0 To MetPoints-1
				If xMetal(i) <= X(m) And xMetal(i)> (X(m-1)+X(m))/2 Then
        			oPlotMaterialLoss(n).AppendXY(X(m),Log((pAccept(m)-dielectricLoss(n,i))/pAccept(m))/Log(10)*10)
        			'oPlotMaterialLoss(n).AppendXY(X(m),Log((pAccept(m)-dielectricLoss(n,i))/pAccept(m))/Log(10)*10)
        			Exit For
        		ElseIf xMetal(i) >= X(m-1) And xMetal(i)< (X(m-1)+X(m))/2 Then
                    oPlotMaterialLoss(n).AppendXY(X(m-1),Log((pAccept(m-1)-dielectricLoss(n,i))/pAccept(m-1))/Log(10)*10)
                    Exit For
				End If

			Next
		Next
		oPlotMaterialLoss(n).xlabel("Frequecy/GHz")
		'oPlotMaterialLoss(n).ylabel("Loss in "+ Left(dielectricList,InStr(dielectricList,"$")-1)+"/dB" )
		oPlotMaterialLoss(n).Title("Loss in "+ Left(dielectricList,InStr(dielectricList,"$")-1)+"/dB" )
		oPlotMaterialLoss(n).ylabel("dB" )
		'Dim temp As String
		'temp = Left(dielectricList,InStr(dielectricList,"$")-1)
		oPlotMaterialLoss(n).Save("RadiationEfficiencyLossIn"+Left(dielectricList,InStr(dielectricList,"$")-1)+ "@Port="+portNumber+".sig")
		oPlotMaterialLoss(n).AddToTree(PowerPath+"\Radiation efficiency loss due to dielectric loss\Loss in "+Left(dielectricList,InStr(dielectricList,"$")-1))
		dielectricList = Right(dielectricList,Len(dielectricList)-InStr(dielectricList,"$"))
	Next

    If filename <> "" Then
        'Plot reflection and coupling loss as well as total metal and dielectric loss
		For n = 0 To nPoints-1
			If (pStimulate(n)-pReflct(n))<=0 Then
				pStimulate(n) = pReflct(n) + 1e-3
			End If
			oPlotRefLoss.AppendXY(X(n),Log((pStimulate(n)-pReflct(n))/pStimulate(n))/Log(10)*10)
			oPlotCouLoss.AppendXY(X(n),Log((pStimulate(n)-pCoupling(n)+pReflct(n))/pStimulate(n))/Log(10)*10)
            If X(n) >= xMetal(0) Then
            	For m = 0 To losPoints-1
            		If xMetal(m) <= X(n) And xMetal(m)> (X(n-1)+X(n))/2 Then
            			oPlotMetLoss.AppendXY(X(n),Log((pStimulate(n)-lossOfMetal(m))/pStimulate(n))/Log(10)*10)
						oPlotDieLoss.AppendXY(X(n),Log((pStimulate(n)-lossOfDielectric(m))/pStimulate(n))/Log(10)*10)
						Exit For
					ElseIf xMetal(m) >= X(n-1) And xMetal(m)< (X(n-1)+X(n))/2 Then
						oPlotMetLoss.AppendXY(X(n-1),Log((pStimulate(n-1)-lossOfMetal(m))/pStimulate(n-1))/Log(10)*10)
						oPlotDieLoss.AppendXY(X(n-1),Log((pStimulate(n-1)-lossOfDielectric(m))/pStimulate(n-1))/Log(10)*10)
						Exit For
					End If
           		 Next
           	End If
		Next

		oPlotRefLoss.Title("Total efficiency Loss due to reflection/dB")
	    oPlotCouLoss.Title("Total efficiency Loss due to coupling/dB")
	    oPlotMetLoss.Title("Total efficiency Loss due to Metal Loss/dB")
	    oPlotDieLoss.Title("Total efficiency Loss due to Dielectric Loss/dB")

	    oPlotRefLoss.ylabel("dB")
	    oPlotCouLoss.ylabel("dB")
	    oPlotMetLoss.ylabel("dB")
	    oPlotDieLoss.ylabel("dB")

	    oPlotRefLoss.xlabel("Frequency/GHz")
	    oPlotCouLoss.xlabel("Frequency/GHz")
	    oPlotMetLoss.xlabel("Frequency/GHz")
	    oPlotDieLoss.xlabel("Frequency/GHz")

		oPlotRefLoss.Save("TotalEfficiencyLossDueToReflection @Port="+portNumber+".sig")
		oPlotCouLoss.Save("TotoalEfficiencyLossDueToCoupling @Port="+portNumber+".sig")
		oPlotMetLoss.Save("TotoalEfficiencyLossDueTometalLoss @Port="+portNumber+".sig")
		oPlotDieLoss.Save("TotoalEfficiencyLossDueTodielectricLoss @Port="+portNumber+".sig")

		oPlotRefLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Reflection")
		oPlotCouLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Coupling")
		oPlotMetLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Metal Loss")
		oPlotDieLoss.AddToTree(PowerPath+"\Total Efficiency Loss\Loss due to Dielectric Loss")
	End If
    'Change Plot Styles
	Dim selectedItem As String
	Dim curveLabel As String
	Dim index As Integer

	selectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation efficiency loss due to dielectric loss")
	While selectedItem <> ""
		SelectTreeItem(selectedItem)
        curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))

	      With Plot1D

		      index =.GetCurveIndexOfcurveLabel(curveLabel)

		     .SetLineStyle(index,"Solid",2) ' thick dashed line

		     '.SetLineColor(index,255,255,0)  ' yellow

		     .Plot ' make changes visible

		End With


		selectedItem = Resulttree.GetNextItemName(selectedItem)
	Wend

	selectedItem = Resulttree.GetFirstChildName(PowerPath+"\Radiation efficiency loss due to metal loss")
	While selectedItem <> ""
		SelectTreeItem(selectedItem)
        curveLabel = Right(selectedItem,Len(selectedItem)-InStrRev(selectedItem,"\"))
        With Plot1D

		      index =.GetCurveIndexOfcurveLabel(curveLabel)

		     .SetLineStyle(index,"Solid",2) ' thick dashed line

		     '.SetLineColor(index,255,255,0)  ' yellow

		     .Plot ' make changes visible

		End With

		selectedItem = Resulttree.GetNextItemName(selectedItem)
	Wend

End Sub


