' Calculate Power loss in dB
' Option Explicit
' 20220828-By Shawn in COROS
' 20260408-Optimized (Conservative version)

Public PowerPath As String

Sub Main()
    ' Efficiency results parent path
    PowerPath = "1D Results\Power"

    ' Get port list
    Dim portArray(100) As String
    Dim ii As Integer
    Dim childItem As String

    childItem = ResultTree.GetFirstChildName(PowerPath)
    ii = 1
    While childItem <> ""
        portArray(ii - 1) = Mid(childItem, InStr(childItem, "[") + 1, InStr(childItem, "]") - InStr(childItem, "[") - 1)
        ii = ii + 1
        childItem = ResultTree.GetNextItemName(childItem)
    Wend

    ' Show port selection dialog
    Begin Dialog UserDialog 180, 91, "Ð§ÂÊËðÊ§·ÖÎö"
        OKButton 10, 63, 80, 21
        CancelButton 110, 63, 70, 21
        GroupBox 10, 7, 160, 49, "Please select a port", .GroupBox1
        DropListBox 50, 28, 70, 14, portArray(), .portNumber
    End Dialog
    Dim dlg As UserDialog

    If Dialog(dlg, -2) = 0 Then Exit All

    Dim portNumber As String
    portNumber = portArray(dlg.portNumber)
    PowerPath = PowerPath + "\Excitation [" + portNumber + "]"

    ' Get all results
    Dim paths As Variant, types As Variant, files As Variant, info As Variant, nResults As Long
    nResults = ResultTree.GetTreeResults(PowerPath, "0D/1D recursive", "", paths, types, files, info)

    ' Data containers
    Dim n As Long, m As Long
    Dim metalList As String, dielectricList As String
    Dim metalNumber As Integer, dielectricNumber As Integer
    Dim metalLoss(100, 1000) As Double
    Dim dielectricLoss(100, 1000) As Double
    metalList = ""
    dielectricList = ""
    metalNumber = 0
    dielectricNumber = 0

    ' Get S-parameter file
    Dim filename As String
    filename = ResultTree.GetFileFromTreeItem("1D Results\S-Parameters\S" + portNumber + "," + portNumber)

    ' Parse all results
    Dim nPoints As Long, losPoints As Long
    Dim X() As Double, pStimulate() As Double, pCoupling() As Double, pAccept() As Double
    Dim lossOfMetal() As Double, lossOfDielectric() As Double
    Dim xMetal() As Double

    For n = 0 To nResults - 1
        ' Power Stimulated
        If InStr(paths(n), "Power Stimulated") <> 0 Then
            Dim opStimulate As Object
            Set opStimulate = Result1DComplex(files(n))
            nPoints = opStimulate.GetN
            ReDim X(nPoints) As Double
            ReDim pStimulate(nPoints) As Double
            For m = 0 To nPoints - 1
                X(m) = opStimulate.GetX(m)
                pStimulate(m) = opStimulate.GetYRe(m)
            Next

        ' Power Outgoing all Ports
        ElseIf InStr(paths(n), "Power Outgoing all Ports") <> 0 Then
            Dim oCoupling As Object
            Set oCoupling = Result1DComplex(files(n))
            nPoints = oCoupling.GetN
            ReDim pCoupling(nPoints) As Double
            For m = 0 To nPoints - 1
                pCoupling(m) = oCoupling.GetYRe(m)
            Next

        ' Power Accepted
        ElseIf (filename <> "" And Right(paths(n), Len(paths(n)) - InStrRev(paths(n), "\")) = "Power Accepted") _
            Or (filename = "" And Right(paths(n), Len(paths(n)) - InStrRev(paths(n), "\")) = "Power Accepted (DS)") Then
            Dim opAccept As Object
            Set opAccept = Result1DComplex(files(n))
            Dim acpPoints As Long
            acpPoints = opAccept.GetN
            ReDim pAccept(acpPoints) As Double
            For m = 0 To acpPoints - 1
                pAccept(m) = opAccept.GetYRe(m)
            Next

        ' Loss in Metals (total)
        ElseIf InStr(paths(n), "Loss in Metals") <> 0 Then
            Dim oPmetalLoss As Object
            Set oPmetalLoss = Result1DComplex(files(n))
            losPoints = oPmetalLoss.GetN
            ReDim lossOfMetal(losPoints) As Double
            ReDim xMetal(losPoints) As Double
            For m = 0 To losPoints - 1
                lossOfMetal(m) = oPmetalLoss.GetYRe(m)
                xMetal(m) = oPmetalLoss.GetX(m)
            Next

        ' Loss in Dielectrics (total)
        ElseIf InStr(paths(n), "Loss in Dielectrics") <> 0 Then
            Dim oPdielectricLoss As Object
            Set oPdielectricLoss = Result1DComplex(files(n))
            losPoints = oPdielectricLoss.GetN
            ReDim lossOfDielectric(losPoints) As Double
            For m = 0 To losPoints - 1
                lossOfDielectric(m) = oPdielectricLoss.GetYRe(m)
            Next

        ' Metal loss (per material)
        ElseIf InStr(paths(n), "Metal loss") <> 0 Then
            metalList = metalList + Right(paths(n), Len(paths(n)) - InStrRev(paths(n), "\") - 14) + "$"
            Dim ometalLoss As Object
            Set ometalLoss = Result1DComplex(files(n))
            Dim MetPoints As Long
            MetPoints = ometalLoss.GetN
            For m = 0 To MetPoints - 1
                metalLoss(metalNumber, m) = ometalLoss.GetYRe(m)
            Next
            metalNumber = metalNumber + 1

        ' Volume loss (per dielectric)
        ElseIf InStr(paths(n), "Volume loss") <> 0 Then
            dielectricList = dielectricList + Right(paths(n), Len(paths(n)) - InStrRev(paths(n), "\") - 15) + "$"
            Dim odielectricLoss As Object
            Set odielectricLoss = Result1DComplex(files(n))
            Dim dielectricPoints As Long
            dielectricPoints = odielectricLoss.GetN
            For m = 0 To dielectricPoints - 1
                dielectricLoss(dielectricNumber, m) = odielectricLoss.GetYRe(m)
            Next
            dielectricNumber = dielectricNumber + 1
        End If
    Next

    ' Calculate reflected power at feeding port
    Dim pReflct() As Double
    If filename <> "" Then
        Dim opReflct As Object
        Set opReflct = Result1DComplex(filename)
        Dim YRe() As Double, YIm() As Double
        ReDim YRe(nPoints) As Double
        ReDim YIm(nPoints) As Double
        ReDim pReflct(nPoints) As Double
        For n = 0 To nPoints - 1
            YRe(n) = opReflct.GetYRe(n)
            YIm(n) = opReflct.GetYIm(n)
            pReflct(n) = (YRe(n) ^ 2 + YIm(n) ^ 2) * pStimulate(n)
        Next
    End If

    ' Plot metal loss (per material)
    Dim oPlotMaterialLoss() As Object
    ReDim oPlotMaterialLoss(metalNumber) As Object
    For n = 0 To metalNumber - 1
        Set oPlotMaterialLoss(n) = Result1D("")
        oPlotMaterialLoss(n).DeleteAt "rebuild"
        For m = 1 To nPoints - 1
            Dim i As Integer
            For i = 0 To MetPoints - 1
                If xMetal(i) <= X(m) And xMetal(i) > (X(m - 1) + X(m)) / 2 Then
                    oPlotMaterialLoss(n).AppendXY X(m), Log((Abs(pAccept(m) - metalLoss(n, i))) / Abs(pAccept(m))) / Log(10) * 10
                    Exit For
                ElseIf xMetal(i) >= X(m - 1) And xMetal(i) < (X(m - 1) + X(m)) / 2 Then
                    oPlotMaterialLoss(n).AppendXY X(m - 1), Log(Abs((pAccept(m - 1) - metalLoss(n, i))) / Abs(pAccept(m - 1))) / Log(10) * 10
                    Exit For
                End If
            Next
        Next
        oPlotMaterialLoss(n).Xlabel "Frequency/GHz"
        oPlotMaterialLoss(n).Title "Loss in " + Left(metalList, InStr(metalList, "$") - 1) + "/dB"
        oPlotMaterialLoss(n).Ylabel "dB"
        oPlotMaterialLoss(n).Save "RadiationEfficiencyLossIn" + Left(metalList, InStr(metalList, "$") - 1) + "@Port=" + portNumber + ".sig"
        oPlotMaterialLoss(n).AddToTree PowerPath + "\Radiation efficiency loss due to metals\Loss in " + Left(metalList, InStr(metalList, "$") - 1)
        metalList = Right(metalList, Len(metalList) - InStr(metalList, "$"))
    Next

    ' Plot dielectric loss (per material)
    If dielectricNumber > 0 Then
        ReDim oPlotMaterialLoss(dielectricNumber) As Object
        For n = 0 To dielectricNumber - 1
            Set oPlotMaterialLoss(n) = Result1D("")
            oPlotMaterialLoss(n).DeleteAt "rebuild"
            For m = 1 To nPoints - 1
                For i = 0 To MetPoints - 1
                    If xMetal(i) <= X(m) And xMetal(i) > (X(m - 1) + X(m)) / 2 Then
                        oPlotMaterialLoss(n).AppendXY X(m), Log(Abs(pAccept(m) - dielectricLoss(n, i)) / Abs(pAccept(m))) / Log(10) * 10
                        Exit For
                    ElseIf xMetal(i) >= X(m - 1) And xMetal(i) < (X(m - 1) + X(m)) / 2 Then
                        oPlotMaterialLoss(n).AppendXY X(m - 1), Log(Abs(pAccept(m - 1) - dielectricLoss(n, i)) / Abs(pAccept(m - 1))) / Log(10) * 10
                        Exit For
                    End If
                Next
            Next
            oPlotMaterialLoss(n).Xlabel "Frequency/GHz"
            oPlotMaterialLoss(n).Title "Loss in " + Left(dielectricList, InStr(dielectricList, "$") - 1) + "/dB"
            oPlotMaterialLoss(n).Ylabel "dB"
            oPlotMaterialLoss(n).Save "RadiationEfficiencyLossIn" + Left(dielectricList, InStr(dielectricList, "$") - 1) + "@Port=" + portNumber + ".sig"
            oPlotMaterialLoss(n).AddToTree PowerPath + "\Radiation efficiency loss due to dielectrics\Loss in " + Left(dielectricList, InStr(dielectricList, "$") - 1)
            dielectricList = Right(dielectricList, Len(dielectricList) - InStr(dielectricList, "$"))
        Next
    End If

    ' Plot total losses (reflection, coupling, metal, dielectric)
    If filename <> "" Then
        Dim oPlotRefLoss As Object, oPlotCouLoss As Object
        Dim oPlotMetLoss As Object, oPlotDieLoss As Object
        Set oPlotRefLoss = Result1D("")
        oPlotRefLoss.DeleteAt "rebuild"
        Set oPlotCouLoss = Result1D("")
        oPlotCouLoss.DeleteAt "rebuild"
        Set oPlotMetLoss = Result1D("")
        oPlotMetLoss.DeleteAt "rebuild"
        Set oPlotDieLoss = Result1D("")
        oPlotDieLoss.DeleteAt "rebuild"

        For n = 1 To nPoints - 1
            If (pStimulate(n) - pReflct(n)) <= 0 Then
                pStimulate(n) = pReflct(n) + 0.001
            End If
            oPlotRefLoss.AppendXY X(n), Log((pStimulate(n) - pReflct(n)) / pStimulate(n)) / Log(10) * 10
            oPlotCouLoss.AppendXY X(n), Log((pStimulate(n) - pCoupling(n) + pReflct(n)) / pStimulate(n)) / Log(10) * 10
            If X(n) >= xMetal(0) Then
                For m = 0 To losPoints - 1
                    If xMetal(m) <= X(n) And xMetal(m) > (X(n - 1) + X(n)) / 2 Then
                        oPlotMetLoss.AppendXY X(n), Log((pStimulate(n) - lossOfMetal(m)) / pStimulate(n)) / Log(10) * 10
                        Exit For
                    ElseIf xMetal(m) >= X(n - 1) And xMetal(m) < (X(n - 1) + X(n)) / 2 Then
                        oPlotMetLoss.AppendXY X(n - 1), Log((pStimulate(n - 1) - lossOfMetal(m)) / pStimulate(n - 1)) / Log(10) * 10
                        Exit For
                    End If
                Next
            End If
            If dielectricNumber > 0 Then
                If X(n) >= xMetal(0) Then
                    For m = 0 To losPoints - 1
                        If xMetal(m) <= X(n) And xMetal(m) > (X(n - 1) + X(n)) / 2 Then
                            oPlotDieLoss.AppendXY X(n), Log((pStimulate(n) - lossOfDielectric(m)) / pStimulate(n)) / Log(10) * 10
                            Exit For
                        ElseIf xMetal(m) >= X(n - 1) And xMetal(m) < (X(n - 1) + X(n)) / 2 And dielectricNumber > 0 Then
                            oPlotDieLoss.AppendXY X(n - 1), Log((pStimulate(n - 1) - lossOfDielectric(m)) / pStimulate(n - 1)) / Log(10) * 10
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        oPlotRefLoss.Title "Total efficiency Loss due to reflection/dB"
        oPlotCouLoss.Title "Total efficiency Loss due to coupling/dB"
        oPlotMetLoss.Title "Total efficiency Loss due to Metal Loss/dB"
        oPlotDieLoss.Title "Total efficiency Loss due to Dielectric Loss/dB"

        oPlotRefLoss.Ylabel "dB"
        oPlotCouLoss.Ylabel "dB"
        oPlotMetLoss.Ylabel "dB"
        oPlotDieLoss.Ylabel "dB"

        oPlotRefLoss.Xlabel "Frequency/GHz"
        oPlotCouLoss.Xlabel "Frequency/GHz"
        oPlotMetLoss.Xlabel "Frequency/GHz"
        oPlotDieLoss.Xlabel "Frequency/GHz"

        oPlotRefLoss.Save "TotalEfficiencyLossDueToReflection @Port=" + portNumber + ".sig"
        oPlotCouLoss.Save "TotalEfficiencyLossDueToCoupling @Port=" + portNumber + ".sig"
        oPlotMetLoss.Save "TotalEfficiencyLossDueToMetalLoss @Port=" + portNumber + ".sig"
        If dielectricNumber > 0 Then
            oPlotDieLoss.Save "TotalEfficiencyLossDueToDielectricLoss @Port=" + portNumber + ".sig"
        End If

        oPlotRefLoss.AddToTree PowerPath + "\Total Efficiency Loss\Loss due to Reflection"
        oPlotCouLoss.AddToTree PowerPath + "\Total Efficiency Loss\Loss due to Coupling"
        oPlotMetLoss.AddToTree PowerPath + "\Total Efficiency Loss\Loss due to Metals"
        If dielectricNumber > 0 Then
            oPlotDieLoss.AddToTree PowerPath + "\Total Efficiency Loss\Loss due to Dielectrics"
        End If
    End If

    ' Change Plot Styles
    Dim selectedItem As String
    Dim curveLabel As String
    Dim index As Integer

    selectedItem = ResultTree.GetFirstChildName(PowerPath + "\Radiation efficiency loss due to dielectrics")
    While selectedItem <> ""
        SelectTreeItem selectedItem
        curveLabel = Right(selectedItem, Len(selectedItem) - InStrRev(selectedItem, "\"))
        With Plot1D
            index = .GetCurveIndexOfcurveLabel(curveLabel)
            .SetLineStyle index, "Solid", 2
            .Plot
        End With
        selectedItem = ResultTree.GetNextItemName(selectedItem)
    Wend

    selectedItem = ResultTree.GetFirstChildName(PowerPath + "\Radiation efficiency loss due to metals")
    While selectedItem <> ""
        SelectTreeItem selectedItem
        curveLabel = Right(selectedItem, Len(selectedItem) - InStrRev(selectedItem, "\"))
        With Plot1D
            index = .GetCurveIndexOfcurveLabel(curveLabel)
            .SetLineStyle index, "Solid", 2
            .Plot
        End With
        selectedItem = ResultTree.GetNextItemName(selectedItem)
    Wend

End Sub
