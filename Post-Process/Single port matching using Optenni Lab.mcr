' 自动运行版本 - 仅首次选择Settings File，之后全默认无交互
Option Explicit

'#include "vba_globals_all.lib"
'#include "mws_ports.lib"
'#include "complex.lib"
'#Include "OptenniLibrary.lib"
'#Include "OptenniLibrary_MWS.lib"

' History of Changes
' ------------------
' 2025-XX-XX: 自动运行版本 - 可直接运行，结果插入结果树
'             仅首次选择Settings File，之后全自动无交互
' ------------------------------------------------------------------------------------

' ============================================================
' 主入口函数 - 可直接运行
' ============================================================
Sub Main()
    ' 仅首次运行检查并选择Settings File
    Call EnsureSettingsFile()

    ' 执行Optenni处理
    Dim res As Boolean
    res = OptenniTransfer()

    If res Then
        ' 将结果添加到结果树
        Call AddResultsToTree()
        'MsgBox "Optenni匹配电路优化完成！", vbOkOnly, "完成"
    Else
        MsgBox "Optenni处理失败，请检查错误信息", vbOkOnly, "错误"
    End If
End Sub

' ============================================================
' 确保Settings File已选择（仅首次）
' ============================================================
Sub EnsureSettingsFile()
    Dim filePath As String
    filePath = GetScriptSetting("filePath", "")
    If MsgBox("Choose last used matching circuit setting file (*.mcs) or not?" + vbNewLine _
    +"The last used *.mcs file is "+filePath, vbOkCancel, "Pick a *.mcs file")=vbCancel Then
		filePath=""
    End If

    If filePath = "" Then
        ' 首次运行，需要用户选择Settings File
        'Dim matchType As String
        'matchType = GetScriptSetting("matchType", "0")


        filePath = GetFilePath("", "mcs", "", _
                               "Select Optenni Lab settings file (首次选择，之后自动)")
        If filePath = "" Then
            MsgBox "未选择Settings File，操作取消", vbOkOnly, "操作取消"
            End
        End If
        StoreScriptSetting "filePath", filePath

        ' 保存默认参数
        StoreScriptSetting "matchType", "0"       ' 单端口匹配
        StoreScriptSetting "resultType", "0"      ' 曲线模式
        StoreScriptSetting "returnCircuit", "1"   ' 不返回电路
        StoreScriptSetting "storeSweepData", "1" ' 存储项目文件
        StoreScriptSetting "UseAR", "0"          ' 不使用AR
        StoreScriptSetting "SendRadPattern", "1"  ' 发送辐射方向图
        StoreScriptSetting "SendRadEffi", "1"     ' 发送辐射效率
        StoreScriptSetting "resolution", "5"     ' 角度分辨率
    End If
End Sub

' ============================================================
' 将结果添加到结果树
' ============================================================
Sub AddResultsToTree()
    Dim OptenniDir As String
    OptenniDir = GetProjectPath("Result") + "OptenniLab"
    Dim output_s As String
    output_s = OptenniDir + "\matchingout.xml"

    If Not IsWindows Then
        output_s = Replace$(output_s, "\", "/")
    End If

    If Dir$(output_s) = "" Then
        ReportInformationToWindow "Optenni输出文件不存在"
        Exit Sub
    End If

    Dim inputLine As String
    Dim spPos As Integer
    Dim numData As Integer
    Dim leftVal As Double
    Dim rightVal As Double
    Dim curveName As String

    On Error Resume Next
    Open output_s For Input As #3
    If Err.Number <> 0 Then
        ReportInformationToWindow "无法读取Optenni输出文件"
        Exit Sub
    End If
    On Error GoTo 0

    Line Input #3, inputLine

    'Dim costValue As Double
    'costValue = GetCostFunctionValue(inputLine)

    spPos = InStr(inputLine, "<optennicurves")
    If spPos = 0 Then
        Close #3
        Exit Sub
    End If

    spPos = InStr(inputLine, " n=")
    If spPos = 0 Then
        Close #3
        Exit Sub
    End If
    inputLine = Mid(inputLine, spPos + 4)
    spPos = InStr(inputLine, """")
    If spPos = 0 Then
        Close #3
        Exit Sub
    End If
    numData = CInt(Left(inputLine, spPos - 1))

    ' 临时存储曲线数据
    Const MAX_CURVES As Integer = 100
    Const MAX_POINTS As Integer = 2000
    Dim curveNames(MAX_CURVES) As String
    Dim curvePointCounts(MAX_CURVES) As Integer
    Dim curveXData(MAX_CURVES, MAX_POINTS) As Double
    Dim curveYData(MAX_CURVES, MAX_POINTS) As Double
    Dim currentCurve As Integer
    Dim pointCount As Integer

    Dim ind As Integer
    For ind = 1 To numData
        Line Input #3, inputLine

        spPos = InStr(inputLine, "<curve")
        If spPos = 0 Then Exit For

        spPos = InStr(inputLine, "name=")
        If spPos = 0 Then Exit For
        inputLine = Mid(inputLine, spPos + 6)
        spPos = InStr(inputLine, """")
        If spPos = 0 Then Exit For

        curveNames(ind) = Left(inputLine, spPos - 1)
        currentCurve = ind
        pointCount = 0

        Dim FoundEnd As Boolean
        FoundEnd = False
        While Not FoundEnd
            Line Input #3, inputLine
            spPos = InStr(inputLine, "</curve")
            If (spPos <> 0) Then
                FoundEnd = True
            Else
                spPos = InStr(inputLine, " ")
                If spPos > 0 And pointCount < MAX_POINTS Then
                    leftVal = CDbl(Left(inputLine, spPos - 1))
                    rightVal = CDbl(Mid(inputLine, spPos + 1))
                    curveXData(currentCurve, pointCount) = leftVal / Units.GetFrequencyUnitToSI
                    curveYData(currentCurve, pointCount) = rightVal
                    pointCount = pointCount + 1
                End If
            End If
        Wend
        curvePointCounts(currentCurve) = pointCount
    Next ind

    Close #3

    ' 使用ResultTree添加结果 (类型必须是有效的: XYSignal, Notefile等)
    'With ResultTree
        ' 添加Optenni Matching组
        '.Name "Optenni Matching"
        '.DeleteAt "truemodelchange"
        '.Type "Notefile"
        '.Add

        ' 添加Matched S-Parameters组
        '.Name "Optenni Matching\Matched S-Parameters"
        '.DeleteAt "truemodelchange"
        '.Type "Notefile"
        '.Add

        ' 添加每条曲线
    Dim o() As Object
	ReDim o(numData)
    For ind = 1 To numData
        If curveNames(ind) <> "" And curvePointCounts(ind) > 0 Then
            Dim resultName As String
            resultName = "1D Results\Matched Results\" + curveNames(ind)

            '.Name resultName
            '.DeleteAt "truemodelchange"
            '.Type "XYSignal"
			Set o(ind) = Result1D("")
            Dim pt As Integer
            For pt = 0 To curvePointCounts(ind) - 1
                'If pt = 0 Then
                '    .StoreAsXYSample curveXData(ind, pt), curveYData(ind, pt)
                'Else
                o(ind).AppendXY curveXData(ind, pt), curveYData(ind, pt)
                'End If
            Next pt
            o(ind).ylabel("dB")

			o(ind).Save(curveNames(ind)+".sig")

			o(ind).AddToTree(resultName)
        End If
    Next ind
    'End With
    '评估出匹配后的辐射效率
    Dim sfilename As String
    Dim Osp As Object
    Dim tfilename As String
    Dim Otot As Object
    Dim Orad As Object
    Dim rItem As String

	sfilename = ResultTree.GetFileFromTreeItem("1D Results\Matched Results\S11")
	tfilename = ResultTree.GetFileFromTreeItem("1D Results\Matched Results\Total efficiency")
	rItem = "1D Results\Matched Results\Radiation efficiency"

	Set Osp = Result1D(sfilename)
	Set Otot = Result1D(tfilename)
	Set Orad = Result1D("")

	Dim nPoints As Integer, n As Integer
	Dim x As Double, ys As Double, yt As Double, yr As Double
	nPoints = Osp.getN

	For n = 0 To nPoints-1

	'read all points, index of first point is zero.

		x = Osp.GetX(n)

		ys = Osp.GetY(n)
		yt = Otot.GetY(n)
		yr = getEffiFromSp_dB(ys,yt, False)
		'print to message window
		Orad.AppendXY x, yr
		'ReportInformationToWindow("x: " + Cstr(x) + " y: " + CStr(y))

	Next n
    Orad.ylabel("dB")

	Orad.Save("Radiation efficiency"+".sig")

	Orad.AddToTree(rItem)
    On Error Resume Next
    Kill output_s
    On Error GoTo 0

    ReportInformationToWindow "Optenni结果已添加到结果树"
End Sub

' ============================================================
' 获取成本函数值
' ============================================================
Function GetCostFunctionValue(inputLineIn As String) As Double
    Dim inputLine As String
    inputLine = inputLineIn
    Dim spPos As Integer
    GetCostFunctionValue = 0

    spPos = InStr(inputLine, "costFunctionValue")
    If spPos = 0 Then Exit Function
    inputLine = Mid(inputLine, spPos + 3)

    spPos = InStr(inputLine, """")
    If spPos = 0 Then Exit Function
    inputLine = Mid(inputLine, spPos + 1)
    spPos = InStr(inputLine, """")
    If spPos = 0 Then Exit Function

    GetCostFunctionValue = CDbl(Left(inputLine, spPos - 1))
End Function

' ============================================================
' 核心传输函数 - 执行Optenni Lab处理
' ============================================================
Function OptenniTransfer() As Boolean
    OptenniTransfer = True

    ' 初始化Optenni库
    Call InitOptenniLibrary()
    Dim optenniPath As String
    optenniPath = GetOptenniLabPath()

    If optenniPath = "" Or optenniPath = "OptenniLabNotInstalled" Then
        MsgBox "Optenni Lab未安装或未找到" + vbCrLf + _
               "请访问 www.optenni.com 获取更多信息", vbOkOnly, "Optenni Lab未安装"
        OptenniTransfer = False
        Exit Function
    End If

    ' 读取配置
    Dim useAR As Boolean
    useAR = CBool(GetScriptSetting("UseAR", "0"))

    Dim filePath As String
    filePath = GetScriptSetting("filePath", "")
    If Not IsWindows Then
        filePath = Replace$(filePath, "\", "/")
    End If

    ' 版本检查
    Dim major As Integer
    Dim minor As Integer
    major = GetOptenniLabMajorVersion()
    minor = GetOptenniLabMinorVersion()

    If (major < 3) Then
        MsgBox "Optenni Lab版本必须至少为3.0，当前版本: " + CStr(major) + "." + CStr(minor), vbOkOnly, "版本不兼容"
        OptenniTransfer = False
        Exit Function
    End If

    ' 检查结果
    If Not ResultTree.DoesTreeItemExist("1D Results\S-Parameters") Then
        MsgBox "没有计算结果。请先运行仿真。", vbOkOnly, "无结果"
        OptenniTransfer = False
        Exit Function
    End If

    ' 创建输出目录
    Dim OptenniDir As String
    OptenniDir = GetProjectPath("Result") + "OptenniLab"
    If Dir(OptenniDir, vbDirectory) = "" Then
        MkDir OptenniDir
    End If

    ' 获取项目信息
    Dim projectFullName As String
    Dim projectName As String
    projectFullName = GetProjectPath("Project")
    Dim slashPos As Integer
    slashPos = InStrRev(projectFullName, "\")
    If slashPos = 0 Then
        MsgBox "无法获取项目名称", vbOkOnly, "错误"
        Exit Function
    End If
    projectName = Mid(projectFullName, slashPos + 1)

    ' 导出Touchstone文件
    With TOUCHSTONE
        .Reset
        .FileName ("OptenniLab\" + projectName)
        .Impedance (50)
        .FrequencyRange ("Full")
        .Renormalize (True)
        If Result1DDataExists("ar^cS1(1)1(1)") And useAR Then
            .UseARResults (True)
        Else
            .UseARResults (False)
        End If
        .Write
    End With

    ' 获取端口数量
    Dim nport As Integer
    nport = GetOptenniPorts()
    If nport = 0 Then
        MsgBox "未找到端口", vbOkOnly, "端口错误"
        OptenniTransfer = False
        Exit Function
    End If

    ' 读取参数设置
    Dim matchType As String
    Dim returnCircuit As String
    Dim storeSweepData As String
    Dim SendRadPattern As String
    Dim SendRadEffi As String
    Dim resolution As Double

    matchType = GetScriptSetting("matchType", "0")
    returnCircuit = GetScriptSetting("returnCircuit", "0")
    storeSweepData = GetScriptSetting("storeSweepData", "0")
    If (major < 3 Or (major = 3 And minor < 3)) Then
        storeSweepData = "0"
    End If
    SendRadPattern = GetScriptSetting("SendRadPattern", "1")
    SendRadEffi = GetScriptSetting("SendRadEffi", "1")
    resolution = CDbl(GetScriptSetting("resolution", "5"))

    ' 转换参数
    Dim matchtypeStr As String
    Dim returnCircuitStr As String
    If (matchType = "0") Then
        matchtypeStr = "single"
    ElseIf (matchType = "1") Then
        matchtypeStr = "multiport"
    Else
        matchtypeStr = "schematic"
    End If

    If (returnCircuit = "0") Then
        returnCircuitStr = "curves"
    Else
        returnCircuitStr = "cstcircuitcurves"
    End If

    ' 辐射效率命令
    Dim efficiencyCommand As String
    If (SendRadEffi = "1") Then
        efficiencyCommand = GetOptenniEfficiencyCommand()
    End If

    ' 辐射方向图命令
    Dim patternCommand As String
    If (major > 4 Or (major = 4 And minor > 1)) Then
        If SendRadPattern = "1" Then
            patternCommand = GetOptenniRadiationPatternCommand(resolution)
        End If
    End If
    If Len(patternCommand) > 0 Then
        efficiencyCommand = ""
    End If

    ' 构建XML文件
    Dim xmlFile As String
    xmlFile = OptenniDir + "\matching.xml"
    If Not IsWindows Then
        xmlFile = Replace$(xmlFile, "\", "/")
    End If

    Dim output_s As String
    output_s = OptenniDir + "\matchingout.xml"
    If Not IsWindows Then
        output_s = Replace$(output_s, "\", "/")
    End If

    ' 获取参数扫描参数
    Dim parStr As String
    parStr = ""
    With ParameterSweep
        Dim numpar As Integer
        numpar = .GetNumberOfVaryingParameters
        Dim ii As Integer
        If numpar > 0 Then
            For ii = 0 To numpar - 1
                If Len(parStr) > 0 Then parStr = parStr + " "
                parStr = parStr + .GetNameOfVaryingParameter(ii) + "=" + CStr(.GetValueOfVaryingParameter(ii))
            Next ii
        End If
    End With

    With Optimizer
        numpar = .GetNumberOfVaryingParameters
        If numpar > 0 Then
            For ii = 0 To numpar - 1
                If (InStr(parStr, .GetNameOfVaryingParameter(ii)) = 0) Then
                    If Len(parStr) > 0 Then parStr = parStr + " "
                    parStr = parStr + .GetNameOfVaryingParameter(ii) + "=" + CStr(.GetValueOfVaryingParameter(ii))
                End If
            Next ii
        End If
    End With

    If (Len(parStr) > 89) Then parStr = Left(parStr, 89)
    If (Len(parStr) = 0) Then parStr = "sweepdata"

    ' 写入XML
    Open xmlFile For Output As #1
    Print #1, "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
    Print #1, "<matching type="""; matchtypeStr; """ returntype="""; returnCircuitStr; """>"
    Print #1, "<settingsfile><![CDATA["; filePath; "]]></settingsfile>"
    Print #1, "<outputfile><![CDATA["; output_s; "]]></outputfile>"

    If (parStr <> "" And storeSweepData = "1") Then
        Dim OptenniSaveDir As String
        OptenniSaveDir = GetProjectPath("Project") + "\OptenniLabData"
        Dim project_s As String
        project_s = OptenniSaveDir + "\" + parStr + ".opr"
        If Not IsWindows Then
            project_s = Replace$(project_s, "\", "/")
        End If
        If Dir(OptenniSaveDir, vbDirectory) = "" Then MkDir OptenniSaveDir
        Print #1, "<projectoutputfile><![CDATA["; project_s; "]]></projectoutputfile>"
    End If

    Print #1, "</matching>"
    Close #1

    ' 创建锁文件
    Dim sFileOptenniRunning As String
    sFileOptenniRunning = OptenniDir + "\running.lok"
    If Not IsWindows Then
        sFileOptenniRunning = Replace$(sFileOptenniRunning, "\", "/")
    End If
    If Dir$(sFileOptenniRunning) <> "" Then Kill sFileOptenniRunning
    Wait 0.01
    Open sFileOptenniRunning For Output As #2
    Close #2

    ' 端口几何信息
    If (major >= 5) Then
        Call StoreOptenniPortGeometryFile()
    End If

    ' CST实例ID
    Dim CSTIdCommand As String
    Dim CSTid As String
    CSTid = DS.GetRegisteredDEString()
    If CSTid <> "" Then
        CSTIdCommand = " -cstinstance " + CSTid
    End If

    ' 启动Optenni Lab
    ReportInformationToWindow "正在启动 Optenni Lab..."

    Dim fname As String
    fname = """" + OptenniDir + "\" + projectName + ".s" + CStr(nport) + "p" + """"
    If Not IsWindows Then
        fname = Replace$(fname, "\", "/")
    End If

    Shell optenniPath + " " + fname + _
          CSTIdCommand + " -matching """ + xmlFile + """" + _
          " -lockfile """ + sFileOptenniRunning + """" + efficiencyCommand + _
          patternCommand, vbNormalFocus

    ' 等待完成
    Dim dRefTime As Double
    Dim dSeconds As Double
    Dim dTimeOut As Double
    dRefTime = Timer
    dTimeOut = 36000

    Do
        Wait 1
        dSeconds = Timer - dRefTime
        If dSeconds < 0 Then dSeconds = dSeconds + 86400
    Loop Until dSeconds > dTimeOut Or Dir$(sFileOptenniRunning) = ""

    If dSeconds > dTimeOut Then
        MsgBox "Optenni Lab处理超时", vbOkOnly, "超时"
        OptenniTransfer = False
        Exit Function
    End If

    ' 保存项目文件
    If (parStr <> "" And storeSweepData = "1") Then
        With ResultTree
            .Name "OptenniLabData" + parStr + ".opr"
            .File OptenniSaveDir + "\" + parStr + ".opr"
            .DeleteAt "truemodelchange"
            .Type "Hidden"
            .Add
        End With
    End If
    ReportInformationToWindow "Optenni Lab处理完成"
End Function
Function getEffiFromSp_dB(s_db As Double, effi_db As Double, Invers As Boolean) As Double
    ' tot to rad
    If Invers = False Then
        getEffiFromSp_dB = 10 * CST_Log10(10 ^ (effi_db / 10) / (1 - (10 ^ (s_db / 20)) ^ 2))
    ' Rad to tot
    Else
        getEffiFromSp_dB = 10 * CST_Log10(10 ^ (effi_db / 10) * (1 - (10 ^ (s_db / 20)) ^ 2))
    End If
End Function
