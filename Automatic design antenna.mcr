'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3D.lib"
'#Uses "AntennaElement.CLS"
'#Uses "Antenna.CLS"
Public antElem_arr() As AntennaElement
Public aSolidArray_CST() As String
Public aMaterialArray_CST() As String
'Public nMaterials_CST As Integer
Public nSolids_CST As Integer
Public ant_material As String
Public sub_material As String
Public ant As Antenna
Public tgtFreq As Double
Public fileNumber As Integer
Public availableElementsStr As String
Public availableElementsNum As Integer
Const macrofile = "C:\Program Files (x86)\CST Studio Suite 2021\Library\Macros\Coros\Post-Process\Store Resonant Frequency and Q.mcr"
Sub Main

	Dim prjPath As String
	Dim logFile As String
	Dim LineRead As String
	Dim LineCount As Integer
	Dim feedSolid As String
	Dim feed As AntennaElement
	Dim tgtLength As Double
	Dim deltaX As Double, deltaY As Double, deltaZ As Double
	Dim resFreq As Double
	Dim Q As Double, totEffi As Double, radEffi As Double
	Dim ifSuccess As Boolean

	prjPath = GetProjectPath("Project")
   	logFile = prjPath + "\Progress log.txt"
   	fileNumber = FreeFile
	Begin Dialog UserDialog 400,119,"Construct a new or Reconstruct an old" ' %GRID:10,7,1,1
		GroupBox 0,7,380,112,"Options:",.GroupBox1
		OKButton 60,91,90,21
		CancelButton 180,91,90,21
		OptionGroup .Group1
			OptionButton 20,28,340,14,"Construct a new antenna from the beginning",.OptionButton1
			OptionButton 20,49,340,14,"Reconstruct a recorded antenna from the log file",.OptionButton2
	End Dialog
	Dim dlg0 As UserDialog
	'Dim logicInformation As String
	If Dialog(dlg0,-2) = 0 Then
		Exit All
	End If
	If dlg0.Group1=1 Then
	'reconstruct an old antenna from the old log file
	   	If Dir(logFile) <> "" Then
			FileCopy logFile, Split(logFile,".")(0)&"_old.txt"
		Else
			MsgBox "%%Log file does not exist.", vbCritical
			Exit All
   		End If
   		Dim antLength As String
   		Dim antRadEffi As String
   		Dim antTotEffi As String
		antennaDesign_initialize(False)
		ReportinformationTowindow "$Main: open file #"&CStr(fileNumber)&" for reconstructing"
		Open Split(logFile,".")(0)&"_old.txt" For Input As #fileNumber
		LineCount = 0
		While Not EOF(fileNumber)
			Line Input #fileNumber, LineRead
			If InStr(LineRead, "Feed element")<>0 Then
				feedSolid = Right(LineRead, Len(LineRead)-InStr(LineRead, ":")-1)
				Set feed = getElementFromSolid(antElem_arr, feedSolid)
				If feed Is Nothing Then
					MsgBox "Set feed part failed!", vbCritical,"Error"
					Exit All
				End If
				Set ant=New Antenna
				ant.initialize(feed, ant_material, sub_material,availableElementsNum, availableElementsStr)
			ElseIf InStr(LineRead, "Antenna Logics") Then
				ant.conLogics = Right(LineRead, Len(LineRead)-InStr(LineRead, ":")-1)
				ant.elementNumber = Int(Len(ant.conLogics)/2)+1
				ant.reconstructor()
				If Not EOF(fileNumber) Then
					Line Input #fileNumber, LineRead
					antLength = Right(LineRead, Len(LineRead)-InStr(LineRead, ":")-1)
				End If
				If Not EOF(fileNumber) Then
					Line Input #fileNumber, LineRead
					antRadEffi = Right(LineRead, Len(LineRead)-InStr(LineRead, ":")-1)
				End If
				If Not EOF(fileNumber) Then
					Line Input #fileNumber, LineRead
					antTotEffi = Right(LineRead, Len(LineRead)-InStr(LineRead, ":")-1)
				End If
				If MsgBox( "Antenna Length: " & antLength & vbCrLf &  _
				"Antenna radiation efficiency: " & antRadEffi & vbCrLf &  _
				"Antenna total efficiency: " & antTotEffi & vbCrLf &  _
				"keep going ?",vbOkCancel, "Simulation results" )= vbOK Then
					ant.destructor()
				Else
					MsgBox "Reconstructing progress teminated!",vbInformation,"Notice"
					Exit All
				End If

			End If

			LineCount += 1
		Wend
		Close # fileNumber
		ReportinformationTowindow "$Main: close file #"&CStr(fileNumber)&" for reconstructing"
		Exit All
	End If
	'Construct a new antenna
	ReportinformationTowindow "$Main: open file #"&CStr(fileNumber)&" for constructing"
   	Open logFile For Output As #fileNumber
   	ReportInformationToWindow "**********Rebuilding the model at " + CStr(Now) +"************"
   	Print #fileNumber, "************Rebuilding the model at " + CStr(Now) +"****************"
   	Rebuild
	ReportInformationToWindow "***********Initializing of antenna begins at " + CStr(Now) +"*************"
   	Print #fileNumber, "************Initializing of antenna begins at " + CStr(Now) +"****************"
	antennaDesign_initialize(True)
	ReportInformationToWindow "**********Initializing of antenna finishes at " + CStr(Now) +"************"
	Print #fileNumber, "************Initializing of antenna finishes at " + CStr(Now) +"************"
	'Select operating frequency and feed point
	Begin Dialog UserDialog 330,140,"Frquency and feeding part settings" ' %GRID:10,7,1,1
		GroupBox 0,0,330,49,"Frequency settings:",.GroupBox1
		TextBox 140,21,50,14,.Freq
		OKButton 60,105,90,21
		CancelButton 200,105,90,21
		Text 10,21,120,14,"Target frequency:",.Text6
		GroupBox 0,49,330,49,"Feeding part selection:",.GroupBox2
		DropListBox 60,70,210,14,aSolidArray_CST(),.feedingPart
		Text 200,21,30,14,"GHz",.Text1
	End Dialog
	Dim dlg As UserDialog
	'Dim tgtFreq As Double


	'Dim ant_material As String
	'ant_material = "Metal/Copper (annealed)"

	dlg.Freq = "1.56"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

	tgtFreq = CDbl(dlg.Freq)
	feedSolid = aSolidArray_CST(dlg.feedingPart)
	ReportInformationToWindow "%% Target frequency: " + dlg.Freq + "GHz"
	Print #fileNumber, "%% Target frequency: " + dlg.Freq + "GHz"
	ReportInformationToWindow "%% Feed element: " + feedSolid
	Print #fileNumber, "%% Feed element: " + feedSolid
	Set feed = getElementFromSolid(antElem_arr, feedSolid)
	If feed Is Nothing Then
		MsgBox "Set feed part failed!", vbCritical,"Error"
		Exit All
	End If

	Set ant=New Antenna
	ant.initialize(feed, ant_material, sub_material,availableElementsNum, availableElementsStr)

	Dim X As Integer
	Dim Y As Integer
	'***************Add a choice to recover antenna routing from the progress log*********************
	feed.getDimensions(deltaX,deltaY,deltaZ)
	X=0
	tgtLength = Round(3e8/(tgtFreq*1e6)/3,1) ' unit mm
	resFreq = 0
	ReportInformationToWindow "*********Routing of antenna starts at " + CStr(Now) +"************"
	Print #fileNumber, "***********Routing of antenna starts at " + CStr(Now) +"**********"
	Do
		X+=1
		ReportInformationToWindow "############### Round#"+CStr(X)+" of antenna routing starts at " + CStr(Now) +" ################"
		Print #fileNumber, "############## Round#"+CStr(X)+" of antenna routing starts at " + CStr(Now) +" ##############"
		ReportInformationToWindow "%% The target length is: "+CStr(Round(tgtLength,2))+"mm"
		Print #fileNumber, "%% The target length is: "+CStr(Round(tgtLength,2))+"mm"
		If X>1 Then
			ant.destructor()
			ant.initialize(feed, ant_material, sub_material,availableElementsNum, availableElementsStr)
		End If
		ifSuccess = False
		Y=0
		While ant.constructor(tgtLength, nSolids_CST, deltaX, deltaY, deltaZ) = True
			Y += 1
			ifSuccess = True
			ReportInformationToWindow "%% Loop#"+CStr(Y)+" of antenna routing ends at " + CStr(Now)
			Print #fileNumber, "%% Loop#"+CStr(Y)+" of antenna routing ends at " + CStr(Now)
			ReportInformationToWindow "%% loop# "+Cstr(Y)+ _
			" is done and the antenna length target is met, simulation will begin now"
			Print #fileNumber, "%% loop# "+Cstr(Y)+ _
			" is done and the antenna length target is met, simulation will begin now"
			'ant.patcher()
			'Rebuild, start simulating
			ReportInformationToWindow "%% Achieved antenna length: " + CStr(Round(ant.length,2))+"mm"
			Print #fileNumber, "%% Achieved antenna length: " + CStr(Round(ant.length,2))+"mm"
			ant.toHistoryList()
			Solver.MeshAdaption(False)
			Solver.SteadyStateLimit(-40)
			Solver.Start
			ReportInformationToWindow "%% Simulation#"+CStr(Y)+" of antenna ends at " + CStr(Now)
			Print #fileNumber, "%% Simulation#"+CStr(Y)+" of antenna ends at " + CStr(Now)
			MacroRun(macrofile)
			resFreq = getResonanceFrequency()(0)
			'Amend the target length when the resonance frequency does not meet the target
			If Abs((resFreq-tgtFreq)/tgtFreq)>=0.035 And Abs((resFreq-tgtFreq)/tgtFreq)<0.5 Then
				ReportInformationToWindow "%% Simulation done, the resonance frequency is: " + CStr(Round(resFreq,2)) + "GHz"
				Print #fileNumber, "%% Simulation done, the resonance frequency is: " + CStr(Round(resFreq,2)) + "GHz"
				tgtLength = (resFreq/tgtFreq)*tgtLength
				ReportInformationToWindow "%%Intermediate Logics: " & ant.conLogics
				Print #fileNumber, "%%Intermediate Logics: " & ant.conLogics
				ReportInformationToWindow "%% The resonance frequency does not meet the target and optimization of length will begin."
				Print #fileNumber, "%% The resonance frequency does not meet the target and optimization of length will begin."
				DeleteResults
				ReportInformationToWindow "%% New target length is: "+CStr(Round(tgtLength,2))+"mm"
				Print #fileNumber, "%% New target length is: "+CStr(Round(tgtLength,2))+"mm"
			ElseIf Abs((resFreq-tgtFreq)/tgtFreq)<0.035 Then
				ReportInformationToWindow "%% Simulation done, the resonance frequency is: " + CStr(Round(resFreq,2)) + "GHz"
				Print #fileNumber, "%% Simulation done, the resonance frequency is: " + CStr(Round(resFreq,2)) + "GHz"
				If False Then	'MsgBox("Go on?",vbOkCancel,"Notice")<>vbOK Then
					Exit Do
				Else
					Q = getQ()(0)
					totEffi = getEfficiencyAtFrequency(resFreq, True)
					radEffi = getEfficiencyAtFrequency(resFreq, False)
					ReportInformationToWindow "%% The resonance frequency has been met @"& CStr(Round(resFreq,2)) & "GHz"
					Print #fileNumber, "%% The resonance frequency has been met @"& CStr(Round(resFreq,2)) & "GHz"
					ReportInformationToWindow "%%Antenna Logics: " & ant.conLogics
					Print #fileNumber, "%% Antenna Logics: " & ant.conLogics
					'ReportInformationToWindow "%% The Q value is "&CStr(Round(Q,2))
					'Print #fileNumber, "%% The Q value is "&CStr(Round(Q,2))
					ReportInformationToWindow "%%Antenna length: " + CStr(Round(ant.length,2))+"mm"
					Print #fileNumber, "%% Antenna length: " + CStr(Round(ant.length,2))+"mm"
					ReportInformationToWindow "%%Radiation efficiency: " & CStr(Round(radEffi, 2))&"dB"
					Print #fileNumber, "%% Radiation efficiency: " & CStr(Round(radEffi, 2))&"dB"
					ReportInformationToWindow "%%Total efficiency: " & CStr(Round(totEffi, 2))&"dB"
					Print #fileNumber, "%% Total efficiency: " & CStr(Round(totEffi, 2))&"dB"
					'*********Record more data such as radiation efficiency and connection logics string to the progress log************
					DeleteResults
					Exit While
				End If
			ElseIf Abs((resFreq-tgtFreq)/tgtFreq)>=0.5 Then
				ReportInformationToWindow "%% Simulation done, the resonance frequency is parsed wrong!"
				Print #fileNumber, "%% Simulation done, the resonance frequency is parsed wrong!"
				DeleteResults
				Exit While
			'Rebuild
			End If
			Plot.Update
			ifSuccess = False
		Wend
		If ifSuccess = False Then
			ReportInformationToWindow "####### Round "+Cstr(X)+ _
				" is done but the antnena Length is not met, another trial starts #######"
			Print #fileNumber, "####### Round "+Cstr(X)+ _
				" is done but the antnena Length is not met, another trial starts ########"
		End If
		'Plot.Update
		'Plot.ExportImage ("E:\image.bmp", 1024, 768)
	Loop Until X>=100
	ReportInformationToWindow "************Routing of antenna ends at " + CStr(Now) +"************"
	Print #fileNumber, "************Routing of antenna ends at " + CStr(Now) +"***************"
	Close #fileNumber
	ReportinformationTowindow "$Main: close file #"&CStr(fileNumber)&" for constructing"
	MsgBox "OOOps"
End Sub
Sub antennaDesign_initialize(cOrR As Boolean)
	availableElementsStr=""
	availableElementsNum=0
	'Rebuild
	Plot.update
	'Dim aSolidArray_CST() As String, nSolids_CST As Integer
	Dim iSolid_CST As Integer, i As Integer
	Dim sCommand As String
	Dim historyCaption As String
	Dim sFullSolidName As String
	Dim iMaterial_CST As Integer
	'Dim sFullMaterialName As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	If MsgBox("Please assign materials for antenna and substrate.",vbYesNo,"Notice") <> vbYes Then
    	Exit All
    End If
	selectMaterial(ant_material,"Pick a material for antenna")
	ReportInformationToWindow "%% Antenna material: " + ant_material
	If cOrR=True Then Print #fileNumber, "%% Antenna material: " + ant_material
	selectMaterial(sub_material,"Pick a material for substrate")
	ReportInformationToWindow "%% Substrate material: " + sub_material
	If cOrR=True Then Print #fileNumber, "%% Substrate material: " + sub_material
	If MsgBox("Please select all solids contained in the antenna.",vbYesNo,"Notice") <> vbYes Then
    	Exit All
    End If
	SelectSolids_Antenna(aSolidArray_CST(), nSolids_CST)
	'For i=0 to Ubound(aSolidArray_CST)
	'SelectSolids_LIB
	'SelectMaterials_LIB(aMaterialArray_CST(), nMaterials_CST)
	ReportInformationToWindow "%% Number of solids for antenna elements: "+CStr(nSolids_CST)
	If cOrR=True Then Print #fileNumber, "%% Number of solids for antenna elements: "+CStr(nSolids_CST)
	ReDim antElem_arr(nSolids_CST)
	sCommand = ""
	'Construct class instances from solids
	availableElementsNum = nSolids_CST
	ReportInformationToWindow "%% Antenna elements initializing...... "
	If (nSolids_CST > 0) Then
		For iSolid_CST = 1 To nSolids_CST
			sFullSolidName = aSolidArray_CST(iSolid_CST-1)
			Solid.GetLooseBoundingBoxOfShape(sFullSolidName,xMin,xMax,yMin,yMax,zMin,zMax)
			Set antElem_arr(iSolid_CST-1) = New AntennaElement
			With antElem_arr(iSolid_CST-1)
				.solidName = sFullSolidName
				availableElementsStr = availableElementsStr & sFullSolidName & "$"
				.solidMaterial = Solid.GetMaterialNameForShape(sFullSolidName)
			If StrComp(.solidMaterial,sub_material)<>0 Then
			'for debug
				 '.setMaterial(sub_material)
				 sCommand = sCommand & .setMaterialPermanently(sub_material)
			End If
				.setStartPoint(xMin,yMin,zMin)
				.setEndPoint(xMax,yMax,zMax)
				.defineVertices()
				.defineEdges()
				.defineFaces()'xMin,yMin,zMin,xMax,yMax,zMax
				.flag = False
			End With
			Plot.Update
			'Debug.Print antennaEle
		Next
	Else
		MsgBox "No solids are selected and the progress shall be terminated!", vbCritical, "Warning"
		Exit All
	End If
	historyCaption = "$IA$"
	If sCommand <> "" Then
		AddToHistory(historyCaption, sCommand)
		Plot.update
	End If
	'Find neighbors for all instances
	Dim ii As Integer
	For iSolid_CST = 1 To nSolids_CST
		For ii = iSolid_CST+1 To nSolids_CST
			'Creating FaceNeighbors has to be done before creating EdgeNeighbors
			'createFaceNeighbors(antElem_arr(iSolid_CST-1),antElem_arr(ii-1))
			antElem_arr(iSolid_CST-1).createFaceNeighborWith(antElem_arr(ii-1))
			'createEdgeNeighbors
			antElem_arr(iSolid_CST-1).createEdgeNeighborWith(antElem_arr(ii-1))
		Next
		Plot.Update
	Next
	MsgBox "Antenann element array initialization completed!!",vbInformation,"Initializing done"
End Sub
Function getElementFromSolid(elementArray() As AntennaElement,solidName As String) As AntennaElement
	Dim i As Integer
	For i = 1 To UBound(elementArray)-LBound(elementArray)+1
		If elementArray(i-1).solidName = solidName Then
			Set getElementFromSolid = elementArray(i-1)
			Exit Function
		End If
	Next
	getElementFromSolid = Nothing
End Function

Function selectMaterial(pickedMaterial As String, notice As String)
	'Dim cst_button As Integer
	'Dim cst_Operation As String
	'Dim nMaterials_LIB As Integer
	'Dim aMaterialArray_LIB() As String
	Dim cst_index As Long
	Dim cst_iii As Long

	'Dim cst_Material_select_name() As String
	Dim cst_Material_unselect_name() As String

	Dim cst_Material As String
	Dim bSelected As Boolean

	'ReDim cst_Material_select_name(0)
	ReDim cst_Material_unselect_name(0)

	For cst_index = 1 To Material.GetNumberOfMaterials
		cst_Material = Material.GetNameOfMaterialFromIndex (cst_index-1)
		ReDim Preserve cst_Material_unselect_name(UBound(cst_Material_unselect_name) + 1)
		cst_Material_unselect_name(UBound(cst_Material_unselect_name)) = cst_Material
		'End If
	Next
	Begin Dialog UserDialog 300,100,notice ',.SelectedMaterialsDialogFunc ' %GRID:10,7,1,1
		OKButton 40,70,90,21
		CancelButton 140,70,90,21
		PushButton 240,189,1,1,"hiddenPB",.hiddenPB
		GroupBox 0,7,300,56,"Pick a material:",.GroupBox1
		DropListBox 30,28,240,21,cst_Material_unselect_name(),.pickedMaterialNumber
	End Dialog
	Dim dlg_Material_CST As UserDialog
	If Dialog(dlg_Material_CST,-2) = 0 Then
		Exit All
	End If
	pickedMaterial = cst_Material_unselect_name(dlg_Material_CST.pickedMaterialNumber+1)
End Function
Sub SelectSolids_Antenna(aSolidArray_LIB() As String, nSolids_LIB As Integer)

	Dim cst_button As Integer
	Dim cst_Operation As String
	Dim cst_index As Long
	Dim cst_iii As Long

	Dim cst_solid_select_name() As String
	Dim cst_solid_unselect_name() As String

	Dim cst_solid As String
	Dim cst_solids() As String
	Dim bSelected As Boolean

	ReDim cst_solid_select_name(0)
	ReDim cst_solid_unselect_name(0)
	ReDim cst_solids(0)
	'filtering out all solids not in "1 Antenna"
	For cst_index = 0 To Solid.GetNumberOfShapes-1
		cst_solid = Solid.GetNameOfShapeFromIndex(cst_index)
		If InStr(cst_solid,"1 Antenna")<>0 Then
			ReDim Preserve cst_solids(UBound(cst_solids)+1)
			cst_solids(UBound(cst_solids)-1) = cst_solid
		End If

	Next
	ReDim Preserve cst_solids(UBound(cst_solids)-1)
	For cst_index = 0 To UBound(cst_solids)
		cst_solid = cst_solids(cst_index)
		bSelected = False
		For cst_iii = 1 To nSolids_LIB
			If (cst_solid = aSolidArray_LIB(cst_iii-1)) Then
				bSelected = True
				Exit For
			End If
		Next cst_iii

		If bSelected Then
			' The entry at index "0" will not be used
			ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name) + 1)
			cst_solid_select_name(UBound(cst_solid_select_name)) = cst_solid
		Else
			' The entry at index "0" will not be used
			ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name) + 1)
			cst_solid_unselect_name(UBound(cst_solid_unselect_name)) = cst_solid
		End If
	Next
	' fsr: A check
	If (nSolids_LIB <> UBound(cst_solid_select_name)) Then
		ReportWarningToWindow("SelectSolids: inconsistent solid count (expected "+CStr(nSolids_LIB)+ _
		", found "+Cstr(UBound(cst_solid_select_name))+"), please contact support.")
	End If

	' Check if the solids in component "1 Antenna"
	Begin Dialog UserDialog 850,238,"Selected Solids" ',.SelectedSolidsDialogFunc ' %GRID:10,7,1,1
		Text 20,7,120,14,"Unselected",.LabelSource
		ComboBox 20,21,350,189,cst_solid_unselect_name(),.Unselect_box
		PushButton 390,49,70,28,"=====>",.AddAll
		PushButton 390,84,70,28,"----->",.Add
		PushButton 390,119,70,28,"<-----",.Remove
		PushButton 390,154,70,28,"<=====",.RemoveAll
		Text 480,7,90,14,"Selected",.LabelTarget
		ComboBox 480,21,350,189,cst_solid_select_name(),.Select_box
		OKButton 20,210,90,21
		CancelButton 120,210,90,21
		PushButton 240,189,1,1,"hiddenPB",.hiddenPB
	End Dialog
	Dim dlg_Solid_CST As UserDialog

	cst_Operation = ""

	Do
		cst_button = Dialog(dlg_Solid_CST,5)	' use button 5 (hiddenPB) as default, for double mouse click

        Select Case cst_button
			Case -1		' OK
				cst_Operation = "OK"
				'
			Case 1		' ====> add all probes
				If (UBound(cst_solid_unselect_name) = 0) Then
					' Do nothing, nothing to add from
				Else
					ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name)+UBound(cst_solid_unselect_name))
					For cst_iii = 1 To UBound(cst_solid_unselect_name)
						cst_solid_select_name(UBound(cst_solid_select_name)-UBound(cst_solid_unselect_name)+cst_iii) = cst_solid_unselect_name(cst_iii)
					Next
					dlg_Solid_CST.unselect_box = ""
					ReDim cst_solid_unselect_name(0)
				End If
				'
			Case 2		' ----> add only one probe
				If (UBound(cst_solid_unselect_name) = 0) Then
					' Do nothing, nothing to add from
				ElseIf(dlg_Solid_CST.unselect_box = "") Then
					dlg_Solid_CST.unselect_box = cst_solid_unselect_name(1)
				Else
					' expand "Select" array
					ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name)+1)
					cst_solid_select_name(UBound(cst_solid_select_name)) = dlg_Solid_CST.Unselect_box
					For cst_iii = 1 To UBound(cst_solid_unselect_name)-1
						' Once the matching item has been found, replace it and every following item by the next item in the array
						If ((cst_solid_unselect_name(cst_iii) = dlg_Solid_CST.unselect_box) Or (dlg_Solid_CST.Unselect_box="")) Then
							cst_solid_unselect_name(cst_iii) = cst_solid_unselect_name(cst_iii+1)
							dlg_Solid_CST.unselect_box = ""
						End If
					Next
					' chop "Unselect" array
					ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name)-1)
					dlg_Solid_CST.unselect_box = ""
				End If
				'
			Case 3	' <----
				If (UBound(cst_solid_select_name) = 0) Then
					' Do nothing, nothing to remove from
				ElseIf(dlg_Solid_CST.select_box = "") Then
					dlg_Solid_CST.select_box = cst_solid_select_name(1)
				Else
					' expand "unselect" array
					ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name)+1)
					cst_solid_unselect_name(UBound(cst_solid_unselect_name)) = dlg_Solid_CST.select_box
					For cst_iii = 1 To UBound(cst_solid_select_name)-1
						' Once the matching item has been found, replace it and every following item by the next item in the array
						If ((cst_solid_select_name(cst_iii) = dlg_Solid_CST.select_box) Or (dlg_Solid_CST.select_box="")) Then
							cst_solid_select_name(cst_iii) = cst_solid_select_name(cst_iii+1)
							dlg_Solid_CST.select_box = ""
						End If
					Next
					' chop "select" array
					ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name)-1)
					dlg_Solid_CST.select_box = ""
				End If
				'
			Case 4	' <====
				If (UBound(cst_solid_select_name) = 0) Then
					' Do nothing, nothing to remove from
				Else
					ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name)+UBound(cst_solid_select_name))
					For cst_iii = 1 To UBound(cst_solid_select_name)
						cst_solid_unselect_name(UBound(cst_solid_unselect_name)-UBound(cst_solid_select_name)+cst_iii) = cst_solid_select_name(cst_iii)
					Next
					dlg_Solid_CST.select_box = ""
					ReDim cst_solid_select_name(0)
				End If
				'
			Case 5 ' hidden button for double mouse click
				If (dlg_Solid_CST.unselect_box <> "") Then
					If (UBound(cst_solid_unselect_name) = 0) Then
						' Do nothing, nothing to add from
					Else
						' expand "Select" array
						ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name)+1)
						cst_solid_select_name(UBound(cst_solid_select_name)) = dlg_Solid_CST.Unselect_box
						For cst_iii = 1 To UBound(cst_solid_unselect_name)-1
							' Once the matching item has been found, replace it and every following item by the next item in the array
							If ((cst_solid_unselect_name(cst_iii) = dlg_Solid_CST.unselect_box) Or (dlg_Solid_CST.Unselect_box="")) Then
								cst_solid_unselect_name(cst_iii) = cst_solid_unselect_name(cst_iii+1)
								dlg_Solid_CST.unselect_box = ""
							End If
						Next
						' chop "Unselect" array
						ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name)-1)
						dlg_Solid_CST.unselect_box = ""
					End If
				ElseIf (dlg_Solid_CST.select_box <> "") Then
					If (UBound(cst_solid_select_name) = 0) Then
						' Do nothing, nothing to remove from
					Else
						' expand "unselect" array
						ReDim Preserve cst_solid_unselect_name(UBound(cst_solid_unselect_name)+1)
						cst_solid_unselect_name(UBound(cst_solid_unselect_name)) = dlg_Solid_CST.select_box
						For cst_iii = 1 To UBound(cst_solid_select_name)-1
							' Once the matching item has been found, replace it and every following item by the next item in the array
							If ((cst_solid_select_name(cst_iii) = dlg_Solid_CST.select_box) Or (dlg_Solid_CST.select_box="")) Then
								cst_solid_select_name(cst_iii) = cst_solid_select_name(cst_iii+1)
								dlg_Solid_CST.select_box = ""
							End If
						Next
						' chop "select" array
						ReDim Preserve cst_solid_select_name(UBound(cst_solid_select_name)-1)
						dlg_Solid_CST.select_box = ""
					End If
				Else
					' Do nothing
				End If
			Case Else		' Cancel
				cst_Operation = "Quit"
		End Select

	Loop Until (cst_Operation <> "")

	If cst_Operation = "OK" Then

		' only update external solid-array, if OK is pressed, Cancel leaves array unchanged
		nSolids_LIB = UBound(cst_solid_select_name)
		If (nSolids_LIB > 0) Then
			ReDim aSolidArray_LIB(nSolids_LIB-1)
			For cst_iii = 1 To nSolids_LIB
				aSolidArray_LIB(cst_iii-1) = cst_solid_select_name(cst_iii)
			Next
		End If

	End If

End Sub
Function getResonanceFrequency() As Variant
    Dim prjPath As String
	Dim dataFile As String
	Dim freq() As Double
	Dim resStr As String
	Dim n As Integer
	Dim LineRead As String
	'Dim fFileNumber As Integer
	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\freq_Q.txt"
   	n = 1
   	'fFileNumber = FreeFile
	Open dataFile For Input As #1
	While Not EOF(1)
		Line Input #1, LineRead
		While Not EOF(1) And Left(LineRead,1)<>"F"
			Line Input #1, LineRead
		Wend

		If Left(LineRead,1) = "F" Then
			ReDim Preserve freq(n)
			resStr = Split(LineRead, "=")(1)
			freq(n-1) = CDbl(resStr) 'realized resonance frequency

		End If
		n = n+1
	Wend
	Close #1
	Return freq
End Function
Function getQ() As Variant
    Dim prjPath As String
	Dim dataFile As String
	Dim Q() As Double
	Dim resStr As String
	Dim n As Integer
	Dim LineRead As String
	'Dim QfileNumber As Integer
	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\freq_Q.txt"
   	n = 1
   	'QfileNumber=FreeFile
	Open dataFile For Input As #2
	While Not EOF(2)
		Line Input #2, LineRead
		While Not EOF(2) And Left(LineRead,1)<>"Q"
			Line Input #2, LineRead
		Wend

		If Left(LineRead,1) = "Q" Then
			ReDim Preserve Q(n)
			resStr = Split(LineRead, "=")(1)
			Q(n-1) = CDbl(resStr) 'Q value

		End If
		n = n+1
	Wend
	Close #2
	Return Q
End Function
Function getEfficiencyAtFrequency(f As Double, totOrRad As Boolean) As Double
	'totOrRad: True-total efficiency; False-radation efficiency
	Dim effiItem As String
	Dim effiPath As String
	Dim effiFile As String
	Dim currentItem As String

    effiPath = "1D Results\Efficiencies"
    effiItem = Resulttree.GetFirstChildName(effiPath)

    If effiItem = "" Then
   	  MsgBox("No Efficiency results found!",vbCritical,"Warning")
   	  Exit All
    End If

    currentItem = Resulttree.GetFirstChildName(effiPath)
    While currentItem <> ""
		Dim EffiType As String, FileName As String
		Dim nPoints As Long, n As Integer, dBValue As Double, X As Double, Y As Double
		Dim O As Object
		EffiType = Mid(currentItem,Len(effiPath)+2,InStr(currentItem,"[")-Len(effiPath)-2)

		If totOrRad=False Then
			If InStr(EffiType, "Rad")<>0 Then
				GoTo getValue
			End If
		Else
			If InStr(EffiType, "Tot")<>0 Then
			getValue:
        		FileName = Resulttree.GetFileFromTreeItem(currentItem)
        		Set O = Result1DComplex(FileName)
        		nPoints = O.GetN

    			For n = 0 To nPoints-2
    				If O.GetX(n)<= f And O.GetX(n+1)>=f Then
    					Y = (O.GetYRe(n+1)-O.GetYRe(n))/(O.GetX(n+1)-O.GetX(n))*(f-O.GetX(n))+O.GetYRe(n)
    					Exit For
    				End If
    			Next
    			'Avg = Ysum/Num
    			dBValue = Log(Y)/Log(10)*10
    			Return dBValue
        	End If
    	End If
	   currentItem = Resulttree.GetNextItemName(currentItem)
    Wend
    Return 0
End Function
