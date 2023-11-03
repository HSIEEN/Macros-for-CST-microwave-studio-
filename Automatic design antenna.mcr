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
Sub Main
	Rebuild
	antennaDesign_initialize()
	'Select operating frequency and feed point
	Begin Dialog UserDialog 330,140,"Frquency and feeding part settings" ' %GRID:10,7,1,1
		GroupBox 0,0,330,49,"Frequency settings:",.GroupBox1
		TextBox 140,21,60,14,.Freq
		OKButton 60,105,90,21
		CancelButton 200,105,90,21
		Text 10,21,120,14,"Target frequency:",.Text6
		GroupBox 0,49,330,49,"Feeding part selection:",.GroupBox2
		DropListBox 60,70,210,14,aSolidArray_CST(),.feedingPart
	End Dialog
	Dim dlg As UserDialog
	'Dim tgtFreq As Double
	Dim feedSolid As String
	Dim feed As AntennaElement
	Dim tgtLength As Double
	'Dim ant_material As String
	'ant_material = "Metal/Copper (annealed)"

	dlg.Freq = "1.56"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

	tgtFreq = CDbl(dlg.Freq)
	feedSolid = aSolidArray_CST(dlg.feedingPart)
	Set feed = getElementFromSolid(antElem_arr, feedSolid)
	If feed Is Nothing Then
		MsgBox "Set feed part failed!", vbCritical,"Error"
	End If
	Set ant=New Antenna
	ant.initialize(feed, ant_material, sub_material)
		'Solid.ChangeMaterial "1 Antenna:Antenna", "Metal/Copper (annealed)"
		'feed.solidMaterial = ant_material
		'Solid.ChangeMaterial(feed.solidName,ant_material)
	Dim X As Integer
	X=0
	tgtLength = Round(3e8/(tgtFreq*1e6)/4,1) ' unit mm
	ReportInformationToWindow "Antenna target length: "+CStr(tgtLength)+"mm"
	'tgtLength = 80
	While X<100
		X+=1
		If X>1 Then
			ant.destructor()
		End If
		If ant.constructor(tgtLength, nSolids_CST, Abs(feed.maxPoint(0)-feed.minPoint(0)), _
		Abs(feed.maxPoint(1)-feed.minPoint(1)), Abs(feed.maxPoint(2)-feed.minPoint(2))) = True Then
			ReportInformationToWindow "loop "+Cstr(X)+ _
			" is done and the antenna length target is met, simulation will begin now"
			ant.toHistoryList()
			'ant.patcher()
			'Rebuild, start simulating
			Solver.MeshAdaption(False)
			Solver.SteadyStateLimit(-40)
			Solver.Start

			If	MsgBox("Go on?",vbOkCancel,"Notice")<>vbOK Then
				Exit While
			End If
			'Rebuild
			Plot.Update
		Else
			ReportInformationToWindow "loop "+Cstr(X)+ _
			" is done but the antnena Length is not met, another trial starts"
			'If	MsgBox("Go on?",vbOkCancel,"Notice")<>vbOK Then
			'	Exit While
			'End If
		End If
		'Plot.Update
		'Plot.ExportImage ("E:\image.bmp", 1024, 768)
	Wend

	MsgBox "OOOps"
End Sub
Sub antennaDesign_initialize()
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
	selectMaterial(sub_material,"Pick a material for substrate")
	If MsgBox("Please select all solids contained in the antenna.",vbYesNo,"Notice") <> vbYes Then
    	Exit All
    End If
	SelectSolids_Antenna(aSolidArray_CST(), nSolids_CST)
	'For i=0 to Ubound(aSolidArray_CST)
	'SelectSolids_LIB
	'SelectMaterials_LIB(aMaterialArray_CST(), nMaterials_CST)
	ReportInformationToWindow "Number of solids for antenna elements: "+CStr(nSolids_CST)
	ReDim antElem_arr(nSolids_CST)
	sCommand = ""
	'Construct class instances from solids
	If (nSolids_CST > 0) Then
		For iSolid_CST = 1 To nSolids_CST
			sFullSolidName = aSolidArray_CST(iSolid_CST-1)
			Solid.GetLooseBoundingBoxOfShape(sFullSolidName,xMin,xMax,yMin,yMax,zMin,zMax)
			Set antElem_arr(iSolid_CST-1) = New AntennaElement
			With antElem_arr(iSolid_CST-1)
				.solidName = sFullSolidName
				.solidMaterial = Solid.GetMaterialNameForShape(sFullSolidName)
			If StrComp(.solidMaterial,sub_material)<>0 Then
			'for debug
				 '.setMaterial(sub_material)
				 sCommand = sCommand + .setMaterialPermanently(sub_material)
			End If
				.setStartPoint(xMin,yMin,zMin)
				.setEndPoint(xMax,yMax,zMax)
				.defineVertices()
				.defineEdges()
				.defineFaces()'xMin,yMin,zMin,xMax,yMax,zMax
			End With
			Plot.Update
			'Debug.Print antennaEle
		Next
	Else
		MsgBox "No solids selected, process exit!", vbCritical, "Warning"
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
	If (nSolids_LIB <> UBound(cst_solid_select_name)) Then ReportWarningToWindow("SelectSolids: inconsistent solid count (expected "+CStr(nSolids_LIB)+", found "+Cstr(UBound(cst_solid_select_name))+"), please contact support.")
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
