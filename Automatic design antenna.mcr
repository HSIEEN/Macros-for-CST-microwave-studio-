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
	ant.antennaInitialize(feed)
		'Solid.ChangeMaterial "1 Antenna:Antenna", "Metal/Copper (annealed)"
		'feed.solidMaterial = ant_material
		'Solid.ChangeMaterial(feed.solidName,ant_material)
	Dim X As Integer
	X=0
	tgtLength = Round(3e8/(tgtFreq*1e6)/1,1) ' unit mm
	ReportInformationToWindow "Antenna initial length: "+CStr(tgtLength)+"mm"
	'tgtLength = 80
	While X<100
		X+=1
		If ant.antennaConstructor(tgtLength, nSolids_CST, 1, 1) = True Then
			ReportInformationToWindow "loop "+Cstr(X)+ _
			" is done and the antenna length target is met, simulation will begin now"
			If	MsgBox("Go on?",vbOkCancel,"Notice")<>vbOK Then
				Exit While
			End If
		Else
			ReportInformationToWindow "loop "+Cstr(X)+ _
			" is done but the antnena Length is not met, another trial starts"
		End If
		'Plot.Update
		'Plot.ExportImage ("E:\image.bmp", 1024, 768)
		ant.antennaDestructor()

	Wend

	MsgBox "OOOps"
End Sub
Sub antennaDesign_initialize()

	'Dim aSolidArray_CST() As String, nSolids_CST As Integer
	Dim iSolid_CST As Integer
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
	SelectSolids_LIB(aSolidArray_CST(), nSolids_CST)
	'SelectSolids_LIB
	'SelectMaterials_LIB(aMaterialArray_CST(), nMaterials_CST)
	ReportInformationToWindow "Number of solids for antenna elements: "+CStr(nSolids_CST)
	ReDim antElem_arr(nSolids_CST)

	'Construct class instances from solids
	If (nSolids_CST > 0) Then
		For iSolid_CST = 1 To nSolids_CST
			sFullSolidName = aSolidArray_CST(iSolid_CST-1)
			Solid.GetLooseBoundingBoxOfShape(sFullSolidName,xMin,xMax,yMin,yMax,zMin,zMax)
			Set antElem_arr(iSolid_CST-1) = New AntennaElement
			With antElem_arr(iSolid_CST-1)
				.solidName = sFullSolidName
				.solidMaterial = Solid.GetMaterialNameForShape(sFullSolidName)
			'If .IsMetal(.solidMaterial) = True Then
				.setMaterial(sub_material)
			'End If
				.setStartPoint(xMin,yMin,zMin)
				.setEndPoint(xMax,yMax,zMax)
				.defineVertices()
				.defineEdges()
				.defineFaces()'xMin,yMin,zMin,xMax,yMax,zMax
			End With

			'Debug.Print antennaEle
		Next
	Else
		MsgBox "No solids selected, process exit!", vbCritical, "Warning"
		Exit All
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
