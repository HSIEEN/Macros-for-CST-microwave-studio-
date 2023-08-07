'#Language "WWB-COM"

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3D.lib"
'#Uses "AntennaElement.CLS"
Public antennaElement_array() As AntennaElement
Public aSolidArray_CST() As String
Public nSolids_CST As Integer
Sub Main
	antennaElement_initialize()
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
	Dim tgtFreq As Double
	Dim feedingSolid As String
	Dim feeding As AntennaElement
	Dim ant_material As String
	ant_material = "Metal/Copper (annealed)"

	dlg.Freq = "1.56"
	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

	tgtFreq = CDbl(dlg.Freq)
	feedingSolid = aSolidArray_CST(dlg.feedingPart)
	Set feeding = getElementFromSolid(antennaElement_array, feedingSolid)
	If feeding Is Nothing Then
		MsgBox "Set feeding part failed!", vbCritical,"Error"
	End If
	'If feeding.solidMaterial = "Vacuum" Then
	If Not feeding.IsMetal(feeding.solidMaterial) Then feeding.setMaterial(ant_material)
		'Solid.ChangeMaterial "1 Antenna:Antenna", "Metal/Copper (annealed)"
		'feeding.solidMaterial = ant_material
		'Solid.ChangeMaterial(feeding.solidName,ant_material)
	antennaGenerator(feeding)
	MsgBox "OOO"
End Sub
Sub antennaElement_initialize()
	If MsgBox("Please select a all solids contained in the antenna.",vbYesNo,"Notice") <> vbYes Then
    	Exit Sub
    End If

	'Dim aSolidArray_CST() As String, nSolids_CST As Integer
	Dim iSolid_CST As Integer
	Dim sFullSolidName As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double

	SelectSolids_LIB(aSolidArray_CST(), nSolids_CST)
	ReDim antennaElement_array(nSolids_CST)

	'Construct class instances from solids
	If (nSolids_CST > 0) Then
		For iSolid_CST = 1 To nSolids_CST
			sFullSolidName = aSolidArray_CST(iSolid_CST-1)
			Solid.GetLooseBoundingBoxOfShape(sFullSolidName,xMin,xMax,yMin,yMax,zMin,zMax)
			Set antennaElement_array(iSolid_CST-1) = New AntennaElement
			With antennaElement_array(iSolid_CST-1)
				.solidName = sFullSolidName
				.solidMaterial = Solid.GetMaterialNameForShape(sFullSolidName)
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
			'createFaceNeighbors(antennaElement_array(iSolid_CST-1),antennaElement_array(ii-1))
			antennaElement_array(iSolid_CST-1).createFaceNeighborWith(antennaElement_array(ii-1))
			'createEdgeNeighbors
			antennaElement_array(iSolid_CST-1).createEdgeNeighborWith(antennaElement_array(ii-1))

		Next
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
Function antennaGenerator(feed As AntennaElement)
	'Int((6*Rnd)+1)
	'mFaceNeighbor-->metal face neighbors; mEdgeNeighbor--> mEdgeNeighbors
	'faceNeighbor--> face Neighbors; edgeNeighbor--> edge Neighbor
	Dim n_neighbor As Integer, n_NMneighbor As Integer, n_Mneighbor As Integer
	Dim i As Integer
	Dim neighbors() As AntennaElement, nonMetalNeighbors() As AntennaElement
	Dim MetalNeighbor() As AntennaElement
	Dim current_element As AntennaElement
	Set current_element = feed
	neighbors = current_element.getFaceNeighbors(n_neighbor)
	nonMetalNeighbors = current_element.getNonMetalFaceNeighbors(n_NMneighbor)
	'While n_NMneighbor >= 1
		'For i = 0 To UBound(nonMetalNeighbor)

		'Next

	'Wend

End Function
