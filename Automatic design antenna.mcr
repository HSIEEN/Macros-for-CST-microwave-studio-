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
	Dim feedSolid As String
	Dim feed As AntennaElement
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
	While X<100

		antennaConstructor(ant)
		X+=1
		If	MsgBox("Go on?",vbOkCancel,"Notice")<>vbOK Then
			Exit While
		End If
		antennaDestructor(ant)

	Wend

	MsgBox "OOO"
End Sub
Sub antennaElement_initialize()
	If MsgBox("Please select a all solids contained in the antenna.",vbYesNo,"Notice") <> vbYes Then
    	Exit Sub
    End If
	'Dim aSolidArray_CST() As String, nSolids_CST As Integer
	Dim iSolid_CST As Integer
	Dim sFullSolidName As String
	Dim iMaterial_CST As Integer
	'Dim sFullMaterialName As String
	Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, zMin As Double, zMax As Double
	selectMaterial(ant_material,"Pick a material for antenna")
	selectMaterial(sub_material,"Pick a material for substrate")
	SelectSolids_LIB(aSolidArray_CST(), nSolids_CST)
	'SelectSolids_LIB
	'SelectMaterials_LIB(aMaterialArray_CST(), nMaterials_CST)
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
Function antennaConstructor(ant As Antenna)
	'Int((6*Rnd)+1)
	'mFaceNeighbor-->metal face neighbors; mEdgeNeighbor--> mEdgeNeighbors
	'faceNeighbor--> face Neighbors; edgeNeighbor--> edge Neighbor
	'Materials of all elements are set to be vacuum
	Dim i As Integer, j As Integer,ii As Integer, jj As Integer
	'If ant.elementNubmer > 1 Then
		'For i=1 To ant.elementNumber-1

		'Next
	'End If

	Dim n_faceNeighbors As Integer
	Dim n_metalFaceNeighbors As Integer, n_nonMetalFaceNeighbors As Integer
	Dim n_metalEdgeNeighbors As Integer

	Dim faceNeighbors() As AntennaElement
	Dim metalFaceneighbors() As AntennaElement, nonMetalFaceNeighbors() As AntennaElement
	Dim metalEdgeNeighbors() As AntennaElement, nonMetalEdgeNeighbors() As AntennaElement
	Dim currentElement As AntennaElement
	Set currentElement = ant.feedElement

	'metalEdgeNeighbors = currentElement.getNonMetalFaceNeighbors(n_NMneighbor)
	Dim n_metalFaceNeighborsOfNonMetalFaceNeighbors As Integer
	Dim n_metalEdgeNeighborsOfNonMetalFaceNeighbors As Integer
	Dim nonMetalFaceNeighborsStr() As String
	Dim metalFaceNeighborsOfNonMetalFaceNeighbors() As AntennaElement
	Dim metalEdgeNeighborsOfNonMetalFaceNeighbors() As AntennaElement
	'number of valid face neighbors
	Dim n_validFaceNeighbors As Integer
	Dim randomNumber As Integer
	'list of validities of face neighbors
	Dim validityOfFaceNeighbors() As Boolean
	'current element has more than 1 non-metal face neighbors
	Do
		'faceNeighbors = currentElement.getFaceNeighbors(n_faceNeighbors)
		nonMetalFaceNeighbors = currentElement.getNonMetalFaceNeighbors(n_nonMetalFaceNeighbors, nonMetalFaceNeighborsStr)
		n_validFaceNeighbors = n_nonMetalFaceNeighbors
		ReDim validityOfFaceNeighbors(n_validFaceNeighbors)
		For i=0 To n_nonMetalFaceNeighbors-1
			validityOfFaceNeighbors(i)=True
		Next
		For i = 0 To n_nonMetalFaceNeighbors-1
			'If there are more than one metal face neighbors, the non metal face neighbors are invalid
			'for being used as an antenna element
			metalFaceNeighborsOfNonMetalFaceNeighbors = _
			nonMetalFaceNeighbors(i).getMetalFaceNeighbors(n_metalFaceNeighborsOfNonMetalFaceNeighbors)
			If n_metalFaceNeighborsOfNonMetalFaceNeighbors>1 Then
				n_validFaceNeighbors -= 1
				validityOfFaceNeighbors(i)=False
			Else
				'Check whether metal edge neighbors of non-metal face neighbors are among
				'face neighbors of current element
				metalEdgeNeighborsOfNonMetalFaceNeighbors = _
				nonMetalFaceNeighbors(i).getMetalEdgeNeighbors(n_metalEdgeNeighborsOfNonMetalFaceNeighbors)
				If n_metalEdgeNeighborsOfNonMetalFaceNeighbors >= 1 Then
					For j=0 To n_metalEdgeNeighborsOfNonMetalFaceNeighbors-1
						If metalEdgeNeighborsOfNonMetalFaceNeighbors(j).isFaceNeighborWith(currentElement)=False Then
							n_validFaceNeighbors -= 1
							validityOfFaceNeighbors(i)=False
							Exit For
						End If
					Next
				End If
			End If
		Next
		'Random pick one of appropriate non metal face neighbors to be among antenna parts
		If n_validFaceNeighbors>0 Then
			randomNumber=Int((n_validFaceNeighbors)*Rnd)
			jj=0
			For i = 0 To n_nonMetalFaceNeighbors-1
				If validityOfFaceNeighbors(i)=True Then
					If randomNumber=jj Then
						'Set antenna element
						nonMetalFaceNeighbors(i).setMaterial(ant_material)
						'Set antnena
						ant.elementNumber+=1
						ant.conLogics+=nonMetalFaceNeighborsStr(i)
						Set ant.tailElement=nonMetalFaceNeighbors(i)
						If InStr(nonMetalFaceNeighborsStr(i),"x")<>0 Then
							ant.length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
						ElseIf InStr(nonMetalFaceNeighborsStr(i),"y")<>0 Then
							ant.length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
						ElseIf InStr(nonMetalFaceNeighborsStr(i),"z")<>0 Then
							ant.length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
						End If
						'ant.length+=Abs(currentElement.-currentElement.)
						Set currentElement = nonMetalFaceNeighbors(i)

					End If
					jj+=1
				End If
			Next
		Else
			Exit Function
		End If
	Loop
End Function
Function antennaDestructor(ant As Antenna)
	Dim A As AntennaElement, i As Integer
	Set A=ant.feedElement
	For i=0 To ant.elementNumber-2
		Select Case Mid(ant.conLogics,2*i+1,2)
		Case "xn"
			A.xnNeighbor.setMaterial(sub_material)
			Set A=A.xnNeighbor
		Case "xp"
			A.xpNeighbor.setMaterial(sub_material)
			Set A=A.xpNeighbor
		Case "yn"
			A.ynNeighbor.setMaterial(sub_material)
			Set A=A.ynNeighbor
		Case "yp"
			A.ypNeighbor.setMaterial(sub_material)
			Set A=A.ypNeighbor
		Case "zn"
			A.znNeighbor.setMaterial(sub_material)
			Set A=A.znNeighbor
		Case "zp"
			A.zpNeighbor.setMaterial(sub_material)
			Set A=A.zpNeighbor
		End Select

	Next
	ant.antennaInitialize(ant.feedElement)
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
