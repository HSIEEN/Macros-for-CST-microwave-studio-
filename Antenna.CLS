VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_PredeclaredId = False
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Name = "Antenna"
'#Language "WWB-COM"
Option Explicit
'#Uses "AntennaElement.CLS"
'antenna trace length
Private alength As Double
Property Get length() As Double
	length = alength
End Property
Property Let length(v As Double)
	alength = v
End Property
'antenna material
Private aMaterial As String
Property Get antMaterial() As String
	antMaterial = aMaterial
End Property
Property Let antMaterial(v As String)
	aMaterial = v
End Property
'substrate material
Private sMaterial As String
Property Get subMaterial() As String
	subMaterial = sMaterial
End Property
Property Let subMaterial(v As String)
	sMaterial = v
End Property

'Element number in Antenna
Private elemNumber As Integer
Property Get elementNumber() As Integer
	elementNumber = elemNumber
End Property
Property Let elementNumber(v As Integer)
	elemNumber = v
End Property

'Element connection logic
'xn,xp,yn,yp,zn,zp
Private connectionLogics As String
Property Get conLogics() As String
	 conLogics = connectionLogics
End Property
Property Let conLogics(v As String)
	connectionLogics = v
End Property
'feeding element in antenna
Private feedElem As AntennaElement
Property Get feedElement() As AntennaElement
	 Set feedElement = feedElem
End Property
Property Set feedElement(v As AntennaElement)
	Set feedElem = v
End Property
'last element in antenna
Private tailElem As AntennaElement
Property Get tailElement() As AntennaElement
	 Set tailElement = tailElem
End Property
Property Set tailElement(v As AntennaElement)
	Set tailElem = v
End Property
Public Sub initialize(feed As AntennaElement, eleMaterial As String, substrateMaterial As String)
	length = 0
	elemNumber = 1
	conLogics=""
	antMaterial = eleMaterial
	subMaterial = substrateMaterial
	Set feedElem=feed
	Set tailElem=feed
	'Set feedElement=feed
	If Not feedElem.IsMetal() Then feedElem.setMaterial(antMaterial)
End Sub
'Public Sub AddElement(elemDirection As String, elemlength As Double)
	'length += elemlength
	'elemNumber += 1
	'conLogics+=elemDirection
'End Sub
Public Function destructor()
	Dim A As AntennaElement, i As Integer
	Set A=feedElement
	For i=0 To elementNumber-2
		Select Case Mid(conLogics,2*i+1,2)
		Case "xn"
			A.xnNeighbor.setMaterial(subMaterial)
			Set A=A.xnNeighbor
		Case "xp"
			A.xpNeighbor.setMaterial(subMaterial)
			Set A=A.xpNeighbor
		Case "yn"
			A.ynNeighbor.setMaterial(subMaterial)
			Set A=A.ynNeighbor
		Case "yp"
			A.ypNeighbor.setMaterial(subMaterial)
			Set A=A.ypNeighbor
		Case "zn"
			A.znNeighbor.setMaterial(subMaterial)
			Set A=A.znNeighbor
		Case "zp"
			A.zpNeighbor.setMaterial(subMaterial)
			Set A=A.zpNeighbor
		End Select
		Plot.Update

	Next
	Dim historyCaption As String
	Dim sCommand As String
	historyCaption = "$SA$"
	sCommand = ""
	AddToHistory(historyCaption, sCommand)
	initialize(feedElement, antMaterial, subMaterial)
End Function
Public Function toHistoryList()
	Dim A As AntennaElement, i As Integer
	Dim historyCaption As String
	Dim sCommand As String
	sCommand = ""
	Set A=feedElement
	sCommand  = sCommand + A.setMaterialPermanently(antMaterial)
	For i=0 To elementNumber-2
		Select Case Mid(conLogics,2*i+1,2)
		Case "xn"
			sCommand  = sCommand + A.xnNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.xnNeighbor
		Case "xp"
			sCommand  = sCommand + A.xpNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.xpNeighbor
		Case "yn"
			sCommand  = sCommand + A.ynNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.ynNeighbor
		Case "yp"
			sCommand  = sCommand + A.ypNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.ypNeighbor
		Case "zn"
			sCommand  = sCommand + A.znNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.znNeighbor
		Case "zp"
			sCommand  = sCommand + A.zpNeighbor.setMaterialPermanently(antMaterial)
			Set A=A.zpNeighbor
		End Select


	Next
	historyCaption  = "$SA$"
	AddToHistory(historyCaption, sCommand)
	Plot.Update
	'initialize(feedElement, antMaterial, subMaterial)
End Function


Public Function constructor(tgtLength As Double, solidNumber As Integer, solidxSize As Double, _
solidySize As Double, solidzSize As Double) As Boolean

	Dim i As Integer, j As Integer, k As Integer, m As Integer
	'Dim n_faceNeighbors As Integer
	Dim n_nonMetalFaceNeighbors As Integer, n_nMFNeighborsOfnMFNeighbors As Integer
	Dim nonMetalFaceNeighbors() As AntennaElement, nMFNeighborsOfnMFNeighbors() As AntennaElement
	Dim nonMetalEdgeNeighbors() As AntennaElement
	Dim currentElement As AntennaElement

	Set currentElement = feedElement

	Dim nonMetalFaceNeighborsStr() As String, nMFNStrOfnMFN() As String
	Dim n_validFaceNeighbors As Integer, n_vFNoFN As Integer
	Dim randomNumber As Integer
	'list of validities of face neighbors
	'Dim validityOfFaceNeighbors() As Boolean, vOFN() As Boolean, NON As Boolean
	Dim validFaceNeighborsArray() As Integer, vOFN() As Integer, NON As Boolean
	Dim iteration As Integer
	Dim remainder As Integer
	Dim randomCycle As Integer
	Dim lastRandomNumber As Integer

	iteration = 0
	lastRandomNumber = 0
	randomCycle = Int(tgtLength^2/(solidxSize+solidySize+solidzSize)*solidNumber/3e5)+1
	ReportInformationToWindow "Random Cycle is "+CStr(randomCycle)

	Do
		'nonMetalFaceNeighbors = _
		'currentElement.getNonMetalFaceNeighbors(n_nonMetalFaceNeighbors, nonMetalFaceNeighborsStr)

		currentElement.getValidNonMetalFaceNeighbors(nonMetalFaceNeighbors, n_nonMetalFaceNeighbors, _
		nonMetalFaceNeighborsStr, validFaceNeighborsArray, n_validFaceNeighbors)
		'Random pick one of appropriate non metal face neighbors to be used as an antenna element
		If n_validFaceNeighbors>0 And ant.length<tgtLength Then
			remainder = iteration Mod (Int(randomCycle*Rnd)+1)
			If remainder = 0 Then
				randomNumber=Int((n_validFaceNeighbors)*Rnd)
				lastRandomNumber = randomNumber
			ElseIf remainder<>0 And n_validFaceNeighbors>lastRandomNumber Then
				randomNumber = lastRandomNumber
			ElseIf remainder<>0 And n_validFaceNeighbors<=lastRandomNumber Then
				randomNumber = 0
			End If
			'Determin whether the condidate neighbor has a valid face neighbor, if not, try
			'to pick another one to be antenna elment
			NON = nonMetalFaceNeighbors(validFaceNeighborsArray(randomNumber)).getValidNonMetalFaceNeighbors( _
			nMFNeighborsOfnMFNeighbors, n_nMFNeighborsOfnMFNeighbors, nMFNStrOfnMFN, vOFN, n_vFNoFN)
			'If current element gets just one valid face neighbor, set the only one face neighbor as an antenna element
			'Or if the random valid face neighbor has at least two valid face neighbors, set the random face neighbor as an antenna element
			'If n_validFaceNeighbors=1 Or (NON=True And n_vFNoFN>1) Then
				'k=validFaceNeighborsArray(randomNumber)
			'If get at least two valid face neighbors, and a valid face neighbor gets no valid neighbors, try to
			'pick another valid face neighbor to be antenna element
			'ElseIf NON=False And n_validFaceNeighbors>1 Then
			If NON=False And n_validFaceNeighbors>1 Then
				'IIf(randomNumber>=1,k=validFaceNeighborsArray(randomNumber-1), _
				'k=validFaceNeighborsArray(randomNumber+1))
				If randomNumber>=1 Then
					k=validFaceNeighborsArray(randomNumber-1)
				Else
					k=validFaceNeighborsArray(randomNumber+1)
				End If
			Else
				k=validFaceNeighborsArray(randomNumber)
			End If
			nonMetalFaceNeighbors(k).setMaterial(antMaterial)
			Plot.Update
			'Set antnena
			elementNumber+=1
			conLogics+=nonMetalFaceNeighborsStr(k)
			Set tailElement=nonMetalFaceNeighbors(k)
			If InStr(nonMetalFaceNeighborsStr(k),"x")<>0 Then
				length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
			ElseIf InStr(nonMetalFaceNeighborsStr(k),"y")<>0 Then
				length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
			ElseIf InStr(nonMetalFaceNeighborsStr(k),"z")<>0 Then
				length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
			End If
			'ant.length+=Abs(currentElement.-currentElement.)
			Set currentElement = nonMetalFaceNeighbors(k)

					'End If
					'j+=1
				'End If
			'Next
		ElseIf ant.length>=tgtLength Then
			ReportInformationToWindow conLogics
			ReportInformationToWindow "Antenna Length: "+CStr(Round(length,2))
			Return True
			'Exit Function
		ElseIf n_validFaceNeighbors<=0 And length<tgtLength-1 Then
			ReportInformationToWindow conLogics
			ReportInformationToWindow "Antenna Length: "+CStr(Round(length,2))
			Return False
		Else
			ReportInformationToWindow conLogics
			ReportInformationToWindow "Antenna Length: "+CStr(Round(length,2))
			Return True
		End If
		iteration += 1
	Loop
End Function
Function patcher()
	Dim A As AntennaElement, i As Integer, j As Integer
	Dim nonMetalFaceNeighbors() As AntennaElement
	Dim n_nonMetalFaceNeighbors As Integer', n_validFaceNeighbors
	Dim nonMetalFaceNeighborsStr() As String
	'Dim validFaceNeighborsArray() As Integer
	Set A=feedElement
	For i=0 To elementNumber-2
		nonMetalFaceNeighbors = _
		A.getNonMetalFaceNeighbors(n_nonMetalFaceNeighbors, nonMetalFaceNeighborsStr)
		If n_nonMetalFaceNeighbors> 0 Then
			For j=0 To n_nonMetalFaceNeighbors-1
				Select Case nonMetalFaceNeighborsStr(j)
				Case "xn"
					'A.xnNeighbor.setMaterial(subMaterial)
					If A.xnNeighbor.isMetal()=False Then
						A.xnNeighbor.setMaterial(antMaterial)
					End If
				Case "xp"
					'A.xpNeighbor.setMaterial(subMaterial)
					If A.xpNeighbor.isMetal()=False Then
						A.xpNeighbor.setMaterial(antMaterial)
					End If
				Case "yn"
					'A.ynNeighbor.setMaterial(subMaterial)
					If A.ynNeighbor.isMetal()=False Then
						A.ynNeighbor.setMaterial(antMaterial)
					End If
				Case "yp"
					'A.ypNeighbor.setMaterial(subMaterial)
					If A.ypNeighbor.isMetal()=False Then
						A.ypNeighbor.setMaterial(antMaterial)
					End If
				Case "zn"
					'A.znNeighbor.setMaterial(subMaterial)
					If A.znNeighbor.isMetal()=False Then
						A.znNeighbor.setMaterial(antMaterial)
					End If
				Case "zp"
					'A.zpNeighbor.setMaterial(subMaterial)
					If A.zpNeighbor.isMetal()=False Then
						A.zpNeighbor.setMaterial(antMaterial)
					End If
				End Select
			Next
		End If
		Select Case Mid(conLogics,2*i+1,2)
		Case "xn"
			'A.xnNeighbor.setMaterial(subMaterial)
			Set A=A.xnNeighbor
		Case "xp"
			'A.xpNeighbor.setMaterial(subMaterial)
			Set A=A.xpNeighbor
		Case "yn"
			'A.ynNeighbor.setMaterial(subMaterial)
			Set A=A.ynNeighbor
		Case "yp"
			'A.ypNeighbor.setMaterial(subMaterial)
			Set A=A.ypNeighbor
		Case "zn"
			'A.znNeighbor.setMaterial(subMaterial)
			Set A=A.znNeighbor
		Case "zp"
			'A.zpNeighbor.setMaterial(subMaterial)
			Set A=A.zpNeighbor
		End Select
		Plot.Update

	Next

End Function
