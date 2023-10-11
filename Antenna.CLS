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
Public Sub antennaInitialize(feed As AntennaElement)
	length = 0
	elemNumber = 1
	conLogics=""
	Set feedElem=feed
	Set tailElem=feed
	'Set feedElement=feed
	If Not feedElem.IsMetal(feedElem.solidMaterial) Then feedElem.setMaterial(ant_material)
End Sub
'Public Sub AddElement(elemDirection As String, elemlength As Double)
	'length += elemlength
	'elemNumber += 1
	'conLogics+=elemDirection
'End Sub
Public Function antennaDestructor()
	Dim A As AntennaElement, i As Integer
	Set A=feedElement
	For i=0 To elementNumber-2
		Select Case Mid(conLogics,2*i+1,2)
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
		Plot.Update

	Next
	antennaInitialize(feedElement)
End Function


Public Function antennaConstructor(tgtLength As Double, solidNumber As Integer, solidxSize As Double, solidySize As Double) As Boolean

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
	randomCycle = Int(tgtLength/(solidNumber*solidxSize*solidySize)*16)+1
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
			NON = _
						nonMetalFaceNeighbors(validFaceNeighborsArray(randomNumber)).getValidNonMetalFaceNeighbors( _
			nMFNeighborsOfnMFNeighbors, n_nMFNeighborsOfnMFNeighbors, nMFNStrOfnMFN, vOFN, n_vFNoFN)
			If NON = True Or n_validFaceNeighbors=1 Then
				k=validFaceNeighborsArray(randomNumber)
			ElseIf NON=False And n_validFaceNeighbors>1 Then
				'IIf(randomNumber>=1,k=validFaceNeighborsArray(randomNumber-1), _
				'k=validFaceNeighborsArray(randomNumber+1))
				If randomNumber>=1 Then
					k=validFaceNeighborsArray(randomNumber-1)
				Else
					k=validFaceNeighborsArray(randomNumber+1)
				End If
			End If
			nonMetalFaceNeighbors(k).setMaterial(ant_material)
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
		End If
		iteration += 1
	Loop
End Function
