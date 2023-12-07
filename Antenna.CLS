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
Public logFile As String
Public prjPath As String
Public log2Window As Boolean
Public fallBackRatio As Double
Public availableStr As String
'Public lengthContainer() As Integer
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
	Set feedElement=feed
	Set tailElement=feed
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
	'length = 0
	'conLogics = ""
	'elementNumber = 1
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
	log2Window=False
	Dim i As Integer, j As Integer, k As Integer, m As Integer
	'Dim n_faceNeighbors As Integer
	Dim n_nonMetalFaceNeighbors As Integer, n_nMFNeighborsOfnMFNeighbors As Integer
	Dim nonMetalFaceNeighbors() As AntennaElement, nMFNeighborsOfnMFNeighbors() As AntennaElement
	Dim nonMetalEdgeNeighbors() As AntennaElement
	Dim currentElement As AntennaElement
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
	'Dim containerIndex As Integer
	Dim lastLength As Double
	Dim tgtIncrement As Double
	Dim maxSize As Double

	If solidxSize>=solidySize And solidxSize>=solidzSize Then
		maxSize = solidxSize
	ElseIf solidySize>=solidxSize And solidySize>=solidzSize Then
		maxSize = solidySize
	ElseIf solidzSize>=solidySize And solidzSize>=solidxSize Then
		maxSize = solidzSize
	End If
	ReportInformationToWindow "Maximum dimension is: " & CStr(Round(maxSize,2))+"mm"
	'containerIndex = 0
	fallBackRatio = 0
	tgtIncrement = 0
	'ReDim lengthContainer(3)

	Set currentElement = tailElement

	prjPath = GetProjectPath("Project")
   	logFile = prjPath + "\Routing log.txt"
   	'If already existed, delete it firstly
   	If Dir(logFile) <> "" Then
		Kill logFile
   	End If
   	Open logFile For Output As #2
	'######## when target length is less than ant.length #####'
	If tgtLength<length Then
		Do
			currentElement.setMaterial(subMaterial)
			elementNumber -= 1
			Select Case Right(conLogics,2)
			Case "xn"
				'A.xnNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.xpNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
				'End If

			Case "xp"
				'A.xpNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.xnNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
				'End If

			Case "yn"
				'A.ynNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.ypNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
				'End If

			Case "yp"
				'A.ypNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.ynNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
				'End If

			Case "zn"
				'A.znNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.zpNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
				'End If

			Case "zp"
				'A.zpNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.znNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
				'End If

			End Select
			conLogics = Left(conLogics, Len(conLogics)-2)
			Plot.Update
		Loop Until tgtLength>=length
		Set tailElement = currentElement
		'Open logFile For Append As #2
		Print #2, conLogics
		Close #2
		Return True
	End If
	'#### When target length is greater than ant.length ####
	iteration = 0
	lastRandomNumber = 0
	Dim pureNonMetalElements As Integer
	Dim needReEstimatingAvailableElements As Boolean
	Dim deleteIndex(8) As Integer
	'Dim ii As Integer
	Dim validFaceNeighbors() As AntennaElement
	Dim validFaceNeighborsStr() As String
	Dim tempObj() As AntennaElement
	Dim tempStr() As String
	'Dim hasBeenCorrected As Boolean
	'Dim availableStr As String

	'Dim cFlag As Boolean

	Do
		randomCycle = Int(tgtLength^2/(solidxSize+solidySize+solidzSize)*solidNumber/4e5)+1
		randomCycle = randomCycle-Int(ant.length/tgtLength*randomCycle) Mod randomCycle
		ReportInformationToWindow "Random Cycle is "+CStr(randomCycle)
		currentElement.getValidNonMetalFaceNeighbors(nonMetalFaceNeighbors, n_nonMetalFaceNeighbors, _
		nonMetalFaceNeighborsStr, validFaceNeighborsArray, n_validFaceNeighbors)
		'Estimate if the candidate neighbor can split all available elements into two parts, if yes
		'estimate all available elements for all neighbors and then pick one meeting the target available elements
		ReDim tempObj(n_validFaceNeighbors)
		ReDim tempStr(n_validFaceNeighbors)
		'reload valid face neighbors to a new array
		For i=0 To n_validFaceNeighbors-1
			Set tempObj(i) = nonMetalFaceNeighbors(validFaceNeighborsArray(i))
			tempStr(i) = nonMetalFaceNeighborsStr(validFaceNeighborsArray(i))
		Next
		'Random pick one of appropriate non metal face neighbors to be used as an antenna element
		If n_validFaceNeighbors>0 And ant.length<tgtLength Then
			remainder = iteration Mod (Int(randomCycle*Rnd)+1)
			If ant.length > tgtLength*0.3 Then
				needReEstimatingAvailableElements = needToReEstimateAvailableElementsOnGoingForward()
				'Determin whether the candidate neighbor has a valid face neighbor, if not, try
				'to pick another one to be antenna elment
				If needReEstimatingAvailableElements=True Then
					ReportInformationToWindow "Need to re-estimate available elements number."
					j=0
					For i=0 To n_validFaceNeighbors-1
						availableStr=""
						'pureNonMetalElements=nonMetalFaceNeighbors(validFaceNeighborsArray(i)).getNumberofPureNonMetalElements(availableStr)
						pureNonMetalElements=tempObj(i).getNumberofPureNonMetalElements(availableStr)
						ReportInformationToWindow "Available elements for neighbor-"&CStr(i)&": "&CStr(pureNonMetalElements)
						'When the number of available elements counting from one neighbor does not meet the target, delete this neighbor
						'from valid face neighbors list(set to be nothing)
						If pureNonMetalElements < CInt((tgtLength-length)*maxSize*2.5) Then
							Set tempObj(i) = Nothing
							tempStr(i)=""
							j+=1
						End If
					Next
					n_validFaceNeighbors-=j
				End If
			End If

			If n_validFaceNeighbors>0 Then
				ReDim validFaceNeighbors(n_validFaceNeighbors)
				ReDim validFaceNeighborsStr(n_validFaceNeighbors)
				j=0
				For i=0 To UBound(tempObj)-1
					If Not tempObj(i) Is Nothing Then
						Set validFaceNeighbors(j)=tempObj(i)
						validFaceNeighborsStr(j)=tempStr(i)
						j+=1
					End If
				Next

				If remainder = 0 Then
					randomNumber=Int((n_validFaceNeighbors)*Rnd)
					lastRandomNumber = randomNumber
				ElseIf remainder<>0 And n_validFaceNeighbors>lastRandomNumber Then
					randomNumber = lastRandomNumber
				ElseIf remainder<>0 And n_validFaceNeighbors<=lastRandomNumber Then
					randomNumber = 0
				End If
				NON = validFaceNeighbors(randomNumber).getValidNonMetalFaceNeighbors( _
				nMFNeighborsOfnMFNeighbors, n_nMFNeighborsOfnMFNeighbors, nMFNStrOfnMFN, vOFN, n_vFNoFN)
				If NON=False And n_validFaceNeighbors>1 Then
					If randomNumber>=1 Then
						k=randomNumber-1
					Else
						k=randomNumber+1
					End If
				Else
					k=randomNumber
				End If
				validFaceNeighbors(k).setMaterial(antMaterial)
				Plot.Update
				'Set antnena
				elementNumber+=1
				conLogics = conLogics & validFaceNeighborsStr(k)
				Set tailElement = validFaceNeighbors(k)
				If InStr(validFaceNeighborsStr(k),"x")<>0 Then
					length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
				ElseIf InStr(validFaceNeighborsStr(k),"y")<>0 Then
					length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
				ElseIf InStr(validFaceNeighborsStr(k),"z")<>0 Then
					length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
				End If
				'ant.length+=Abs(currentElement.-currentElement.)
				Set currentElement = validFaceNeighbors(k)
				If log2Window=True Then ReportInformationToWindow "0$Antenna Length: "+CStr(Round(length,2))+"mm"
			Else
				GoTo Backward
			End If

		ElseIf ant.length>=tgtLength Then

			'Open logFile For Append As #2
			Print #2, conLogics
			Close #2
			ReportInformationToWindow conLogics
			ReportInformationToWindow "1$Antenna Length: "+CStr(Round(length,2))+"mm"
			Return True
			'Exit Function
		ElseIf n_validFaceNeighbors<=0 And length<tgtLength-1 Then
			Backward:
				'Open logFile For Append As #2
				Print #2, conLogics'+tailElement.solidName
				'Close #2
				'Call corrector to correct the routing of tail part
				ReportInformationToWindow "Iteration #"+CStr(iteration)+"-----------------"
				ReportInformationToWindow conLogics
				ReportInformationToWindow "2$Antenna Length: "+CStr(Round(length,2))+"mm"
				'tgtIncrement = (tgtLength-length)*0.2+2
				'tgtIncrement = 15
				'If (containerIndex > 0) And (containerIndex Mod 3 =0) Then
				If ant.length>0.3*tgtLength Then
					'If fallBackRatio<0.8 Then fallBackRatio+=0.04 Else fallBackRatio=0.8
					fallBackRatio = 0
				Else
					If fallBackRatio<0.4 Then fallBackRatio+=0.05 Else fallBackRatio=0.4
				End If

				'End If
				'lastLength=length
				'containerIndex += 1
				Set tailElement=corrector(tgtLength, maxSize, fallBackRatio)
				If Not tailElement Is Nothing Then
					'Nothing to do
				Else
					Close #2
					Return False
				End If
				Set currentElement = tailElement
			'Return False
		Else
			'Open logFile For Append As #2
			Print #2, conLogics
			Close #2
			ReportInformationToWindow conLogics
			ReportInformationToWindow "3$Antenna Length: "+CStr(Round(length,2))+"mm"
			Return True
		End If
		iteration += 1
	Loop Until iteration>=10000
	ReportInformationToWindow "Over 10000 times trials, return failure"
	Close #2
	Return False

End Function
Function corrector(tgtLength As Double, maxSize As Double, fBR As Double) As AntennaElement
	log2Window=False
	availableStr = ""
	Dim currentElement As AntennaElement
	Dim i As Integer, ii As Integer, iteration As Integer, jj As Integer
	Dim nonMetalFaceNeighbors() As AntennaElement
	Dim n_nonMetalFaceNeighbors As Integer, nonMetalFaceNeighborsStr() As String
	Dim validFaceNeighborsArray() As Integer, n_validFaceNeighbors As Integer
	Dim currentNeighborStr As String

	Dim lineRead As String, lineCount As Integer, lineNumber As Integer
	Dim usedNeighborStr As String
	Dim logicLength As Integer
	Dim initialLength As Double

	Dim pureNonMetalElements As Integer
	'Dim currentConnectionStatus As Boolean
	Dim needToReestimate As Boolean
	Dim tgtCondition As Boolean
	'Dim previousConnectionStatus As Boolean

	lineCount = 0
	Open logFile For Input As #1
	'Count line number of Routing log.txt
	While Not EOF(1)
		Line Input #1, lineRead
		lineCount += 1
	Wend
	Close #1
	'tailElement.setMaterial(subMaterial)
	Set currentElement = tailElement
	'i = Len(conLogics)
	iteration = -1
	needToReestimate = False
	'previousConnectionStatus = False
	pureNonMetalElements = 0
	logicLength = Len(conLogics)
	initialLength = length

	While Not currentElement Is Nothing ' And iteration*2<logicLength/4
		Do

			If Len(conLogics)<2 Then
				Return Nothing
				'Exit Function
			End If
			currentElement.setMaterial(subMaterial)
			elementNumber -= 1
			currentNeighborStr = Right(conLogics,2)
			Plot.Update
			If fBR=0 Then
				'needToReestimate = currentElement.needToEstimateAvailableElementsAllOverAgain(currentNeighborStr, availableStr)
				needToReestimate = needToReEstimateAvailableElementsOnGoingBackward()
				'Estimate if need to re-estimate the pure-non metal elements all over again
				If needToReestimate=True Then
					availableStr = ""
					pureNonMetalElements=currentElement.getNumberofPureNonMetalElements(availableStr)
				Else
					pureNonMetalElements+=currentElement.getNumberofPureNonMetalNeighbors(availableStr)
				End If
			End If
			'previousConnectionStatus = currentConnectionStatus
			'cFlag=currentElement.flag
			conLogics = Left(conLogics, Len(conLogics)-2)
			Select Case currentNeighborStr
			Case "xn"
				'A.xnNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.xpNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
				'End If

			Case "xp"
				'A.xpNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.xnNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
				'End If

			Case "yn"
				'A.ynNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.ypNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
				'End If

			Case "yp"
				'A.ypNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.ynNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
				'End If

			Case "zn"
				'A.znNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.zpNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
				'End If

			Case "zp"
				'A.zpNeighbor.setMaterial(subMaterial)
				Set currentElement=currentElement.znNeighbor
				'If iteration>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
				'End If

			End Select
			Set tailElement = currentElement
			ReportInformationToWindow "4$Antenna Length: "+CStr(Round(length,2))+"mm"
			'fall back until the difference between current lenghth and the inital length over the initial
			'lenght is greater than what is requred, or fBR
			'currentElement.setMaterial(subMaterial)
			'pureNonMetalElements+=currentElement.getNumberofPureNonMetalNeighbors()
			ReportInformationToWindow "4$Target available antenna Elements is " & CInt((tgtLength-length)*maxSize*2.5)
			ReportInformationToWindow "4$Available antenna Elements is " & CStr(pureNonMetalElements)
			iteration += 1
			If fBR > 0 Then
				tgtCondition=((initialLength-length)/initialLength > fBR)
			Else
				tgtCondition=(pureNonMetalElements >= CInt((tgtLength-length)*maxSize*2.5))
			End If
		Loop Until tgtCondition'pureNonMetalElements >= CInt((tgtLength-length)*maxSize*2.5)'Until (initialLength-length)/initialLength > fBR

		currentElement.getValidNonMetalFaceNeighbors(nonMetalFaceNeighbors, n_nonMetalFaceNeighbors, _
		nonMetalFaceNeighborsStr, validFaceNeighborsArray, n_validFaceNeighbors)
		If n_validFaceNeighbors>1 Then
			'Check routing log to make sure if the lineCount is greater than 1. If so, collect all connection logics
			usedNeighborStr = currentNeighborStr
			lineNumber = lineCount
			Open logFile For Input As #1
			While lineNumber > 1
				Line Input #1, lineRead
				If Len(lineRead) >= logicLength-iteration*2  _
				And StrComp(conLogics, Left(lineRead, logicLength-iteration*2-2))=0 And _
				InStr(usedNeighborStr, Mid(lineRead, logicLength-iteration*2-1, 2))=0 Then
					usedNeighborStr = usedNeighborStr & Mid(lineRead, logicLength-iteration*2-1, 2)
				End If
				lineNumber -=1
			Wend
			Close #1
			'need change it to be random instead of sequential!!!!!
			For ii=0 To n_validFaceNeighbors-1

				'Only if the connnection logic is not in the used connection logic collection, do the next step
				If InStr(usedNeighborStr, nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)))=0 Then
					Select Case nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii))
					Case "xn"
						'A.xnNeighbor.setMaterial(subMaterial)
						length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
						currentElement.xnNeighbor.setMaterial(antMaterial)
						conLogics += "xn"
						elementNumber += 1
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.xnNeighbor

					Case "xp"
						length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
						currentElement.xpNeighbor.setMaterial(antMaterial)
						conLogics += "xp"
						elementNumber += 1
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.xpNeighbor
					Case "yn"
						length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
						currentElement.ynNeighbor.setMaterial(antMaterial)
						conLogics = conLogics+"yn"
						elementNumber += 1
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.ynNeighbor
					Case "yp"
						length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
						currentElement.ypNeighbor.setMaterial(antMaterial)
						conLogics += "yp"
						elementNumber += 1
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.ypNeighbor
					Case "zn"
						length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
						currentElement.znNeighbor.setMaterial(antMaterial)
						elementNumber += 1
						conLogics += "zn"
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.znNeighbor
					Case "zp"
						length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
						currentElement.zpNeighbor.setMaterial(antMaterial)
						conLogics += "zp"
						elementNumber += 1
						If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
						Return currentElement.zpNeighbor
					End Select
				End If
			Next
		End If
		'iteration += 1
	Wend
	Return Nothing
End Function
'Estimate if neighbors on both sides of the element are possible to be connected by at least one routing path
'connectionStr is the connection relationship string between current element and the last one
Public Function needToReEstimateAvailableElementsOnGoingBackward() As Boolean
	'Dim NeighborNum As Integer, i As Integer, n As Integer
	'Dim neighbors() As AntennaElement
	If tailElement.isMetal()=True Then
		ReportInformationToWindow "The tail element is metal, please switch to a nonMetal one"
		Return False
	End If

	Select Case Right(conLogics,2)
	Case "xn"
		If tailElement.xnNeighbor Is Nothing Then Return False
		If tailElement.xnNeighbor.isPureNonMetal()=False Then Return False
		If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$")=0 _
			And tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True _
			And tailElement.xnNeighbor.xnNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		If Not tailElement.ypNeighbor Is Nothing Then
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$")=0 _
				And tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True _
				And tailElement.ypNeighbor.ypNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ynNeighbor Is Nothing Then
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$")=0 _
				And tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True _
				And tailElement.ynNeighbor.ynNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.znNeighbor Is Nothing Then
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$")=0 _
				And tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True _
				And tailElement.znNeighbor.znNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.zpNeighbor Is Nothing Then
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If  InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$")=0 _
				And tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True _
				And tailElement.zpNeighbor.zpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	Case "xp"
		If tailElement.xpNeighbor Is Nothing Then Return False
		If tailElement.xpNeighbor.isPureNonMetal()=False Then Return False
		'If Not xpNeihgbor Is Nothing Then
		If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$")=0 _
			And tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True _
			And tailElement.xpNeighbor.xpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		'End If
		If Not tailElement.ypNeighbor Is Nothing Then
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$")=0 _
				And tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True _
				And tailElement.ypNeighbor.ypNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ynNeighbor Is Nothing Then
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$")=0 _
				And tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True _
				And tailElement.ynNeighbor.ynNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.znNeighbor Is Nothing Then
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$")=0 _
				And tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True _
				And tailElement.znNeighbor.znNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.zpNeighbor Is Nothing Then
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$")=0 _
				And tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True _
				And tailElement.zpNeighbor.zpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	Case "yn"
		If tailElement.ynNeighbor Is Nothing Then Return False
		If tailElement.ynNeighbor.isPureNonMetal()=False Then Return False
		'If Not ynNeihgbor Is Nothing Then
		If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$")=0 _
			And tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True _
			And tailElement.ynNeighbor.ynNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		'End If
		If Not tailElement.xpNeighbor Is Nothing Then
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$")=0 _
				And tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True _
				And tailElement.xpNeighbor.xpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.xnNeighbor Is Nothing Then
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$")=0 _
				And tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True _
				And tailElement.xnNeighbor.xnNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.znNeighbor Is Nothing Then
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$")=0 _
				And tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True _
				And tailElement.znNeighbor.znNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.zpNeighbor Is Nothing Then
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$")=0 _
				And tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True _
				And tailElement.zpNeighbor.zpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	Case "yp"
		If tailElement.ypNeighbor Is Nothing Then Return False
		If tailElement.ypNeighbor.isPureNonMetal()=False Then Return False
		'If Not ynNeihgbor Is Nothing Then
		If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$")=0 _
			And tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True _
			And tailElement.ypNeighbor.ypNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		'End If
		If Not tailElement.xpNeighbor Is Nothing Then
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$")=0 _
				And tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True _
				And tailElement.xpNeighbor.xpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.xnNeighbor Is Nothing Then
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$")=0 _
				And tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True _
				And tailElement.xnNeighbor.xnNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.znNeighbor Is Nothing Then
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$")=0 _
				And tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True _
				And tailElement.znNeighbor.znNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.zpNeighbor Is Nothing Then
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$")=0 _
				And tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True _
				And tailElement.zpNeighbor.zpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	Case "zn"
		If tailElement.znNeighbor Is Nothing Then Return False
		If tailElement.znNeighbor.isPureNonMetal()=False Then Return False
		'If Not ynNeihgbor Is Nothing Then
		If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$")=0 _
			And tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True _
			And tailElement.znNeighbor.znNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		'End If
		If Not tailElement.xpNeighbor Is Nothing Then
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$")=0 _
				And tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True _
				And tailElement.xpNeighbor.xpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.xnNeighbor Is Nothing Then
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$")=0 _
				And tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True _
				And tailElement.xnNeighbor.xnNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ynNeighbor Is Nothing Then
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$")=0 _
				And tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True _
				And tailElement.ynNeighbor.ynNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ypNeighbor Is Nothing Then
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$")=0 _
				And tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True _
				And tailElement.ypNeighbor.ypNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	Case "zp"
		If tailElement.zpNeighbor Is Nothing Then Return False
		If tailElement.zpNeighbor.isPureNonMetal()=False Then Return False
		'If Not ynNeihgbor Is Nothing Then
		If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
			If InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$")=0 _
			And tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True _
			And tailElement.zpNeighbor.zpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
				Return True
			End If
		End If
		'End If
		If Not tailElement.xpNeighbor Is Nothing Then
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$")=0 _
				And tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True _
				And tailElement.xpNeighbor.xpNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.xnNeighbor Is Nothing Then
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$")=0 _
				And tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True _
				And tailElement.xnNeighbor.xnNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ynNeighbor Is Nothing Then
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$")=0 _
				And tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True _
				And tailElement.ynNeighbor.ynNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
		If Not tailElement.ypNeighbor Is Nothing Then
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$")=0 _
				And tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True _
				And tailElement.ypNeighbor.ypNeighbor.getNumberOfNonMetalFaceNeighbors()>=4 Then
					Return True
				End If
			End If
		End If
	End Select
	Return False
	'getNumberofPureNonMetalElements = n
End Function
Public Function needToReEstimateAvailableElementsOnGoingForward() As Boolean
	'Dim NeighborNum As Integer, i As Integer, n As Integer
	'Dim neighbors() As AntennaElement
	If tailElement.isMetal()=False Then
		ReportInformationToWindow "The tail element is Non-Metal, please switch to a Metal one"
		Return False
	End If

	Select Case Right(conLogics,2)
	Case "xn"
		If tailElement.xnNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.ynNeighbor Is Nothing) Or (Not tailElement.ypNeighbor Is Nothing) Then
			If tailElement.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			If tailElement.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			If Not tailElement.ynNeighbor.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.ypNeighbor.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.znNeighbor Is Nothing) Or (Not tailElement.zpNeighbor Is Nothing) Then
			If tailElement.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor.isMetal()=True Then Return False
			If tailElement.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			If Not tailElement.znNeighbor.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.znNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.zpNeighbor.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.xnNeighbor.xnNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.xnNeighbor.xnNeighbor.isMetal()=False And _
			Not tailElement.xnNeighbor.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.xnNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	Case "xp"
		If tailElement.xpNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.ynNeighbor Is Nothing) Or (Not tailElement.ypNeighbor Is Nothing) Then
			If tailElement.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			If tailElement.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			If Not tailElement.ynNeighbor.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.ypNeighbor.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.znNeighbor Is Nothing) Or (Not tailElement.zpNeighbor Is Nothing) Then
			If tailElement.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor.isMetal()=True Then Return False
			If tailElement.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			If Not tailElement.znNeighbor.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.znNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.zpNeighbor.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.xpNeighbor.xpNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.xpNeighbor.xpNeighbor.isMetal()=False And _
			Not tailElement.xpNeighbor.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.xpNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	Case "yn"
		If tailElement.ynNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.xnNeighbor Is Nothing) Or (Not tailElement.xpNeighbor Is Nothing) Then
			If tailElement.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			If tailElement.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			If Not tailElement.xnNeighbor.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.xpNeighbor.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.znNeighbor Is Nothing) Or (Not tailElement.zpNeighbor Is Nothing) Then
			If tailElement.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor.isMetal()=True Then Return False
			If tailElement.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			If Not tailElement.znNeighbor.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.znNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.zpNeighbor.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.ynNeighbor.ynNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.ynNeighbor.ynNeighbor.isMetal()=False And _
			Not tailElement.ynNeighbor.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.ynNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	Case "yp"
		If tailElement.ypNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.xnNeighbor Is Nothing) Or (Not tailElement.xpNeighbor Is Nothing) Then
			If tailElement.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			If tailElement.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			If Not tailElement.xnNeighbor.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.xpNeighbor.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.znNeighbor Is Nothing) Or (Not tailElement.zpNeighbor Is Nothing) Then
			If tailElement.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor Is Nothing Then Return False
			If tailElement.zpNeighbor.zpNeighbor Is Nothing Then Return False
			If tailElement.znNeighbor.znNeighbor.isMetal()=True Then Return False
			If tailElement.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			If Not tailElement.znNeighbor.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.znNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.zpNeighbor.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.zpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.ypNeighbor.ypNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.ypNeighbor.ypNeighbor.isMetal()=False And _
			Not tailElement.ypNeighbor.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.ypNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	Case "zn"
		If tailElement.znNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.xnNeighbor Is Nothing) Or (Not tailElement.xpNeighbor Is Nothing) Then
			If tailElement.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			If tailElement.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			If Not tailElement.xnNeighbor.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.xpNeighbor.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.ynNeighbor Is Nothing) Or (Not tailElement.ypNeighbor Is Nothing) Then
			If tailElement.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			If tailElement.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			If Not tailElement.ynNeighbor.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.ypNeighbor.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.znNeighbor.znNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.znNeighbor.znNeighbor.isMetal()=False And _
			Not tailElement.znNeighbor.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.znNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	Case "zp"
		If tailElement.zpNeighbor Is Nothing Then Return False
		'If tailElement.nnn2nnpNeighbor Is Nothing Then Return False
		'If tailElement.npn2nppNeighbor Is Nothing Then Return False
		If (Not tailElement.xnNeighbor Is Nothing) Or (Not tailElement.xpNeighbor Is Nothing) Then
			If tailElement.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor Is Nothing Then Return False
			If tailElement.xpNeighbor.xpNeighbor Is Nothing Then Return False
			If tailElement.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			If tailElement.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			If Not tailElement.xnNeighbor.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.xnNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.xpNeighbor.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.xpNeighbor.isMetal()=True Then Return False
			End If
		End If
		If (Not tailElement.ynNeighbor Is Nothing) Or (Not tailElement.ypNeighbor Is Nothing) Then
			If tailElement.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor Is Nothing Then Return False
			If tailElement.ypNeighbor.ypNeighbor Is Nothing Then Return False
			If tailElement.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			If tailElement.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			If Not tailElement.ynNeighbor.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.ynNeighbor.isMetal()=True Then Return False
			End If
			If Not tailElement.ypNeighbor.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.ypNeighbor.isMetal()=True Then Return False
			End If
		End If
		If tailElement.zpNeighbor.zpNeighbor Is Nothing Then
			Return True
		Else
			If tailElement.zpNeighbor.zpNeighbor.isMetal()=False And _
			Not tailElement.zpNeighbor.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.zpNeighbor.isMetal()=True Then
					Return True
				End If
			End If
		End If
		Return False
	End Select
	'getNumberofPureNonMetalElements = n
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
Private Function deleteElementAt(ByVal index As Integer, ByRef arr As Variant)
	Dim i As Integer
	For i=index+1 To UBound(arr)
		arr(i-1)=arr(i)
	Next
	ReDim Preserve arr(LBound(arr) To UBound(arr)-1)
End Function
