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
Public availableNum As Integer
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
Public Sub initialize(feed As AntennaElement, eleMaterial As String, _
substrateMaterial As String, availableElementsNum As Integer, availableElementsStr As String)
	length = 0
	elemNumber = 1
	conLogics=""
	availableStr = availableElementsStr
	availableNum = availableElementsNum
	antMaterial = eleMaterial
	subMaterial = substrateMaterial
	Set feedElement=feed
	Set tailElement=feed
	'Set feedElement=feed
	If Not feedElem.IsMetal() Then feedElem.setMaterial(antMaterial)
	removeAvailableElements()
End Sub
'When going forward, remove available elements which are not available any more from availableStr
'The progress starts from the current tailElement
Public Sub removeAvailableElements()
	Dim nonMetalNeighbors() As AntennaElement
	Dim nonMetalNeighborsNum As Integer, i As Integer
	Dim nonMetalNeighborsStr() As String, neighborStr As String
	neighborStr = tailElement.solidName & "$"
	If InStr(availableStr, neighborStr)<>0 Then
		availableStr = Left(availableStr,InStr(availableStr,neighborStr)-1) _
		& Right(availableStr, Len(availableStr)-InStr(availableStr,neighborStr)-Len(neighborStr)+1)
		availableNum -= 1
	End If
	nonMetalNeighbors = tailElement.getNonMetalFaceNeighbors(nonMetalNeighborsNum, nonMetalNeighborsStr)
	For i=0 To nonMetalNeighborsNum-1
		neighborStr = nonMetalNeighbors(i).solidName & "$"
		If InStr(availableStr, neighborStr)<>0 Then
			availableStr = Left(availableStr,InStr(availableStr,neighborStr)-1) _
			& Right(availableStr, Len(availableStr)-InStr(availableStr,neighborStr)-Len(neighborStr)+1)
			availableNum -= 1
		End If
	Next
	nonMetalNeighbors = tailElement.getNonMetalEdgeNeighbors(nonMetalNeighborsNum)
	For i=0 To nonMetalNeighborsNum-1
		neighborStr = nonMetalNeighbors(i).solidName & "$"
		If InStr(availableStr, neighborStr)<>0 Then
			availableStr = Left(availableStr,InStr(availableStr,neighborStr)-1) _
			& Right(availableStr, Len(availableStr)-InStr(availableStr,neighborStr)-Len(neighborStr)+1)
			availableNum -= 1
		End If
	Next
	nonMetalNeighbors = tailElement.getNonMetalVerticeNeighbors(nonMetalNeighborsNum)
	For i=0 To nonMetalNeighborsNum-1
		neighborStr = nonMetalNeighbors(i).solidName & "$"
		If InStr(availableStr, neighborStr)<>0 Then
			availableStr = Left(availableStr,InStr(availableStr,neighborStr)-1) _
			& Right(availableStr, Len(availableStr)-InStr(availableStr,neighborStr)-Len(neighborStr)+1)
			availableNum -= 1
		End If
	Next
End Sub
'When going backward, append available elements which are not in availableStr
'The progress starts from the last tailElement, the connection string between current
'tailElement and the last tailElement is a necessary parameter
Public Sub appendAvailableElements(conStr As String)
	Dim nonMetalNeighbors() As AntennaElement
	Dim currentElement As AntennaElement
	Dim nonMetalNeighborsNum As Integer, i As Integer
	Dim nonMetalNeighborsStr() As String, neighborStr As String
	Set currentElement = tailElement.getFaceNeighborFromString(conStr, False)
	availableNum += currentElement.getNumberofPureNonMetalNeighbors(availableStr)
End Sub
Public Function constructor(tgtLength As Double, solidNumber As Integer, solidxSize As Double, _
solidySize As Double, solidzSize As Double) As Boolean
	log2Window=False
	Dim i As Integer, j As Integer, k As Integer, m As Integer, ii As Integer
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
	ReportInformationToWindow "Maximum dimension is: " & cStr(Round(maxSize,2))+"mm"
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
			Set currentElement = currentElement.getFaceNeighborFromString(Right(conLogics,2), True)
			If InStr(Right(conLogics,2), "x")<>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
			ElseIf InStr(Right(conLogics,2), "y")<>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
			ElseIf InStr(Right(conLogics,2), "z")<>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
			End If
			Set tailElement = currentElement
			appendAvailableElements(Right(conLogics,2))
			conLogics = Left(conLogics, Len(conLogics)-2)
			Plot.Update
		Loop Until tgtLength>=length
		'Set tailElement = currentElement
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
	'Dim ifOmitFirtEstimationInCorrector As Boolean
	'Dim totAvailableElementsNumber As Integer
	Dim neighborAvailableElementsStr() As String
	Dim neighborAvailableElementsNum() As Integer
	Dim neighborAvailableElementsLogic() As String
	Dim invStr As String
	'Dim ifAvailable As Boolean

	Do
		randomCycle = Int(tgtLength^2/(solidxSize+solidySize+solidzSize)*solidNumber/6e5)+1
		randomCycle = randomCycle-Int(ant.length/tgtLength*randomCycle) Mod randomCycle
		ReportInformationToWindow "Random Cycle is "+cStr(randomCycle)
		currentElement.getValidNonMetalFaceNeighbors(nonMetalFaceNeighbors, n_nonMetalFaceNeighbors, _
		nonMetalFaceNeighborsStr, validFaceNeighborsArray, n_validFaceNeighbors)
		'Estimate if the candidate neighbor can split all available elements into two parts, if yes
		'estimate all available elements for all neighbors and then pick one meeting the target available elements
		ReDim tempObj(n_validFaceNeighbors)
		ReDim tempStr(n_validFaceNeighbors)
		'ReDim tempIndex(n_validFaceNeighbors)
		ReDim neighborAvailableElementsStr(2)
		ReDim neighborAvailableElementsNum(2)
		ReDim neighborAvailableElementsLogic(2)
		'reload valid face neighbors to a new array
		For i=0 To n_validFaceNeighbors-1
			Set tempObj(i) = nonMetalFaceNeighbors(validFaceNeighborsArray(i))
			tempStr(i) = nonMetalFaceNeighborsStr(validFaceNeighborsArray(i))
		Next

		'Random pick one of appropriate non metal face neighbors to be used as an antenna element
		If n_validFaceNeighbors>0 And ant.length<tgtLength Then
			If ant.length > tgtLength*0.2 Then
				needReEstimatingAvailableElements=False
				If n_validFaceNeighbors>1 Then
					needReEstimatingAvailableElements = _
					needToReEstimateAvailableElementsOnGoingForward(tempStr)
				End If

				'Determin whether the candidate neighbor has a valid face neighbor, if not, try
				'to pick another one to be antenna elment
				If needReEstimatingAvailableElements=True Then
					For i=0 To 1
						neighborAvailableElementsStr(i)=""
						neighborAvailableElementsNum(i)=0
						neighborAvailableElementsLogic(i)=""
					Next
					ReportInformationToWindow "4$Need to re-estimate available elements number on going forward."
					ReportInformationToWindow "4$Target available antenna Elements is " & CInt((tgtLength-length)*maxSize*2.5)
					availableNum=0
					availableStr = ""
					'get the number ID of the face neighbor that uses the same connection string by the tailElement
					For i=0 To n_validFaceNeighbors-1
						If StrComp(tempStr(i), Right(conLogics,2))=0 Then
							m=i
							Exit For
						End If

					Next

					For i=0 To n_validFaceNeighbors-1
						If StrComp(tempStr(i), tempStr(m))<>0 Then
							neighborAvailableElementsLogic(0)= tempStr(i)
							neighborAvailableElementsNum(0)= _
							tempObj(m).getFaceNeighborFromString(tempStr(i), False).getFaceNeighborFromString(tempStr(i), False).getNumberofPureNonMetalElements(neighborAvailableElementsStr(0))
							ReportInformationToWindow "4$Available elements for neighbor-"&neighborAvailableElementsLogic(0)&": "&CStr(neighborAvailableElementsNum(0))
							If InStr(tempStr(i),"n")<>0 Then
								invStr = Left(tempStr(i),1) & "p"
							Else
								invStr = Left(tempStr(i),1) & "n"
							End If
							neighborAvailableElementsLogic(1)= invStr
							neighborAvailableElementsNum(1)= _
							tempObj(m).getFaceNeighborFromString(tempStr(i), True).getFaceNeighborFromString(tempStr(i), True).getNumberofPureNonMetalElements(neighborAvailableElementsStr(1))
							ReportInformationToWindow "4$Available elements for neighbor-"&neighborAvailableElementsLogic(1)&": "&CStr(neighborAvailableElementsNum(1))
							Exit For
						End If
					Next
					'***********************************************************************************************************************
					'kick out the neighbor that does not meet the target number of available elements
					If neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.5) Or _
					neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.5) Then
						j=0
						For i=0 To 1
							If neighborAvailableElementsNum(i) < CInt((tgtLength-length)*maxSize*2.5) Then
								For ii=0 To n_validFaceNeighbors-1
									If StrComp(neighborAvailableElementsLogic(i), tempStr(ii))=0 Then
										Set tempObj(ii) = Nothing
										tempStr(ii)=""
										j+=1
									End If
								Next
							End If
						Next
						n_validFaceNeighbors-=j
					Else
						availableStr = neighborAvailableElementsStr(0) & neighborAvailableElementsStr(1)
						availableNum = neighborAvailableElementsNum(0) + neighborAvailableElementsNum(1)
						GoTo BackWard
					End If
				End If
			Else
				needReEstimatingAvailableElements=False
			End If
			'******************************************************************************************************************************
			If n_validFaceNeighbors>0 Then
				'ifOmitFirtEstimationInCorrector=False
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
				remainder = iteration Mod (Int(randomCycle*Rnd)+1)
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
				'**************************************************************************************
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
				Set currentElement = tailElement
				If log2Window=True Then ReportInformationToWindow "0$Antenna Length: "+cStr(Round(length,2))+"mm"
				'***************************************************************************************************
				If needReEstimatingAvailableElements=True Then
					If StrComp(validFaceNeighborsStr(k), tempStr(m))<>0 Then
						If StrComp(validFaceNeighborsStr(k), neighborAvailableElementsLogic(0))=0 Then
							availableStr = neighborAvailableElementsStr(0)
							availableNum = neighborAvailableElementsNum(0)
						ElseIf StrComp(validFaceNeighborsStr(k), neighborAvailableElementsLogic(1))=0 Then
							availableStr = neighborAvailableElementsStr(1)
							availableNum = neighborAvailableElementsNum(1)
						End If
						removeAvailableElements()
					Else
						If neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.5) And _
						neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.5) Then
							k = Int(2*Rnd)
						ElseIf neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.5) And _
						neighborAvailableElementsNum(1)<CInt((tgtLength-length)*maxSize*2.5) Then
							k = 0
						ElseIf neighborAvailableElementsNum(0)<CInt((tgtLength-length)*maxSize*2.5) And _
						neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.5) Then
							k = 1
						End If
						tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(k), False).setMaterial(antMaterial)
						Plot.Update
						elementNumber+=1
						conLogics = conLogics & neighborAvailableElementsLogic(k)
						Set tailElement = tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(k), False)
						If InStr(neighborAvailableElementsLogic(k),"x")<>0 Then
							length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
						ElseIf InStr(neighborAvailableElementsLogic(k),"y")<>0 Then
							length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
						ElseIf InStr(neighborAvailableElementsLogic(k),"z")<>0 Then
							length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
						End If
						'ant.length+=Abs(currentElement.-currentElement.)
						Set currentElement = tailElement
						availableStr = neighborAvailableElementsStr(k)
						availableNum = neighborAvailableElementsNum(k)
						removeAvailableElements()
					End If
				Else
					removeAvailableElements()
				End If
			Else
				availableStr = neighborAvailableElementsStr(0) & neighborAvailableElementsStr(1)
				availableNum = neighborAvailableElementsNum(0) + neighborAvailableElementsNum(1)
				'ifOmitFirtEstimationInCorrector=True
				GoTo Backward
			End If

		ElseIf ant.length>=tgtLength Then

			'Open logFile For Append As #2
			Print #2, conLogics
			Close #2
			ReportInformationToWindow conLogics
			ReportInformationToWindow "1$Antenna Length: "+cStr(Round(length,2))+"mm"
			Return True
			'Exit Function
		ElseIf n_validFaceNeighbors<=0 And length<tgtLength-1 Then
			'when going into a dead end
			availableStr=""
			availableNum=0
			Backward:
				'Open logFile For Append As #2
				Print #2, conLogics'+tailElement.solidName
				'Close #2
				'Call corrector to correct the routing of tail part
				ReportInformationToWindow "Iteration #"+cStr(iteration)+"-----------------"
				ReportInformationToWindow conLogics
				ReportInformationToWindow "2$Antenna Length: "+cStr(Round(length,2))+"mm"
				'tgtIncrement = (tgtLength-length)*0.2+2
				'tgtIncrement = 15
				'If (containerIndex > 0) And (containerIndex Mod 3 =0) Then
				If ant.length>0.2*tgtLength Then
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
				removeAvailableElements()
			'Return False
		Else
			'Open logFile For Append As #2
			Print #2, conLogics
			Close #2
			ReportInformationToWindow conLogics
			ReportInformationToWindow "3$Antenna Length: "+cStr(Round(length,2))+"mm"
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

	'Dim pureNonMetalElements As Integer
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
	logicLength = Len(conLogics)
	initialLength = length

	While Not currentElement Is Nothing ' And iteration*2<logicLength/4
		Do
			If Len(conLogics)<2 Then
				Return Nothing
				'Exit Function
			End If
			currentElement.setMaterial(subMaterial)
			currentNeighborStr = Right(conLogics,2)
			'Set currentElement = currentElement.getFaceNeighborFromString(currentNeighborStr, True)
			elementNumber -= 1
			availableNum+=currentElement.getNumberofPureNonMetalNeighbors(availableStr)
			Plot.Update
			If fBR=0 Then
				'needToReestimate = currentElement.needToEstimateAvailableElementsAllOverAgain(currentNeighborStr, availableStr)
				'availableNum+=currentElement.getNumberofPureNonMetalNeighbors(availableStr)
				needToReestimate = needToReEstimateAvailableElementsOnGoingBackward()
				'Estimate if need to re-estimate the pure-non metal elements all over again
				If needToReestimate=True Then
					ReportInformationToWindow "Need to re-estimate available elements number on going backward."
					availableStr = ""
					availableNum=currentElement.getNumberofPureNonMetalElements(availableStr)
				End If
			End If
			'previousConnectionStatus = currentConnectionStatus
			'cFlag=currentElement.flag
			conLogics = Left(conLogics, Len(conLogics)-2)

			Set currentElement = currentElement.getFaceNeighborFromString(currentNeighborStr, True)
			If InStr(currentNeighborStr, "x")<>0 Then
				length-=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
			ElseIf InStr(currentNeighborStr, "y")<>0 Then
				length-=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
			ElseIf InStr(currentNeighborStr, "z")<>0 Then
				length-=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
			End If
			Set tailElement = currentElement
			'appendAvailableElements(currentNeighborStr)
			ReportInformationToWindow "4$Antenna Length: "+CStr(Round(length,2))+"mm"
			'fall back until the difference between current lenghth and the inital length over the initial
			'lenght is greater than what is requred, or fBR
			'currentElement.setMaterial(subMaterial)
			'pureNonMetalElements+=currentElement.getNumberofPureNonMetalNeighbors()
			ReportInformationToWindow "4$Target available antenna Elements is " & CInt((tgtLength-length)*maxSize*2.5)
			ReportInformationToWindow "4$Available antenna Elements is " & CStr(availableNum)
			iteration += 1
			If fBR > 0 Then
				tgtCondition=((initialLength-length)/initialLength > fBR)
			Else
				tgtCondition=(availableNum >= CInt((tgtLength-length)*maxSize*2.5))
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
					currentElement.getFaceNeighborFromString(nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)), False).setMaterial(antMaterial)
					conLogics = conLogics & nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii))
					elementNumber += 1
					If InStr(nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)), "x")<>0 Then
						length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
					ElseIf InStr(nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)), "y")<>0 Then
						length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
					ElseIf InStr(nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)), "z")<>0 Then
						length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
					End If
					If log2Window=True Then ReportInformationToWindow "5$Antenna Length: "+CStr(Round(length,2))
					Return currentElement.getFaceNeighborFromString(nonMetalFaceNeighborsStr(validFaceNeighborsArray(ii)), False)
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
	'Return value:
	'0-do not need to re-estimate available elements number
	'1-need to re-estimate available elements number only once for all
	'2-need to re-estimate available elements number for each non-metal face neighbor

	If tailElement.isMetal()=True Then
		ReportInformationToWindow "The tail element is metal, please switch to a nonMetal one"
		Return False
	End If

	Dim i As Integer, neighborsStr As String
	Dim nMFneighorsStr() As String
	Dim nMFneighbors() As AntennaElement
	Dim nMFneighborsNum As Integer

	neighborsStr = ""
	nMFneighbors = tailElement.getNonMetalFaceNeighbors(nMFneighborsNum, nMFneighorsStr)
	For i=0 To nMFneighborsNum-1
		neighborsStr = neighborsStr & nMFneighorsStr(i)
	Next

	Dim pElement As AntennaElement
	Dim nElement As AntennaElement

	Set pElement = tailElement.getFaceNeighborFromString(Right(conLogics,2),True)
	Set nElement = tailElement.getFaceNeighborFromString(Right(conLogics,2),False)

	If nElement Is Nothing Then Return False
	If nElement.isPureNonMetal()=False Then Return False
	Select Case Right(conLogics,2)
	Case "xn"
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.nnLNeighbor Is Nothing Then
				If	Not tailElement.nnLNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.nnLNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nnLNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nnLNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nnLNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nnLNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.npLNeighbor Is Nothing Then
				If	Not tailElement.npLNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.npLNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.npLNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.npLNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.npLNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.npLNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.nLnNeighbor Is Nothing Then
				If	Not tailElement.nLnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLnNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.nLpNeighbor Is Nothing Then
				If	Not tailElement.nLpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLpNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case "xp"
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.pnLNeighbor Is Nothing Then
				If	Not tailElement.pnLNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.pnLNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pnLNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.pnLNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.pnLNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pnLNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.ppLNeighbor Is Nothing Then
				If	Not tailElement.ppLNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.ppLNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.ppLNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.ppLNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.ppLNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.ppLNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.nLnNeighbor Is Nothing Then
				If	Not tailElement.nLnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLnNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.nLpNeighbor Is Nothing Then
				If	Not tailElement.nLpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLpNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case "yn"
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.nnLNeighbor Is Nothing Then
				If	Not tailElement.nnLNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.nnLNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nnLNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nnLNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nnLNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nnLNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.pnLNeighbor Is Nothing Then
				If	Not tailElement.pnLNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.pnLNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pnLNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.pnLNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.pnLNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pnLNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.LnnNeighbor Is Nothing Then
				If	Not tailElement.LnnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.LnnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LnnNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.LnnNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnnNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.LnpNeighbor Is Nothing Then
				If	Not tailElement.LnpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.LnpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LnpNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.LnpNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnpNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case "yp"
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.npLNeighbor Is Nothing Then
				If	Not tailElement.npLNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.npLNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.npLNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.npLNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.npLNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.npLNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.ppLNeighbor Is Nothing Then
				If	Not tailElement.ppLNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.ppLNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.ppLNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.ppLNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.ppLNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.ppLNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.LpnNeighbor Is Nothing Then
				If	Not tailElement.LpnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.LpnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LpnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LpnNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.LpnNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LpnNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.LppNeighbor Is Nothing Then
				If	Not tailElement.LppNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.LppNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LppNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LppNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.LppNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LppNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case "zn"
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.nLnNeighbor Is Nothing Then
				If	Not tailElement.nLnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLnNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLnNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLnNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.pLnNeighbor Is Nothing Then
				If	Not tailElement.pLnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.pLnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pLnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.pLnNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.pLnNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pLnNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.LnnNeighbor Is Nothing Then
				If	Not tailElement.LnnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.LnnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LnnNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.LnnNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnnNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.LpnNeighbor Is Nothing Then
				If	Not tailElement.LpnNeighbor.znNeighbor Is Nothing Then
					If	tailElement.LpnNeighbor.znNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LpnNeighbor.znNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LpnNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.LpnNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LpnNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case "zp"
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.nLpNeighbor Is Nothing Then
				If	Not tailElement.nLpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.nLpNeighbor.xnNeighbor Is Nothing Then
					If	tailElement.nLpNeighbor.xnNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.nLpNeighbor.xnNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.pLpNeighbor Is Nothing Then
				If	Not tailElement.pLpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.pLpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pLpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.pLpNeighbor.xpNeighbor Is Nothing Then
					If	tailElement.pLpNeighbor.xpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.pLpNeighbor.xpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.LnpNeighbor Is Nothing Then
				If	Not tailElement.LnpNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.LnpNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnpNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LnpNeighbor.ynNeighbor Is Nothing Then
					If	tailElement.LnpNeighbor.ynNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LnpNeighbor.ynNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
			If Not tailElement.LppNeighbor Is Nothing Then
				If	Not tailElement.LppNeighbor.zpNeighbor Is Nothing Then
					If	tailElement.LppNeighbor.zpNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LppNeighbor.zpNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
				If	Not tailElement.LppNeighbor.ypNeighbor Is Nothing Then
					If	tailElement.LppNeighbor.ypNeighbor.isPureNonMetal()=True And _
					InStr(availableStr, tailElement.LppNeighbor.ypNeighbor.SolidName & "$" )=0 Then
						Return True
					End If
				End If
			End If
		End If
	Case Else
		Return False
	End Select
	Return False
	'getNumberofPureNonMetalElements = n
End Function
Public Function needToReEstimateAvailableElementsOnGoingForward(validNeighborsStr() As String) As Boolean
	Dim i As Integer
	Dim neighborsStr As String
	neighborsStr=""
	Dim nElement As AntennaElement

	For i=0 To UBound(validNeighborsStr)-1
		neighborsStr = neighborsStr & validNeighborsStr(i)
	Next

	If tailElement.isMetal()=False Then
		ReportInformationToWindow "The tail element is Non-Metal, please switch to a Metal one"
		Return False
	End If

	'Set pElement = tailElement.getFaceNeighborFromString(Right(conLogics,2),True)
	Set nElement = tailElement.getFaceNeighborFromString(Right(conLogics,2),False)
	If nElement Is Nothing Then Return False
	'If nElement.isPureNonMetal()=False Then Return False
	Select Case Right(conLogics,2)
	Case "xn"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.ynNeighbor Is Nothing Then
				If Not tailElement.ynNeighbor.nnLNeighbor Is Nothing Then
					If tailElement.ynNeighbor.nnLNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.ynNeighbor.nnLNeighbor.solidName & "$")<>0 Then
						If Not tailElement.ypNeighbor Is Nothing Then
							If Not tailElement.ypNeighbor.npLNeighbor Is Nothing Then
								If tailElement.ypNeighbor.npLNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.ypNeighbor.npLNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.xnNeighbor.xnNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.nnLNeighbor Is Nothing Then
										If tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.npLNeighbor Is Nothing Then
										If tailElement.npLNeighbor.npLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.znNeighbor Is Nothing Then
				If Not tailElement.znNeighbor.nLnNeighbor Is Nothing Then
					If tailElement.znNeighbor.nLnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.znNeighbor.nLnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.zpNeighbor Is Nothing Then
							If Not tailElement.zpNeighbor.nLpNeighbor Is Nothing Then
								If tailElement.zpNeighbor.nLpNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.zpNeighbor.nLpNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.xnNeighbor.xnNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.nLnNeighbor Is Nothing Then
										If tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.nLpNeighbor Is Nothing Then
										If tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	Case "xp"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.ynNeighbor Is Nothing Then
				If Not tailElement.ynNeighbor.pnLNeighbor Is Nothing Then
					If tailElement.ynNeighbor.pnLNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.ynNeighbor.pnLNeighbor.solidName & "$")<>0 Then
						If Not tailElement.ypNeighbor Is Nothing Then
							If Not tailElement.ypNeighbor.ppLNeighbor Is Nothing Then
								If tailElement.ypNeighbor.ppLNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.ypNeighbor.ppLNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.xpNeighbor.xpNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.pnLNeighbor Is Nothing Then
										If tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.ppLNeighbor Is Nothing Then
										If tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.znNeighbor Is Nothing Then
				If Not tailElement.znNeighbor.pLnNeighbor Is Nothing Then
					If tailElement.znNeighbor.pLnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.znNeighbor.pLnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.zpNeighbor Is Nothing Then
							If Not tailElement.zpNeighbor.pLpNeighbor Is Nothing Then
								If tailElement.zpNeighbor.pLpNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.zpNeighbor.pLpNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.xpNeighbor.xpNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.pLnNeighbor Is Nothing Then
										If tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.pLpNeighbor Is Nothing Then
										If tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	Case "yn"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.xnNeighbor Is Nothing Then
				If Not tailElement.xnNeighbor.nnLNeighbor Is Nothing Then
					If tailElement.xnNeighbor.nnLNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.xnNeighbor.nnLNeighbor.solidName & "$")<>0 Then
						If Not tailElement.xpNeighbor Is Nothing Then
							If Not tailElement.xpNeighbor.pnLNeighbor Is Nothing Then
								If tailElement.xpNeighbor.pnLNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.xpNeighbor.pnLNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.ynNeighbor.ynNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.nnLNeighbor Is Nothing Then
										If tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.pnLNeighbor Is Nothing Then
										If tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.znNeighbor Is Nothing Then
				If Not tailElement.znNeighbor.LnnNeighbor Is Nothing Then
					If tailElement.znNeighbor.LnnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.znNeighbor.LnnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.zpNeighbor Is Nothing Then
							If Not tailElement.zpNeighbor.LnpNeighbor Is Nothing Then
								If tailElement.zpNeighbor.LnpNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.zpNeighbor.LnpNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.ynNeighbor.ynNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.LnnNeighbor Is Nothing Then
										If tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.LnpNeighbor Is Nothing Then
										If tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	Case "yp"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.xnNeighbor Is Nothing Then
				If Not tailElement.xnNeighbor.npLNeighbor Is Nothing Then
					If tailElement.xnNeighbor.npLNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.xnNeighbor.npLNeighbor.solidName & "$")<>0 Then
						If Not tailElement.xpNeighbor Is Nothing Then
							If Not tailElement.xpNeighbor.ppLNeighbor Is Nothing Then
								If tailElement.xpNeighbor.ppLNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.xpNeighbor.ppLNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.ypNeighbor.ypNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.npLNeighbor Is Nothing Then
										If tailElement.npLNeighbor.npLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.ppLNeighbor Is Nothing Then
										If tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.znNeighbor Is Nothing Then
				If Not tailElement.znNeighbor.LpnNeighbor Is Nothing Then
					If tailElement.znNeighbor.LpnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.znNeighbor.LpnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.zpNeighbor Is Nothing Then
							If Not tailElement.zpNeighbor.LppNeighbor Is Nothing Then
								If tailElement.zpNeighbor.LppNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.zpNeighbor.LppNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.ynNeighbor.ynNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.LpnNeighbor Is Nothing Then
										If tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.LppNeighbor Is Nothing Then
										If tailElement.LppNeighbor.LppNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	Case "zn"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.xnNeighbor Is Nothing Then
				If Not tailElement.xnNeighbor.nLnNeighbor Is Nothing Then
					If tailElement.xnNeighbor.nLnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.xnNeighbor.nLnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.xpNeighbor Is Nothing Then
							If Not tailElement.xpNeighbor.pLnNeighbor Is Nothing Then
								If tailElement.xpNeighbor.pLnNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.xpNeighbor.pLnNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.znNeighbor.znNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.znNeighbor.znNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.nLnNeighbor Is Nothing Then
										If tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.pLnNeighbor Is Nothing Then
										If tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.ynNeighbor Is Nothing Then
				If Not tailElement.ynNeighbor.LnnNeighbor Is Nothing Then
					If tailElement.ynNeighbor.LnnNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.ynNeighbor.LnnNeighbor.solidName & "$")<>0 Then
						If Not tailElement.ypNeighbor Is Nothing Then
							If Not tailElement.ypNeighbor.LpnNeighbor Is Nothing Then
								If tailElement.ypNeighbor.LpnNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.ypNeighbor.LpnNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.znNeighbor.znNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.znNeighbor.znNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.LnnNeighbor Is Nothing Then
										If tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.LpnNeighbor Is Nothing Then
										If tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	Case "zp"
		'edge neighbor of face neighbor should be pure nonmetal on both sides
		If InStr(neighborsStr,"x")<>0 Then
			If Not tailElement.xnNeighbor Is Nothing Then
				If Not tailElement.xnNeighbor.nLpNeighbor Is Nothing Then
					If tailElement.xnNeighbor.nLpNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.xnNeighbor.nLpNeighbor.solidName & "$")<>0 Then
						If Not tailElement.xpNeighbor Is Nothing Then
							If Not tailElement.xpNeighbor.pLpNeighbor Is Nothing Then
								If tailElement.xpNeighbor.pLpNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.xpNeighbor.pLpNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.zpNeighbor.zpNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.nLpNeighbor Is Nothing Then
										If tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.pLpNeighbor Is Nothing Then
										If tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.ynNeighbor Is Nothing Then
				If Not tailElement.ynNeighbor.LnpNeighbor Is Nothing Then
					If tailElement.ynNeighbor.LnpNeighbor.isPureNonMetal()=True _
					And InStr(availableStr,tailElement.ynNeighbor.LnpNeighbor.solidName & "$")<>0 Then
						If Not tailElement.ypNeighbor Is Nothing Then
							If Not tailElement.ypNeighbor.LppNeighbor Is Nothing Then
								If tailElement.ypNeighbor.LppNeighbor.isPureNonMetal()=True _
								And InStr(availableStr,tailElement.ypNeighbor.LppNeighbor.solidName & "$")<>0 Then
									'estimate if the path is blocked, if yes, return true
									If tailElement.zpNeighbor.zpNeighbor Is Nothing Then
										Return True
									ElseIf tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=False Then
										Return True
									End If
									If Not tailElement.LnpNeighbor Is Nothing Then
										If tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
									If Not tailElement.LppNeighbor Is Nothing Then
										If tailElement.LppNeighbor.LppNeighbor Is Nothing Then
											Return True
										ElseIf tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=False Then
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End Select
	Return False
	'getNumberofPureNonMetalElements = n
End Function
Public Function destructor()
	Dim A As AntennaElement, B As AntennaElement
	Dim i As Integer
	Set A=feedElement
	For i=0 To elementNumber-2
		Set B = A.getFaceNeighborFromString(Mid(conLogics,2*i+1,2),False)
		B.setMaterial(subMaterial)
		Set A = B
		Plot.Update
	Next
	Dim historyCaption As String
	Dim sCommand As String
	historyCaption = "$SA$"
	sCommand = ""
	AddToHistory(historyCaption, sCommand)
	'initialize(feedElement, antMaterial, subMaterial)
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
Public Function toHistoryList()
	Dim A As AntennaElement, i As Integer
	Dim historyCaption As String
	Dim sCommand As String
	sCommand = ""
	Set A=feedElement
	sCommand  = sCommand + A.setMaterialPermanently(antMaterial)
	For i=0 To elementNumber-2
		sCommand = sCommand +  _
		A.getFaceNeighborFromString(Mid(conLogics,2*i+1,2), False).setMaterialPermanently(antMaterial)
		Set A = A.getFaceNeighborFromString(Mid(conLogics,2*i+1,2), False)
	Next
	historyCaption  = "$SA$"
	AddToHistory(historyCaption, sCommand)
	Plot.Update
	'initialize(feedElement, antMaterial, subMaterial)
End Function
