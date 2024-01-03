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
substrateMaterial As String, ByVal availableElementsNum As Integer, ByVal availableElementsStr As String)
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
	Dim validFaceNeighbors() As AntennaElement, validFaceNeighborsStr() As String
	Dim n_validFaceNeighbors As Integer, validNeighborsString As String
	neighborStr = tailElement.solidName & "$"
	If InStr(availableStr, neighborStr)<>0 Then
		availableStr = Left(availableStr,InStr(availableStr,neighborStr)-1) _
		& Right(availableStr, Len(availableStr)-InStr(availableStr,neighborStr)-Len(neighborStr)+1)
		availableNum -= 1
	End If
	nonMetalNeighbors = tailElement.getNonMetalFaceNeighbors(nonMetalNeighborsNum, nonMetalNeighborsStr)
	'tailElement.getValidFaceNeighbors(validFaceNeighbors, validFaceNeighborsStr, n_validFaceNeighbors)
	'For i=0 To n_validFaceNeighbors-1
	'	validNeighborsString = validNeighborsString & validFaceNeighborsStr(i)
	'Next
	For i=0 To nonMetalNeighborsNum-1
		neighborStr = nonMetalNeighbors(i).solidName & "$"
		If InStr(availableStr, neighborStr)<>0 Then 'And InStr(validNeighborsString, nonMetalNeighborsStr(i))=0 Then
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
	Dim currentElement As AntennaElement
	Dim validFaceNeighbors() As AntennaElement, validFaceNeighborsStr() As String
	Dim n_validFaceNeighbors As Integer, validNeighborsString As String
	Set currentElement = tailElement.getFaceNeighborFromString(conStr, False)
	availableNum += currentElement.getNumberofPureNonMetalNeighbors(availableStr)
	'tailElement.getValidFaceNeighbors(validFaceNeighbors,  validFaceNeighborsStr, n_validFaceNeighbors)
	'For i=0 To n_validFaceNeighbors-1
	'	If InStr(availableStr, validFaceNeighbors(i).solidName & "$") Then
	'		availableNum += 1
	'		availableStr = availableStr & validFaceNeighbors(i).solidName & "$"
	'	End If
	'Next
End Sub
Public Function constructor(tgtLength As Double, solidNumber As Integer, solidxSize As Double, _
solidySize As Double, solidzSize As Double) As Boolean
	log2Window=False
	Dim i As Integer, j As Integer, k As Integer, m As Integer, ii As Integer
	'Dim n_faceNeighbors As Integer
	Dim n_nonMetalFaceNeighbors As Integer
	Dim nonMetalFaceNeighbors() As AntennaElement, vFNeighborsOfVFNeighbors() As AntennaElement
	Dim nonMetalEdgeNeighbors() As AntennaElement
	Dim currentElement As AntennaElement
	Dim nonMetalFaceNeighborsStr() As String, vFNStrOfvFN() As String
	Dim n_validFaceNeighbors As Integer, n_vFNoVFN As Integer
	Dim randomNumber As Integer
	'list of validities of face neighbors
	'Dim validityOfFaceNeighbors() As Boolean, vOFN() As Boolean, NON As Boolean
	Dim validFaceNeighborsArray() As Integer,  NON As Boolean
	Dim iteration As Integer
	Dim fallBackTimes As Integer
	Dim remainder As Integer
	Dim randomCycle As Integer
	Dim lastRandomNumber As Integer
	'Dim containerIndex As Integer
	Dim lastLength As Double
	Dim tgtIncrement As Double
	Dim maxSize As Double
	Dim cFileNumber As Integer
	cFileNumber = FreeFile
	If solidxSize>=solidySize And solidxSize>=solidzSize Then
		maxSize = solidxSize
	ElseIf solidySize>=solidxSize And solidySize>=solidzSize Then
		maxSize = solidySize
	ElseIf solidzSize>=solidySize And solidzSize>=solidxSize Then
		maxSize = solidzSize
	End If
	ReportInformationToWindow "$Constructor: Maximum dimension is: " & cStr(Round(maxSize,2))+"mm"
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
   	Open logFile For Output As #cFileNumber
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
		Print #cFileNumber, conLogics
		Close #cFileNumber
		Return True
	End If
	'#### When target length is greater than ant.length ####
	iteration = 0
	fallBackTimes = 0
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
	Dim flag As Boolean
	'Dim ifAvailable As Boolean

	Do
		randomCycle = Int(tgtLength^2/(solidxSize+solidySize+solidzSize)*solidNumber/6e5)+1
		randomCycle = randomCycle-Int(ant.length/tgtLength*randomCycle) Mod randomCycle
		ReportInformationToWindow "$Constructor: Random Cycle is "&CStr(randomCycle)
		ReportInformationToWindow "$Constructor: Antenna length is "&CStr(Round(length,2))&"mm"
		currentElement.getValidFaceNeighbors(tempObj, tempStr, n_validFaceNeighbors)
		'Estimate if the candidate neighbor can split all available elements into two parts, if yes
		'estimate all available elements for all neighbors and then pick one meeting the target available elements
		'ReDim tempObj(n_validFaceNeighbors)
		'ReDim tempStr(n_validFaceNeighbors)
		'ReDim tempIndex(n_validFaceNeighbors)
		'The third element stores the information of the m-th elements
		ReDim neighborAvailableElementsStr(3)
		ReDim neighborAvailableElementsNum(3)
		ReDim neighborAvailableElementsLogic(3)
		'reload valid face neighbors to a new array
		'If an available face neighbor is Not among available String, set it to Invalid
		'For i=0 To n_validFaceNeighbors-1
			'If InStr(availableStr, )
			'i+=1
		'Next
		'Random pick one of appropriate non metal face neighbors to be used as an antenna element
		If n_validFaceNeighbors>0 And ant.length<tgtLength Then
			needReEstimatingAvailableElements=False
			If ant.length > tgtLength*0.2 Then
				If n_validFaceNeighbors>1 Then
					needReEstimatingAvailableElements = _
					needToReEstimateAvailableElementsOnGoingForward(tempStr)
				End If

				'Determin whether the candidate neighbor has a valid face neighbor, if not, try
				'to pick another one to be antenna elment
				If needReEstimatingAvailableElements=True Then
					For i=0 To n_validFaceNeighbors-1
						neighborAvailableElementsStr(i)=""
						neighborAvailableElementsNum(i)=-1
						neighborAvailableElementsLogic(i)=""
					Next
					ReportInformationToWindow "$Constructor: Need to re-estimate available elements number on going forward."
					ReportInformationToWindow "$Constructor: Target available antenna Elements is " & CInt((tgtLength-length)*maxSize*2.6)
					availableNum=0
					availableStr = ""
					'get the number ID of the face neighbor that uses the same connection string by the tailElement
					For i=0 To n_validFaceNeighbors-1
						If StrComp(tempStr(i), Right(conLogics,2))=0 Then
							m=i
							Exit For
						End If

					Next
					neighborAvailableElementsNum(2)=0
					neighborAvailableElementsLogic(2)= tempStr(m)
					If Not tempObj(m).getFaceNeighborFromString(tempStr(m), False) Is Nothing Then
						If tempObj(m).getFaceNeighborFromString(tempStr(m), False).isPureNonMetal()=True Then
							neighborAvailableElementsNum(2)= _
							tempObj(m).getFaceNeighborFromString(tempStr(m), False).getNumberofPureNonMetalElements(neighborAvailableElementsStr(2))
							ReportInformationToWindow "$Constructor: Available elements for neighbor-"&neighborAvailableElementsLogic(2)&": "&CStr(neighborAvailableElementsNum(2))
						End If

					End If
					'ii=0
					'get all available elements for side neighbors
					For i=0 To 2
						flag=False
						If StrComp(tempStr(i), tempStr(m))<>0 Then
							neighborAvailableElementsLogic(0)= tempStr(i)
							If Not tempObj(m).getFaceNeighborFromString(tempStr(i), False) Is Nothing Then
								If  Not tempObj(m).getFaceNeighborFromString(tempStr(i), False).getFaceNeighborFromString(tempStr(i), False) Is Nothing Then
									If InStr(neighborAvailableElementsStr(2), _
									tempObj(m).getFaceNeighborFromString(tempStr(i), False).getFaceNeighborFromString(tempStr(i), False).solidName & "$")=0 Then
										'neighborAvailableElementsLogic(ii)= tempStr(i)
										flag=True
										neighborAvailableElementsNum(0)= _
										tempObj(m).getFaceNeighborFromString(tempStr(i),  _
										False).getFaceNeighborFromString(tempStr(i), False).getNumberofPureNonMetalElements(neighborAvailableElementsStr(0))
										ReportInformationToWindow "$Constructor: Available elements for neighbor-"&neighborAvailableElementsLogic(0)&": "&CStr(neighborAvailableElementsNum(0))
									Else
										neighborAvailableElementsNum(0)=neighborAvailableElementsNum(2)
										neighborAvailableElementsStr(0)=neighborAvailableElementsStr(2)
									End If
								End If
							End If
							If InStr(tempStr(i),"n")<>0 Then
								invStr = Left(tempStr(i),1) & "p"
							Else
								invStr = Left(tempStr(i),1) & "n"
							End If
							If Not tempObj(m).getFaceNeighborFromString(tempStr(i), True) Is Nothing Then
								If Not tempObj(m).getFaceNeighborFromString(tempStr(i), True).getFaceNeighborFromString(tempStr(i), True) Is Nothing Then
									If tempObj(m).getFaceNeighborFromString(tempStr(i),True).getFaceNeighborFromString(tempStr(i), True).isPureNonMetal()=True Then
										neighborAvailableElementsLogic(1)= invStr
										If InStr(neighborAvailableElementsStr(2), _
										tempObj(m).getFaceNeighborFromString(tempStr(i), True).getFaceNeighborFromString(tempStr(i), True).solidName & "$")=0 Then
											flag=True
											neighborAvailableElementsNum(1)= _
											tempObj(m).getFaceNeighborFromString(tempStr(i), _
											True).getFaceNeighborFromString(tempStr(i), True).getNumberofPureNonMetalElements(neighborAvailableElementsStr(1))
											ReportInformationToWindow "$Constructor: Available elements for neighbor-"&neighborAvailableElementsLogic(1)&": "&CStr(neighborAvailableElementsNum(1))
										Else
											neighborAvailableElementsNum(1)=neighborAvailableElementsNum(2)
											neighborAvailableElementsStr(1)=neighborAvailableElementsStr(2)
										End If
									End If
								End If
							End If
							If flag=True Then
								Exit For
							End If
							'ii += 1
						End If
					Next
					'***********************************************************************************************************************
					Dim isAnyNeighborAvailable As Boolean
					isAnyNeighborAvailable=False
					For i=0 To 2
						If neighborAvailableElementsNum(i)>=CInt((tgtLength-length)*maxSize*2.6) Then
							isAnyNeighborAvailable=True
							Exit For
						End If
					Next
					'kick out the neighbor that does not meet the target number of available elements
					If isAnyNeighborAvailable=True Then
						j=0
						For i=0 To 2
							If neighborAvailableElementsNum(i) < CInt((tgtLength-length)*maxSize*2.6) And neighborAvailableElementsNum(i)<>0 Then
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
						'Have three different availbleElementsStr()
						If StrComp(neighborAvailableElementsStr(0), neighborAvailableElementsStr(2))<>0 And _
						StrComp(neighborAvailableElementsStr(1), neighborAvailableElementsStr(2))<>0 Then
							For i=0 To 2
								availableStr = availableStr & neighborAvailableElementsStr(i)
								availableNum = availableNum + neighborAvailableElementsNum(i)
							Next
						Else
							For j=0 To 1
								'one of the side neighbor has the same number of availableElements as the straight one
								If StrComp(neighborAvailableElementsStr(j), neighborAvailableElementsStr(2))=0 Then
									availableStr = neighborAvailableElementsStr((1+j) Mod 2) & neighborAvailableElementsStr(2)
									availableNum = neighborAvailableElementsNum((1+j) Mod 2) + neighborAvailableElementsNum(2)
									Exit For
								End If
							Next
						End If
						GoTo BackWard
					End If
				End If
			End If
			'******************************************************************************************************************************
			If n_validFaceNeighbors>0 Then
				'ifOmitFirtEstimationInCorrector=False
				ReDim validFaceNeighbors(n_validFaceNeighbors)
				ReDim validFaceNeighborsStr(n_validFaceNeighbors)
				j=0
				If needReEstimatingAvailableElements=True Then
					For i=0 To UBound(tempObj)-1
						If Not tempObj(i) Is Nothing Then
							Set validFaceNeighbors(j)=tempObj(i)
							validFaceNeighborsStr(j)=tempStr(i)
							j+=1
						End If
					Next
				Else
					For i=0 To UBound(tempObj)-1
						'If InStr(availableStr, tempObj(i).solidName & "$")<>0 Then
						Set validFaceNeighbors(j)=tempObj(i)
						validFaceNeighborsStr(j)=tempStr(i)
						j+=1
						'End If
					Next
				End If
				remainder = iteration Mod (Int(randomCycle*Rnd)+1)
				If remainder = 0 Then
					randomNumber=Int((n_validFaceNeighbors)*Rnd)
					lastRandomNumber = randomNumber
				ElseIf remainder<>0 And n_validFaceNeighbors>lastRandomNumber Then
					randomNumber = lastRandomNumber
				ElseIf remainder<>0 And n_validFaceNeighbors<=lastRandomNumber Then
					randomNumber = 0
				End If
				NON = validFaceNeighbors(randomNumber).getValidFaceNeighbors( _
				vFNeighborsOfVFNeighbors,vFNStrOfvFN, n_vFNoVFN)
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
				ReportInformationToWindow "$Constructor: Set element: "+validFaceNeighbors(k).solidName
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
				If log2Window=True Then ReportInformationToWindow "$Constructor: Antenna Length: "+cStr(Round(length,2))+"mm"
				'***************************************************************************************************
				If needReEstimatingAvailableElements=True Then
					If StrComp(validFaceNeighborsStr(k), neighborAvailableElementsLogic(2))<>0 Then
						'When taking a step forward from the side path
						j=0
						For i=0 To 1
							If StrComp(validFaceNeighborsStr(k), neighborAvailableElementsLogic(i))=0 Then
								availableStr = neighborAvailableElementsStr(i)
								availableNum = neighborAvailableElementsNum(i)
								j=i
								Exit For
							End If
						Next
						removeAvailableElements()
						'When the straight one is valid face neighbor but with 1-N (N is less than target number)
						If neighborAvailableElementsNum(2)>0 And _
						neighborAvailableElementsNum(2)<CInt((tgtLength-length)*maxSize*2.6) Then
							ReportInformationToWindow "$Constructor: Set element: "+tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(j), False).solidName
							tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(j), False).setMaterial(antMaterial)
							Plot.update
							elementNumber+=1
							conLogics = conLogics & neighborAvailableElementsLogic(j)
							Set tailElement = tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(j), False)
							If InStr(neighborAvailableElementsLogic(j),"x")<>0 Then
								length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
							ElseIf InStr(neighborAvailableElementsLogic(j),"y")<>0 Then
								length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
							ElseIf InStr(neighborAvailableElementsLogic(j),"z")<>0 Then
								length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
							End If
							'ant.length+=Abs(currentElement.-currentElement.)
							Set currentElement = tailElement
							'availableStr = neighborAvailableElementsStr(k)
							'availableNum = neighborAvailableElementsNum(k)
							removeAvailableElements()
						End If

					Else
						'when the strait one has no available alements
						If neighborAvailableElementsNum(2)=0 Then
							If neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.6) And _
							neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.6) Then
								k = Int(2*Rnd)
							ElseIf neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.6) And _
							neighborAvailableElementsNum(1)<CInt((tgtLength-length)*maxSize*2.6) Then
								k = 0
							ElseIf neighborAvailableElementsNum(0)<CInt((tgtLength-length)*maxSize*2.6) And _
							neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.6) Then
								k = 1
							End If
							ReportInformationToWindow "$Constructor: Set element: "+tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(k), False).solidName
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
						Else
							availableStr = neighborAvailableElementsStr(2)
							availableNum = neighborAvailableElementsNum(2)
							removeAvailableElements()
							'When having one side neighbor is valid but with number of valid elements less than target, take another step forward to
							'straight
							If (neighborAvailableElementsNum(0)>-1 And neighborAvailableElementsNum(0)<CInt((tgtLength-length)*maxSize*2.6)) And _
							(neighborAvailableElementsNum(1)>-1 And neighborAvailableElementsNum(1)<CInt((tgtLength-length)*maxSize*2.6)) Then
								GoTo Rerouting
							ElseIf (neighborAvailableElementsNum(0)>=CInt((tgtLength-length)*maxSize*2.6)) And _
							(neighborAvailableElementsNum(1)>-1 And neighborAvailableElementsNum(1)<CInt((tgtLength-length)*maxSize*2.6)) Then
								i=Int(2*Rnd)
								If i=0 Then
									k=0
								Else
									k=2
								End If
								GoTo Rerouting
							ElseIf (neighborAvailableElementsNum(1)>=CInt((tgtLength-length)*maxSize*2.6)) And _
							(neighborAvailableElementsNum(0)>-1 And neighborAvailableElementsNum(0)<CInt((tgtLength-length)*maxSize*2.6)) Then
								i=Int(2*Rnd)
								If i=0 Then
									k=1
								Else
									k=2
								End If

								Rerouting:
								ReportInformationToWindow "$Constructor: Set element: "+tailElement.getFaceNeighborFromString(neighborAvailableElementsLogic(k), False).solidName
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
								'availableStr = neighborAvailableElementsStr(k)
								'availableNum = neighborAvailableElementsNum(k)
								removeAvailableElements()
							End If
						End If
					End If
				Else
					removeAvailableElements()
				End If
			End If

		ElseIf ant.length>=tgtLength Then

			'Open logFile For Append As #2
			Print #cFileNumber, conLogics
			Close #cFileNumber
			ReportInformationToWindow "$Constructor: Antenna Logics: "&conLogics
			ReportInformationToWindow "$Constructor: Antenna Length: "&cStr(Round(length,2))&"mm and is over target length."
			Return True
			'Exit Function
		ElseIf n_validFaceNeighbors<=0 And length<tgtLength-1 Then
			'when going into a dead end
			availableStr=""
			availableNum=0
			'*****************************************
			Backward:
			fallBackTimes += 1
			ReportInformationToWindow "$Constructor: fall back counter--"&CStr(fallBackTimes)
			'Open logFile For Append As #2
			Print #cFileNumber, conLogics'+tailElement.solidName
			'Close #2
			'Call corrector to correct the routing of tail part
			ReportInformationToWindow "$Constructor: Iteration #"+cStr(iteration)+"-----------------"
			ReportInformationToWindow "$Constructor: Antenna Logics: "&conLogics
			ReportInformationToWindow "$Constructor: Antenna Length: "&cStr(Round(length,2))&"mm and is still below target length. Backward progress shall begin"
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
				ReportInformationToWindow "$Constructor: Corrector returns nothing object!"
				Close #cFileNumber
				Return False
			End If
			Set currentElement = tailElement
			removeAvailableElements()
			'Return False
		Else
			'Open logFile For Append As #2
			Print #cFileNumber, conLogics
			ReportInformationToWindow "$Constructor: Antenna Logics: "&conLogics
			ReportInformationToWindow "$Constructor: Antenna Length: "&cStr(Round(length,2))&"mm in other conditions, go to next loop"
			'Print #cFileNumber, "$Constructor: Antenna Length: "+cStr(Round(length,2))+"mm in other conditions, go to next loop"
			Close #cFileNumber
			Return True
		End If
		iteration += 1
	Loop Until fallBackTimes>100
	ReportInformationToWindow "$Constructor: Over 100 times falling back trials, return failure"
	'Print #cFileNumber, "$Constructor: Over 10000 times trials, return failure"
	Close #cFileNumber
	Return False
End Function
Function corrector(tgtLength As Double, maxSize As Double, fBR As Double) As AntennaElement
	log2Window=False
	Dim currentElement As AntennaElement
	Dim i As Integer, ii As Integer, iteration As Integer, jj As Integer
	Dim validFaceNeighbors() As AntennaElement
	Dim n_validFaceNeighbors As Integer, validFaceNeighborsStr() As String
	Dim validFaceNeighborsArray() As Integer
	Dim currentNeighborStr As String

	Dim lineRead As String, lineCount As Integer, lineNumber As Integer
	Dim usedNeighborStr As String
	Dim logicLength As Integer
	Dim initialLength As Double

	Dim coFileNumber As Integer

	'Dim pureNonMetalElements As Integer
	'Dim currentConnectionStatus As Boolean
	Dim needToReestimate As Boolean
	Dim tgtCondition As Boolean
	'Dim previousConnectionStatus As Boolean
	coFileNumber=FreeFile
	lineCount = 0
	Open logFile For Input As #coFileNumber
	'Count line number of Routing log.txt
	While Not EOF(coFileNumber)
		Line Input #coFileNumber, lineRead
		lineCount += 1
	Wend
	Close #coFileNumber
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
					ReportInformationToWindow "$Corrector: Need to re-estimate available elements number on going backward."
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
			ReportInformationToWindow "$Corrector: Antenna Length: "+CStr(Round(length,2))+"mm"
			'fall back until the difference between current lenghth and the inital length over the initial
			'lenght is greater than what is requred, or fBR
			'currentElement.setMaterial(subMaterial)
			'pureNonMetalElements+=currentElement.getNumberofPureNonMetalNeighbors()
			ReportInformationToWindow "$Corrector: Target available antenna Elements is " & CInt((tgtLength-length)*maxSize*2.6)
			ReportInformationToWindow "$Corrector: Available antenna Elements is " & CStr(availableNum)
			iteration += 1
			If fBR > 0 Then
				tgtCondition=((initialLength-length)/initialLength > fBR)
			Else
				tgtCondition=(availableNum >= CInt((tgtLength-length)*maxSize*2.6))
			End If
		Loop Until tgtCondition'pureNonMetalElements >= CInt((tgtLength-length)*maxSize*2.6)'Until (initialLength-length)/initialLength > fBR

		currentElement.getValidFaceNeighbors(validFaceNeighbors, validFaceNeighborsStr, n_validFaceNeighbors)
		If n_validFaceNeighbors>1 Then
			'Check routing log to make sure if the lineCount is greater than 1. If so, collect all connection logics
			usedNeighborStr = currentNeighborStr
			lineNumber = lineCount
			coFileNumber=FreeFile
			Open logFile For Input As #coFileNumber
			While lineNumber > 1
				Line Input #coFileNumber, lineRead
				If Len(lineRead) >= logicLength-iteration*2  _
				And StrComp(conLogics, Left(lineRead, logicLength-iteration*2-2))=0 And _
				InStr(usedNeighborStr, Mid(lineRead, logicLength-iteration*2-1, 2))=0 Then
					usedNeighborStr = usedNeighborStr & Mid(lineRead, logicLength-iteration*2-1, 2)
				End If
				lineNumber -=1
			Wend
			Close #coFileNumber
			'need change it to be random instead of sequential!!!!!
			For ii=0 To n_validFaceNeighbors-1

				'Only if the connnection logic is not in the used connection logic collection, do the next step
				If InStr(usedNeighborStr, validFaceNeighborsStr(ii))=0 Then
					currentElement.getFaceNeighborFromString(validFaceNeighborsStr(ii), False).setMaterial(antMaterial)
					Plot.Update
					conLogics = conLogics & validFaceNeighborsStr(ii)
					elementNumber += 1
					If InStr(validFaceNeighborsStr(ii), "x")<>0 Then
						length+=Abs(currentElement.minPoint(0)-currentElement.maxPoint(0))
					ElseIf InStr(validFaceNeighborsStr(ii), "y")<>0 Then
						length+=Abs(currentElement.minPoint(1)-currentElement.maxPoint(1))
					ElseIf InStr(validFaceNeighborsStr(ii), "z")<>0 Then
						length+=Abs(currentElement.minPoint(2)-currentElement.maxPoint(2))
					End If
					If log2Window=True Then ReportInformationToWindow "$Corrector: Antenna Length: "+CStr(Round(length,2))
					Return currentElement.getFaceNeighborFromString(validFaceNeighborsStr(ii), False)
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
			If Not tailElement.nnLNeighbor Is Nothing Then
				If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
					If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=False _
					And tailElement.nnLNeighbor.nnLNeighbor.isMetal()=False Then
						GoTo Checkxnyn
					End If
				Else
				Checkxnyn:
					If Not tailElement.ynNeighbor Is Nothing Then
						If Not tailElement.ynNeighbor.nnLNeighbor Is Nothing Then
							If tailElement.ynNeighbor.nnLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.ynNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
									If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnyn-xnxn"
										Return True
									End If
								End If
								If Not tailElement.ypNeighbor Is Nothing Then
									If Not tailElement.ypNeighbor.npLNeighbor Is Nothing Then
										If tailElement.ypNeighbor.nPLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ypNeighbor.nPLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnyn-ypnpL"
											Return True
										End If
									End If
								End If
								If Not tailElement.npLNeighbor Is Nothing Then
									If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
										If tailElement.npLNeighbor.nPLNeighbor.isPureNonMetal()=True And _
											InStr(availableStr, tailElement.npLNeighbor.nPLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnyn-npLnpL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.npLNeighbor Is Nothing Then
				If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
					If tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=False _
					And tailElement.npLNeighbor.npLNeighbor.isMetal()=False Then
						GoTo Checkxnyp
					End If
				Else
				Checkxnyp:
					If Not tailElement.ypNeighbor Is Nothing Then
						If Not tailElement.ypNeighbor.npLNeighbor Is Nothing Then
							If tailElement.ypNeighbor.npLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.ypNeighbor.npLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
									If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xnNeighbor.xnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnyp-xnxn"
										Return True
									End If
								End If
								If Not tailElement.ynNeighbor Is Nothing Then
									If Not tailElement.ynNeighbor.nnLNeighbor Is Nothing Then
										If tailElement.ynNeighbor.nnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ynNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnyp-ynnnL"
											Return True
										End If
									End If
								End If
								If Not tailElement.nnLNeighbor Is Nothing Then
									If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
										If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.nnLNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnyp-nnLnnL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=False And _
				tailElement.xnNeighbor.xnNeighbor.isMetal()=False Then
					GoTo Checkxnxny
				End If
			Else
			Checkxnxny:
				If Not tailElement.ynNeighbor Is Nothing Then
					If Not tailElement.ynNeighbor.nnLNeighbor Is Nothing Then
						If tailElement.ynNeighbor.nnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.ynNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
							GoTo xnxnL
						End If
					End If
				End If
				If Not tailElement.nnLNeighbor Is Nothing Then
					If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
						If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.nnLNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
						xnxnL:
							If Not tailElement.ypNeighbor Is Nothing Then
								If Not tailElement.ypNeighbor.npLNeighbor Is Nothing Then
									If tailElement.ypNeighbor.nPLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.ypNeighbor.nPLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnxny-ypnpL"
										Return True
									End If
								End If
							End If
							If Not tailElement.npLNeighbor Is Nothing Then
								If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
									If tailElement.npLNeighbor.nPLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.npLNeighbor.nPLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnxny-npLnpL"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.nLnNeighbor Is Nothing Then
				If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
					If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=False _
					And tailElement.nLnNeighbor.nLnNeighbor.isMetal()=False Then
						GoTo Checkxnzn
					End If
				Else
				Checkxnzn:
					If Not tailElement.znNeighbor Is Nothing Then
						If Not tailElement.znNeighbor.nLnNeighbor Is Nothing Then
							If tailElement.znNeighbor.nLnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.znNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
									If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.xnNeighbor.xnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnzn-xnxn"
										Return True
									End If
								End If
								If Not tailElement.zpNeighbor Is Nothing Then
									If Not tailElement.zpNeighbor.nLpNeighbor Is Nothing Then
										If tailElement.zpNeighbor.nLpNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.zpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnzn-zpnLp"
											Return True
										End If
									End If
								End If
								If Not tailElement.nLpNeighbor Is Nothing Then
									If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
										If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.nLpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnzn-nLpnLp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.nLpNeighbor Is Nothing Then
				If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
					If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=False _
					And tailElement.nLpNeighbor.nLpNeighbor.isMetal()=False Then
						GoTo Checkxnzp
					End If
				Else
				Checkxnzp:
					If Not tailElement.zpNeighbor Is Nothing Then
						If Not tailElement.zpNeighbor.nLpNeighbor Is Nothing Then
							If tailElement.zpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.zpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
									If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.xnNeighbor.xnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnzp-xnxn"
										Return True
									End If
								End If
								If Not tailElement.znNeighbor Is Nothing Then
									If Not tailElement.znNeighbor.nLnNeighbor Is Nothing Then
										If tailElement.znNeighbor.nLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.znNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnzp-znnLn"
											Return True
										End If
									End If
								End If
								If Not tailElement.nLnNeighbor Is Nothing Then
									If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
										If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.nLnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xnzp-nLnnLn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.xnNeighbor.xnNeighbor Is Nothing Then
				If tailElement.xnNeighbor.xnNeighbor.isPureNonMetal()=False And _
				tailElement.xnNeighbor.xnNeighbor.isMetal()=False Then
					GoTo Checkxnxnz
				End If
			Else
			Checkxnxnz:
				If Not tailElement.znNeighbor Is Nothing Then
					If Not tailElement.znNeighbor.nLnNeighbor Is Nothing Then
						If tailElement.znNeighbor.nLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.znNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckxnLxn
						End If
					End If
				End If
				If Not tailElement.nLnNeighbor Is Nothing Then
					If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
						If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.nLnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
							CheckxnLxn:
							If Not tailElement.zpNeighbor Is Nothing Then
								If Not tailElement.zpNeighbor.nLpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnxnz-zpnLp"
										Return True
									End If
								End If
							End If
							If Not tailElement.nLpNeighbor Is Nothing Then
								If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
									If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.nLpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xnxnz-nLpnLp"
										Return True
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
			If Not tailElement.pnLNeighbor Is Nothing Then
				If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
					If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=False _
					And tailElement.pnLNeighbor.pnLNeighbor.isMetal()=False Then
						GoTo Checkxpyn
					End If
				Else
				Checkxpyn:
					If Not tailElement.ynNeighbor Is Nothing Then
						If Not tailElement.ynNeighbor.pnLNeighbor Is Nothing Then
							If tailElement.ynNeighbor.pnLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.ynNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
									If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$" )<>0 Then

										ReportInformationToWindow "$Forward check: xpyn-xpxp"
										Return True
									End If
								End If
								If Not tailElement.ypNeighbor Is Nothing Then
									If Not tailElement.ypNeighbor.ppLNeighbor Is Nothing Then
										If tailElement.ypNeighbor.ppLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ypNeighbor.ppLNeighbor.solidName & "$" )<>0 Then

											ReportInformationToWindow "$Forward check: xpyn-ypppL"
											Return True
										End If
									End If
								End If
								If Not tailElement.ppLNeighbor Is Nothing Then
									If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
										If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ppLNeighbor.ppLNeighbor.solidName & "$" )<>0 Then

											ReportInformationToWindow "$Forward check: xpyn-ppLppL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ppLNeighbor Is Nothing Then
				If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
					If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=False _
					And tailElement.ppLNeighbor.ppLNeighbor.isMetal()=False Then
						GoTo Checkxpyp
					End If
				Else
				Checkxpyp:
					If Not tailElement.ypNeighbor Is Nothing Then
						If Not tailElement.ypNeighbor.ppLNeighbor Is Nothing Then
							If tailElement.ypNeighbor.ppLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.ypNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
									If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.xpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpyp-xpxp"
										Return True
									End If
								End If
								If Not tailElement.ynNeighbor Is Nothing Then
									If Not tailElement.ynNeighbor.pnLNeighbor Is Nothing Then
										If tailElement.ynNeighbor.pnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ynNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpyp-ynpnL"
											Return True
										End If
									End If
								End If
								If Not tailElement.pnLNeighbor Is Nothing Then
									If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
										If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.pnLNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpyp-pnLpnL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=False And _
				tailElement.xpNeighbor.xpNeighbor.isMetal()=False Then
					GoTo Checkxpxpy
				End If
			Else
			Checkxpxpy:
				If Not tailElement.ynNeighbor Is Nothing Then
					If Not tailElement.ynNeighbor.pnLNeighbor Is Nothing Then
						If tailElement.ynNeighbor.pnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.ynNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
							GoTo CheckxpxpL
						End If
					End If
				End If
				If Not tailElement.pnLNeighbor Is Nothing Then
					If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
						If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.pnLNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
							CheckxpxpL:
							If Not tailElement.ypNeighbor Is Nothing Then
								If Not tailElement.ypNeighbor.ppLNeighbor Is Nothing Then
									If tailElement.ypNeighbor.ppLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ypNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpxpy-ypppL"
										Return True
									End If
								End If
							End If
							If Not tailElement.ppLNeighbor Is Nothing Then
								If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
									If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ppLNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpxpy-ppLppL"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.pLnNeighbor Is Nothing Then
				If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
					If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=False _
					And tailElement.pLnNeighbor.pLnNeighbor.isMetal()=False Then
						GoTo Checkxpzn
					End If
				Else
				Checkxpzn:
					If Not tailElement.znNeighbor Is Nothing Then
						If Not tailElement.znNeighbor.pLnNeighbor Is Nothing Then
							If tailElement.znNeighbor.pLnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.znNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
									If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.xpNeighbor.xpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpzn-xpxp"
										Return True
									End If
								End If
								If Not tailElement.zpNeighbor Is Nothing Then
									If Not tailElement.zpNeighbor.pLpNeighbor Is Nothing Then
										If tailElement.zpNeighbor.pLpNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.zpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpzn-zppLp"
											Return True
										End If
									End If
								End If
								If Not tailElement.pLpNeighbor Is Nothing Then
									If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
										If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.pLpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpzn-pLppLp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.pLpNeighbor Is Nothing Then
				If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
					If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=False _
					And tailElement.pLpNeighbor.pLpNeighbor.isMetal()=False Then
						GoTo Checkxpzp
					End If
				Else
				Checkxpzp:
					If Not tailElement.zpNeighbor Is Nothing Then
						If Not tailElement.zpNeighbor.pLpNeighbor Is Nothing Then
							If tailElement.zpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.zpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
									If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.xpNeighbor.xpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpzp-xpxp"
										Return True
									End If
								End If
								If Not tailElement.znNeighbor Is Nothing Then
									If Not tailElement.znNeighbor.pLnNeighbor Is Nothing Then
										If tailElement.znNeighbor.pLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.znNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpzp-znpLn"
											Return True
										End If
									End If
								End If
								If Not tailElement.pLnNeighbor Is Nothing Then
									If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
										If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.pLnNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: xpzp-pLnpLn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.xpNeighbor.xpNeighbor Is Nothing Then
				If tailElement.xpNeighbor.xpNeighbor.isPureNonMetal()=False And _
				tailElement.xpNeighbor.xpNeighbor.isMetal()=False Then
					GoTo Checkxpxpz
				End If
			Else
			Checkxpxpz:
				If Not tailElement.znNeighbor Is Nothing Then
					If Not tailElement.znNeighbor.pLnNeighbor Is Nothing Then
						If tailElement.znNeighbor.pLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.znNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckxpLxp
						End If
					End If
				End If
				If Not tailElement.pLnNeighbor Is Nothing Then
					If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
						If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.pLnNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
							CheckxpLxp:
							If Not tailElement.zpNeighbor Is Nothing Then
								If Not tailElement.zpNeighbor.pLpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpxpz-zppLp"
										Return True
									End If
								End If
							End If
							If Not tailElement.pLpNeighbor Is Nothing Then
								If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
									If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.pLpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: xpxpz-pLppLp"
										Return True
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
			If Not tailElement.nnLNeighbor Is Nothing Then
				If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
					If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=False _
					And tailElement.nnLNeighbor.nnLNeighbor.isMetal()=False Then
						GoTo Checkynxn
					End If
				Else
				Checkynxn:
					If Not tailElement.xnNeighbor Is Nothing Then
						If Not tailElement.xnNeighbor.nnLNeighbor Is Nothing Then
							If tailElement.xnNeighbor.nnLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xnNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
									If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynxn-ynyn"
										Return True
									End If
								End If
								If Not tailElement.xpNeighbor Is Nothing Then
									If Not tailElement.xpNeighbor.pnLNeighbor Is Nothing Then
										If tailElement.xpNeighbor.pnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xpNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynxn-xppnL"
											Return True
										End If
									End If
								End If
								If Not tailElement.pnLNeighbor Is Nothing Then
									If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
										If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.pnLNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynxn-pnLpnL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.pnLNeighbor Is Nothing Then
				If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
					If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=False _
					And tailElement.pnLNeighbor.pnLNeighbor.isMetal()=False Then
						GoTo Checkynxp
					End If
				Else
				Checkynxp:
					If Not tailElement.xpNeighbor Is Nothing Then
						If Not tailElement.xpNeighbor.pnLNeighbor Is Nothing Then
							If tailElement.xpNeighbor.pnLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xpNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
									If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ynNeighbor.ynNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynxp-ynyn"
										Return True
									End If
								End If
								If Not tailElement.xnNeighbor Is Nothing Then
									If Not tailElement.xnNeighbor.nnLNeighbor Is Nothing Then
										If tailElement.xnNeighbor.nnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xnNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynxp-xnnnL"
											Return True
										End If
									End If
								End If
								If Not tailElement.nnLNeighbor Is Nothing Then
									If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
										If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.nnLNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynxp-nnLnnL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=False And _
				tailElement.ynNeighbor.ynNeighbor.isMetal()=False Then
					GoTo Checkynynx
				End If
			Else
			Checkynynx:
				If Not tailElement.xnNeighbor Is Nothing Then
					If Not tailElement.xnNeighbor.nnLNeighbor Is Nothing Then
						If tailElement.xnNeighbor.nnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.xnNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
							GoTo ynynL
						End If
					End If
				End If
				If Not tailElement.nnLNeighbor Is Nothing Then
					If Not tailElement.nnLNeighbor.nnLNeighbor Is Nothing Then
						If tailElement.nnLNeighbor.nnLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.nnLNeighbor.nnLNeighbor.solidName & "$" )<>0 Then
						ynynL:
							If Not tailElement.xpNeighbor Is Nothing Then
								If Not tailElement.xpNeighbor.pnLNeighbor Is Nothing Then
									If tailElement.xpNeighbor.pnLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynynx-xppnL"
										Return True
									End If
								End If
							End If
							If Not tailElement.pnLNeighbor Is Nothing Then
								If Not tailElement.pnLNeighbor.pnLNeighbor Is Nothing Then
									If tailElement.pnLNeighbor.pnLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.pnLNeighbor.pnLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynynx-pnLpnL"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.LnnNeighbor Is Nothing Then
				If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
					If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=False _
					And tailElement.LnnNeighbor.LnnNeighbor.isMetal()=False Then
						GoTo Checkynzn
					End If
				Else
				Checkynzn:
					If Not tailElement.znNeighbor Is Nothing Then
						If Not tailElement.znNeighbor.LnnNeighbor Is Nothing Then
							If tailElement.znNeighbor.LnnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.znNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
									If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ynNeighbor.ynNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynzn-ynyn"
										Return True
									End If
								End If
								If Not tailElement.zpNeighbor Is Nothing Then
									If Not tailElement.zpNeighbor.LnpNeighbor Is Nothing Then
										If tailElement.zpNeighbor.LnpNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.zpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynzn-zpLnp"
											Return True
										End If
									End If
								End If
								If Not tailElement.LnpNeighbor Is Nothing Then
									If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
										If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LnpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynzn-LnpLnp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.LnpNeighbor Is Nothing Then
				If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
					If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=False _
					And tailElement.LnpNeighbor.LnpNeighbor.isMetal()=False Then
						GoTo Checkynzp
					End If
				Else
				Checkynzp:
					If Not tailElement.zpNeighbor Is Nothing Then
						If Not tailElement.zpNeighbor.LnpNeighbor Is Nothing Then
							If tailElement.zpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.zpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
									If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ynNeighbor.ynNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynzp-ynyn"
										Return True
									End If
								End If
								If Not tailElement.znNeighbor Is Nothing Then
									If Not tailElement.znNeighbor.LnnNeighbor Is Nothing Then
										If tailElement.znNeighbor.LnnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.znNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynzp-znLnn"
											Return True
										End If
									End If
								End If
								If Not tailElement.LnnNeighbor Is Nothing Then
									If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
										If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LnnNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ynzp-LnnLnn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ynNeighbor.ynNeighbor Is Nothing Then
				If tailElement.ynNeighbor.ynNeighbor.isPureNonMetal()=False And _
				tailElement.ynNeighbor.ynNeighbor.isMetal()=False Then
					GoTo Checkynynz
				End If
			Else
			Checkynynz:
				If Not tailElement.znNeighbor Is Nothing Then
					If Not tailElement.znNeighbor.LnnNeighbor Is Nothing Then
						If tailElement.znNeighbor.LnnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.znNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckLynyn
						End If
					End If
				End If
				If Not tailElement.LnnNeighbor Is Nothing Then
					If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
						If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.LnnNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
							CheckLynyn:
							If Not tailElement.zpNeighbor Is Nothing Then
								If Not tailElement.zpNeighbor.LnpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynynz-zpLnp"
										Return True
									End If
								End If
							End If
							If Not tailElement.LnpNeighbor Is Nothing Then
								If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
									If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.LnpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ynynz-LnpLnp"
										Return True
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
			If Not tailElement.npLNeighbor Is Nothing Then
				If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
					If tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=False _
					And tailElement.npLNeighbor.npLNeighbor.isMetal()=False Then
						GoTo Checkypxn
					End If
				Else
				Checkypxn:
					If Not tailElement.xnNeighbor Is Nothing Then
						If Not tailElement.xnNeighbor.npLNeighbor Is Nothing Then
							If tailElement.xnNeighbor.npLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xnNeighbor.npLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
									If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypxn-ypyp"
										Return True
									End If
								End If
								If Not tailElement.xpNeighbor Is Nothing Then
									If Not tailElement.xpNeighbor.ppLNeighbor Is Nothing Then
										If tailElement.xpNeighbor.ppLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xpNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypxn-xpppL"
											Return True
										End If
									End If
								End If
								If Not tailElement.ppLNeighbor Is Nothing Then
									If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
										If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ppLNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypxn-ppLppL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ppLNeighbor Is Nothing Then
				If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
					If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=False _
					And tailElement.ppLNeighbor.ppLNeighbor.isMetal()=False Then
						GoTo Checkypxp
					End If
				Else
				Checkypxp:
					If Not tailElement.xpNeighbor Is Nothing Then
						If Not tailElement.xpNeighbor.ppLNeighbor Is Nothing Then
							If tailElement.xpNeighbor.ppLNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xpNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
									If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.ypNeighbor.ypNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypxp-ypyp"
										Return True
									End If
								End If
								If Not tailElement.xnNeighbor Is Nothing Then
									If Not tailElement.xnNeighbor.npLNeighbor Is Nothing Then
										If tailElement.xnNeighbor.npLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xnNeighbor.npLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypxp-xnnpL"
											Return True
										End If
									End If
								End If
								If Not tailElement.npLNeighbor Is Nothing Then
									If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
										If tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.npLNeighbor.npLNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypxp-npLnpL"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=False And _
				tailElement.ypNeighbor.ypNeighbor.isMetal()=False Then
					GoTo Checkypypx
				End If
			Else
			Checkypypx:
				If Not tailElement.xnNeighbor Is Nothing Then
					If Not tailElement.xnNeighbor.npLNeighbor Is Nothing Then
						If tailElement.xnNeighbor.npLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.xnNeighbor.npLNeighbor.solidName & "$" )<>0 Then
							GoTo ypypL
						End If
					End If
				End If
				If Not tailElement.npLNeighbor Is Nothing Then
					If Not tailElement.npLNeighbor.npLNeighbor Is Nothing Then
						If tailElement.npLNeighbor.npLNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.npLNeighbor.npLNeighbor.solidName & "$" )<>0 Then
						ypypL:
							If Not tailElement.xpNeighbor Is Nothing Then
								If Not tailElement.xpNeighbor.ppLNeighbor Is Nothing Then
									If tailElement.xpNeighbor.ppLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypypx-xpppL"
										Return True
									End If
								End If
							End If
							If Not tailElement.ppLNeighbor Is Nothing Then
								If Not tailElement.ppLNeighbor.ppLNeighbor Is Nothing Then
									If tailElement.ppLNeighbor.ppLNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ppLNeighbor.ppLNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypypx-ppLppL"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"z")<>0 Then
			If Not tailElement.LpnNeighbor Is Nothing Then
				If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
					If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=False _
					And tailElement.LpnNeighbor.LpnNeighbor.isMetal()=False Then
						GoTo Checkypzn
					End If
				Else
				Checkypzn:
					If Not tailElement.znNeighbor Is Nothing Then
						If Not tailElement.znNeighbor.LpnNeighbor Is Nothing Then
							If tailElement.znNeighbor.LpnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.znNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
									If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ypNeighbor.ypNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypzn-ypyp"
										Return True
									End If
								End If
								If Not tailElement.zpNeighbor Is Nothing Then
									If Not tailElement.zpNeighbor.LppNeighbor Is Nothing Then
										If tailElement.zpNeighbor.LppNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.zpNeighbor.LppNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypzn-zpLpp"
											Return True
										End If
									End If
								End If
								If Not tailElement.LppNeighbor Is Nothing Then
									If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
										If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LppNeighbor.LppNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypzn-LppLpp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.LppNeighbor Is Nothing Then
				If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
					If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=False _
					And tailElement.LppNeighbor.LppNeighbor.isMetal()=False Then
						GoTo Checkypzp
					End If
				Else
				Checkypzp:
					If Not tailElement.zpNeighbor Is Nothing Then
						If Not tailElement.zpNeighbor.LppNeighbor Is Nothing Then
							If tailElement.zpNeighbor.LppNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.zpNeighbor.LppNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
									If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ypNeighbor.ypNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypzp-ypyp"
										Return True
									End If
								End If
								If Not tailElement.znNeighbor Is Nothing Then
									If Not tailElement.znNeighbor.LpnNeighbor Is Nothing Then
										If tailElement.znNeighbor.LpnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.znNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
												ReportInformationToWindow "$Forward check: ypzp-znLpn"
												Return True
										End If
									End If
								End If
								If Not tailElement.LpnNeighbor Is Nothing Then
									If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
										If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LpnNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: ypzp-LpnLpn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.ypNeighbor.ypNeighbor Is Nothing Then
				If tailElement.ypNeighbor.ypNeighbor.isPureNonMetal()=False And _
				tailElement.ypNeighbor.ypNeighbor.isMetal()=False Then
					GoTo Checkypypz
				End If
			Else
			Checkypypz:
				If Not tailElement.znNeighbor Is Nothing Then
					If Not tailElement.znNeighbor.LpnNeighbor Is Nothing Then
						If tailElement.znNeighbor.LpnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.znNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckLypyp
						End If
					End If
				End If
				If Not tailElement.LpnNeighbor Is Nothing Then
					If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
						If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.LpnNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
							CheckLypyp:
							If Not tailElement.zpNeighbor Is Nothing Then
								If Not tailElement.zpNeighbor.LppNeighbor Is Nothing Then
									If tailElement.zpNeighbor.LppNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.LppNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypypz-zpLpp"
										Return True
									End If
								End If
							End If
							If Not tailElement.LppNeighbor Is Nothing Then
								If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
									If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.LppNeighbor.LppNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: ypypz-LppLpp"
										Return True
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
			If Not tailElement.nLnNeighbor Is Nothing Then
				If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
					If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=False _
					And tailElement.nLnNeighbor.nLnNeighbor.isMetal()=False Then
						GoTo Checkznxn
					End If
				Else
				Checkznxn:
					If Not tailElement.xnNeighbor Is Nothing Then
						If Not tailElement.xnNeighbor.nLnNeighbor Is Nothing Then
							If tailElement.xnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
									If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znxn-znzn"
										Return True
									End If
								End If
								If Not tailElement.xpNeighbor Is Nothing Then
									If Not tailElement.xpNeighbor.pLnNeighbor Is Nothing Then
										If tailElement.xpNeighbor.pLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xpNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znxn-xppLn"
											Return True
										End If
									End If
								End If
								If Not tailElement.pLnNeighbor Is Nothing Then
									If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
										If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.pLnNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znxn-pLnpLn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.pLnNeighbor Is Nothing Then
				If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
					If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=False _
					And tailElement.pLnNeighbor.pLnNeighbor.isMetal()=False Then
						GoTo Checkznxp
					End If
				Else
				Checkznxp:
					If Not tailElement.xpNeighbor Is Nothing Then
						If Not tailElement.xpNeighbor.pLnNeighbor Is Nothing Then
							If tailElement.xpNeighbor.pLnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xpNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
									If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.znNeighbor.znNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znxp-znzn"
										Return True
									End If
								End If
								If Not tailElement.xnNeighbor Is Nothing Then
									If Not tailElement.xnNeighbor.nLnNeighbor Is Nothing Then
										If tailElement.xnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znxp-xnnLn"
											Return True
										End If
									End If
								End If
								If Not tailElement.nLnNeighbor Is Nothing Then
									If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
										If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.nLnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znxp-nLnnLn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=False And _
				tailElement.znNeighbor.znNeighbor.isMetal()=False Then
					GoTo Checkznznx
				End If
			Else
			Checkznznx:
				If Not tailElement.xnNeighbor Is Nothing Then
					If Not tailElement.xnNeighbor.nLnNeighbor Is Nothing Then
						If tailElement.xnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.xnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckznznL
						End If
					End If
				End If
				If Not tailElement.nLnNeighbor Is Nothing Then
					If Not tailElement.nLnNeighbor.nLnNeighbor Is Nothing Then
						If tailElement.nLnNeighbor.nLnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.nLnNeighbor.nLnNeighbor.solidName & "$" )<>0 Then
						CheckznznL:
							If Not tailElement.xpNeighbor Is Nothing Then
								If Not tailElement.xpNeighbor.pLnNeighbor Is Nothing Then
									If tailElement.xpNeighbor.pLnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znznx-xppLn"
										Return True
									End If
								End If
							End If
							If Not tailElement.pLnNeighbor Is Nothing Then
								If Not tailElement.pLnNeighbor.pLnNeighbor Is Nothing Then
									If tailElement.pLnNeighbor.pLnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.pLnNeighbor.pLnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znznx-pLnpLn"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.LnnNeighbor Is Nothing Then
				If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
					If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=False _
					And tailElement.LnnNeighbor.LnnNeighbor.isMetal()=False Then
						GoTo Checkznyn
					End If
				Else
				Checkznyn:
					If Not tailElement.ynNeighbor Is Nothing Then
						If Not tailElement.ynNeighbor.LnnNeighbor Is Nothing Then
							If tailElement.ynNeighbor.LnnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.ynNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
									If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.znNeighbor.znNeighbor.solidName & "$" )<>0 Then
									ReportInformationToWindow "$Forward check: znyn-znzn"
									Return True
									End If
								End If
								If Not tailElement.ypNeighbor Is Nothing Then
									If Not tailElement.ypNeighbor.LpnNeighbor Is Nothing Then
										If tailElement.ypNeighbor.LpnNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.ypNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znyn-ypLpn"
											Return True
										End If
									End If
								End If
								If Not tailElement.LpnNeighbor Is Nothing Then
									If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
										If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LpnNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znyn-LpnLpn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.LpnNeighbor Is Nothing Then
				If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
					If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=False _
					And tailElement.LpnNeighbor.LpnNeighbor.isMetal()=False Then
						GoTo Checkznyp
					End If
				Else
				Checkznyp:
					If Not tailElement.ypNeighbor Is Nothing Then
						If Not tailElement.ypNeighbor.LpnNeighbor Is Nothing Then
							If tailElement.ypNeighbor.LpnNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.ypNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
									If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.znNeighbor.znNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znyp-znzn"
										Return True
									End If
								End If
								If Not tailElement.ynNeighbor Is Nothing Then
									If Not tailElement.ynNeighbor.LnnNeighbor Is Nothing Then
										If tailElement.ynNeighbor.LnnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.ynNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znyp-ynLnn"
											Return True
										End If
									End If
								End If
								If Not tailElement.LnnNeighbor Is Nothing Then
									If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
										If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LnnNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: znyp-LnnLnn"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.znNeighbor.znNeighbor Is Nothing Then
				If tailElement.znNeighbor.znNeighbor.isPureNonMetal()=False And _
				tailElement.znNeighbor.znNeighbor.isMetal()=False Then
					GoTo Checkznzny
				End If
			Else
			Checkznzny:
				If Not tailElement.ynNeighbor Is Nothing Then
					If Not tailElement.ynNeighbor.LnnNeighbor Is Nothing Then
						If tailElement.ynNeighbor.LnnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.ynNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
							GoTo CheckLznzn
						End If
					End If
				End If
				If Not tailElement.LnnNeighbor Is Nothing Then
					If Not tailElement.LnnNeighbor.LnnNeighbor Is Nothing Then
						If tailElement.LnnNeighbor.LnnNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.LnnNeighbor.LnnNeighbor.solidName & "$" )<>0 Then
							CheckLznzn:
							If Not tailElement.ypNeighbor Is Nothing Then
								If Not tailElement.ypNeighbor.LpnNeighbor Is Nothing Then
									If tailElement.ypNeighbor.LpnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ypNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znzny-ypLpn"
										Return True
									End If
								End If
							End If
							If Not tailElement.LpnNeighbor Is Nothing Then
								If Not tailElement.LpnNeighbor.LpnNeighbor Is Nothing Then
									If tailElement.LpnNeighbor.LpnNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.LpnNeighbor.LpnNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: znzny-LpnLpn"
										Return True
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
			If Not tailElement.nLpNeighbor Is Nothing Then
				If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
					If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=False _
					And tailElement.nLpNeighbor.nLpNeighbor.isMetal()=False Then
						GoTo Checkzpxn
					End If
				Else
				Checkzpxn:
					If Not tailElement.xnNeighbor Is Nothing Then
						If Not tailElement.xnNeighbor.nLpNeighbor Is Nothing Then
							If tailElement.xnNeighbor.nLpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xnNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpxn-zpzp"
										Return True
									End If
								End If
								If Not tailElement.xpNeighbor Is Nothing Then
									If Not tailElement.xpNeighbor.pLpNeighbor Is Nothing Then
										If tailElement.xpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpxn-xppLp"
											Return True
										End If
									End If
								End If
								If Not tailElement.pLpNeighbor Is Nothing Then
									If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
										If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.pLpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpxn-pLppLp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.pLpNeighbor Is Nothing Then
				If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
					If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=False _
					And tailElement.pLpNeighbor.pLpNeighbor.isMetal()=False Then
						GoTo Checkzpxp
					End If
				Else
				Checkzpxp:
					If Not tailElement.xpNeighbor Is Nothing Then
						If Not tailElement.xpNeighbor.pLpNeighbor Is Nothing Then
							If tailElement.xpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr, tailElement.xpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.zpNeighbor.zpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpxp-zpzp"
										Return True
									End If
								End If
								If Not tailElement.xnNeighbor Is Nothing Then
									If Not tailElement.xnNeighbor.nLpNeighbor Is Nothing Then
										If tailElement.xnNeighbor.nLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.xnNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpxp-xnnLp"
											Return True
										End If
									End If
								End If
								If Not tailElement.nLpNeighbor Is Nothing Then
									If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
										If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr, tailElement.nLpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpxp-nLpnLp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=False And _
				tailElement.zpNeighbor.zpNeighbor.isMetal()=False Then
					GoTo Checkzpzpx
				End If
			Else
			Checkzpzpx:
				If Not tailElement.xnNeighbor Is Nothing Then
					If Not tailElement.xnNeighbor.nLpNeighbor Is Nothing Then
						If tailElement.xnNeighbor.nLpNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.xnNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
							GoTo CheckzpzpL
						End If
					End If
				End If
				If Not tailElement.nLpNeighbor Is Nothing Then
					If Not tailElement.nLpNeighbor.nLpNeighbor Is Nothing Then
						If tailElement.nLpNeighbor.nLpNeighbor.isPureNonMetal()=True And _
						InStr(availableStr, tailElement.nLpNeighbor.nLpNeighbor.solidName & "$" )<>0 Then
						CheckzpzpL:
							If Not tailElement.xpNeighbor Is Nothing Then
								If Not tailElement.xpNeighbor.pLpNeighbor Is Nothing Then
									If tailElement.xpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr, tailElement.xpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpzpx-xppLp"
										Return True
									End If
								End If
							End If
							If Not tailElement.pLpNeighbor Is Nothing Then
								If Not tailElement.pLpNeighbor.pLpNeighbor Is Nothing Then
									If tailElement.pLpNeighbor.pLpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.pLpNeighbor.pLpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpzpx-pLppLp"
										Return True
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If InStr(neighborsStr,"y")<>0 Then
			If Not tailElement.LnpNeighbor Is Nothing Then
				If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
					If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=False _
					And tailElement.LnpNeighbor.LnpNeighbor.isMetal()=False Then
						GoTo Checkzpyn
					End If
				Else
				Checkzpyn:
					If Not tailElement.ynNeighbor Is Nothing Then
						If Not tailElement.ynNeighbor.LnpNeighbor Is Nothing Then
							If tailElement.ynNeighbor.LnpNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.ynNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.zpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpyn-zpzp"
										Return True
									End If
								End If
								If Not tailElement.ypNeighbor Is Nothing Then
									If Not tailElement.ypNeighbor.LppNeighbor Is Nothing Then
										If tailElement.ypNeighbor.LppNeighbor.isPureNonMetal()=True  And _
										InStr(availableStr,tailElement.ypNeighbor.LppNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpyn-ypLpp"
											Return True
										End If
									End If
								End If
								If Not tailElement.LppNeighbor Is Nothing Then
									If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
										If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LppNeighbor.LppNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpyn-LppLpp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.LppNeighbor Is Nothing Then
				If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
					If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=False _
					And tailElement.LppNeighbor.LppNeighbor.isMetal()=False Then
						GoTo Checkzpyp
					End If
				Else
				Checkzpyp:
					If Not tailElement.ypNeighbor Is Nothing Then
						If Not tailElement.ypNeighbor.LppNeighbor Is Nothing Then
							If tailElement.ypNeighbor.LppNeighbor.isPureNonMetal()=True And _
							InStr(availableStr,tailElement.ypNeighbor.LppNeighbor.solidName & "$" )<>0 Then
								If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
									If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.zpNeighbor.zpNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpyp-zpzp"
										Return True
									End If
								End If
								If Not tailElement.ynNeighbor Is Nothing Then
									If Not tailElement.ynNeighbor.LnpNeighbor Is Nothing Then
										If tailElement.ynNeighbor.LnpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.ynNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpyp-ynLnp"
											Return True
										End If
									End If
								End If
								If Not tailElement.LnpNeighbor Is Nothing Then
									If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
										If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
										InStr(availableStr,tailElement.LnpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
											ReportInformationToWindow "$Forward check: zpyp-LnpLnp"
											Return True
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
			If Not tailElement.zpNeighbor.zpNeighbor Is Nothing Then
				If tailElement.zpNeighbor.zpNeighbor.isPureNonMetal()=False And _
				tailElement.zpNeighbor.zpNeighbor.isMetal()=False Then
					GoTo Checkzpzpy
				End If
			Else
			Checkzpzpy:
				If Not tailElement.ynNeighbor Is Nothing Then
					If Not tailElement.ynNeighbor.LnpNeighbor Is Nothing Then
						If tailElement.ynNeighbor.LnpNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.ynNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
							GoTo CheckLzpzp
						End If
					End If
				End If
				If Not tailElement.LnpNeighbor Is Nothing Then
					If Not tailElement.LnpNeighbor.LnpNeighbor Is Nothing Then
						If tailElement.LnpNeighbor.LnpNeighbor.isPureNonMetal()=True And _
						InStr(availableStr,tailElement.LnpNeighbor.LnpNeighbor.solidName & "$" )<>0 Then
							CheckLzpzp:
							If Not tailElement.ypNeighbor Is Nothing Then
								If Not tailElement.ypNeighbor.LppNeighbor Is Nothing Then
									If tailElement.ypNeighbor.LppNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.ypNeighbor.LppNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpzpy-ypLpp"
										Return True
									End If
								End If
							End If
							If Not tailElement.LppNeighbor Is Nothing Then
								If Not tailElement.LppNeighbor.LppNeighbor Is Nothing Then
									If tailElement.LppNeighbor.LppNeighbor.isPureNonMetal()=True And _
									InStr(availableStr,tailElement.LppNeighbor.LppNeighbor.solidName & "$" )<>0 Then
										ReportInformationToWindow "$Forward check: zpzpy-LppLpp"
										Return True
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
Public Function reconstructor() As Boolean
	Dim A As AntennaElement, B As AntennaElement
	Dim i As Integer
	Set A=feedElement
	For i=0 To elementNumber-2
		Set B = A.getFaceNeighborFromString(Mid(conLogics,2*i+1,2),False)
		B.setMaterial(antMaterial)
		Set A = B
		Plot.Update
	Next

	Return True
	'Dim historyCaption As String
	'Dim sCommand As String
	'historyCaption = "$SA$"
	'sCommand = ""
	'AddToHistory(historyCaption, sCommand)
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
