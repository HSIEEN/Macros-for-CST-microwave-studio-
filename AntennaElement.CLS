VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_PredeclaredId = False
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Name = "AntennaElement"
'#Language "WWB-COM"
Option Explicit
Private sName As String
Property Get solidName() As String
	solidName = sName
End Property
Property Let solidName(value As String)
		sName = value
End Property
Private aMaterial As String
Property Get solidMaterial() As String
	solidMaterial = aMaterial
End Property
Property Let solidMaterial(value As String)
		aMaterial = value
End Property
Private ele_flag As Boolean
Property Get flag() As Boolean
	flag = ele_flag
End Property
Property Let flag(value As Boolean)
		ele_flag = value
End Property
Private startPoint() As Double
Property Get minPoint() As Double()
	minPoint = startPoint
End Property

'endPoint--> xMax, yMax, zMax
Private endPoint() As Double
Property Get maxPoint() As Double()
	maxPoint = endPoint
End Property
'vertices-->xnynzn,xnynzp,xnypzn,xnypzp,xpynzn,xpynzp,xpypzn,xpypzp
'Private xnynzn() As Double
'Private xnynzp() As Double
'Private xnypzn() As Double
'Private xnypzp() As Double
'Private xpynzn() As Double
'Private xpynzp() As Double
'Private xpypzn() As Double
'Private xpypzp() As Double
Private vertices() As Double
Property Get nnn() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(0,i)
	Next
	nnn = vertice
End Property
Property Get nnp() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(1,i)
	Next
	nnp = vertice
End Property
Property Get npn() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(2,i)
	Next
	npn = vertice
End Property
Property Get npp() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(3,i)
	Next
	npp = vertice
End Property
Property Get pnn() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(4,i)
	Next
	pnn = vertice
End Property
Property Get pnp() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(5,i)
	Next
	pnp = vertice
End Property
Property Get ppn() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(6,i)
	Next
	ppn = vertice
End Property
Property Get ppp() As Double()
	Dim vertice() As Double, i As Integer
	ReDim vertice(3)
	For i=0 To 2
		vertice(i)=vertices(7,i)
	Next
	ppp = vertice
End Property
'edges: nnn2nnp,nnn2npn,nnn2pnn### npp2npn, npp2nnp,npp2ppp###pnp2pnn,pnp2nnp,pnp2ppp###ppn2npn,ppn2pnn,ppn2ppp
'Private nnn2nnp() As Double, nnn2npn() As Double, nnn2pnn() As Double
'Private npp2nnp() As Double, npp2npn() As Double, npp2ppp() As Double
'Private pnp2nnp() As Double, pnp2pnn() As Double, pnp2ppp() As Double
'Private ppn2npn() As Double, ppn2pnn() As Double, ppn2ppp() As Double
Private edges() As Double
'nnn2nnp
Property Get nnL() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(0,i,j)
		Next
	Next
	nnL = edge
End Property
'nnn2npn
Property Get nLn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(1,i,j)
		Next
	Next
	nLn = edge
End Property
'nnn2pnn
Property Get Lnn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(2,i,j)
		Next
	Next
	Lnn = edge
End Property
'npp2nnp
Property Get nLp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(3,i,j)
		Next
	Next
	nLp = edge
End Property
'npp2npn
Property Get npL() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(4,i,j)
		Next
	Next
	npL = edge
End Property
'npp2ppp
Property Get Lpp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(5,i,j)
		Next
	Next
	Lpp = edge
End Property
'pnp2nnp
Property Get Lnp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(6,i,j)
		Next
	Next
	Lnp = edge
End Property
'pnp2pnn
Property Get pnL() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(7,i,j)
		Next
	Next
	pnL = edge
End Property
'pnp2ppp
Property Get pLP() As Double()
	'pLPEdge = pLP
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(8,i,j)
		Next
	Next
	pLP = edge
End Property
'ppn2npn
Property Get Lpn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(9,i,j)
		Next
	Next
	Lpn = edge
End Property
'ppn2pnn
Property Get pLn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(10,i,j)
		Next
	Next
	pLn = edge
End Property
'ppn2ppp
Property Get ppL() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(11,i,j)
		Next
	Next
	ppL = edge
End Property

'faces--> xnFace, xpFace, ynFace, ypFace, znFace, zpFace
'Private faceXn() As Double
'Private faceXp() As Double
'Private faceYn() As Double
'Private faceYp() As Double
'Private faceZn() As Double
'Private faceZp() As Double
Private faces() As Double
Property Get xnFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(0,i,j)
		Next
	Next
	xnFace = nFace
End Property
Property Get xpFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(1,i,j)
		Next
	Next
	xpFace = nFace
End Property
Property Get ynFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(2,i,j)
		Next
	Next
	ynFace = nFace
End Property
Property Get ypFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(3,i,j)
		Next
	Next
	ypFace = nFace
End Property
Property Get znFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(4,i,j)
		Next
	Next
	znFace = nFace
End Property
Property Get zpFace() As Double()
	Dim nFace() As Double
	Dim i As Integer, j As Integer
	ReDim nFace(4,3)
	For i=0 To 3
		For j=0 To 2
			nFace(i,j)=faces(5,i,j)
		Next
	Next
	zpFace = nFace
End Property
'vertice neighbors
Private verticeNeighbors(8) As AntennaElement
'nnnNeighbor
Property Get nnnNeighbor() As AntennaElement
	Set nnnNeighbor = verticeNeighbors(0)
End Property
Property Set nnnNeighbor(value As AntennaElement)
	Set verticeNeighbors(0) = value
End Property
'nnpNeighbor
Property Get nnpNeighbor() As AntennaElement
	Set nnpNeighbor = verticeNeighbors(1)
End Property
Property Set nnpNeighbor(value As AntennaElement)
	Set verticeNeighbors(1) = value
End Property
'npnNeighbor

Property Get npnNeighbor() As AntennaElement
	Set npnNeighbor = verticeNeighbors(2)
End Property
Property Set npnNeighbor(value As AntennaElement)
	Set verticeNeighbors(2) = value
End Property
'nppNeighbor
Property Get nppNeighbor() As AntennaElement
	Set nppNeighbor = verticeNeighbors(3)
End Property
Property Set nppNeighbor(value As AntennaElement)
	Set verticeNeighbors(3) = value
End Property
'pnnNeighbor

Property Get pnnNeighbor() As AntennaElement
	Set pnnNeighbor = verticeNeighbors(4)
End Property
Property Set pnnNeighbor(value As AntennaElement)
	Set verticeNeighbors(4) = value
End Property
'pnpNeighbor
Property Get pnpNeighbor() As AntennaElement
	Set pnpNeighbor = verticeNeighbors(5)
End Property
Property Set pnpNeighbor(value As AntennaElement)
	Set verticeNeighbors(5) = value
End Property
'ppnNeighbor

Property Get ppnNeighbor() As AntennaElement
	Set ppnNeighbor = verticeNeighbors(6)
End Property
Property Set ppnNeighbor(value As AntennaElement)
	Set verticeNeighbors(6) = value
End Property
'pppNeighbor
Property Get pppNeighbor() As AntennaElement
	Set pppNeighbor = verticeNeighbors(7)
End Property
Property Set pppNeighbor(value As AntennaElement)
	Set verticeNeighbors(7) = value
End Property

Private faceNeighbors(6) As AntennaElement
'Private faceNeighborXn As AntennaElement
'Private ReDim faceNeighbors(6)
Property Get xnNeighbor() As AntennaElement
	Set xnNeighbor = faceNeighbors(0)
End Property
Property Set xnNeighbor(value As AntennaElement)
	Set faceNeighbors(0) = value
End Property

Property Get xpNeighbor() As AntennaElement
	Set xpNeighbor = faceNeighbors(1)
End Property
Property Set xpNeighbor(value As AntennaElement)
	Set faceNeighbors(1) = value
End Property

Property Get ynNeighbor() As AntennaElement
	Set ynNeighbor = faceNeighbors(2)
End Property
Property Set ynNeighbor(value As AntennaElement)
	Set faceNeighbors(2) = value
End Property
Property Get ypNeighbor() As AntennaElement
	Set ypNeighbor = faceNeighbors(3)
End Property
Property Set ypNeighbor(value As AntennaElement)
	Set faceNeighbors(3) = value
End Property

Property Get znNeighbor() As AntennaElement
	Set znNeighbor = faceNeighbors(4)
End Property
Property Set znNeighbor(value As AntennaElement)
	Set faceNeighbors(4) = value
End Property

Property Get zpNeighbor() As AntennaElement
	Set zpNeighbor = faceNeighbors(5)
End Property
Property Set zpNeighbor(value As AntennaElement)
	Set faceNeighbors(5) = value
End Property


'===========================edge Neighbors============================
	' nnL(2,3)
	' nLn(2,3)
	' Lnn(2,3)
	' nLp(2,3)
	' npL(2,3)
	' Lpp(2,3)
	' Lnp(2,3)
	' pnL(2,3)
	' pLP(2,3)
	' Lpn(2,3)
	' pLn(2,3)
	' ppL(2,3)
'=================================nnL
Private edgeNeighbors(12) As AntennaElement
'ReDim edgeNeighbors(12)
'Private edgeNeighbornnL As AntennaElement
Property Get nnLNeighbor() As AntennaElement
	Set nnLNeighbor = edgeNeighbors(0)
End Property
Property Set nnLNeighbor(value As AntennaElement)
	Set edgeNeighbors(0) = value
End Property
'=================================nLn
Property Get nLnNeighbor() As AntennaElement
	Set nLnNeighbor = edgeNeighbors(1)
End Property
Property Set nLnNeighbor(value As AntennaElement)
	Set edgeNeighbors(1) = value
End Property
'=================================Lnn
Property Get LnnNeighbor() As AntennaElement
	Set LnnNeighbor = edgeNeighbors(2)
End Property
Property Set LnnNeighbor(value As AntennaElement)
	Set edgeNeighbors(2) = value
End Property
'=================================nLp
Property Get nLpNeighbor() As AntennaElement
	Set nLpNeighbor = edgeNeighbors(3)
End Property
Property Set nLpNeighbor(value As AntennaElement)
	Set edgeNeighbors(3) = value
End Property
'===================================npL
Property Get npLNeighbor() As AntennaElement
	Set npLNeighbor = edgeNeighbors(4)
End Property
Property Set npLNeighbor(value As AntennaElement)
	Set edgeNeighbors(4) = value
End Property
'=====================================Lpp
Property Get LppNeighbor() As AntennaElement
	Set LppNeighbor = edgeNeighbors(5)
End Property
Property Set LppNeighbor(value As AntennaElement)
	Set edgeNeighbors(5) = value
End Property
'========================================Lnp
Property Get LnpNeighbor() As AntennaElement
	Set LnpNeighbor = edgeNeighbors(6)
End Property
Property Set LnpNeighbor(value As AntennaElement)
	Set edgeNeighbors(6) = value
End Property
'========================================pnL
Property Get pnLNeighbor() As AntennaElement
	Set pnLNeighbor = edgeNeighbors(7)
End Property
Property Set pnLNeighbor(value As AntennaElement)
	Set edgeNeighbors(7) = value
End Property
'========================================pLP
Property Get pLPNeighbor() As AntennaElement
	Set pLPNeighbor = edgeNeighbors(8)
End Property
Property Set pLPNeighbor(value As AntennaElement)
	Set edgeNeighbors(8) = value
End Property
'========================================Lpn
Property Get LpnNeighbor() As AntennaElement
	Set LpnNeighbor = edgeNeighbors(9)
End Property
Property Set LpnNeighbor(value As AntennaElement)
	Set edgeNeighbors(9) = value
End Property
'========================================pLn
Property Get pLnNeighbor() As AntennaElement
	Set pLnNeighbor = edgeNeighbors(10)
End Property
Property Set pLnNeighbor(value As AntennaElement)
	Set edgeNeighbors(10) = value
End Property
'========================================ppL
Property Get ppLNeighbor() As AntennaElement
	Set ppLNeighbor = edgeNeighbors(11)
End Property
Property Set ppLNeighbor(value As AntennaElement)
	Set edgeNeighbors(11) = value
End Property
'===============================set Material==================================
Public Sub setMaterial(mMaterial As String)
	Me.solidMaterial = mMaterial
	Solid.ChangeMaterial(Me.solidName,mMaterial)
End Sub
Public Function setMaterialPermanently(mMaterial As String) As String
	'Dim historyCaption As String
	Dim sCommand As String
	Me.solidMaterial = mMaterial
	sCommand = ""
	sCommand = sCommand + "Solid.ChangeMaterial """ + Me.solidName + """,""" + mMaterial + """"+vbLf
	'historyCaption = "$SMP" + " " + Me.solidName + "---" + mMaterial
	'AddToHistory(historyCaption,sCommand)
	Return sCommand
End Function

Public Sub setStartPoint(xMin As Double, yMin As Double, zMin As Double)
	ReDim startPoint(3)
	startPoint(0) = xMin
	startPoint(1) = yMin
	startPoint(2) = zMin
End Sub
Public Sub setEndPoint(xMax As Double, yMax As Double, zMax As Double)
	ReDim endPoint(3)
	endPoint(0) = xMax
	endPoint(1) = yMax
	endPoint(2) = zMax
End Sub
Public Sub defineVertices()
	ReDim vertices(8,3)
	'nnn
	vertices(0,0) = startPoint(0)
	vertices(0,1) = startPoint(1)
	vertices(0,2) = startPoint(2)
	'nnp
	vertices(1,0) = startPoint(0)
	vertices(1,1) = startPoint(1)
	vertices(1,2) = endPoint(2)
	'npn
	vertices(2,0) = startPoint(0)
	vertices(2,1) = endPoint(1)
	vertices(2,2) = startPoint(2)
	'npp
	vertices(3,0) = startPoint(0)
	vertices(3,1) = endPoint(1)
	vertices(3,2) = endPoint(2)
	'pnn
	vertices(4,0) = endPoint(0)
	vertices(4,1) = startPoint(1)
	vertices(4,2) = startPoint(2)
	'pnp
	vertices(5,0) = endPoint(0)
	vertices(5,1) = startPoint(1)
	vertices(5,2) = endPoint(2)
	'ppn
	vertices(6,0) = endPoint(0)
	vertices(6,1) = endPoint(1)
	vertices(6,2) = startPoint(2)
	'ppp
	vertices(7,0) = endPoint(0)
	vertices(7,1) = endPoint(1)
	vertices(7,2) = endPoint(2)
End Sub
Public Sub defineEdges()
	ReDim edges(12,2,3)
	Dim i As Integer, j As Integer
	'nnL
	For i=0 To 2
		edges(0,0,i)=nnn(i)
		edges(0,1,i)=nnp(i)
	Next
	'nLn
	For i=0 To 2
		edges(1,0,i)=nnn(i)
		edges(1,1,i)=npn(i)
	Next
	'Lnn
	For i=0 To 2
		edges(2,0,i)=nnn(i)
		edges(2,1,i)=pnn(i)
	Next
	'nLp
	For i=0 To 2
		edges(3,0,i)=npp(i)
		edges(3,1,i)=nnp(i)
	Next
	'npL
	For i=0 To 2
		edges(4,0,i)=npp(i)
		edges(4,1,i)=npn(i)
	Next
	'Lpp
	For i=0 To 2
		edges(5,0,i)=npp(i)
		edges(5,1,i)=ppp(i)
	Next
	'Lnp
	For i=0 To 2
		edges(6,0,i)=pnp(i)
		edges(6,1,i)=nnp(i)
	Next
	'pnL
	For i=0 To 2
		edges(7,0,i)=pnp(i)
		edges(7,1,i)=pnn(i)
	Next
	'pLP
	For i=0 To 2
		edges(8,0,i)=pnp(i)
		edges(8,1,i)=ppp(i)
	Next
	'Lpn
	For i=0 To 2
		edges(9,0,i)=ppn(i)
		edges(9,1,i)=npn(i)
	Next
	'pLn
	For i=0 To 2
		edges(10,0,i)=ppn(i)
		edges(10,1,i)=pnn(i)
	Next
	'ppL
	For i=0 To 2
		edges(11,0,i)=ppn(i)
		edges(11,1,i)=ppp(i)
	Next

End Sub

'The changes made include using the   +=   operator To concatenate the strings, And using   IsNot   instead of   Not   For Null/Nothing checks.
Public Sub defineFaces()
	ReDim faces(6,4,3)
	Dim i As Integer, j As Integer, m As Integer
	'faceXn --> (xmin, ymin, zmin),(xmin,ymin, zmax),(xmin,ymax,zmin),(xmin,ymax,zmax)
	'faceXp --> (xmax, ymin, zmin),(xmax,ymin, zmax),(xmax,ymax,zmin),(xmax,ymax,zmax)
	'...
	'etc
	'===========xnface, face(0,:,:), nnn,nnp,npn,npp===============
	For i=0 To 3
		For j=0 To 2
			faces(0,i,j)=vertices(i,j)
		Next

	Next
	'===========xpface, face(1,:,:), pnn,pnp,ppn,ppp===============
	For i=0 To 3
		For j=0 To 2
			faces(1,i,j)=vertices(i+4,j)
		Next

	Next
	'===========ynface, face(2,:,:), nnn,nnp,pnn,pnp===============
	For i=0 To 3
		For j=0 To 2
			faces(2,i,j)=vertices(i+Int(i/2)*2,j)
		Next

	Next
	'===========ypface, face(3,:,:), npn,npp,ppn,ppp===============
	For i=0 To 3
		For j=0 To 2
			faces(3,i,j)=vertices(i+2+Int(i/2)*2,j)
		Next

	Next
	'===========znface, face(4,:,:), nnn,npn,pnn,ppn===============
	For i=0 To 3
		For j=0 To 2
			faces(4,i,j)=vertices(i*2,j)
		Next

	Next
	'===========zpface, face(5,:,:), nnp,npp,pnp,ppp===============
	For i=0 To 3
		For j=0 To 2
			faces(5,i,j)=vertices(2*i+1,j)
		Next

	Next

End Sub
Public Sub createVerticeNeighborWith(A As AntennaElement)
	If doesTouchWith(A) = False Then
		Exit Sub
	End If
    'edgeNeighbors-->nnn,nnp,npn,npp,pnn,pnp,ppn,ppp
	If StrComp(A.solidName, solidName) <> 0 Then
		'Debug.Print TypeName(A.xnNeighbor)
		If (A.nnnNeighbor Is Nothing) And isArrayIdentical(A.nnn, ppp)=True Then
			'A.xn<--> xp
			A.nnnNeighbor = Me
			 pppNeighbor = A
		ElseIf (A.nnpNeighbor Is Nothing) And isArrayIdentical(A.nnp, ppn)=True Then
			'A.xp<--> xn
			A.nnpNeighbor = Me
			 ppnNeighbor = A
		ElseIf (A.npnNeighbor Is Nothing) And isArrayIdentical(A.npn, pnp)=True Then
			'A.yp<--> yn
			A.npnNeighbor = Me
			 pnpNeighbor = A
		ElseIf (A.nppNeighbor Is Nothing) And isArrayIdentical(A.npp, pnn)=True Then
			'A.yn<--> yp
			A.nppNeighbor = Me
			 pnnNeighbor = A
		ElseIf (A.pnnNeighbor Is Nothing) And isArrayIdentical(A.pnn, npp)=True Then
			'A.zp<--> zn
			A.pnnNeighbor= Me
			 nppNeighbor = A
		ElseIf (A.pnpNeighbor Is Nothing) And isArrayIdentical(A.pnp, npn)=True Then
			'A.zn<--> zp
			A.pnpNeighbor = Me
			 npnNeighbor = A
		ElseIf (A.ppnNeighbor Is Nothing) And isArrayIdentical(A.ppn, nnp)=True Then
			'A.zn<--> zp
			A.ppnNeighbor = Me
			 nnpNeighbor = A
		ElseIf (A.pppNeighbor Is Nothing) And isArrayIdentical(A.ppp, nnn)=True Then
			'A.zn<--> zp
			A.pppNeighbor = Me
			 nnnNeighbor = A
		End If
	End If
End Sub
Public Function isVerticeNeighborWith(A As AntennaElement) As Boolean
	If doesTouchWith(A) = False Then
		Return False
		Exit Function
	End If
	If StrComp(A.solidName, solidName) <> 0 Then
		'Debug.Print TypeName(A.xnNeighbor)
		If isArrayIdentical(A.nnn, ppp)=True Or isArrayIdentical(A.nnp, ppn)=True _
		Or isArrayIdentical(A.npn, pnp)=True Or isArrayIdentical(A.npp, pnn)=True _
		Or isArrayIdentical(A.pnn, npp)=True Or isArrayIdentical(A.pnp, npn)=True _
		Or isArrayIdentical(A.ppn, nnp)=True Or isArrayIdentical(A.ppp, nnn)=True Then
			Return True
		Else
			Return False
		End If

	End If
End Function

Public Sub createFaceNeighborWith(A As AntennaElement)
	'ReDim faceNeighbors(6)
	If doesTouchWith(A) = False Then
		Exit Sub
	End If
    'faceNeighbors initialize, xn, xp, yn, yp, zn, zp
	If StrComp(A.solidName, solidName) <> 0 Then
		'Debug.Print TypeName(A.xnNeighbor)
		If (A.xnNeighbor Is Nothing) And isArrayIdentical(A.xnFace, xpFace)=True Then
			'A.xn<--> xp
			A.xnNeighbor = Me
			 xpNeighbor = A
		ElseIf (A.xpNeighbor Is Nothing) And isArrayIdentical(A.xpFace, xnFace)=True Then
			'A.xp<--> xn
			A.xpNeighbor = Me
			 xnNeighbor = A
		ElseIf (A.ypNeighbor Is Nothing) And isArrayIdentical(A.ypFace, ynFace)=True Then
			'A.yp<--> yn
			A.ypNeighbor = Me
			 ynNeighbor = A
		ElseIf (A.ynNeighbor Is Nothing) And isArrayIdentical(A.ynFace, ypFace)=True Then
			'A.yn<--> yp
			A.ynNeighbor = Me
			 ypNeighbor = A
		ElseIf (A.zpNeighbor Is Nothing) And isArrayIdentical(A.zpFace, znFace)=True Then
			'A.zp<--> zn
			A.zpNeighbor = Me
			 znNeighbor = A
		ElseIf (A.znNeighbor Is Nothing) And isArrayIdentical(A.znFace, zpFace)=True Then
			'A.zn<--> zp
			A.znNeighbor = Me
			 zpNeighbor = A
		End If
	End If
End Sub
Public Function isFaceNeighborWith(A As AntennaElement) As Boolean
	If doesTouchWith(A) = False Then
		Return False
		Exit Function
	End If
	If StrComp(A.solidName, solidName) <> 0 Then
		'Debug.Print TypeName(A.xnNeighbor)
		If isArrayIdentical(A.xnFace, xpFace)=True Or isArrayIdentical(A.xpFace, xnFace)=True _
		Or isArrayIdentical(A.ypFace, ynFace)=True Or isArrayIdentical(A.ynFace, ypFace)=True _
		Or isArrayIdentical(A.zpFace, znFace)=True Or isArrayIdentical(A.znFace, zpFace)=True Then
			Return True
		Else
			Return False
		End If

	End If
End Function

Public Sub createEdgeNeighborWith(A As AntennaElement)
	' 0-nnL(2,3) 1-nLn(2,3) 2-Lnn(2,3) 3-nLp(2,3) 4-npL(2,3) 5-Lpp(2,3)
	' 6-Lnp(2,3) 7-pnL(2,3) 8-pLP(2,3) 9-Lpn(2,3) 10-pLn(2,3) 11-ppL(2,3)
	'A is my FaceNeighbor
	If 	doesTouchWith(A) = False Then
		Exit Sub
	End If
	If StrComp(A.solidName, solidName) <> 0 Then
		If (A.nnLNeighbor Is Nothing) And isArrayIdentical(A.nnL, Me.ppL)=True Then
			'A.nnL<--> ppL
			A.nnLNeighbor = Me
			Me.ppLNeighbor = A
		ElseIf (A.nLnNeighbor Is Nothing) And isArrayIdentical(A.nLn,Me.pLP)=True Then
			'A.nLn<--> pLP
			A.nLnNeighbor  = Me
			Me.pLPNeighbor = A
		ElseIf (A.LnnNeighbor Is Nothing) And isArrayIdentical(A.Lnn,Me.Lpp)=True Then
			'A.Lnn<--> Lpp
			A.LnnNeighbor = Me
			Me.LppNeighbor = A
		ElseIf (A.nLpNeighbor Is Nothing) And isArrayIdentical(A.nLp,Me.pLn)=True Then
			'A.nLp<--> pLn
			A.nLpNeighbor = Me
			Me.pLnNeighbor = A
		ElseIf (A.npLNeighbor Is Nothing) And isArrayIdentical(A.npL,Me.pnL)=True Then
			'A.npL<--> pnL
			A.npLNeighbor = Me
			Me.pnLNeighbor = A
		ElseIf (A.LppNeighbor Is Nothing) And isArrayIdentical(A.Lpp,Me.Lnn)=True Then
			'A.Lpp<--> Lnn
			A.LppNeighbor = Me
			Me.LnnNeighbor = A
		ElseIf (A.LnpNeighbor Is Nothing) And isArrayIdentical(A.Lnp,Me.Lpn)=True Then
			'A.Lnp<--> Lpn
			A.LnpNeighbor = Me
			Me.LpnNeighbor = A
		ElseIf (A.pnLNeighbor Is Nothing) And isArrayIdentical(A.pnL,Me.npL)=True Then
			'A.pnL<--> npL
			A.pnLNeighbor = Me
			Me.npLNeighbor = A
		ElseIf (A.pLPNeighbor Is Nothing) And isArrayIdentical(A.pLP,Me.nLn)=True Then
			'A.pLP<--> nLn
			A.pLPNeighbor = Me
			Me.nLnNeighbor = A
		ElseIf (A.LpnNeighbor Is Nothing) And isArrayIdentical(A.Lpn,Me.Lnp)=True Then
			'A.Lpn<--> Lnp
			A.LpnNeighbor = Me
			Me.LnpNeighbor = A
		ElseIf (A.pLnNeighbor Is Nothing) And isArrayIdentical(A.pLn,Me.nLp)=True Then
			'A.pLn<--> nLp
			A.pLnNeighbor = Me
			Me.nLpNeighbor = A
		ElseIf (A.ppLNeighbor Is Nothing) And isArrayIdentical(A.ppL,Me.nnL)=True Then
			'A.ppL<--> nnL
			A.ppLNeighbor = Me
			Me.nnLNeighbor = A
		End If
	End If
End Sub
Public Function iSEdgeNeighborWith(A As AntennaElement) As Boolean
	If doesTouchWith(A) = False Then
		Return False
		Exit Function
	End If
	If StrComp(A.solidName, solidName) <> 0 Then
		'Debug.Print TypeName(A.xnNeighbor)
		If isArrayIdentical(A.nnL, Me.ppL)=True Or isArrayIdentical(A.nLn,Me.pLP)=True _
		Or isArrayIdentical(A.Lnn,Me.Lpp)=True Or isArrayIdentical(A.nLp,Me.pLn)=True _
		Or isArrayIdentical(A.Lpp,Me.Lnn)=True Or isArrayIdentical(A.Lnp,Me.Lpn)=True _
		Or isArrayIdentical(A.npL,Me.pnL)=True Or isArrayIdentical(A.pnL,Me.npL)=True _
		Or isArrayIdentical(A.pLP,Me.nLn)=True Or isArrayIdentical(A.Lpn,Me.Lnp)=True _
		Or isArrayIdentical(A.pLn,Me.nLp)=True Or isArrayIdentical(A.ppL,Me.nnL)=True Then
			Return True
		Else
			Return False
		End If

	End If
End Function

Private Function isArrayIdentical(A() As Double, B() As Double) As Boolean
	Dim i As Integer, j As Integer
	isArrayIdentical = True
	For i=0 To UBound(A,1)-1'points
		For j=0 To UBound(A,2)-1'coordiantes of points
			If A(i,j) <> B(i,j) Then
				isArrayIdentical = False
				Exit Function
			End If
		Next
	Next
End Function
'get all vertice neighbors
Public Function getVerticeNeighbors(ByRef NeighborNum As Integer) As Variant

	Dim i As Integer
	Dim neighbor() As AntennaElement
	'Dim vNeighborStr() As Variant
	'vNeighborStr()=Array("nnn","nnp","npn","npp","pnn","pnp","ppn","ppp")
    ReDim neighbor(8)
    'neighborStr=""
    NeighborNum = 0

    For i=0 To 7
    	If verticeNeighbors(i) IsNot Nothing Then
			Set neighbor(NeighborNum)=verticeNeighbors(i)
			'neighborStr(neighborNum)=vNeighborStr(i)
			NeighborNum += 1
    	End If
    Next
    'neighbor = "", old codes

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get number of all vertice neighbors
Public Function getNumberOfVerticeNeighbors() As Integer

	Dim i As Integer, NeighborNum As Integer
    NeighborNum = 0

    For i=0 To 7
    	If verticeNeighbors(i) IsNot Nothing Then
			NeighborNum += 1
    	End If
    Next
    'neighbor = "", old codes
	Return NeighborNum
End Function
'get all edge neighbors
Public Function getEdgeNeighbors(ByRef NeighborNum As Integer) As Variant

	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(12)
    NeighborNum = 0

    For i=0 To 11
    	If edgeNeighbors(i) IsNot Nothing Then

			Set neighbor(NeighborNum)=edgeNeighbors(i)
			NeighborNum += 1

    	End If
    Next
    'neighbor = "", old codes

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get the number of all edge neighbors
Public Function getNumberOfEdgeNeighbors() As Integer

	Dim i As Integer, NeighborNum As Integer
    NeighborNum = 0

    For i=0 To 11
    	If edgeNeighbors(i) IsNot Nothing Then
			NeighborNum += 1
    	End If
    Next
    'neighbor = "", old codes
	Return NeighborNum
End Function
'get all face neighbors
Public Function getFaceNeighbors(ByRef NeighborNum As Integer) As Variant
	Dim i As Integer
    Dim neighbor() As AntennaElement
	'Dim fNeighborStr() As Variant
	'fNeighborStr()=Array("xn","xp","yn","yp","zn","zp")
    ReDim neighbor(6)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
    	If faceNeighbors(i) IsNot Nothing Then

			Set neighbor(NeighborNum)=faceNeighbors(i)
			'neighborStr(neighborNum)=fNeihgborStr(i)
			NeighborNum += 1

    	End If
    Next

    ReDim Preserve neighbor(NeighborNum)

    getFaceNeighbors = neighbor
End Function
'get the number of all face neighbors
Public Function getNumberOfFaceNeighbors() As Integer
	Dim i As Integer, NeighborNum As Integer
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
    	If faceNeighbors(i) IsNot Nothing Then

			NeighborNum += 1

    	End If
    Next

	Return NeighborNum
    'ReDim Preserve neighbor(NeighborNum)

    'getFaceNeighbors = neighbor
End Function
'get all non metal edge neighbors
Public Function getNonMetalVerticeNeighbors(ByRef NeighborNum As Integer) As Variant

	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(8)
    NeighborNum = 0

    For i=0 To 7
	    If verticeNeighbors(i) IsNot Nothing Then
	    	If verticeNeighbors(i).isMetal()=False Then
				Set neighbor(NeighborNum)=verticeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next
    'neighbor = "", old codes

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get Number of all non metal edge neighbors
Public Function getNumberOfNonMetalVerticeNeighbors() As Integer

	Dim i As Integer, NeighborNum As Integer
	'Dim neighbor() As AntennaElement
   ' ReDim neighbor(8)
    NeighborNum = 0

    For i=0 To 7
	    If verticeNeighbors(i) IsNot Nothing Then
	    	If verticeNeighbors(i).isMetal()=False Then
				'Set neighbor(NeighborNum)=verticeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next
    'neighbor = "", old codes
	Return NeighborNum
    'ReDim Preserve neighbor(NeighborNum)
    'Return neighbor
End Function

'get  all Non-Metal edge neighbors
Public Function getNonMetalEdgeNeighbors(ByRef NeighborNum As Integer) As Variant
	' 0-nnL(2,3) 1-nLn(2,3) 2-Lnn(2,3) 3-nLp(2,3) 4-npL(2,3) 5-Lpp(2,3)
	' 6-Lnp(2,3) 7-pnL(2,3) 8-pLP(2,3) 9-Lpn(2,3) 10-pLn(2,3) 11-ppL(2,3)
	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(12)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 11
	    If edgeNeighbors(i) IsNot Nothing Then
	    	If edgeNeighbors(i).isMetal()=False Then
				Set neighbor(NeighborNum)=edgeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get Number of all Non-Metal edge neighbors
Public Function getNumberOfNonMetalEdgeNeighbors() As Integer
	' 0-nnL(2,3) 1-nLn(2,3) 2-Lnn(2,3) 3-nLp(2,3) 4-npL(2,3) 5-Lpp(2,3)
	' 6-Lnp(2,3) 7-pnL(2,3) 8-pLP(2,3) 9-Lpn(2,3) 10-pLn(2,3) 11-ppL(2,3)
	Dim i As Integer, NeighborNum As Integer
	'Dim neighbor() As AntennaElement
    'ReDim neighbor(12)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 11
	    If edgeNeighbors(i) IsNot Nothing Then
	    	If edgeNeighbors(i).isMetal()=False Then
				'Set neighbor(NeighborNum)=edgeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next
	Return NeighborNum
    'ReDim Preserve neighbor(NeighborNum)
    'Return neighbor
End Function
'get all Non-Metal face neighbors
Public Function getNonMetalFaceNeighbors(ByRef NeighborNum As Integer, ByRef neighborStr() As String) As Variant
	Dim i As Integer
	Dim neighbor() As AntennaElement
	Dim fNeighborStr() As Variant
	fNeighborStr()=Array("xn","xp","yn","yp","zn","zp")
    ReDim neighbor(6)
    ReDim neighborStr(6)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
	    If faceNeighbors(i) IsNot Nothing Then
	    	If faceNeighbors(i).isMetal()=False Then
				Set neighbor(NeighborNum)=faceNeighbors(i)
				neighborStr(NeighborNum)=fNeighborStr(i)
				NeighborNum += 1
			End If
	    End If
    Next
    ReDim Preserve neighbor(NeighborNum)
    ReDim Preserve neighborStr(NeighborNum)
    Return neighbor
End Function
'get number of all Non-Metal face neighbors
Public Function getNumberOfNonMetalFaceNeighbors() As Integer
	Dim i As Integer, NeighborNum As Integer
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
	    If faceNeighbors(i) IsNot Nothing Then
	    	If faceNeighbors(i).isMetal()=False Then
				'Set neighbor(NeighborNum)=faceNeighbors(i)
				'neighborStr(NeighborNum)=fNeighborStr(i)
				NeighborNum += 1
			End If
	    End If
    Next
    Return NeighborNum
End Function
'get all metal edge neighbors
Public Function getMetalVerticeNeighbors(ByRef NeighborNum As Integer) As Variant

	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(8)
    NeighborNum = 0

    For i=0 To 7
	    If verticeNeighbors(i) IsNot Nothing Then
	    	If verticeNeighbors(i).isMetal()=True Then
				Set neighbor(NeighborNum)=verticeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next
    'neighbor = "", old codes

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get number of all metal edge neighbors
Public Function getNumberOfMetalVerticeNeighbors() As Integer

	Dim i As Integer, NeighborNum As Integer
	'Dim neighbor() As AntennaElement
    'ReDim neighbor(8)
    NeighborNum = 0

    For i=0 To 7
	    If verticeNeighbors(i) IsNot Nothing Then
	    	If verticeNeighbors(i).isMetal()=True Then
				'Set neighbor(NeighborNum)=verticeNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next
    'neighbor = "", old codes
	Return NeighborNum
    'ReDim Preserve neighbor(NeighborNum)
    'Return neighbor
End Function
'get all Metal edge neighbors
Public Function getMetalEdgeNeighbors(ByRef NeighborNum As Integer) As Variant
	' 0-nnL(2,3) 1-nLn(2,3) 2-Lnn(2,3) 3-nLp(2,3) 4-npL(2,3) 5-Lpp(2,3)
	' 6-Lnp(2,3) 7-pnL(2,3) 8-pLP(2,3) 9-Lpn(2,3) 10-pLn(2,3) 11-ppL(2,3)
	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(12)
    NeighborNum = 0
    For i=0 To 11
	    If edgeNeighbors(i) IsNot Nothing Then
	    	If edgeNeighbors(i).isMetal()=True Then
				Set neighbor(NeighborNum)=edgeNeighbors(i)
				NeighborNum += 1
			End If
	    End If

    Next
    'neighbor = ""

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get Number of all Metal edge neighbors
Public Function getNumberOfMetalEdgeNeighbors() As Integer
	' 0-nnL(2,3) 1-nLn(2,3) 2-Lnn(2,3) 3-nLp(2,3) 4-npL(2,3) 5-Lpp(2,3)
	' 6-Lnp(2,3) 7-pnL(2,3) 8-pLP(2,3) 9-Lpn(2,3) 10-pLn(2,3) 11-ppL(2,3)
	Dim i As Integer, NeighborNum As Integer
	'Dim neighbor() As AntennaElement
    'ReDim neighbor(12)
    NeighborNum = 0
    For i=0 To 11
	    If edgeNeighbors(i) IsNot Nothing Then
	    	If edgeNeighbors(i).isMetal()=True Then
				'Set neighbor(neighborNum)=edgeNeighbors(i)
				NeighborNum += 1
			End If
	    End If

    Next
    'neighbor = ""
	Return NeighborNum
    'ReDim Preserve neighbor(neighborNum)
    'Return neighbor
End Function
'get all Metal face neighbors
Public Function getMetalFaceNeighbors(ByRef NeighborNum As Integer) As Variant
	Dim i As Integer
	Dim neighbor() As AntennaElement
    ReDim neighbor(6)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
	    If faceNeighbors(i) IsNot Nothing Then
	    	If faceNeighbors(i).isMetal()=True Then
				Set neighbor(NeighborNum)=faceNeighbors(i)
				NeighborNum += 1
			End If
	    End If
    Next

    ReDim Preserve neighbor(NeighborNum)
    Return neighbor
End Function
'get Number Of all Metal face neighbors
Public Function getNumberOfMetalFaceNeighbors() As Integer
	Dim i As Integer, NeighborNum As Integer
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
	    If faceNeighbors(i) IsNot Nothing Then
	    	If faceNeighbors(i).isMetal()=True Then
				NeighborNum += 1
			End If
	    End If
    Next
	Return NeighborNum
End Function
'get dimensions of elements in majority
Public Function getDimensions(deltaX As Double, deltaY As Double, deltaZ As Double) As Boolean
	Dim NeighborNum As Integer, i As Integer, ii As Integer
	Dim neighbors() As AntennaElement
	Dim neighborsOfNeighbors() As AntennaElement
	neighbors=getFaceNeighbors(NeighborNum)
	If NeighborNum=4 Then
		deltaX = Abs(maxPoint(0)-minPoint(0))
		deltaY = Abs(maxPoint(1)-minPoint(1))
		deltaZ = Abs(maxPoint(2)-minPoint(2))
		Return True
	Else
		For i=0 To NeighborNum-1
			neighborsOfNeighbors=neighbors(i).getFaceNeighbors(ii)
			If ii=4 Then
				deltaX = Abs(neighbors(i).maxPoint(0)-neighbors(i).minPoint(0))
				deltaY = Abs(neighbors(i).maxPoint(1)-neighbors(i).minPoint(1))
				deltaZ = Abs(neighbors(i).maxPoint(2)-neighbors(i).minPoint(2))
				Return True
			End If
		Next
	End If
	Return False
End Function
Public Function isMetal() As Boolean
	If Material.GetTypeOfMaterial(solidMaterial) = "PEC" Or _
	Material.GetTypeOfMaterial(solidMaterial) = "Lossy metal" Then
		Return True
	Else
		Return False
	End If
End Function
Public Function doesTouchWith(A As AntennaElement) As Boolean
	'verticeNum = 0
	'If isArrayIdentical(Me.xnynzn, A.) Then
	doesTouchWith = True
	If Me.minPoint(0)> A.maxPoint(0) Or Me.maxPoint(0)< A.minPoint(0) Or _
	Me.minPoint(1)> A.maxPoint(1) Or Me.maxPoint(1)< A.minPoint(1) Or _
	Me.minPoint(2)> A.maxPoint(2) Or Me.maxPoint(2)< A.minPoint(2) Then
		doesTouchWith = False
	End If
	'End If
End Function
'Pure nonMetallic: is not metallic and has no any metal neighbors
Public Function isPureNonMetal() As Boolean
	If isMetal()=False And (getNumberOfMetalFaceNeighbors()+ _
	getNumberOfMetalEdgeNeighbors()+getNumberOfMetalVerticeNeighbors())=0 Then
		Return True
	Else
		Return False
	End If

End Function
'AvailableStr is the name string of all solid that has been counted as one pure Nonmetal element
Public Function getNumberofPureNonMetalElements(availableElements As String) As Integer
	Dim NeighborNum As Integer, i As Integer, n As Integer
	'Dim cFlag As Boolean
	Dim neighborStr() As String
	'Dim availableElements As String
	Dim neighbors() As AntennaElement
	Dim hasPureNonMetalNeighbor As Boolean

	hasPureNonMetalNeighbor = False

	If isMetal()=True Then
		ReportInformationToWindow "The element is metal, please switch to a nonMetal one"
		Return 0
	End If
	n = 0
	neighbors = getNonMetalFaceNeighbors(NeighborNum, neighborStr)
	'availableElements = ""
	'cFlag = flag
	For i=0 To NeighborNum-1
		'ReportInformationToWindow "Current solid under check: " & neighbors(i).solidName
		If neighbors(i).isPureNonMetal()=True And InStr(availableElements, neighbors(i).solidName & "$")=0 Then
			'neighbors(i).flag=True
			hasPureNonMetalNeighbor = True
			availableElements = availableElements & neighbors(i).solidName & "$"
			n += 1
			'ReportInformationToWindow "#Number of pure NonMetal: " & CStr(n)
			n += neighbors(i).getNumberofPureNonMetalElements(availableElements)
			'neighbors(i).flag=False
			'ReportInformationToWindow "##Number of pure NonMetal: " & CStr(n)
		End If
	Next
	If hasPureNonMetalNeighbor = False And InStr(availableElements, solidName & "$")=0 Then
		n += 1
		availableElements = availableElements & solidName & "$"
	End If
	getNumberofPureNonMetalElements = n
End Function
'availableElements is name string of all solids that are pure nonMetal elements
Public Function getNumberofPureNonMetalNeighbors(availableElementsStr As String) As Integer
	Dim NeighborNum As Integer, i As Integer, n As Integer
	Dim neighborStr() As String
	'Dim availableElementsStr As String
	Dim neighbors() As AntennaElement
	n=0
	If isMetal()=True Then
		ReportInformationToWindow "The element is metal, please switch to a nonMetal one"
		Return 0
	End If
	neighbors = getNonMetalFaceNeighbors(NeighborNum, neighborStr)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True And InStr(availableElementsStr, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElementsStr = availableElementsStr & neighbors(i).solidName & "$"
		End If

	Next
	neighbors = getNonMetalEdgeNeighbors(NeighborNum)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True And InStr(availableElementsStr, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElementsStr = availableElementsStr & neighbors(i).solidName & "$"
		End If
	Next
	neighbors = getNonMetalVerticeNeighbors(NeighborNum)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True And InStr(availableElementsStr, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElementsStr = availableElementsStr & neighbors(i).solidName & "$"
		End If
	Next
	getNumberofPureNonMetalNeighbors=n
End Function

Public Function getFaceNeighborFromString(conStr As String, invflag As Boolean) As AntennaElement
	Select Case conStr
	Case "xn"
		If invflag=False Then
			If Not xnNeighbor Is Nothing Then Return xnNeighbor
		Else
			If Not xpNeighbor Is Nothing Then Return xpNeighbor
		End If
	Case "xp"
		If invflag=False Then
			If Not xpNeighbor Is Nothing Then Return xpNeighbor
		Else
			If Not xnNeighbor Is Nothing Then Return xnNeighbor
		End If
	Case "yn"
		If invflag=False Then
			If Not ynNeighbor Is Nothing Then Return ynNeighbor
		Else
			If Not ypNeighbor Is Nothing Then Return ypNeighbor
		End If
	Case "yp"
		If invflag=False Then
			If Not ypNeighbor Is Nothing Then Return ypNeighbor
		Else
			If Not ynNeighbor Is Nothing Then Return ynNeighbor
		End If
	Case "zn"
		If invflag=False Then
			If Not znNeighbor Is Nothing Then Return znNeighbor
		Else
			If Not zpNeighbor Is Nothing Then Return zpNeighbor
		End If
	Case "zp"
		If invflag=False Then
			If Not zpNeighbor Is Nothing Then Return zpNeighbor
		Else
			If Not znNeighbor Is Nothing Then Return znNeighbor
		End If
	Case Else
		Return Nothing
	End Select
	Return Nothing
End Function

Public Function getValidFaceNeighbors(validFaceNeighbors() As AntennaElement, validFaceNeighborsStr() As String,  _
n_validFaceNeighbors As Integer) As Boolean

	Dim nonMetalFaceNeighbors() As AntennaElement, n_nonMetalFaceNeighbors As Integer
	Dim nonMetalFaceNeighborsStr() As String, validFaceNeighborsArray() As Integer
	Dim n_metalFaceNeighborsOfNonMetalFaceNeighbors As Integer
	Dim n_metalEdgeNeighborsOfNonMetalFaceNeighbors As Integer
	Dim n_metalVerticeNeighborsOfNonMetalFaceNeighbors As Integer
	'Dim nonMetalFaceNeighborsStr() As String
	Dim metalFaceNeighborsOfNonMetalFaceNeighbors() As AntennaElement
	Dim metalEdgeNeighborsOfNonMetalFaceNeighbors() As AntennaElement
	Dim metalVerticeNeighborsOfNonMetalFaceNeighbors() As AntennaElement
	'Dim n_nonMetalFaceNeighbors As Integer
	Dim i As Integer, j As Integer, k As Integer

	nonMetalFaceNeighbors = Me.getNonMetalFaceNeighbors(n_nonMetalFaceNeighbors, nonMetalFaceNeighborsStr)
	n_validFaceNeighbors = n_nonMetalFaceNeighbors
	ReDim validFaceNeighborsArray(n_validFaceNeighbors-1)
	'Initialize the validity of nonMetalFaceNeighbors
	For i=0 To n_nonMetalFaceNeighbors-1
		validFaceNeighborsArray(i)=-1
	Next
	'Re-Check the validity of nonMetalFaceNeighgors
	k = 0
	For i = 0 To n_nonMetalFaceNeighbors-1
		'If there are more than one metal face neighbors, the non metal face neighbors are invalid
		If nonMetalFaceNeighbors(i).getNumberOfMetalFaceNeighbors()>1 And n_validFaceNeighbors>=1 Then
			n_validFaceNeighbors -= 1
			GoTo Goon
		End If

		'Check whether metal edge neighbors of non-metal face neighbors are among
		'face neighbors of current element, if not, set invalid
		metalEdgeNeighborsOfNonMetalFaceNeighbors = _
		nonMetalFaceNeighbors(i).getMetalEdgeNeighbors(n_metalEdgeNeighborsOfNonMetalFaceNeighbors)
		If n_metalEdgeNeighborsOfNonMetalFaceNeighbors >= 1 And n_validFaceNeighbors>=1 Then
			For j=0 To n_metalEdgeNeighborsOfNonMetalFaceNeighbors-1
				'If metalEdgeNeighborsOfNonMetalFaceNeighbors(j).iSFaceNeighborWith(Me)=False Then
				If Me.isFaceNeighborWith(metalEdgeNeighborsOfNonMetalFaceNeighbors(j))=False Then
					n_validFaceNeighbors -= 1
					'validityOfFaceNeighbors(i)=False
					'ReDim Preserve validFaceNeighborsArray(n_validFaceNeighbors-1)
					GoTo Goon
				End If
			Next
		End If
		'Check whether a non_metal face neighbor gets a metal vertice neighbor
		'if yes, set invalid
		If nonMetalFaceNeighbors(i).getNumberOfMetalVerticeNeighbors() >= 1 And n_validFaceNeighbors>=1 Then
			n_validFaceNeighbors -= 1
			'ReDim Preserve validFaceNeighborsArray(n_validFaceNeighbors-1)
			GoTo Goon
		End If
		validFaceNeighborsArray(k) = i
		k += 1
	Goon:
	Next
	If n_validFaceNeighbors >=1 Then
		'ReDim Preserve validFaceNeighborsArray(n_validFaceNeighbors)
		ReDim validFaceNeighbors(n_validFaceNeighbors)
		ReDim validFaceNeighborsStr(n_validFaceNeighbors)
		For i=0 To n_validFaceNeighbors-1
			Set validFaceNeighbors(i) = nonMetalFaceNeighbors(validFaceNeighborsArray(i))
			validFaceNeighborsStr(i) = nonMetalFaceNeighborsStr(validFaceNeighborsArray(i))
		Next
	End If

	If n_validFaceNeighbors = 0 Then
		Return False
	Else
		Return True
	End If
End Function
