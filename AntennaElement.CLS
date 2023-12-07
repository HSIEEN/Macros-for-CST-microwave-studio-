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
Property Get nnn2nnp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(0,i,j)
		Next
	Next
	nnn2nnp = edge
End Property
'nnn2npn
Property Get nnn2npn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(1,i,j)
		Next
	Next
	nnn2npn = edge
End Property
'nnn2pnn
Property Get nnn2pnn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(2,i,j)
		Next
	Next
	nnn2pnn = edge
End Property
'npp2nnp
Property Get npp2nnp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(3,i,j)
		Next
	Next
	npp2nnp = edge
End Property
'npp2npn
Property Get npp2npn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(4,i,j)
		Next
	Next
	npp2npn = edge
End Property
'npp2ppp
Property Get npp2ppp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(5,i,j)
		Next
	Next
	npp2ppp = edge
End Property
'pnp2nnp
Property Get pnp2nnp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(6,i,j)
		Next
	Next
	pnp2nnp = edge
End Property
'pnp2pnn
Property Get pnp2pnn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(7,i,j)
		Next
	Next
	pnp2pnn = edge
End Property
'pnp2ppp
Property Get pnp2ppp() As Double()
	'pnp2pppEdge = pnp2ppp
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(8,i,j)
		Next
	Next
	pnp2ppp = edge
End Property
'ppn2npn
Property Get ppn2npn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(9,i,j)
		Next
	Next
	ppn2npn = edge
End Property
'ppn2pnn
Property Get ppn2pnn() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(10,i,j)
		Next
	Next
	ppn2pnn = edge
End Property
'ppn2ppp
Property Get ppn2ppp() As Double()
	Dim i As Integer, j As Integer
	Dim edge() As Double
	ReDim edge(2,3)
	For i=0 To 1
		For j=0 To 2
			edge(i,j)=edges(11,i,j)
		Next
	Next
	ppn2ppp = edge
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
	' nnn2nnp(2,3)
	' nnn2npn(2,3)
	' nnn2pnn(2,3)
	' npp2nnp(2,3)
	' npp2npn(2,3)
	' npp2ppp(2,3)
	' pnp2nnp(2,3)
	' pnp2pnn(2,3)
	' pnp2ppp(2,3)
	' ppn2npn(2,3)
	' ppn2pnn(2,3)
	' ppn2ppp(2,3)
'=================================nnn2nnp
Private edgeNeighbors(12) As AntennaElement
'ReDim edgeNeighbors(12)
'Private edgeNeighbornnn2nnp As AntennaElement
Property Get nnn2nnpNeighbor() As AntennaElement
	Set nnn2nnpNeighbor = edgeNeighbors(0)
End Property
Property Set nnn2nnpNeighbor(value As AntennaElement)
	Set edgeNeighbors(0) = value
End Property
'=================================nnn2npn
Property Get nnn2npnNeighbor() As AntennaElement
	Set nnn2npnNeighbor = edgeNeighbors(1)
End Property
Property Set nnn2npnNeighbor(value As AntennaElement)
	Set edgeNeighbors(1) = value
End Property
'=================================nnn2pnn
Property Get nnn2pnnNeighbor() As AntennaElement
	Set nnn2pnnNeighbor = edgeNeighbors(2)
End Property
Property Set nnn2pnnNeighbor(value As AntennaElement)
	Set edgeNeighbors(2) = value
End Property
'=================================npp2nnp
Property Get npp2nnpNeighbor() As AntennaElement
	Set npp2nnpNeighbor = edgeNeighbors(3)
End Property
Property Set npp2nnpNeighbor(value As AntennaElement)
	Set edgeNeighbors(3) = value
End Property
'===================================npp2npn
Property Get npp2npnNeighbor() As AntennaElement
	Set npp2npnNeighbor = edgeNeighbors(4)
End Property
Property Set npp2npnNeighbor(value As AntennaElement)
	Set edgeNeighbors(4) = value
End Property
'=====================================npp2ppp
Property Get npp2pppNeighbor() As AntennaElement
	Set npp2pppNeighbor = edgeNeighbors(5)
End Property
Property Set npp2pppNeighbor(value As AntennaElement)
	Set edgeNeighbors(5) = value
End Property
'========================================pnp2nnp
Property Get pnp2nnpNeighbor() As AntennaElement
	Set pnp2nnpNeighbor = edgeNeighbors(6)
End Property
Property Set pnp2nnpNeighbor(value As AntennaElement)
	Set edgeNeighbors(6) = value
End Property
'========================================pnp2pnn
Property Get pnp2pnnNeighbor() As AntennaElement
	Set pnp2pnnNeighbor = edgeNeighbors(7)
End Property
Property Set pnp2pnnNeighbor(value As AntennaElement)
	Set edgeNeighbors(7) = value
End Property
'========================================pnp2ppp
Property Get pnp2pppNeighbor() As AntennaElement
	Set pnp2pppNeighbor = edgeNeighbors(8)
End Property
Property Set pnp2pppNeighbor(value As AntennaElement)
	Set edgeNeighbors(8) = value
End Property
'========================================ppn2npn
Property Get ppn2npnNeighbor() As AntennaElement
	Set ppn2npnNeighbor = edgeNeighbors(9)
End Property
Property Set ppn2npnNeighbor(value As AntennaElement)
	Set edgeNeighbors(9) = value
End Property
'========================================ppn2pnn
Property Get ppn2pnnNeighbor() As AntennaElement
	Set ppn2pnnNeighbor = edgeNeighbors(10)
End Property
Property Set ppn2pnnNeighbor(value As AntennaElement)
	Set edgeNeighbors(10) = value
End Property
'========================================ppn2ppp
Property Get ppn2pppNeighbor() As AntennaElement
	Set ppn2pppNeighbor = edgeNeighbors(11)
End Property
Property Set ppn2pppNeighbor(value As AntennaElement)
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
	'nnn2nnp
	For i=0 To 2
		edges(0,0,i)=nnn(i)
		edges(0,1,i)=nnp(i)
	Next
	'nnn2npn
	For i=0 To 2
		edges(1,0,i)=nnn(i)
		edges(1,1,i)=npn(i)
	Next
	'nnn2pnn
	For i=0 To 2
		edges(2,0,i)=nnn(i)
		edges(2,1,i)=pnn(i)
	Next
	'npp2nnp
	For i=0 To 2
		edges(3,0,i)=npp(i)
		edges(3,1,i)=nnp(i)
	Next
	'npp2npn
	For i=0 To 2
		edges(4,0,i)=npp(i)
		edges(4,1,i)=npn(i)
	Next
	'npp2ppp
	For i=0 To 2
		edges(5,0,i)=npp(i)
		edges(5,1,i)=ppp(i)
	Next
	'pnp2nnp
	For i=0 To 2
		edges(6,0,i)=pnp(i)
		edges(6,1,i)=nnp(i)
	Next
	'pnp2pnn
	For i=0 To 2
		edges(7,0,i)=pnp(i)
		edges(7,1,i)=pnn(i)
	Next
	'pnp2ppp
	For i=0 To 2
		edges(8,0,i)=pnp(i)
		edges(8,1,i)=ppp(i)
	Next
	'ppn2npn
	For i=0 To 2
		edges(9,0,i)=ppn(i)
		edges(9,1,i)=npn(i)
	Next
	'ppn2pnn
	For i=0 To 2
		edges(10,0,i)=ppn(i)
		edges(10,1,i)=pnn(i)
	Next
	'ppn2ppp
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
	'ReDim edgeNeighbors(12)
			' 0-nnn2nnp(2,3)
			' 1-nnn2npn(2,3)
			' 2-nnn2pnn(2,3)
			' 3-npp2nnp(2,3)
			' 4-npp2npn(2,3)
			' 5-npp2ppp(2,3)
			' 6-pnp2nnp(2,3)
			' 7-pnp2pnn(2,3)
			' 8-pnp2ppp(2,3)
			' 9-ppn2npn(2,3)
			' 10-ppn2pnn(2,3)
			' 11-ppn2ppp(2,3)
	'A is my FaceNeighbor
	If 	doesTouchWith(A) = False Then
		Exit Sub
	End If
	If StrComp(A.solidName, solidName) <> 0 Then
		If (A.nnn2nnpNeighbor Is Nothing) And isArrayIdentical(A.nnn2nnp, Me.ppn2ppp)=True Then
			'A.nnn2nnp<--> ppn2ppp
			A.nnn2npnNeighbor = Me
			Me.ppn2pppNeighbor = A
		ElseIf (A.nnn2npnNeighbor Is Nothing) And isArrayIdentical(A.nnn2npn,Me.pnp2ppp)=True Then
			'A.nnn2npn<--> pnp2ppp
			A.nnn2npnNeighbor  = Me
			Me.pnp2pppNeighbor = A
		ElseIf (A.nnn2pnnNeighbor Is Nothing) And isArrayIdentical(A.nnn2pnn,Me.npp2ppp)=True Then
			'A.nnn2pnn<--> npp2ppp
			A.nnn2pnnNeighbor = Me
			Me.npp2pppNeighbor = A
		ElseIf (A.npp2nnpNeighbor Is Nothing) And isArrayIdentical(A.npp2nnp,Me.ppn2pnn)=True Then
			'A.npp2nnp<--> ppn2pnn
			A.npp2nnpNeighbor = Me
			Me.ppn2pnnNeighbor = A
		ElseIf (A.npp2npnNeighbor Is Nothing) And isArrayIdentical(A.npp2npn,Me.pnp2pnn)=True Then
			'A.npp2npn<--> pnp2pnn
			A.npp2npnNeighbor = Me
			Me.pnp2pnnNeighbor = A
		ElseIf (A.npp2pppNeighbor Is Nothing) And isArrayIdentical(A.npp2ppp,Me.nnn2pnn)=True Then
			'A.npp2ppp<--> nnn2pnn
			A.npp2pppNeighbor = Me
			Me.nnn2pnnNeighbor = A
		ElseIf (A.pnp2nnpNeighbor Is Nothing) And isArrayIdentical(A.pnp2nnp,Me.ppn2npn)=True Then
			'A.pnp2nnp<--> ppn2npn
			A.pnp2nnpNeighbor = Me
			Me.ppn2npnNeighbor = A
		ElseIf (A.pnp2pnnNeighbor Is Nothing) And isArrayIdentical(A.pnp2pnn,Me.npp2npn)=True Then
			'A.pnp2pnn<--> npp2npn
			A.pnp2pnnNeighbor = Me
			Me.npp2npnNeighbor = A
		ElseIf (A.pnp2pppNeighbor Is Nothing) And isArrayIdentical(A.pnp2ppp,Me.nnn2npn)=True Then
			'A.pnp2ppp<--> nnn2npn
			A.pnp2pppNeighbor = Me
			Me.nnn2npnNeighbor = A
		ElseIf (A.ppn2npnNeighbor Is Nothing) And isArrayIdentical(A.ppn2npn,Me.pnp2nnp)=True Then
			'A.ppn2npn<--> pnp2nnp
			A.ppn2npnNeighbor = Me
			Me.pnp2nnpNeighbor = A
		ElseIf (A.ppn2pnnNeighbor Is Nothing) And isArrayIdentical(A.ppn2pnn,Me.npp2nnp)=True Then
			'A.ppn2pnn<--> npp2nnp
			A.ppn2pnnNeighbor = Me
			Me.npp2nnpNeighbor = A
		ElseIf (A.ppn2pppNeighbor Is Nothing) And isArrayIdentical(A.ppn2ppp,Me.nnn2nnp)=True Then
			'A.ppn2ppp<--> nnn2nnp
			A.ppn2pppNeighbor = Me
			Me.nnn2nnpNeighbor = A
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
		If isArrayIdentical(A.nnn2nnp, Me.ppn2ppp)=True Or isArrayIdentical(A.nnn2npn,Me.pnp2ppp)=True _
		Or isArrayIdentical(A.nnn2pnn,Me.npp2ppp)=True Or isArrayIdentical(A.npp2nnp,Me.ppn2pnn)=True _
		Or isArrayIdentical(A.npp2ppp,Me.nnn2pnn)=True Or isArrayIdentical(A.pnp2nnp,Me.ppn2npn)=True _
		Or isArrayIdentical(A.npp2npn,Me.pnp2pnn)=True Or isArrayIdentical(A.pnp2pnn,Me.npp2npn)=True _
		Or isArrayIdentical(A.pnp2ppp,Me.nnn2npn)=True Or isArrayIdentical(A.ppn2npn,Me.pnp2nnp)=True _
		Or isArrayIdentical(A.ppn2pnn,Me.npp2nnp)=True Or isArrayIdentical(A.ppn2ppp,Me.nnn2nnp)=True Then
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
    'Dim neighbor() As AntennaElement
	'Dim fNeighborStr() As Variant
	'fNeighborStr()=Array("xn","xp","yn","yp","zn","zp")
    'ReDim neighbor(6)
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
		' 0-nnn2nnp(2,3)
		' 1-nnn2npn(2,3)
		' 2-nnn2pnn(2,3)
		' 3-npp2nnp(2,3)
		' 4-npp2npn(2,3)
		' 5-npp2ppp(2,3)
		' 6-pnp2nnp(2,3)
		' 7-pnp2pnn(2,3)
		' 8-pnp2ppp(2,3)
		' 9-ppn2npn(2,3)
		' 10-ppn2pnn(2,3)
		' 11-ppn2ppp(2,3)
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
		' 0-nnn2nnp(2,3)
		' 1-nnn2npn(2,3)
		' 2-nnn2pnn(2,3)
		' 3-npp2nnp(2,3)
		' 4-npp2npn(2,3)
		' 5-npp2ppp(2,3)
		' 6-pnp2nnp(2,3)
		' 7-pnp2pnn(2,3)
		' 8-pnp2ppp(2,3)
		' 9-ppn2npn(2,3)
		' 10-ppn2pnn(2,3)
		' 11-ppn2ppp(2,3)
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
	'Dim neighbor() As AntennaElement
	'Dim fNeighborStr() As Variant
	'fNeighborStr()=Array("xn","xp","yn","yp","zn","zp")
    'ReDim neighbor(6)
    'ReDim neighborStr(6)
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
    'ReDim Preserve neighbor(NeighborNum)
    'ReDim Preserve neighborStr(NeighborNum)
    'Return neighbor
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
		' 0-nnn2nnp(2,3)
		' 1-nnn2npn(2,3)
		' 2-nnn2pnn(2,3)
		' 3-npp2nnp(2,3)
		' 4-npp2npn(2,3)
		' 5-npp2ppp(2,3)
		' 6-pnp2nnp(2,3)
		' 7-pnp2pnn(2,3)
		' 8-pnp2ppp(2,3)
		' 9-ppn2npn(2,3)
		' 10-ppn2pnn(2,3)
		' 11-ppn2ppp(2,3)
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
		' 0-nnn2nnp(2,3)
		' 1-nnn2npn(2,3)
		' 2-nnn2pnn(2,3)
		' 3-npp2nnp(2,3)
		' 4-npp2npn(2,3)
		' 5-npp2ppp(2,3)
		' 6-pnp2nnp(2,3)
		' 7-pnp2pnn(2,3)
		' 8-pnp2ppp(2,3)
		' 9-ppn2npn(2,3)
		' 10-ppn2pnn(2,3)
		' 11-ppn2ppp(2,3)
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
	'Dim fNeighborStr() As Variant
	'fNeighborStr()=Array("xn","xp","yn","yp","zn","zp")
    ReDim neighbor(6)
    NeighborNum = 0
    'neighbor = ""
    For i=0 To 5
	    If faceNeighbors(i) IsNot Nothing Then
	    	If faceNeighbors(i).isMetal()=True Then
				Set neighbor(NeighborNum)=faceNeighbors(i)
				'neighborStr(neighborNum)=fNeighborStr(i)
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
'Count how many elements left that can be used as antenna elements
'AvailableStr is the name string of all solid that has been counted as one pure Nonmetal element
Public Function getNumberofPureNonMetalElements(availableElements As String) As Integer
	Dim NeighborNum As Integer, i As Integer, n As Integer
	'Dim cFlag As Boolean
	Dim neighborStr() As String
	'Dim availableElements As String
	Dim neighbors() As AntennaElement

	If isMetal()=True Then
		ReportInformationToWindow "The element is metal, please switch to a nonMetal one"
		Return 0
	End If
	neighbors = getNonMetalFaceNeighbors(NeighborNum, neighborStr)
	n = 0
	'availableElements = ""
	'cFlag = flag
	For i=0 To NeighborNum-1
		'ReportInformationToWindow "Current solid under check: " & neighbors(i).solidName
		If neighbors(i).isPureNonMetal()=True And InStr(availableElements, neighbors(i).solidName & "$")=0 Then
			'neighbors(i).flag=True
			availableElements = availableElements & neighbors(i).solidName & "$"
			n += 1
			'ReportInformationToWindow "#Number of pure NonMetal: " & CStr(n)
			n += neighbors(i).getNumberofPureNonMetalElements(availableElements)
			'neighbors(i).flag=False
			'ReportInformationToWindow "##Number of pure NonMetal: " & CStr(n)
		End If
	Next
	getNumberofPureNonMetalElements = n
End Function
'availableElements is names of all solids that are pure nonMetal elements
Public Function getNumberofPureNonMetalNeighbors(availableElements As String) As Integer
	Dim NeighborNum As Integer, i As Integer, n As Integer
	Dim neighborStr() As String
	'Dim availableElements As String
	Dim neighbors() As AntennaElement
	n=0
	If isMetal()=True Then
		ReportInformationToWindow "The element is metal, please switch to a nonMetal one"
		Return 0
	End If
	neighbors = getNonMetalFaceNeighbors(NeighborNum, neighborStr)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True Then'And InStr(availableElements, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElements = availableElements & neighbors(i).solidName & "$"
		End If

	Next
	neighbors = getNonMetalEdgeNeighbors(NeighborNum)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True Then' And InStr(availableElements, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElements = availableElements & neighbors(i).solidName & "$"
		End If
	Next
	neighbors = getNonMetalVerticeNeighbors(NeighborNum)
	For i=0 To NeighborNum-1
		If neighbors(i).isPureNonMetal()=True Then' And InStr(availableElements, neighbors(i).solidName & "$")=0 Then
			n+=1
			availableElements = availableElements & neighbors(i).solidName & "$"
		End If
	Next
	getNumberofPureNonMetalNeighbors=n
End Function

'Get valid face neighhors that could be used as possible antnenaElements
Function getValidNonMetalFaceNeighbors(nonMetalFaceNeighbors() As AntennaElement, n_nonMetalFaceNeighbors As Integer, _
nonMetalFaceNeighborsStr() As String, validFaceNeighborsArray() As Integer, n_validFaceNeighbors As Integer) As Boolean

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
			GoTo continueLoop
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
					GoTo continueLoop
				End If
			Next
		End If
		'Check whether a non_metal face neighbor gets a metal vertice neighbor
		'if yes, set invalid
		If nonMetalFaceNeighbors(i).getNumberOfMetalVerticeNeighbors() >= 1 And n_validFaceNeighbors>=1 Then
			n_validFaceNeighbors -= 1
			'ReDim Preserve validFaceNeighborsArray(n_validFaceNeighbors-1)
			GoTo continueLoop
		End If
		validFaceNeighborsArray(k) = i
		k += 1
	continueLoop:
	Next
	If n_validFaceNeighbors >=1 Then
		ReDim Preserve validFaceNeighborsArray(n_validFaceNeighbors-1)
	End If
	If n_validFaceNeighbors = 0 Then
		Return False
	Else
		Return True
	End If
End Function
