'#Language "WWB-COM"
Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

Public nx As Long
Public ny As Long
Public nz As Long
'Public fieldIndex As Long
'-----------------------------------------------------------------------------------------------------------------------------
'Const CLight= 299792448
Const epsilon = 8.854187817e-12
Const mu = 1.2566370614e-6


Sub Main
    Dim startTime As Double
    Dim freq As String

	startTime = Timer

    Dim ix As Integer, iy As Integer, iz As Integer
    Dim i As Integer

    Dim xp(8) As Double
    Dim yp(8) As Double
    Dim zp(8) As Double


    nx = Mesh.GetNx
    ny = Mesh.GetNy
    nz = Mesh.GetNz


    Dim nxnyMul As Long
	nxnyMul = nx*ny

    Dim nxMinus1 As Long, nyMinus1 As Long, nzMinus1 As Long
    nxMinus1 = nx - 1
    nyMinus1 = ny - 1
    nzMinus1 = nz - 1

	Dim fieldType As String, field As Object, fileName As String

	If InStr(GetselectedTreeItem, "2D/3D Results\E-Field\e-field") Or InStr(GetselectedTreeItem, "2D/3D Results\H-Field\h-field") Then

		fileName  = ResultTree.GetFileFromTreeItem(GetselectedTreeItem)
	Else
		MsgBox "Please pick a 3D e-field or a 3D h-field result!", vbCritical
		Exit All
	End If

	freq = Mid(GetselectedTreeItem, InStr(GetselectedTreeItem, "=")+1, InStr(GetselectedTreeItem, ")")-InStr(GetselectedTreeItem, "=")-1) & Units.GetUnit("Frequency")
	Set field = result3D(fileName)

	fieldType = field.getType

	Dim energyFactor(2) As Double, energyIndex As Integer, ehType(2) As String

	energyFactor(0) = epsilon
	energyFactor(1) = mu

	ehType(0) = "E-field"
	ehType(1) = "H-field"

	If InStr(fieldType, "E-Field") Then
		energyIndex = 0
	Else
		energyIndex = 1
	End If

    Dim totEnergy As Double, cellEnergy As Double

    Dim vxre(8) As Double, vxim(8) As Double
    Dim vyre(8) As Double, vyim(8) As Double
    Dim vzre(8) As Double, vzim(8) As Double
    Dim vertexEnergy(8) As Double
    Dim cellEnergyDensity As Double

    '=====<\Test codes====
    Dim fx() As Double, fy() As Double, fz() As Double
    Dim fxi() As Double, fyi() As Double, fzi() As Double
    fx = field.GetArray("xre")
    fy = field.GetArray("yre")
    fz = field.GetArray("zre")

    fxi = field.GetArray("xim")
    fyi = field.GetArray("yim")
    fzi = field.GetArray("zim")
    '====Test codes/>====

	Dim geometryUnitToSI As Double
	geometryUnitToSI = Units.GetGeometryUnitToSI()

	Dim xCoords() As Double, yCoords() As Double, zCoords() As Double
	ReDim xCoords(nxMinus1), yCoords(nyMinus1), zCoords(nzMinus1)

	For ix = 0 To nxMinus1
	    xCoords(ix) = Mesh.GetX(ix) * geometryUnitToSI
	Next

	For iy = 0 To nyMinus1
	    yCoords(iy) = Mesh.GetY(iy) * geometryUnitToSI
	Next

	For iz = 0 To nzMinus1
	    zCoords(iz) = Mesh.GetZ(iz) * geometryUnitToSI
	Next

	Dim baseIndex As Long
    Dim totalCells As Long, processedCells As Long
    Dim updateInterval As Long
    totalCells = nxMinus1 * nyMinus1 * nzMinus1
    processedCells = 0
    updateInterval = IIf(totalCells > 10000, totalCells \ 100, 100)

    'Initialize mesh nodes energy
    Dim energyNodes() As Double
    ReDim energyNodes(nx*ny*nz-1)
    Dim  iznxnyMul As Long
    Dim iynxMul As Long
	ReportInformationToWindow "Mesh nodes energy initializing..."
	For iz = 0 To nzMinus1
		'Dim iznxnyMul As Long
		iznxnyMul = iz * nxnyMul
		For iy = 0 To nyMinus1

			iynxMul = iy * nx
			For ix = 0 To nxMinus1
				baseIndex = ix + iynxMul + iznxnyMul
				energyNodes(baseIndex) = fx(baseIndex) ^ 2 + fy(baseIndex) ^ 2 + fz(baseIndex) ^ 2 + fxi(baseIndex) ^ 2 + fyi(baseIndex) ^ 2 + fzi(baseIndex) ^ 2
			Next
		Next
	Next

    Dim progress As Double
	Dim x0 As Double, x1 As Double
	Dim y0 As Double, y1 As Double
	Dim z0 As Double, z1 As Double

	For iz = 0 To nz-2
        z0 = zCoords(iz)
        z1 = zCoords(iz + 1)
		iznxnyMul = nxnyMul*iz

		For iy = 0 To ny-2
            y0 = yCoords(iy)
            y1 = yCoords(iy + 1)
            iynxMul = iy*nx

			For ix = 0 To nx-2
		        x0 = xCoords(ix)
		        x1 = xCoords(ix + 1)

                xp(0) = x0: yp(0) = y0: zp(0) = z0
                xp(1) = x1: yp(1) = y0: zp(1) = z0
                xp(2) = x0: yp(2) = y1: zp(2) = z0
                xp(3) = x1: yp(3) = y1: zp(3) = z0
                xp(4) = x0: yp(4) = y0: zp(4) = z1
                xp(5) = x1: yp(5) = y0: zp(5) = z1
                xp(6) = x0: yp(6) = y1: zp(6) = z1
                xp(7) = x1: yp(7) = y1: zp(7) = z1

           		cellEnergyDensity = (energyNodes(ix + iynxMul + iznxnyMul) + energyNodes(ix+1 + iynxMul + iznxnyMul) + _
                energyNodes(ix + iynxMul+nx + iznxnyMul) + energyNodes(ix + iynxMul + iznxnyMul + nxnyMul) + _
                energyNodes(ix + iynxMul+nx + iznxnyMul + nxnyMul) + energyNodes(ix+1 + iynxMul+nx + iznxnyMul) + _
           		energyNodes(ix+1 + iynxMul + iznxnyMul + nxnyMul) + energyNodes(ix+1 + iynxMul+nx + iznxnyMul + nxnyMul)) / 8

				cellEnergy = Abs(xp(1)-xp(0))*Abs(yp(2)-yp(0))*Abs(zp(4)-zp(0)) _
				*energyFactor(energyIndex)*cellEnergyDensity

				totEnergy = totEnergy + cellEnergy

                processedCells = processedCells + 1

                If processedCells Mod updateInterval = 0 Then
                    progress = processedCells / totalCells * 100
                    ReportInformationToWindow "Progress: " & Format(progress, "0.00") & "%£¬total " & ehType(energyIndex)& " energy: " & Cstr(totEnergy/4)
                End If

			Next
		Next
	Next
	Dim elapsedTime As Double
	elapsedTime = Timer - startTime
	ReportInformationToWindow "Progress: 100%, total " & ehType(energyIndex)& " energy: " & Cstr(totEnergy/4) &", time elapsed: "&CStr(elapsedTime)& "s."
	MsgBox "Total stored " & ehType(energyIndex)& " energy: " & Format(totEnergy/4, "0.00E-00"),vbInformation,"Statistics @" & freq

End Sub
