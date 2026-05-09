'#Language "WWB-COM"
'Search for the resonant frequencies and Q value, and then write them in a target file

Option Explicit
'#include "vba_globals_all.lib"
Dim spectrum As Object

Sub Main ()
	Dim Spath As String
	Dim FirstItem As String
	Spath = "1D Results\Z Matrix"
	FirstItem = ResultTree.GetFirstChildName(Spath)

	'Debug.Print FirstItem
	'Debug.Clear
    If FirstItem = "" Then
	   	 MsgBox("No Z-parameter results found!",vbCritical,"Error")
	   	 Exit All
    End If

    Dim currentItem As String
    Dim curveName As String
    Dim i As Integer

	Dim prjPath As String
	Dim dataFile As String
	'Dim fNumber As Integer

	prjPath = GetProjectPath("Project")
   	dataFile = prjPath + "\freq_Q.txt"
   	'fNumber = FreeFile
	Open dataFile For Output As #1

	currentItem = FirstItem

	While currentItem <> ""
		curveName = basename(currentItem)
		Dim temStr As String
		Dim flag As Boolean
		flag = False
		For i = 1 To 10 STEP 1
			temStr = CStr(i)&","&CStr(i)
			If InStr(curveName,temStr) <> 0 Then
				Print #1, "%"
				Print #1, "P =" +CStr(i)
				flag = True
				Exit For
			End If
		Next
		If  flag = True Then
			Dim  FileName As String
			Dim nPoints As Long, n As Integer, Num As Integer, Ysum As Double, Avg As Double, dBAvg As Double, X As Double, Y As Double
			Dim calcQ As Double
			'Dim spectrum As Object
			Dim nRes As Long
			Dim fileType As String
			Dim spectrum_copy As Object

			'Dim calcQ As Double
			'EffiType = Mid(currentItem,Len(Spath)+2,InStr(currentItem,"[")-Len(Spath)-2)
			FileName = ResultTree.GetFileFromTreeItem(currentItem)
			'Debug.Print FileName

			fileType =  GetFileType(FileName)
			If fileType = "complex" Then
				Set spectrum = Result1DComplex(FileName)
				Set spectrum = spectrum.Imaginary
				Set spectrum_copy = Result1D("")
			End If
			Dim ii As Integer
			'Get resonance frequency
			With spectrum
				For ii = 10 To .GetN-1

					X = spectrum.GetX(ii)
					Y = spectrum.GetY(ii)
					spectrum_copy.Appendxy(X,Abs(Y))
					If Abs(Y)<1e-1 Then

						Print #1,"F"+CStr(i)+"=" + CStr(Round(X,2))' + vbNewLine + "Q"+CStr(i)+"=" + CStr(Round(calcQ,2))
						Exit For
					End If
				Next
				'No points meet the resonance conditions in last part
				If ii=.GetN Then
					ii=spectrum_copy.GetGlobalMinimum
					X = spectrum_copy.GetX(ii)
					Print #1,"F"+CStr(i)+"=" + CStr(Round(X,2))' + vbNewLine + "Q"+CStr(i)+"=" + CStr(Round(calcQ,2))
				End If
			End With
		End If
		currentItem = ResultTree.GetNextItemName(currentItem)
    Wend
	Close #1
End Sub

Function CalculateQ(nRes As Long) As Single

        Dim vres As Double, v1 As Double, v2 As Double, v3dB As Double, vpre As Double
        Dim fres As Double, f1 As Double, f2 As Double
        Dim nfrq As Long, ii As Long
        Dim noLeft As Boolean, noRight As Boolean

        nfrq = spectrum.GetN
        vres = spectrum.GetY(nRes)
        fres = spectrum.GetX(nRes)

    	v3dB = Sqr(0.5*(vres*vres+1))
        v3dB = vres + 3

        noLeft = True
        vpre = vres
        For ii = nRes-1 To 0 STEP -1
                v1 = spectrum.GetY(ii)
                If v1 < vpre Then Exit For
				If v1 > v3dB Then
						noLeft = False
                        f1 = CalculateX(v3dB, spectrum.GetX(ii+1), spectrum.GetX(ii), vpre, v1)
                        Exit For
				End If
                vpre = v1
        Next ii

        noRight = True
        vpre = vres
        For ii = nRes+1 To nfrq-1
                v2 = spectrum.GetY(ii)
                If v2 < vpre Then Exit For
				If v2 > v3dB Then
						noRight = False
                        f2 = CalculateX(v3dB, spectrum.GetX(ii-1), spectrum.GetX(ii), vpre, v2)
                        Exit For
				End If
                vpre = v2
        Next ii

        If noLeft Then
				'MsgBox "Lower Frequency point not found; used symmetrical point "
                CalculateQ = fres / (2*Abs(fres - f2))
        ElseIf noRight Then
				'MsgBox "Upper Frequency point not found; used symmetrical point "
				CalculateQ = fres / (2*Abs(fres - f1))
		Else
				CalculateQ = fres / Abs(f2 - f1)
        End If

End Function

' linear regression and calculate x from y (yco7 May-27-2020)
Function CalculateX(Y As Double, x1 As Double, x2 As Double, y1 As Double, y2 As Double) As Double

		Dim a As Double, b As Double

		b = (y2 - y1)/(x2 - x1)		' the slope b will not be zero because this operates at a peak or velley (yco7 May-27-2020)
		a = y1 - b * x1

		CalculateX = (Y - a)/b

End Function
