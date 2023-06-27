' MeshViewAtPickedPoint

Sub Main ()

	Dim x As Double, y As Double, z As Double
	Dim N As Integer
	Dim WCSStr As String

	WCSStr = WCS.IsWCSActive

	N = Pick.GetNumberOfPickedPoints
	If N > 0 Then
		WCS.ActivateWCS "global"
	    Pick.GetPickpointCoordinates(0,x,y,z)
	    Plot.DefinePlane(0,1,0,x,y,z)
	    Plot.DefinePlane(0,0,1,x,y,z)
	    Plot.DefinePlane(1,0,0,x,y,z)
	    'Plot.ShowCutplane(True)
	    Mesh.ViewMeshMode(True)
	    WCS.ActivateWCS WCSStr
	Else
		MsgBox("Please pick one point before running the macro",vbCritical,"Warning")
		Exit Sub
    End If

End Sub
