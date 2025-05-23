' *Construct / Discrete Ports / Multiple discrete Ports
' !!! Do not change the line above !!!

Option Explicit

'#include "vba_globals_all.lib"

Sub Main

 Dim port_nr_offset As Integer
 Dim n_of_ppoints As Integer
 Dim i As Integer
 Dim Array_x () As Double
 Dim Array_y () As Double
 Dim Array_z () As Double

 
 n_of_ppoints = Pick.GetNumberOfPickedPoints

 If n_of_ppoints = 0 Then
		MsgBox _
			"No Points are picked - aborting Macro", _
			vbOkOnly + vbCritical, _
			"Error!"
		Exit All
ElseIf n_of_ppoints Mod 2 <> 0 Then
		MsgBox _
			"Odd number of points are picked - aborting Macro", _
			vbOkOnly + vbCritical, _
			"Error!"
		Exit All

 End If

 ReDim Array_x(n_of_ppoints)
 ReDim Array_y(n_of_ppoints)
 ReDim Array_z(n_of_ppoints)
	
 For i = 1 To n_of_ppoints
       If Pick.GetPickpointCoordinates (i, Array_x(i-1), Array_y(i-1), Array_z(i-1))  = True Then  
       Else 
        MsgBox "failed to get pickpoint coordinates"
       End If
 Next i

 port_nr_offset=Solver.GetNumberOfPorts

 Dim sCommand As String

 Dim isLocalWCSActive As Boolean

  isLocalWCSActive = False
  sCommand=""
  If WCS.IsWCSActive() = "local" Then
	  'WCS.ActivateWCS("global")
	  sCommand = "WCS.ActivateWCS ""global"""
	  AddToHistory "Activate global coordinates system", sCommand
	  isLocalWCSActive = True
  End If


 For i = 0 To n_of_ppoints/2-1 STEP 1

	sCommand = ""
	sCommand = sCommand + "With DiscretePort" + vbLf
	sCommand = sCommand + "     .Reset" + vbLf
	sCommand = sCommand + "     .Portnumber """+ CStr(i+1+port_nr_offset)+ """" + vbLf
	sCommand = sCommand + "     .Type ""Sparameter""" + vbLf
	sCommand = sCommand + "     .Voltage """+ "1.0"+"""" + vbLf
	sCommand = sCommand + "     .Current """+ "1.0"+"""" + vbLf
	sCommand = sCommand + "     .SetP1 False,""" +StrValue(Array_x(i*2))+""", """+StrValue(Array_y(i*2))+""", """ +StrValue(Array_z(i*2)) + """" + vbLf
	sCommand = sCommand + "     .SetP2 False,""" + StrValue(Array_x(i*2+1)) + """, """+StrValue(Array_y(i*2+1))+""", """+StrValue( Array_z(i*2+1)) + """" + vbLf
    sCommand = sCommand + "     .Impedance """ + "50.0"+"""" + vbLf
    sCommand = sCommand + "     .Monitor True" + vbLf
    sCommand = sCommand + "     .Create" + vbLf
    sCommand = sCommand + "End With" + vbLf
    AddToHistory "define discrete port: "+ CStr(i+1+port_nr_offset), sCommand

	sCommand = ""
    sCommand = sCommand + "With Brick" + vbLf
    sCommand = sCommand + "     .Reset" + vbLf
    sCommand = sCommand + "     .Name ""PortMesh_"+CStr(i+1+port_nr_offset)+"""" + vbLf
    sCommand = sCommand + "     .Component """ + "9 Mesh box"+""""+ vbLf
    sCommand = sCommand + "     .Material """ + "Vacuum"+""""+ vbLf
    'check if coordinates of two points are along an axis
	Dim deltax As Double, deltay As Double, deltaz As Double

	deltax = Abs(Array_x(i*2)-Array_x(i*2+1))
	deltay = Abs(Array_y(i*2)-Array_y(i*2+1))
	deltaz = Abs(Array_z(i*2)-Array_z(i*2+1))

    If Abs(Array_x(i*2)-Array_x(i*2+1))< 0.1 And Abs(Array_y(i*2)-Array_y(i*2+1))< 0.1 And deltaz >= deltay And deltaz >= deltax  Then
    'along axis Z
		Array_x(i*2) = Array_x(i*2)-0.05
		Array_x(i*2+1) = Array_x(i*2+1)+0.05
		Array_y(i*2) = Array_y(i*2)-0.05
		Array_y(i*2+1) = Array_y(i*2+1)+0.05
	ElseIf Abs(Array_x(i*2)-Array_x(i*2+1))< 0.1 And Abs(Array_z(i*2)-Array_z(i*2+1))< 0.1 And deltay >= deltaz And deltay >= deltax  Then
		Array_x(i*2) = Array_x(i*2)-0.05
		Array_x(i*2+1) = Array_x(i*2+1)+0.05
		Array_z(i*2) = Array_z(i*2)-0.05
		Array_z(i*2+1) = Array_z(i*2+1)+0.05
	ElseIf Abs(Array_y(i*2)-Array_y(i*2+1))< 0.1 And Abs(Array_z(i*2)-Array_z(i*2+1))< 0.1 And deltax >= deltay And deltax >= deltaz  Then
		Array_z(i*2) = Array_z(i*2)-0.05
		Array_z(i*2+1) = Array_z(i*2+1)+0.05
		Array_y(i*2) = Array_y(i*2)-0.05
		Array_y(i*2+1) = Array_y(i*2+1)+0.05
	End If
    sCommand = sCommand + "     .Xrange """ +StrValue(Array_x(i*2))+""", """+StrValue(Array_x(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Yrange """ +StrValue(Array_y(i*2))+""", """+StrValue(Array_y(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Zrange""" +StrValue(Array_z(i*2))+""", """+StrValue(Array_z(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Create" + vbLf
    sCommand = sCommand + "End With" + vbLf
    AddToHistory "define Port mesh: "+ CStr(i+1+port_nr_offset), sCommand

	sCommand = "Group.AddItem ""solid$9 Mesh box:PortMesh_"+CStr(i+1+port_nr_offset)+""", ""Excluded from Simulation"""+ vbLf
	AddToHistory "Exclude PortMesh_"+ CStr(i+1+port_nr_offset)+" from simulation", sCommand

	sCommand = ""

	If deltax >= deltay And deltax >= deltaz Then
		sCommand = "Group.AddItem ""solid$9 Mesh box:PortMesh_"+CStr(i+1+port_nr_offset)+""", ""7 XBox"""+ vbLf
		AddToHistory "Add Portmesh_"+ CStr(i+1+port_nr_offset)+" to 7 XBox", sCommand
	ElseIf deltay >= deltax And deltay >= deltaz Then
		sCommand = "Group.AddItem ""solid$9 Mesh box:PortMesh_"+CStr(i+1+port_nr_offset)+""", ""8 YBox"""+ vbLf
		AddToHistory "Add Portmesh_"+ CStr(i+1+port_nr_offset)+" to 8 YBox", sCommand
	ElseIf deltaz >= deltay And deltaz >= deltax Then
		sCommand = "Group.AddItem ""solid$9 Mesh box:PortMesh_"+CStr(i+1+port_nr_offset)+""", ""9 ZBox"""+ vbLf
		AddToHistory "Add Portmesh_"+ CStr(i+1+port_nr_offset)+" to 9 ZBox", sCommand
	End If

 Next i

 sCommand = ""
 sCommand = sCommand + "Pick.ClearAllPicks"
 AddToHistory "Clear all picked points", sCommand

 'Pick.ClearAllPicks

 'Mesh.Update
 'Mesh.ViewMeshMode(False)

 sCommand=""
 If isLocalWCSActive = True Then
 	'WCS.ActivateWCS("local")
	sCommand="WCS.ActivateWCS ""local"""
	AddToHistory "Activate local coordinates system", sCommand
 End If

End Sub
'---------------------------------------------------------------------
Function StrValue (getthedouble As Double) As String
  StrValue = Replace(CStr(getthedouble),",",".")
End Function
