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

 Dim N As Integer
 Dim selectedItem As String
 N = 0
 selectedItem = ResultTree.GetFirstChildName("Components\9 Mesh box")
 While selectedItem <> ""
 	N = N+1
 	selectedItem = ResultTree.GetNextItemName(selectedItem)
 Wend
 port_nr_offset = N

 Dim sCommand As String

 Dim isLocalWCSActive As Boolean

  isLocalWCSActive = False
  'sCommand=""
  If WCS.IsWCSActive() = "local" Then
	  'WCS.ActivateWCS("global")
	  sCommand = "WCS.ActivateWCS ""global"""
	  AddToHistory "Activate global coordinates system", sCommand
	  isLocalWCSActive = True
  End If


 For i = 0 To n_of_ppoints/2-1 STEP 1
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
	sCommand = ""
    sCommand = sCommand + "With Brick" + vbLf
    sCommand = sCommand + "     .Reset" + vbLf
    sCommand = sCommand + "     .Name ""Box_"+CStr(i+1+port_nr_offset)+"""" + vbLf
    sCommand = sCommand + "     .Component """ + "9 Others"+""""+ vbLf
    sCommand = sCommand + "     .Material """ + "Vacuum"+""""+ vbLf
    sCommand = sCommand + "     .Xrange """ +StrValue(Array_x(i*2))+""", """+StrValue(Array_x(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Yrange """ +StrValue(Array_y(i*2))+""", """+StrValue(Array_y(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Zrange""" +StrValue(Array_z(i*2))+""", """+StrValue(Array_z(i*2+1))+"""" + vbLf
    sCommand = sCommand + "     .Create" + vbLf
    sCommand = sCommand + "End With" + vbLf
    AddToHistory "define Box_"+ CStr(i+1+port_nr_offset), sCommand

	'sCommand = "Group.AddItem ""solid$9 Mesh box:MeshBox_"+CStr(i+1+port_nr_offset)+""", ""Excluded from Simulation"""+ vbLf
	'AddToHistory "Exclude MeshBox_"+ CStr(i+1+port_nr_offset)+" from simulation", sCommand

	'Dim deltax As Double, deltay As Double, deltaz As Double

	'deltax = Abs(Array_x(i*2)-Array_x(i*2+1))
	'deltay = Abs(Array_y(i*2)-Array_y(i*2+1))
	'deltaz = Abs(Array_z(i*2)-Array_z(i*2+1))
	'sCommand = ""

	'If deltax >= deltay And deltax >= deltaz Then
	'	sCommand = "Group.AddItem ""solid$9 Mesh box:MeshBox_"+CStr(i+1+port_nr_offset)+""", ""7 XBox"""+ vbLf
	'	AddToHistory "Add MeshBox_"+ CStr(i+1+port_nr_offset)+" to 7 XBox", sCommand
	'ElseIf deltay >= deltax And deltay >= deltaz Then
	'	sCommand = "Group.AddItem ""solid$9 Mesh box:MeshBox_"+CStr(i+1+port_nr_offset)+""", ""8 YBox"""+ vbLf
	'	AddToHistory "Add MeshBox_"+ CStr(i+1+port_nr_offset)+" to 8 YBox", sCommand
	'ElseIf deltaz >= deltay And deltaz >= deltax Then
	'	sCommand = "Group.AddItem ""solid$9 Mesh box:MeshBox_"+CStr(i+1+port_nr_offset)+""", ""9 ZBox"""+ vbLf
	'	AddToHistory "Add MeshBox_"+ CStr(i+1+port_nr_offset)+" to 9 ZBox", sCommand
	'End If

 Next i

 'Mesh.Update
 'Mesh.ViewMeshMode(False)
 sCommand = ""
 sCommand = sCommand + "Pick.ClearAllPicks"
 AddToHistory "Clear all picked points", sCommand
' Pick.ClearAllPicks
 'sCommand=""
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
