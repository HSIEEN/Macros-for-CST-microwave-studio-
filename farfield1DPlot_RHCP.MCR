' Copy farfield 1d results to 1D results
'2023-02-10 by Shawn Shi
'Option Explicit

'#include "vba_globals_all.lib"

Sub Main ()
    'Get current farfield plot mode
    Dim CurrentPlotMode As String

    CurrentPlotMode = FarfieldPlot.GetPlotMode

    FarfieldPlot.StoreSettings

    'FarfieldPlot.Plot

    'FarField Calculation

    Dim SelectedItem As String
    'Dim n As Integer

    'Dim Frequency As Double, FrequencyStr As String
    'Dim PortStr As String

    SelectedItem = GetSelectedTreeItem

    If (InStr(SelectedItem,"farfield (") = 0) Then

        MsgBox("Please select a farfield result before runing this macro.",vbCritical,"Warning")

        Exit All
    End If

	Begin Dialog UserDialog 200,105,"Please choose a cut plane to plot",.DialogFunction ' %GRID:10,7,1,1
		Text 10,7,90,14,"Coordinate:",.Text1
		OptionGroup .Group1
			OptionButton 10,21,60,14,"Theta",.OptionButton1
			OptionButton 80,21,50,14,"Phi",.OptionButton2
		Text 10,42,60,14,"Angle:",.Text2
		TextBox 70,42,50,14,.Angle
		OKButton 10,70,90,21
		CancelButton 110,70,70,21
	End Dialog
	Dim dlg As UserDialog

	dlg.Group1 = 0
	dlg.Angle = "90"

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Dim parameterFile As String
   	Dim prjPath As String

   	prjPath = GetProjectPath("Project")
   	parameterFile = prjPath + "\dialog_parameter.txt"

	Select Case Action%
	Case 1 ' Dialog box initialization
		ReStoreAllDialogSettings_LIB(parameterFile)
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box
		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
   			'parameterFile = "D:\Simulation\SXW\Research\Basics\a.txt"
			StoreAllDialogSettings_LIB(parameterFile)
			'Dim angleStr As String
			Dim angle As Double
			Dim groupValue As Double
			Dim PortStr As String
			Dim FrequencyStr As String

			angle = CDbl(DlgText("Angle"))
			groupValue = DlgValue("Group1")

			Dim SelectedItem As String
    		SelectedItem = GetSelectedTreeItem


			PortStr = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)
			FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)

			FarfieldPlot.Reset
			'FarfieldPlot.SetAntennaType("directional_linear")
			'FarfieldPlot.Plot
			FarfieldPlot.SetAntennaType("directional_circular")
			'FarfieldPlot.Plot
			FarfieldPlot.SetPlotMode("directivity")
			'FarfieldPlot.Plot
			FarfieldPlot.SelectComponent("Abs")
			FarfieldPlot.PlotType("polar")
			If groupValue = 0 Then

				FarfieldPlot.Vary("angle2")
				FarfieldPlot.Theta(angle)

			Else

				FarfieldPlot.Vary("angle1")
				FarfieldPlot.Phi(angle)

			End If

			FarfieldPlot.Plot
			FarfieldPlot.CopyFarfieldTo1DResults("farfield1D_CP_"+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Abs")
			FarfieldPlot.SelectComponent("Right")
			'FarfieldPlot.PlotType("polar")
			FarfieldPlot.Plot
			FarfieldPlot.CopyFarfieldTo1DResults("farfield1D_CP_"+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Right")
			FarfieldPlot.SelectComponent("Left")
			'FarfieldPlot.PlotType("polar")
			FarfieldPlot.Plot
			FarfieldPlot.CopyFarfieldTo1DResults("farfield1D_CP_"+FrequencyStr+"GHz","farfield (f="+FrequencyStr+")["+PortStr+"]_Left")

			SelectTreeItem("1D Results\farfield1D_CP_"+FrequencyStr+"GHz")

            CurrentItem = FirstChildItem

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
