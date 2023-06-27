' Copy farfield 1d results to 1D results
'2023-02-10 by Shawn Shi
Option Explicit

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

	Begin Dialog UserDialog 210,154,"Please choose a cut plane to plot",.DialogFunction ' %GRID:10,7,1,1
		OKButton 20,126,80,21
		CancelButton 110,126,70,21
		GroupBox 10,63,190,56,"Cut plane settings:",.GroupBox1
		GroupBox 10,0,190,63,"Farfield plot mode:",.GroupBox2
		OptionGroup .Group2
			OptionButton 30,77,60,14,"Theta",.OptionButton1
			OptionButton 110,77,60,14,"Phi",.OptionButton2
		Text 40,98,50,14,"Angle:",.Text1
		TextBox 90,98,60,14,.Angle
		OptionGroup .Group1
			OptionButton 20,21,100,14,"Directivity",.OptionButton3
			OptionButton 130,21,60,14,"Gain",.OptionButton4
			OptionButton 20,42,110,14,"Realized gain",.OptionButton5
	End Dialog
	Dim dlg As UserDialog

	dlg.Group1 = 0
	dlg.Group2 = 1
	dlg.Angle = "270"

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
			Dim componentValue As Integer, planeValue As Integer
			Dim PortStr As String
			Dim FrequencyStr As String


			componentValue = DlgValue("Group1")
			angle = CDbl(DlgText("Angle"))
			planeValue = DlgValue("Group2")

			Dim SelectedItem As String
    		SelectedItem = GetSelectedTreeItem


			PortStr = Mid$(SelectedItem$,InStr(SelectedItem,"[")+1,InStr(SelectedItem,"]")-InStr(SelectedItem,"[")-1)
			FrequencyStr = Mid$(SelectedItem$,InStr(SelectedItem,"=")+1,InStr(SelectedItem,")")-InStr(SelectedItem,"=")-1)

			FarfieldPlot.Reset



			'FarfieldPlot.Plot

			'FarfieldPlot.Plot
			FarfieldPlot.SelectComponent("Abs")
			FarfieldPlot.PlotType("polar")
			If planeValue = 0 Then

				FarfieldPlot.Vary("angle2")
				FarfieldPlot.Theta(angle)

			Else

				FarfieldPlot.Vary("angle1")
				FarfieldPlot.Phi(angle)

			End If

			FarfieldPlot.SetAxesType("currentwcs")
			FarfieldPlot.SetAntennaType("unknown")
			Select Case componentValue
		   	Case 0
				FarfieldPlot.SetPlotMode("directivity")
			Case 1
				FarfieldPlot.SetPlotMode("gain")
			Case 2
				FarfieldPlot.SetPlotMode("realized gain")
		   	End Select
			'FarfieldPlot.SetAntennaType("directional_linear")
			'FarfieldPlot.SetAntennaType("directional_circular")
			FarfieldPlot.SetCoordinateSystemType("ludwig3")
			FarfieldPlot.SetAutomaticCoordinateSystem("True")
			FarfieldPlot.SetPolarizationType("Circular")

			FarfieldPlot.StoreSettings

			FarfieldPlot.Plot
			Dim DirName As String
			Dim ChildItem As String
			Select Case componentValue
		   	Case 0
				DirName = "CP directivity\@"+FrequencyStr+"GHz"
			Case 1
				DirName = "CP gain\@"+FrequencyStr+"GHz"
			Case 2
				DirName = "CP realized gain\@"+FrequencyStr+"GHz"
		   	End Select

			If Resulttree.DoesTreeItemExist("1D Results\"+DirName) Then
				ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName)

				While ChildItem <> ""
					With Resulttree
						.Name ChildItem
						.Delete
					End With
					ChildItem = Resulttree.GetFirstChildName("1D Results\"+DirName)
				Wend

			End If
			FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Abs")
			FarfieldPlot.SelectComponent("Right")
			'FarfieldPlot.PlotType("polar")
			FarfieldPlot.Plot
			FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Right")
			FarfieldPlot.SelectComponent("Left")
			'FarfieldPlot.PlotType("polar")
			FarfieldPlot.Plot
			FarfieldPlot.CopyFarfieldTo1DResults(DirName,"farfield (f="+FrequencyStr+")["+PortStr+"]_Left")

			SelectTreeItem("1D Results\"+DirName)

            'CurrentItem = FirstChildItem
            		Dim curveLabel As String
		Dim index As Integer
		'Dim SelectedItem As String

		SelectedItem = Resulttree.GetFirstChildName("1D Results\"+DirName)
		While SelectedItem <> ""
			'SelectTreeItem(selectedItem)
			curveLabel = Right(SelectedItem,Len(SelectedItem)-InStrRev(SelectedItem,"\"))

		   With Plot1D
		      index =.GetCurveIndexOfCurveLabel(curveLabel)
		     .SetLineStyle(index,"Solid",3) ' thick dashed line
		     .SetFont("Tahoma","bold","16")
		     '.SetLineColor(index,255,255,0)  ' yellow
		     .Plot ' make changes visible
			End With

			SelectedItem = Resulttree.GetNextItemName(SelectedItem)
		Wend

		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
