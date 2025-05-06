' Calculate the average efficiency over specified bands
'Option Explicit
'20220828-By Shawn in COROS

Public SelectedItem As String
'#include "vba_globals_all.lib"

Sub Main ()
  'Efficiency results parent path
   SelectedItem = GetSelectedTreeItem

   If SelectedItem = "" Then
   	 MsgBox("No item is selected, please select at least one 1D curve!",vbCritical,"Warning")
   	 Exit All
   	Else
   		Dim nResults As Integer
   		nResults = ResultTree.GetTreeResults(SelectedItem,"folder 0D/1D recursive","",paths,types,files,info)
   		If InStr(SelectedItem,"1D Results\") = 0 Then
			MsgBox("Selected item is not in 1D Results, please select at least one 1D curve!",vbCritical,"Warning")
   			Exit All
   		'Else
	   	'	If nResults <> 1 Or types(0) <> "XYSIGNAL" Then
	    '        MsgBox("Selected item is not a curve, please select at least one 1D curve!",vbCritical,"Warning")
	   	'		Exit All
   		'	End If
   		End If
   End If

    Dim InformationStr As String
    InformationStr = "请输入要标注的频段（eg L1 L5 W2 W5）或者直接输入频点（eg 1.56 2.4）："
	Begin Dialog UserDialog 320,77,"根据输入频段添加Marker",.DialogFunction ' %GRID:10,7,1,1
		Text 10,7,310,14,InformationStr,.Text1
		TextBox 50,28,160,14,.Band
		OKButton 20,56,90,21
		CancelButton 170,56,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.Band = "L1 L5"

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

End Sub

Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean
	Dim parameterFile As String
	Dim prjPath As String

	prjPath = GetProjectPath("Project")
	parameterFile = prjPath + "\marker_dialog_parameter.txt"

	Select Case Action
	Case 1 ' Dialog box initialization
		RestoreAllDialogSettings_LIB(parameterFile)
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box

		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
			DialogFunction = False
			StoreAllDialogSettings_LIB(parameterFile)
			Dim Bands As String
		    Dim CurrentItem As String
		    Dim Label As String
		    Dim index As Integer
		    Bands = DlgText("Band")

            While SelectedItem <> ""
                    SelectTreeItem(SelectedItem)
                    
            		If InStr(Bands,"L1")<> 0 Then
						'Label = Right(SelectedItem,Len(SelectedItem)-InStrRev(SelectedItem,"\"))
               			With Plot1D

						     .AddMarker(1.559) '
						     .AddMarker(1.610) '
						     '.ShowMarkerAtMin
						     .Plot ' make changes visible

						End With


            		End If
            		If InStr(Bands,"L5")<> 0 Then
            			With Plot1D

						     .AddMarker(1.164) '
						     .AddMarker(1.189) '
						     '.ShowMarkerAtMin
						     .Plot ' make changes visible

						End With

            		End If
            		If InStr(Bands,"W2")<> 0 Then
						With Plot1D

						     .AddMarker(2.40) '
						     .AddMarker(2.48) '
						     '.ShowMarkerAtMin
						     .Plot ' make changes visible

						End With
            		End If
            		If InStr(Bands,"W5")<> 0 Then
						With Plot1D

						     .AddMarker(5.15) '
						     .AddMarker(5.85) '
						     '.ShowMarkerAtMin
						     .Plot ' make changes visible

						End With
            		End If

            	SelectedItem = GetNextSelectedTreeItem
            Wend


		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : DialogFunction = True ' Continue getting idle actions
	Case 6 ' Function key

	End Select



End Function

