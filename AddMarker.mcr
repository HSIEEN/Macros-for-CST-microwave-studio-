' Calculate the average efficiency over specified bands
'Option Explicit
'20220828-By Shawn in COROS

Public SelectedItem As String

Sub Main ()
  'Efficiency results parent path
   SelectedItem = GetSelectedTreeItem

   If SelectedItem = "" Then
   	 MsgBox("No item is selected, please select at least one 1D curve!",vbCritical,"Warning")
   	 Exit All
   	Else
   		Dim nResults As Integer
   		nResults = Resulttree.GetTreeResults(SelectedItem,"folder 0D/1D recursive","",paths,types,files,info)
   		If InStr(SelectedItem,"1D Results\") = 0 Then
			MsgBox("Selected item is not in 1D Results, please select at least one 1D curve!",vbCritical,"Warning")
   			Exit All
   		Else
	   		If nResults <> 1 Or types(0) <> "XYSIGNAL" Then
	            MsgBox("Selected item is not a curve, please select at least one 1D curve!",vbCritical,"Warning")
	   			Exit All
   			End If
   		End If
   End If

    Dim InformationStr As String
    InformationStr = "请输入要标注的频段（eg L1 L5 W2 W5）："
	Begin Dialog UserDialog 410,98,"根据输入频段添加Marker",.DialogFunction ' %GRID:10,7,1,1
		Text 10,28,340,14,InformationStr,.Text1
		TextBox 20,49,90,14,.Band
		OKButton 20,77,90,21
		CancelButton 120,77,90,21
	End Dialog
	Dim dlg As UserDialog

	dlg.Band = "L1"

	If Dialog(dlg,-2) = 0 Then
		Exit All
	End If

End Sub

Private Function DialogFunction(DlgItem$, Action%, SuppValue?) As Boolean

	Select Case Action
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem DialogFunction = True ' Prevent button press from closing the dialog box

		Select Case DlgItem
		Case "Cancle"
			Exit All
		Case "OK"
			DialogFunction = False

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
						     .AddMarker(1.214) '
						     '.ShowMarkerAtMin
						     .Plot ' make changes visible

						End With

            		End If
            		If InStr(Bands,"W2")<> 0 Then
						With Plot1D

						     .AddMarker(2.4) '
						     .AddMarker(2.483) '
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

