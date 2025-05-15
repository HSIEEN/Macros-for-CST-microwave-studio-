' Create Parameters for switching on-off selected solids
'--------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3D.lib"

Sub Main ()

	Dim aSolidArray_CST() As String, nSolids_CST As Integer
	Dim iSolid_CST As Integer, bSolids_CST As Boolean

	Dim sParaName As String, sFullSolidName As String
	Dim sParaNameTemp As String
	Dim ParaNameMaxLength As Long
    Dim sCommand As String
    Dim SelectedItem As String
    Dim group_sname As String

   	Dim cst_parnames_pretty As String
	Dim cst_parvalues_pretty As String

	Dim nParameter As Integer
	nParameter = 0

	cst_parnames_pretty = ""
	cst_parvalues_pretty = ""

	'SelectSolids_LIB(aSolidArray_CST(), nSolids_CST)
	MsgBox "Please make sure you have selected all solids for creating parameters.", vbInformation,"Notice"
	nSolids_CST = GetNumberOfSelectedTreeItems
    SelectedItem = GetSelectedTreeItem
	'MsgBox cstr(nSolids_CST)

	' now step through all the solids
	If (nSolids_CST > 0) Then

		sCommand = ""
		' create the group to add the solids to; additionally set up the meshing special
		'when a variable is set to be 0, the corresponding solid is moved into this group
		sCommand = sCommand + "Group.Add ""materialindependent-meshing"", ""mesh""" + vbLf
		sCommand = sCommand + "With MeshSettings" + vbLf
		sCommand = sCommand + "	With .ItemMeshSettings (""group$materialindependent-meshing"")" + vbLf
		sCommand = sCommand + "		.SetMeshType ""Tet""" + vbLf
		sCommand = sCommand + "		.Set ""MaterialIndependent"", 1" + vbLf
		sCommand = sCommand + "	End With" + vbLf + "End With" + vbLf
		
		ParaNameMaxLength = 1

		While SelectedItem <> ""
			group_sname = GetQualifiedNameFromTreeName(SelectedItem)
			sFullSolidName = Right(group_sname,Len(group_sname)-InStrRev(group_sname, "$"))
			'sFullSolidName = aSolidArray_CST(iSolid_CST-1)

			If (sFullSolidName <> "") Then
				nParameter = nParameter + 1
				'sParaName = "AA_" + Mid(sFullSolidName,1+InStrRev(sFullSolidName,":"))
				sParaNameTemp = Replace(sFullSolidName," ","_") ' parameternames with a blank are not allowed!
				'MsgBox sParaNameTemp
				sParaName = "para_" + Replace(Replace(sParaNameTemp,":","_"),"/","_")
				If Len(sParaName)>ParaNameMaxLength Then ParaNameMaxLength = Len(sParaName)

				sCommand = sCommand + "MakeSureParameterExists(""" + sParaName + """, ""1"")" + vbLf

				MakeSureParameterExists(sParaName, "1")
				SetParameterDescription(sParaName, "(0=off / 1=on) consider " + sFullSolidName + " for simulation")

				'sCommand = sCommand + "Solid.SetUseForSimulation "
				'sCommand = sCommand + """" + sFullSolidName + """"  + " , "
				'sCommand = sCommand + "IIf(" + sParaName + "," + """True""" + "," + """False""" + ")" + vbLf
				' add something to change the material if the vaiables are set
				sCommand = sCommand + "If (" + sParaName + "=0) Then" + vbLf
				sCommand = sCommand + "	Solid.ChangeMaterial(""" + sFullSolidName + """, ""Vacuum"")" + vbLf
				sCommand = sCommand + "	Group.AddItem(""solid$" + sFullSolidName + """, ""materialindependent-meshing"")" + vbLf + "End If" + vbLf

				cst_parnames_pretty  = cst_parnames_pretty  + CST_Print(sParaName, ParaNameMaxLength+1)
				cst_parvalues_pretty = cst_parvalues_pretty + CST_Print("1", ParaNameMaxLength+1)

				'MsgBox sCommand
				'Solid.SetUseForSimulation sFullSolidName,IIf(sParaName,"True","False")
			End If
			SelectedItem = GetNextSelectedTreeItem
		Wend

		sCommand = sCommand + "Solid.FastModelUpdate """ + "False"""

		AddToHistory "Macro: Parametric Switch on-off for Consider Solids for Simulation", sCommand
		Rebuild
		If ((MsgBox Cstr(nParameter) + " Parameters have been successfully defined now." + vbCrLf  + vbCrLf + _
				"Do you want to open a dummy tabular text file for the parameters?",vbYesNo)=vbYes) Then

			Dim stxtfileName As String, txtFile As Long
			txtFile = FreeFile
			stxtfileName = GetProjectPath("Temp")+"\tmp-parametersweep-input-file.txt"
			Open stxtfileName For Output As #txtFile
				Print #txtFile, cst_parnames_pretty
				Print #txtFile, cst_parvalues_pretty
				Print #txtFile, cst_parvalues_pretty
				Print #txtFile, cst_parvalues_pretty
				Print #txtFile, cst_parvalues_pretty
				Print #txtFile, cst_parvalues_pretty
			Close #txtFile
			Shell("notepad.exe " + stxtfileName, 1)

		End If
	Else
		MsgBox "No solids are selected!",vbCritical,"Error"
		Exit All
	End If

End Sub
