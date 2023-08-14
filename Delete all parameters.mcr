'#Language "WWB-COM"

'Option Explicit

Sub Main
	n=GetNumberOfParameters
	While n>0
		para=GetParameterName(0)
		DeleteParameter(para)
		n=GetNumberOfParameters
	Wend
End Sub
