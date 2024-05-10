'#Language "WWB-COM"
'Convert Cartesian coordinates to Spherical coordiantes or vice verse

Option Explicit

Sub Main
	'While True
	Begin Dialog UserDialog 500,98,"Coordinates converter (angle is in degree)" ' %GRID:10,7,1,1
		GroupBox 0,0,500,70,"Cartesian <---> Spherical",.GroupBox1
		OKButton 90,70,90,21,.Calculate
		CancelButton 220,70,90,21
		Text 210,21,30,14,"x-u:",.Text1
		Text 210,49,40,14," ¦Ñ:",.Text4
		Text 310,21,30,14,"y-v:",.Text2
		Text 310,49,30,14,"¦È:",.Text5
		Text 400,21,40,14,"z-w:",.Text3
		Text 410,49,30,14,"¦Õ:",.Text6
		TextBox 250,21,50,14,.xu
		TextBox 250,49,50,14,.rou
		TextBox 340,21,50,14,.yv
		TextBox 340,49,50,14,.theta
		TextBox 440,21,50,14,.zw
		TextBox 440,49,50,14,.phi
		OptionGroup .Group1
			OptionButton 20,21,180,14,"Cartesian-->Spherical",.OptionButton1
			OptionButton 20,49,170,14,"Spherical-->Cartesian",.OptionButton2
	End Dialog
		Dim dlg As UserDialog
		dlg.xu="0"
		dlg.yv="0"
		dlg.zw="0"
		dlg.rou="0"
		dlg.theta="0"
		dlg.phi="0"
		While True
			If Dialog(dlg,-2) = 0 Then
				Exit All
			End If
			'Estimate if converting Cartesian to Spherical or converting Spherical to Cartesian
			Dim xu As Double, yv As Double, zw As Double
			Dim rou As Double, theta As Double, phi As Double
			xu = CDbl(dlg.xu)
			yv = CDbl(dlg.yv)
			zw = CDbl(dlg.zw)
			rou = CDbl(dlg.rou)
			theta = CDbl(dlg.theta)/180*3.1416
			phi = CDbl(dlg.phi)/180*3.1416
			'xu = CDbl(dlg.xu)
			If dlg.Group1 = 1 Then
				'Convert spherical coordinates to cartesian coordinates
				dlg.xu = CStr(Round(rou*Sin(theta)*Cos(phi),2))
				dlg.yv = CStr(Round(rou*Sin(theta)*Sin(phi),2))
				dlg.zw = CStr(Round(rou*Cos(theta),2))
			ElseIf dlg.Group1 = 0 Then
				'Convert cartesian coordinates to spherical coordinates
				dlg.rou = CStr(Round(Sqr(xu^2+yv^2+zw^2),2))
				If zw >= 0  Then
					dlg.theta = CStr(Round(Atn(Sqr(xu^2+yv^2)/(zw+1e-5))/3.1416*180,2))
				Else
					dlg.theta = CStr(Round(Atn(Sqr(xu^2+yv^2)/(zw+1e-5))/3.1416*180+180,2))
				End If

				If yv >= 0 And xu >= 0 Then
					dlg.phi = CStr(Round(Atn(yv/(xu+1e-5))/3.1416*180,2))
				ElseIf yv < 0 And xu >= 0 Then
					dlg.phi = CStr(Round(Atn(yv/(xu+1e-5))/3.1416*180+360,2))
				ElseIf xu < 0 Then
					dlg.phi = CStr(Round(Atn(yv/(xu+1e-5))/3.1416*180+180,2))
				End If
			End If
		Wend
End Sub
