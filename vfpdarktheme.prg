*-----------------------------------------------------------------------------------
* Created by Daniel Nogales
* ver 1.001 - 03/09/2018 
*-----------------------------------------------------------------------------------
SET PATH TO HOME() + "FFC\" ADDITIVE
#INCLUDE "Registry.H"
Clear

SetDarkThemeEditorColors()

Messagebox([Vuelva a ingresar a Microsoft Visual FoxPro para aplicar los cambios de tema en el editor.];
		  , 4096 + 64;
		  , [DarkTheme for Microsoft Visual FoxPro])


Function SetDarkThemeEditorColors()
Local lnLoop    As Integer;
	, lcRegKey  As String;
	, lcValue   As String;
	, regConst  As String

Dimension aEditorColors[7, 3]
aEditorColors[1, 1] = "EditorCommentColor"
aEditorColors[1, 2] = "RGB(87,166,74,30,30,30), NoAuto, NoAuto"
aEditorColors[2, 1] = "EditorKeywordColor"
aEditorColors[2, 2] = "RGB(57,135,214,30,30,30), NoAuto, NoAuto"
aEditorColors[3, 1] = "EditorConstantColor"
aEditorColors[3, 2] = "RGB(177,130,177,30,30,30), NoAuto, NoAuto"
aEditorColors[4, 1] = "EditorNormalColor"
aEditorColors[4, 2] = "RGB(0,255,0,30,30,30), NoAuto, NoAuto"
aEditorColors[5, 1] = "EditorOperatorColor"
aEditorColors[5, 2] = "RGB(220,220,220,30,30,30), NoAuto, NoAuto"
aEditorColors[6, 1] = "EditorStringColor"
aEditorColors[6, 2] = "RGB(214,136,82,30,30,30), NoAuto, NoAuto"
aEditorColors[7, 1] = "EditorVariableColor"
aEditorColors[7, 2] = "RGB(127,216,170,30,30,30), NoAuto, NoAuto"

lcRegKey = "Software\Microsoft\VisualFoxPro\9.0\Options"
regConst = HKEY_CURRENT_USER
For lnLoop = 1 To Alen(aEditorColors, 1)
	lcValue = SetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1], aEditorColors[lnLoop, 2])
Endfor
Endfunc

Function SetRegKey()
Lparameters hkey As Character, keyPath As Character, entry As Character, regValue As Character
GetRegistryApi()
keyPath = Addbs(Transform(Evl(keyPath, "")))
entry = Transform(Evl(entry, ""))
regValue = Transform(regValue)
Return (oRegApi.SetRegKey(entry, regValue, keyPath, hkey, .T.) = 0)

Function GetRegistryApi() AS Registry
If Vartype(oRegApi) # "O"
	Public oRegApi AS Registry
	oRegApi = NewObject("Registry",;
	                    HOME() + "FFC\Registry.VCX")
Endif
Return oRegApi
Endfunc