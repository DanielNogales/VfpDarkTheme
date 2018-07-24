*-----------------------------------------------------------------------------------
* Created by Daniel Nogales
* ver 1.000 - 20/07/2018 
*-----------------------------------------------------------------------------------
*!* --------------------------------------------------------------------------------
*!*	Inspiración para esta funcionalidad en Microsoft Visual FoxPro.
*!*	https://myvfpblog.blogspot.com.ar/2008/03/vfp-editor-code-rtf2html-part-4.html
*!* --------------------------------------------------------------------------------

Set Asserts On
Assert .F.

Local pathExecution As String
pathExecution = Left(Sys(16),Rat('\', Sys(16)))
Set Path To (pathExecution) Additive
#INCLUDE registry.H

Clear
*!*	 Save User config in registry values. [User]<RegistryEntry>
SaveUserThemeEditorColors()

SetDarkThemeEditorColors()

Messagebox([Vuelva a ingresar a Microsoft Visual FoxPro para aplicar los cambios de tema en el editor.];
		  , 4096 + 64;
		  , [DarkTheme for Microsoft Visual FoxPro])


Function SaveUserThemeEditorColors()
Local lnLoop    As Integer;
	, lcRegKey  As String;
	, lcValue   As String;
	, regConst  As String

Dimension aEditorColors[7, 3]
aEditorColors[1, 1] = 'EditorCommentColor'
aEditorColors[2, 1] = 'EditorKeywordColor'
aEditorColors[3, 1] = 'EditorConstantColor'
aEditorColors[4, 1] = 'EditorNormalColor'
aEditorColors[5, 1] = 'EditorOperatorColor'
aEditorColors[6, 1] = 'EditorStringColor'
aEditorColors[7, 1] = 'EditorVariableColor'

Dimension aUserEditorColors[7, 3]
aUserEditorColors[1, 1] = 'UserEditorCommentColor'
aUserEditorColors[2, 1] = 'UserEditorKeywordColor'
aUserEditorColors[3, 1] = 'UserEditorConstantColor'
aUserEditorColors[4, 1] = 'UserEditorNormalColor'
aUserEditorColors[5, 1] = 'UserEditorOperatorColor'
aUserEditorColors[6, 1] = 'UserEditorStringColor'
aUserEditorColors[7, 1] = 'UserEditorVariableColor'

lcRegKey = 'Software\Microsoft\VisualFoxPro\9.0\Options'
regConst = HKEY_CURRENT_USER
For lnLoop = 1 To Alen(aUserEditorColors, 1)
	lcValue = GetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1])
	lcValueUser = GetRegKey(regConst, lcRegKey, aUserEditorColors[lnLoop, 1])

	If Empty(lcValueUser) Or lcValueUser = [.F.]
		lcValue = SetRegKey(regConst, lcRegKey, aUserEditorColors[lnLoop, 1], lcValue)
	Endif
Endfor
Endfunc


Function SetUserThemeEditorColors()
Local lnLoop    As Integer;
	, lcRegKey  As String;
	, lcValue   As String;
	, regConst  As String

Dimension aEditorColors[7, 3]
aEditorColors[1, 1] = 'EditorCommentColor'
aEditorColors[2, 1] = 'EditorKeywordColor'
aEditorColors[3, 1] = 'EditorConstantColor'
aEditorColors[4, 1] = 'EditorNormalColor'
aEditorColors[5, 1] = 'EditorOperatorColor'
aEditorColors[6, 1] = 'EditorStringColor'
aEditorColors[7, 1] = 'EditorVariableColor'

Dimension aUserEditorColors[7, 3]
aUserEditorColors[1, 1] = 'UserEditorCommentColor'
aUserEditorColors[2, 1] = 'UserEditorKeywordColor'
aUserEditorColors[3, 1] = 'UserEditorConstantColor'
aUserEditorColors[4, 1] = 'UserEditorNormalColor'
aUserEditorColors[5, 1] = 'UserEditorOperatorColor'
aUserEditorColors[6, 1] = 'UserEditorStringColor'
aUserEditorColors[7, 1] = 'UserEditorVariableColor'

lcRegKey = 'Software\Microsoft\VisualFoxPro\9.0\Options'
regConst = HKEY_CURRENT_USER

For lnLoop = 1 To Alen(aEditorColors, 1)
	lcValue = GetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1])
	lcValueUser = GetRegKey(regConst, lcRegKey, aUserEditorColors[lnLoop, 1])
	lcValue = SetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1], lcValueUser)
Endfor
Endfunc

Function SetDefaultThemeEditorColors()
Local lnLoop    As Integer;
	, lcRegKey  As String;
	, lcValue   As String;
	, regConst  As String

Dimension aEditorColors[7, 3]
aEditorColors[1, 1] = 'EditorCommentColor'
aEditorColors[1, 2] = 'RGB(0,128,0,255,255,255), NoAuto, Auto'
aEditorColors[2, 1] = 'EditorKeywordColor'
aEditorColors[2, 2] = 'RGB(0,0,255,255,255,255), NoAuto, Auto'
aEditorColors[3, 1] = 'EditorConstantColor'
aEditorColors[3, 2] = 'RGB(0,0,0,255,255,255), Auto, Auto'
aEditorColors[4, 1] = 'EditorNormalColor'
aEditorColors[4, 2] = 'RGB(0,0,0,255,255,255), Auto, Auto'
aEditorColors[5, 1] = 'EditorOperatorColor'
aEditorColors[5, 2] = 'RGB(0,0,0,255,255,255), Auto, Auto'
aEditorColors[6, 1] = 'EditorStringColor'
aEditorColors[6, 2] = 'RGB(0,0,0,255,255,255), Auto, Auto'
aEditorColors[7, 1] = 'EditorVariableColor'
aEditorColors[7, 2] = 'RGB(0,0,0,255,255,255), Auto, Auto'

lcRegKey = 'Software\Microsoft\VisualFoxPro\9.0\Options'
regConst = HKEY_CURRENT_USER
For lnLoop = 1 To Alen(aEditorColors, 1)
	lcValue = SetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1], aEditorColors[lnLoop, 2])
Endfor
Endfunc

Function SetDarkThemeEditorColors()
Local lnLoop    As Integer;
	, lcRegKey  As String;
	, lcValue   As String;
	, regConst  As String

Dimension aEditorColors[7, 3]
aEditorColors[1, 1] = 'EditorCommentColor'
aEditorColors[1, 2] = 'RGB(87,166,74,30,30,30), NoAuto, NoAuto'
aEditorColors[2, 1] = 'EditorKeywordColor'
aEditorColors[2, 2] = 'RGB(57,135,214,30,30,30), NoAuto, NoAuto'
aEditorColors[3, 1] = 'EditorConstantColor'
aEditorColors[3, 2] = 'RGB(177,130,177,30,30,30), NoAuto, NoAuto'
aEditorColors[4, 1] = 'EditorNormalColor'
aEditorColors[4, 2] = 'RGB(0,255,0,30,30,30), NoAuto, NoAuto'
aEditorColors[5, 1] = 'EditorOperatorColor'
aEditorColors[5, 2] = 'RGB(220,220,220,30,30,30), NoAuto, NoAuto'
aEditorColors[6, 1] = 'EditorStringColor'
aEditorColors[6, 2] = 'RGB(214,136,82,30,30,30), NoAuto, NoAuto'
aEditorColors[7, 1] = 'EditorVariableColor'
aEditorColors[7, 2] = 'RGB(127,216,170,30,30,30), NoAuto, NoAuto'

lcRegKey = 'Software\Microsoft\VisualFoxPro\9.0\Options'
regConst = HKEY_CURRENT_USER
For lnLoop = 1 To Alen(aEditorColors, 1)
	lcValue = SetRegKey(regConst, lcRegKey, aEditorColors[lnLoop, 1], aEditorColors[lnLoop, 2])
Endfor
Endfunc


Function GetRegKey()
Lparameters hkey As Character, keyPath As Character, entry As Character, defValue As Character
Local regValue As Character
GetRegistryApi()
keyPath = Addbs(Transform(Evl(keyPath, '')))
entry = Transform(Evl(entry, ''))
hkey = Upper(Transform(Evl(hkey, HKEY_CURRENT_USER)))
defValue = Transform(defValue)
regValue = ''
oRegApi.GetRegKey(entry, @regValue, keyPath, hkey)
If Empty(regValue)
	If Empty(oRegApi.setregkey(entry, defValue, keyPath, hkey, .T.))
		regValue = defValue
	Endif
Endif
Return regValue

Function SetRegKey()
Lparameters hkey As Character, keyPath As Character, entry As Character, regValue As Character
GetRegistryApi()
keyPath = Addbs(Transform(Evl(keyPath, '')))
entry = Transform(Evl(entry, ''))
hkey = Upper(Transform(Evl(hkey, HKEY_CURRENT_USER)))
regValue = Transform(regValue)
Return (oRegApi.SetRegKey(entry, regValue, keyPath, hkey, .T.) = 0)


Function GetRegistryApi() AS Registry
If Vartype(oRegApi) # 'O'
	Public oRegApi AS Registry
	Set Classlib To registry.vcx Alias MiRegistry Additive
	oRegApi = Createobject('registry')
Endif
Return oRegApi
Endfunc


