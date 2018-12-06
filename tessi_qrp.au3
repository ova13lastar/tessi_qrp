; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
#AutoIt3Wrapper_Res_ProductName=tessi_qrp
#AutoIt3Wrapper_Res_Description=Outil de remplissage automatique du driver TessiPOST à partir d'un courrier QRP
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_FileVersion=1.0.0
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/APPLINAT
#AutoIt3Wrapper_Res_LegalCopyright=yann.daniel@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Icon="static\icon.ico"
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Run_Au3Stripper=N
#Au3Stripper_Parameters=/MO /RSLN
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#NoTrayIcon
; Includes Constants
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
#include <GUIConstants.au3>
; Includes
#include <Misc.au3>
#include <Array.au3>
#include <File.au3>
#include <YDLogger.au3>
#include <YDTool.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDLogger_Init()
Global $g_sAppName       = _YDTool_GetAppWrapperRes("ProductName")
Global $g_sAppDesc       = _YDTool_GetAppWrapperRes("Description")
Global $g_sAppVersion    = _YDTool_GetAppWrapperRes("FileVersion")
Global $g_sAppVersionV   = "v" & $g_sAppVersion
Global $g_sAppTitle      = $g_sAppName & " - " & $g_sAppVersionV
Global $g_sAppDataPath   = @ScriptDir & "\data"
Global $g_sAppStaticPath = @ScriptDir & "\static"
Global $g_sAppLogsPath   = @ScriptDir & "\logs"
Global $g_sAppVendorPath = @ScriptDir & "\vendor"
Global $g_sAppIconPath   = $g_sAppStaticPath & "\icon.ico"
Global $g_sPrinterName   = "TessiPOST_CPAM62"
Global $hPdf
; ===============================================================================================================================

; #VARIABLES DEBUG# =============================================================================================================
_YDTool_DebugGlobals()
_YDLogger_Var("$g_sAppName", $g_sAppName)
_YDLogger_Var("$g_sAppDesc", $g_sAppDesc)
_YDLogger_Var("$g_sAppVersion", $g_sAppVersion)
_YDLogger_Var("$g_sAppVersionV", $g_sAppVersionV)
_YDLogger_Var("$g_sAppTitle", $g_sAppTitle)
_YDLogger_Var("$g_sAppDataPath", $g_sAppDataPath)
_YDLogger_Var("$g_sAppStaticPath", $g_sAppStaticPath)
_YDLogger_Var("$g_sAppLogsPath", $g_sAppLogsPath)
_YDLogger_Var("$g_sAppVendorPath", $g_sAppVendorPath)
_YDLogger_Var("$g_sAppIconPath", $g_sAppIconPath)
_YDLogger_Var("$g_sPrinterName", $g_sPrinterName)
If _Singleton($g_sAppName, 1) = 0 Then
    MsgBox($MB_SYSTEMMODAL, "Warning", "L'application " & $g_sAppName & " est déjà en cours d'exécution !")
    Exit
EndIf
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
; On gere l'affichage de l'icone dans le tray
TraySetIcon($g_sAppIconPath)
TraySetToolTip($g_sAppTitle)
Global $idController = TrayCreateItem("Lancer le traitement (Ctrl+Alt+t)")
TrayCreateItem("")
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
_YDTool_SetTrayTip($g_sAppTitle, "L'application " & $g_sAppName & " est lancée en tâche de fond", 3, 1)

HotKeySet("^!t", "_Controller")
HotKeySet("^!T", "_Controller")
; ===============================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
    Switch TrayGetMsg()
        Case $idTrayExit
            ExitLoop
        Case $idController
            _Controller()
        Case $idTrayAbout
            _About()
    EndSwitch
    Sleep(50)
WEnd
; ===============================================================================================================================


; #FUNCTION# ====================================================================================================================
; Name...........: _Controller
; Description ...: Fonction principale de controle de l'action
; Syntax.........: _Controller()
; Parameters ....:
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================
Func _Controller()
    Local $sFuncName = "_Controller"
    Local $sKeyword = "_questionnaire"
    Local $sPdfFullPath = $g_sAppDataPath & "\" & $sKeyword & ".pdf"
    Local $sTxtFullPath = StringReplace($sPdfFullPath, ".pdf", ".txt")
    Local $aTxt, $aEmptyLines[0], $aAdr[0], $aNir[0]

    _YDLogger_Sep()
    _YDLogger_Log("Debut fonction : " & $sFuncName)
    _YDLogger_Var("$sPdfFullPath", $sPdfFullPath)
    _YDLogger_Var("$sTxtFullPath", $sTxtFullPath)

    ; On ne peut lancer l'application que si pdf actif
    _YDTool_SetTrayTip($g_sAppTitle, "Analyse du contexte ...")
    $hPdf = WinActivate("[CLASS:AcrobatSDIWindow]")
    If Not $hPdf Then
        Return _YDTool_SetMsgBoxError("Le fichier .pdf issu de QRP doit etre ouvert et actif !", $sFuncName)
    EndIf

    ; Le pdf doit comporter le mot "_questionnaire" dans son titre
    Local $sPdfTile = WinGetTitle($hPdf)
    _YDLogger_Var("$sPdfTile", $sPdfTile)
    If StringInStr($sPdfTile, $sKeyword) = 0 Then
        _YDLogger_Log("Titre du pdf = " & $sPdfTile)
        Return _YDTool_SetMsgBoxError('Le fichier .pdf doit avoir le mot-clé "' & $sKeyword & '" dans son titre !', $sFuncName)
    ; Le pdf "_questionnaire.pdf" ne doit pas etre deja ouvert
    ElseIf StringInStr($sPdfTile, $sKeyword & ".pdf") > 0 Then
        WinActivate($hPdf)
        Send("^w")
        Return _YDTool_SetMsgBoxError('Le fichier .pdf temporaire était déjà ouvert !', $sFuncName)
    EndIf

    ; On enregistre le pdf dans le dossier de travail
    _YDTool_SetTrayTip($g_sAppTitle, "Copie du pdf ...")
    _YDTool_CopyFile(_YDTool_GetAcrobatReaderLastRecentFile(), $sPdfFullPath)

    ; On convertit le pdf en txt
    _YDTool_SetTrayTip($g_sAppTitle, "Conversion du pdf ...")
    _YDTool_CreateFolderIfNotExist(@ScriptDir & "/vendor")
    FileInstall(".\vendor\pdftotext.exe", $g_sAppVendorPath & "\pdftotext.exe", 1)
    If _XPDF_ToText($sPdfFullPath, $sTxtFullPath, 1, 3, True) = 0 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la conversion !", $sFuncName)
    EndIf
    ; On traite les donnees dans un Array
    _YDTool_SetTrayTip($g_sAppTitle, "Analyse du pdf ...")
    _FileReadToArray($sTxtFullPath, $aTxt)
    For $i = 1 to UBound($aTxt) -1
        ; On recupere les lignes vides
        If $aTxt[$i] = "" Then
            _ArrayAdd($aEmptyLines, $i)
        EndIf
        ; On recupere la ligne avec le NIR
        Local $sNirLinePosition = StringInStr($aTxt[$i], " sociale ", 0)
        If $sNirLinePosition > 0 Then
            _ArrayAdd($aNir, $sNirLinePosition)
            _ArrayAdd($aNir, $aTxt[$i])
        Endif
    Next
    If UBound($aNir) < 2 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation du NIR !", $sFuncName)
    EndIf
    If UBound($aEmptyLines) < 2 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation des lignes vides !", $sFuncName)
    EndIf
    ; On recupere le NIR
    _YDLogger_Var("$aNir[0]", $aNir[0])
    _YDLogger_Var("$aNir[1]", $aNir[1])
    Local $sNir = StringMid($aNir[1], $aNir[0]+9, 15)
    _YDLogger_Var("$sNir (avant)", $sNir)
    If StringLen($sNir) <> 15 Then
        Return _YDTool_SetMsgBoxError("Un NIR doit faire 15 caracteres ! (" & StringLen($sNir) & ")", $sFuncName)
    EndIf
    $sNir = StringMid($sNir,1,1) & " " & StringMid($sNir,2,2) & " " & StringMid($sNir,4,2) & " " & StringMid($sNir,6,2) & " " & StringMid($sNir,8,3) & " " & StringMid($sNir,11,3) & " " & StringMid($sNir,14,2)
    _YDLogger_Var("$sNir (apres)", $sNir)
    ; On rempli le tableau $aAdr avec les lignes d adresse
    For $i = 1 to UBound($aEmptyLines) -1
        If $i > $aEmptyLines[0] And $i < $aEmptyLines[1] Then
            _ArrayAdd($aAdr, StringStripWS($aTxt[$i], 3))
        EndIf
    Next
    If UBound($aAdr) < 2 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation de l'adresse !", $sFuncName)
    EndIf
    ;_ArrayDisplay($aAdr)

    ; On lance l'impression via TessiPOST
    _YDTool_SetTrayTip($g_sAppTitle, "Lancement de l'impression via " & $g_sPrinterName & " ...")
    Send("^p")
    If WinWaitActive("Imprimer", "", 15) = 0 Then
        Return _YDTool_SetMsgBoxError("Impossible de lancer l'impression !", $sFuncName)
    EndIf
    ControlFocus("Imprimer", "", "[CLASS:ComboBox; INSTANCE:1]")
    Send("t")
    Sleep(1000)
    Send("{ENTER}")
    Local $hTessiPrinter = WinWaitActive($g_sPrinterName, "", 30)
    If $hTessiPrinter = 0 Then
        Return _YDTool_SetMsgBoxError("L'imprimante " & $g_sPrinterName & " ne semble pas accessible !", $sFuncName)
    EndIf
    _YDTool_SetTrayTip($g_sAppTitle, "Remplissage des donnees ...")
    ; On vide le pave d adresse
    Local $aAdrLines[0]
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:21]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:18]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:17]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:16]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:15]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:14]")
    _ArrayAdd($aAdrLines, "[CLASS:ThunderRT6TextBox; INSTANCE:13]")
    For $i = 0 To UBound($aAdrLines)-1
        _TESSI_DeleteTextInControl($hTessiPrinter, $aAdrLines[$i])
    Next
    ; On rempli le pave d adresse
    Local $j=0
    For $i = UBound($aAdrLines)-UBound($aAdr) To UBound($aAdrLines)-1
        ;_YDLogger_Log("$i=" & $i)
        ;_YDLogger_Log("$j=" & $j)
        ;_YDLogger_Log("$aAdr[$j]=" & $aAdr[$j])
        _TESSI_SetTextInControl($hTessiPrinter, $aAdrLines[$i], $aAdr[$j])
        $j = $j + 1
    Next
    ; On rempli le NIR
    _TESSI_SetTextInControl($hTessiPrinter, "[CLASS:ThunderRT6TextBox; INSTANCE:19]", $sNir)
    Sleep(1000)
    ; On coche le R1 (LRAR)
    ControlFocus($hTessiPrinter, "", "[CLASS:ThunderRT6OptionButton; INSTANCE:6]")
    If @error Then
        Return _YDTool_SetMsgBoxError("L'application " & $g_sAppName & " a rencontré un problème inconnu !", $sFuncName)
    Else
        _YDTool_SetTrayTip($g_sAppTitle, "Fin du traitement.", 3, 1)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _About
; Description ...: Fonction qui renvoi une GUI "A propos"
; Syntax.........: _About()
; Parameters ....:
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================
Func _About()
    Local $sFuncName = "_About"
    Local $iAboutWidth = 550
    Local $iAboutHeight = 160
    Local $iAboutOkButtonWidth = 30
    Local $font = "Verdana"
    _YDLogger_Var("$iAboutWidth", $iAboutWidth, $sFuncName)

    Local $hAboutGUI = GUICreate("A propos", $iAboutWidth, $iAboutHeight, -1, -1, BitOR($WS_POPUP,$WS_CAPTION))
    ; Titre
    GUISetFont(12, $iAboutWidth*2, 0, $font)
    GUICtrlCreateLabel($g_sAppName, 0, 0, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    ; Description + version
    GUISetFont(9, $iAboutWidth, 0, $font)
    GUICtrlCreateLabel($g_sAppDesc, 0, 40, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    GUICtrlCreateLabel($g_sAppVersionV, 0, 80, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    ; Bouton OK
    Local $idOkButton = GUICtrlCreateButton("OK", $iAboutWidth/2-$iAboutOkButtonWidth/2, 120, $iAboutOkButtonWidth, 25, BitOr($BS_MULTILINE,$BS_CENTER))
    ; Affichage GUI
    GUISetState(@SW_SHOW, $hAboutGUI)
    ; Loop GUI
    While 1
        If GUIGetMsg() = $idOkButton Then
            GUIDelete($hAboutGUI)
            ExitLoop
        EndIf
        Sleep(50)
    WEnd
    Return True
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _XPDF_ToText
; Description....: Converts a PDF file to plain  text.
; Syntax.........: _XPDF_ToText ( "PDFFile" , "TxtFile" [ , FirstPage [, LastPage [, Layout ]]] )
; Parameters.....: PDFFile    - PDF Input File.
;                  TxtFile    - Plain text file to convert to
;                  FirstPage  - First page to convert (default is 1)
;                  LastPage   - Last page to convert (default is last page of the document)
;                  Layout     - If true, maintains (as  best as possible) the original physical layout of the text
;                               If false, the behavior is to 'undo'  physical  layout  (columns, hyphenation, etc.)
;                                 and output the text in reading order.
;                               Default is True
; Return values..: Success - 1
;                  Failure - 0, and sets @error to :
;                   1 - PDF File not found
;                   2 - Unable to find the external program
; ===============================================================================================================================
Func _XPDF_ToText($sPDFFile, $sTXTFile, $iFirstPage = 1, $iLastPage = 0, $bLayout = True)
    Local $sXPDFToText = $g_sAppVendorPath & "\pdftotext.exe"
    Local $sOptions

    If NOT FileExists($sPDFFile) Then Return SetError(1, 0, 0)
    If NOT FileExists($sXPDFToText) Then Return SetError(2, 0, 0)

    If $iFirstPage <> 1 Then $sOptions &= " -f " & $iFirstPage
    If $iLastPage <> 0 Then $sOptions &= " -l " & $iLastPage
    If $bLayout = True Then $sOptions &= " -layout"

    Local $iReturn = ShellExecuteWait ( $sXPDFToText , $sOptions & ' "' & $sPDFFile & '" "' & $sTXTFile & '"', @ScriptDir, "", @SW_HIDE)
    If $iReturn = 0 Then Return 1

    Return 0

EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _TESSI_DeleteTextInControl
; Description ...: Permet de supprimer le contenu d'un controle de type zone de texte
; Syntax.........: _TESSI_DeleteTextInControl($_hTessiPrinter, $_sControlClassName)
; Parameters ....: $_hTessiPrinter       - Handle de la fenetre principale
;                  $_sControlClassName   - Nom de la classe du controle
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================
Func _TESSI_DeleteTextInControl($_hTessiPrinter, $_sControlClassName)
    ControlFocus($_hTessiPrinter, "", $_sControlClassName)
    Send("{HOME}{SHIFTDOWN}{END}{SHIFTUP}")
    Send("{DELETE}")
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _TESSI_SetTextInControl
; Description ...: Permet de remplir le contenu d'un controle de type zone de texte
; Syntax.........: _TESSI_SetTextInControl($_hTessiPrinter, $_sControlClassName, $_sText)
; Parameters ....: $_hTessiPrinter       - Handle de la fenetre principale
;                  $_sControlClassName   - Nom de la classe du controle
;                  $_sText               - Texte à placer dans la zone de texte
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================
Func _TESSI_SetTextInControl($_hTessiPrinter, $_sControlClassName, $_sText)
    ControlFocus($_hTessiPrinter, "", $_sControlClassName)
    ClipPut($_sText)
    Send("^v")
    Return True
EndFunc
