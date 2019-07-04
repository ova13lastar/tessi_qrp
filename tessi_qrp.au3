; #INDEX# =======================================================================================================================
; Title .........: tessi_qrp
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
#AutoIt3Wrapper_Res_ProductName=tessi_qrp
#AutoIt3Wrapper_Res_Description=Outil de remplissage automatique du driver TessiPOST à partir d'un courrier QRP
#AutoIt3Wrapper_Res_ProductVersion=1.0.5
#AutoIt3Wrapper_Res_FileVersion=1.0.5
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
; Includes YD
#include "D:\Autoit_dev\Include\YDGVars.au3"
#include "D:\Autoit_dev\Include\YDLogger.au3"
#include "D:\Autoit_dev\Include\YDTool.au3"
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("FileVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)
_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On gere l'affichage de l'icone dans le tray
TraySetIcon(_YDGVars_Get("sAppIconPath"))
TraySetToolTip(_YDGVars_Get("sAppTitle"))
Global $idController = TrayCreateItem("Lancer le traitement (Ctrl+Alt+t)")
TrayCreateItem("")
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "L'application " & _YDGVars_Get("sAppName") & " est lancée en tâche de fond", 3, 1)
;------------------------------
Global $g_hPdf
;------------------------------
; On recupere les valeurs de conf.ini
Global $g_sPrinterName = _YDTool_GetAppConfValue("general", "printer")
;------------------------------
; On definit les raccourcis clavier
HotKeySet("^!t", "_Controller")
HotKeySet("^!T", "_Controller")
; ===============================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
    Global $iMsg = TrayGetMsg()
    Select
		Case $iMsg = $idTrayExit
			_YDTool_ExitConfirm()
		Case $iMsg = $idTrayAbout
			_YDTool_GUIShowAbout()
		Case  $iMsg = $idController
			_Controller()
	EndSelect
    ;------------------------------
    Sleep(10)
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
    Local $sPdfFullPath = _YDGVars_Get("sAppDirDataPath") & "\" & $sKeyword & ".pdf"
    Local $sTxtFullPath = StringReplace($sPdfFullPath, ".pdf", ".txt")
    Local $aTxt, $aAdr[0], $aNir[0], $sNir

    ; On ne peut lancer l'application que si pdf actif
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Analyse du contexte ...")
    $g_hPdf = WinActivate("[CLASS:AcrobatSDIWindow]")
    If Not $g_hPdf Then
        Return _YDTool_SetMsgBoxError("Le fichier .pdf issu de QRP doit etre ouvert et actif !", $sFuncName)
    EndIf

    ; Le pdf doit comporter le mot "_questionnaire" dans son titre
    Local $sPdfTitle = WinGetTitle($g_hPdf)
    _YDLogger_Var("$sPdfTitle", $sPdfTitle, $sFuncName, 2)
    If StringInStr($sPdfTitle, $sKeyword) = 0 Then
        Return _YDTool_SetMsgBoxError('Le fichier .pdf doit avoir le mot-clé "' & $sKeyword & '" dans son titre !', $sFuncName)
    ; Le pdf "_questionnaire.pdf" ne doit pas etre deja ouvert
    ElseIf StringInStr($sPdfTitle, $sKeyword & ".pdf") > 0 Then
        WinActivate($g_hPdf)
        Send("^w")
        Return _YDTool_SetMsgBoxError('Le fichier .pdf temporaire était déjà ouvert !', $sFuncName)
    EndIf

    ; On recupere le chemin du fichier pdf
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Récuperation du chemin du pdf ...")
    WinActivate($g_hPdf)
    Send("^d")
    If WinWaitActive("Propriétés du document", "", 5) = 0 Then
         _YDLogger_Error("Fenêtre <Propriétés du document> non trouvee !", $sFuncName)
    EndIf
    Send("^d")
    Local $sPdfTempPath = ControlGetText("Propriétés du document", "", "[CLASS:Static; INSTANCE:18]")
    _YDLogger_Var("$sPdfTempPath (avant)", $sPdfTempPath, $sFuncName, 2)
    ControlClick("Propriétés du document", "", "[CLASS:Button; INSTANCE:5]")    
    If $sPdfTempPath = "" Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation du chemin du pdf !", $sFuncName)
    EndIf
    $sPdfTempPath = $sPdfTempPath & StringLeft($sPdfTitle, StringInStr($sPdfTitle, ".pdf") + 4)
    _YDLogger_Var("$sPdfTempPath (apres)", $sPdfTempPath, $sFuncName, 2)
    If $sPdfTempPath = $sPdfFullPath Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation du chemin du pdf ($sPdfTempPath = $sPdfFullPath) !", $sFuncName)
    EndIf
    
    ; On enregistre le pdf dans le dossier de travail
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Copie locale du pdf ...")
    If Not _YDTool_CopyFile($sPdfTempPath, $sPdfFullPath) Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la copie locale du pdf !", $sFuncName)
    EndIf

    ; On convertit le pdf en txt
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Conversion du pdf ...")
    _YDTool_CreateFolderIfNotExist(_YDGVars_Get("sAppDirVendorPath"))
    FileInstall(".\vendor\pdftotext.exe", _YDGVars_Get("sAppDirVendorPath") & "\pdftotext.exe", 1)
    If _XPDF_ToText($sPdfFullPath, $sTxtFullPath, 1, 3, True) = 0 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la conversion !", $sFuncName)
    EndIf
    ; On traite les donnees dans un Array
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Analyse du pdf ...")
    _FileReadToArray($sTxtFullPath, $aTxt)
    Local $bAdrEnd = False
    Local $bTemoin = False
    For $i = 1 to UBound($aTxt) -1        
        ; On recupere l'adresse dans un tableau        
        If $bAdrEnd = False Then 
            _ArrayAdd($aAdr, StringStripWS($aTxt[$i], 3))
            If $aTxt[$i] = "" Then
                $bAdrEnd = True
            EndIf 
        EndIf
        ; On recupere la ligne avec le NIR
        Local $sNirLinePosition = StringInStr($aTxt[$i], " sécurité sociale ", 0)
        If $sNirLinePosition > 0 Then
            _ArrayAdd($aNir, $sNirLinePosition)
            _ArrayAdd($aNir, $aTxt[$i])
            ExitLoop
        Endif
        ; On verifie qu'il ne s'agit pas d'un questionnaire temoin (car pas de NIR)
        If StringInStr($aTxt[$i], "QUESTIONNAIRE TÉMOIN", 0) > 0 Then
            $bTemoin = True
        EndIf
    Next
    ;~ _ArrayDisplay($aNir)
    ;~ _ArrayDisplay($aAdr)
    If UBound($aAdr) < 2 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation de l'adresse !", $sFuncName)
    EndIf    
    ; On verifie si temoin et si NIR valide
    If $bTemoin = True Then
        $sNir = ""
    Elseif UBound($aNir) < 2 Then
        Return _YDTool_SetMsgBoxError("Une erreur est survenue lors de la recuperation du NIR !", $sFuncName)
    Else
        _YDLogger_Var("$aNir[0]", $aNir[0], $sFuncName, 2)
        _YDLogger_Var("$aNir[1]", $aNir[1], $sFuncName, 2)
        $sNir = StringStripWS(StringRegExpReplace($aNir[1], "[^[:digit:]]", ""), 3)
        _YDLogger_Var("$sNir (avant)", $sNir, $sFuncName, 2)
        If StringLen($sNir) <> 15 Then
            Return _YDTool_SetMsgBoxError("Un NIR doit faire 15 caracteres ! (" & StringLen($sNir) & ")", $sFuncName)
        EndIf
        $sNir = StringMid($sNir,1,1) & " " & StringMid($sNir,2,2) & " " & StringMid($sNir,4,2) & " " & StringMid($sNir,6,2) & " " & StringMid($sNir,8,3) & " " & StringMid($sNir,11,3) & " " & StringMid($sNir,14,2)
        _YDLogger_Var("$sNir (apres)", $sNir, $sFuncName, 2)
    EndIf
    
    ; On lance l'impression via TessiPOST
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Lancement de l'impression via " & $g_sPrinterName & " ...")
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
    _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Remplissage des donnees ...")
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
    _TESSI_DeleteTextInControl($hTessiPrinter, "[CLASS:ThunderRT6TextBox; INSTANCE:19]")
    _TESSI_SetTextInControl($hTessiPrinter, "[CLASS:ThunderRT6TextBox; INSTANCE:19]", $sNir)
    Sleep(1000)
    ; On change le FDP sur TESSI
    ControlCommand($hTessiPrinter, "", "[CLASS:ThunderRT6ComboBox; INSTANCE:5]", "SelectString", "FDP_CPAM624")
    ControlCommand($hTessiPrinter, "", "[CLASS:ThunderRT6ListBox; INSTANCE:3]", "SelectString", "Première page")    
    Sleep(1000)
    ; On coche le R1 (LRAR)
    ControlFocus($hTessiPrinter, "", "[CLASS:ThunderRT6OptionButton; INSTANCE:6]")
    If @error Then
        Return _YDTool_SetMsgBoxError("L'application " & _YDGVars_Get("sAppName") & " a rencontré un problème inconnu !", $sFuncName)
    Else
        _YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fin du traitement.", 3, 1)
        Return True
    EndIf
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
    Local $sXPDFToText = _YDGVars_Get("sAppDirVendorPath") & "\pdftotext.exe"
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
