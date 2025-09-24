#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=o.ico
#AutoIt3Wrapper_Res_Description=ToolCaiOffice_v1.1 ; Tên hiển thị trong Task Manager
#AutoIt3Wrapper_Outfile=ToolCaiOffice_v1.1.exe ; Tên file đầu ra (.exe) cho ứng dụng
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
#AutoIt3Wrapper_Res_Companyname=Copyright@lqviet_24.09.2025
#AutoIt3Wrapper_Res_Language=1066 ; Vietnamese
#AutoIt3Wrapper_Run_Obfuscator=y  ; Sử dụng bộ làm rối mã nguồn
#AutoIt3Wrapper_UseUpx=y   ; Sử dụng công cụ UPX để nén file .exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

_CHECK_WINDOWS()

;Lịch sử ODT (chỉ hỗ trợ Win 10 và Win 11):
;https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18925-20138.exe
;https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19231-20072.exe

Global $SODT_URL = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19231-20072.exe"
Global $SODT_EXE    = @ScriptDir & "\odt.exe"
Global $SCONFIGFILE = @ScriptDir & "\config.xml"

; ==== Danh sách ứng dụng ====
Global $AOFFICEAPPS[7] = ["Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote"]
Global $ACHKAPP[UBound($AOFFICEAPPS)]

; ==================== GUI ====================
Func CaiOffice()
	#RequireAdmin
    Local $HGUI = GUICreate("Cài đặt Office Online cho Win 10 và Win 11 v1.1", 520, 400)

    ; ==== Label giới thiệu tác giả ====
    Local $LBLAUTHOR = GUICtrlCreateLabel("Phát triển bởi Lê Quốc Việt", 20, 10, 480, 20)
    GUICtrlSetFont($LBLAUTHOR, 10, 2, 0)
    GUICtrlSetColor($LBLAUTHOR, 0xFF0000)

    ; ==== Chọn phiên bản Office ====
    GUICtrlCreateGroup("📌 Phiên bản Office", 20, 35, 480, 50)
    Local $OPT2019 = GUICtrlCreateRadio("2019", 40, 55, 80, 18)
    Local $OPT2021 = GUICtrlCreateRadio("2021", 140, 55, 80, 18)
    Local $OPT2024 = GUICtrlCreateRadio("2024", 240, 55, 80, 18)
    Local $OPT365  = GUICtrlCreateRadio("365", 340, 55, 80, 18)
    GUICtrlSetState($OPT2024, $GUI_CHECKED)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; ==== Kiến trúc ====
    GUICtrlCreateGroup("⚙️ Kiến trúc", 20, 90, 480, 45)
    Local $OPT64BIT = GUICtrlCreateRadio("64 bit", 60, 110, 100, 18)
    Local $OPT32BIT = GUICtrlCreateRadio("32 bit", 200, 110, 100, 18)
    GUICtrlSetState($OPT64BIT, $GUI_CHECKED)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; ==== Ứng dụng ====
    GUICtrlCreateGroup("✅ Ứng dụng", 20, 140, 480, 100)
    Local $ILEFT1 = 40, $ILEFT2 = 260, $ITOP = 160
    For $I = 0 To UBound($AOFFICEAPPS) - 1
        If Mod($I, 2) = 0 Then
            $ACHKAPP[$I] = GUICtrlCreateCheckbox($AOFFICEAPPS[$I], $ILEFT1, $ITOP, 200, 18)
        Else
            $ACHKAPP[$I] = GUICtrlCreateCheckbox($AOFFICEAPPS[$I], $ILEFT2, $ITOP, 200, 18)
            $ITOP += 22
        EndIf
        GUICtrlSetFont(-1, 9, 400)

        ; Tick mặc định Word, Excel, PowerPoint, Access
        If $AOFFICEAPPS[$I] = "Word" Or _
           $AOFFICEAPPS[$I] = "Excel" Or _
           $AOFFICEAPPS[$I] = "PowerPoint" Or _
           $AOFFICEAPPS[$I] = "Access" Then
            GUICtrlSetState(-1, $GUI_CHECKED)
        EndIf
    Next
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; ==== Sản phẩm riêng ====
    GUICtrlCreateGroup("📇 Sản phẩm riêng", 20, 245, 480, 40)
    Local $CHKPROJECT = GUICtrlCreateCheckbox("Project", 60, 265, 150, 18)
    Local $CHKVISIO   = GUICtrlCreateCheckbox("Visio", 260, 265, 150, 18)
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    ; ==== Ngôn ngữ ====
    Local $LBLLANG = GUICtrlCreateLabel("🌍 Ngôn ngữ:", 30, 295, 100, 20)
    GUICtrlSetFont($LBLLANG, 9, 400)
    Local $CMBLANGUAGE = GUICtrlCreateCombo("", 120, 293, 160, 22)
    GUICtrlSetData($CMBLANGUAGE, "en-us|en-gb|vi-vn|fr-fr|de-de|ja-jp", "en-us")

    ; ==== Nút ====
    Local $BTNSELECTALL   = GUICtrlCreateButton("📝 Chọn tất cả", 40, 330, 120, 30)
    Local $BTNINSTALL     = GUICtrlCreateButton("▶ Cài đặt", 200, 330, 120, 30)
	GUICtrlSetBkColor($BTNINSTALL, 0xFFA500) ; Cam
    Local $BTNSHORTCUTS   = GUICtrlCreateButton("🖥️ Tạo shortcut", 360, 330, 120, 30)

    ; ==== Trạng thái ====
    Local $LBLSTATUS = GUICtrlCreateLabel("", 20, 370, 480, 20)
    GUICtrlSetFont($LBLSTATUS, 9, 400, 2)
    GUICtrlSetColor($LBLSTATUS, 0x009999)

    GUISetState(@SW_SHOW)

    ; ==== Vòng lặp GUI ====
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                Exit
            Case $BTNSELECTALL
                For $I = 0 To UBound($ACHKAPP) - 1
                    GUICtrlSetState($ACHKAPP[$I], $GUI_CHECKED)
                Next
            Case $BTNINSTALL
                GUICtrlSetData($LBLSTATUS, "🔧 Đang tạo file config.xml...")
                Local $SCFG = _MAKECONFIG($OPT2019, $OPT2021, $OPT2024, $OPT365, $OPT64BIT, $CMBLANGUAGE, $ACHKAPP, $CHKPROJECT, $CHKVISIO)
                Local $RES = _DEPLOY_SETUP($SCFG, $LBLSTATUS)
                If $RES = 0 Then
                    GUICtrlSetData($LBLSTATUS, "✅ Hoàn tất cài đặt!")
                Else
                    GUICtrlSetData($LBLSTATUS, "❌ Cài đặt thất bại. Mã: " & $RES)
                EndIf
            Case $BTNSHORTCUTS
                _CREATESHORTCUTS()
                GUICtrlSetData($LBLSTATUS, "🖥️ Đã tạo shortcut trên Desktop.")
        EndSwitch
    WEnd
EndFunc

; ==================== MAIN ====================
CaiOffice()

; ==================== HÀM HỖ TRỢ ====================

Func _URLDOWNLOAD($SURL, $SDEST)
	#RequireAdmin
    InetGet($SURL, $SDEST, $INET_FORCERELOAD, $INET_LOCALCACHE)
    If @error = 0 And FileExists($SDEST) And FileGetSize($SDEST) > 200000 Then Return True
    If FileExists($SDEST) Then FileDelete($SDEST)

    Local $PS = 'powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -Uri ''' & $SURL & ''' -OutFile ''' & $SDEST & ''' -UseBasicParsing -ErrorAction Stop } catch { exit 1 }"'
    Local $EXIT = RunWait(@ComSpec & " /c " & $PS, "", @SW_HIDE)
    If $EXIT = 0 And FileExists($SDEST) And FileGetSize($SDEST) > 200000 Then Return True
    If FileExists($SDEST) Then FileDelete($SDEST)

    $PS = 'powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Start-BitsTransfer -Source ''' & $SURL & ''' -Destination ''' & $SDEST & ''' -ErrorAction Stop } catch { exit 1 }"'
    $EXIT = RunWait(@ComSpec & " /c " & $PS, "", @SW_HIDE)
    If $EXIT = 0 And FileExists($SDEST) And FileGetSize($SDEST) > 200000 Then Return True
    If FileExists($SDEST) Then FileDelete($SDEST)

    $EXIT = RunWait(@ComSpec & ' /c curl.exe -L -o "' & $SDEST & '" "' & $SURL & '"', "", @SW_HIDE)
    If $EXIT = 0 And FileExists($SDEST) And FileGetSize($SDEST) > 200000 Then Return True
    If FileExists($SDEST) Then FileDelete($SDEST)

    Return False
EndFunc

; ==================== HÀM GIẢI NÉN ODT ====================
Func _EXTRACTODT($SEXEPATH, $SOUTDIR)
	#RequireAdmin
    If Not FileExists($SOUTDIR) Then DirCreate($SOUTDIR)

    ;ConsoleWrite("[EXTRACT] Bắt đầu thử các lệnh giải nén..." & @CRLF)

    ; Lấy tên file thực thi từ đường dẫn
    ;Local $sFileName = StringRegExpReplace($SEXEPATH, ".*\\(.*)", "$1")

    ; Thử Lệnh 1: /extract:"path"
    ;ConsoleWrite("[EXTRACT] Thử lệnh 1: /extract:..." & @CRLF)
    ;Local $sCmd1 = '"' & $SEXEPATH & '" /extract:"' & $SOUTDIR & '"'
    ;Run(@ComSpec & ' /c "' & $sCmd1 & '"', $SOUTDIR, @SW_HIDE)

    ; Thay RunWait bằng ProcessWaitClose để chờ tiến trình kết thúc
    ;ProcessWaitClose($sFileName)

    ;If FileExists($SOUTDIR & "\setup.exe") Then
    ;    ConsoleWrite("[EXTRACT] ✅ Lệnh 1 thành công!" & @CRLF)
    ;    Return True
    ;EndIf

    ; Thử Lệnh 2: /quiet /extract:"path"
    ;ConsoleWrite("[EXTRACT] Thử lệnh 2: /quiet /extract:..." & @CRLF)
    ;Local $sCmd2 = '"' & $SEXEPATH & '" /quiet /extract:"' & $SOUTDIR & '"'
    ;Run(@ComSpec & ' /c "' & $sCmd2 & '"', $SOUTDIR, @SW_HIDE)
    ;ProcessWaitClose($sFileName)
    ;If FileExists($SOUTDIR & "\setup.exe") Then
    ;    ConsoleWrite("[EXTRACT] ✅ Lệnh 2 thành công!" & @CRLF)
    ;    Return True
    ;EndIf

    ; Thử Lệnh 3: /extract:"path" /quiet
    Local $sCmd3 = '"' & $SEXEPATH & '" /extract:"' & $SOUTDIR & '" /quiet'
    ;  Chạy trực tiếp mà không thông qua CMD
    Local $iResult = RunWait($sCmd3, $SOUTDIR, @SW_HIDE)
	;ProcessWaitClose($iResult)
    If FileExists($SOUTDIR & "\setup.exe") Then
		; Xóa file odt.exe
        FileDelete($SEXEPATH)
		_REMOVEXML()
        Return True
    Else
		MsgBox(16, "Lỗi giải nén", _
        "❌ Không thể giải nén tự động." & @CRLF & _
        "Vui lòng tự giải nén file:" & @CRLF & '"' & $SEXEPATH & '"' & @CRLF & _
        "ngay trong thư mục Tool, sau đó nhấn lại nút Cài đặt.")
    Return False
	EndIf

    ; Thử Lệnh 4: /passive /extract:"path"
    ;ConsoleWrite("[EXTRACT] Thử lệnh 4: /passive /extract:..." & @CRLF)
    ;Local $sCmd4 = '"' & $SEXEPATH & '" /passive /extract:"' & $SOUTDIR & '"'
    ;Run(@ComSpec & ' /c "' & $sCmd4 & '"', $SOUTDIR, @SW_HIDE)
    ;ProcessWaitClose($sFileName)
	;If FileExists($SOUTDIR & "\setup.exe") Then
    ;    ConsoleWrite("[EXTRACT] ✅ Lệnh 4 thành công!" & @CRLF)
    ;    Return True
    ;EndIf
EndFunc

; ==================== TRIỂN KHAI ODT ====================
Func _DEPLOY_SETUP($SCONFIGFILE, $LBLSTATUS)
	#RequireAdmin
    Local $LOCAL_SETUP = @ScriptDir & "\setup.exe"

    ; Nếu có setup.exe sẵn
    If FileExists($LOCAL_SETUP) Then
        GUICtrlSetData($LBLSTATUS, "Đang cài đặt bằng setup.exe có sẵn...")
		Local $IRET = RunWait('"' & $LOCAL_SETUP & '" /configure "' & $SCONFIGFILE & '"', "", @SW_HIDE)
		; Sau khi cài đặt xong thì xoá config.xml
        _REMOVEXML()
        Return $IRET
    EndIf

    ; Nếu chưa có thì tải odt.exe
    Local $SODT_EXE = @ScriptDir & "\odt.exe"
    GUICtrlSetData($LBLSTATUS, "Đang tải Office Deployment Tool...")
    If Not _URLDOWNLOAD($SODT_URL, $SODT_EXE) Then
        GUICtrlSetData($LBLSTATUS, "❌ Lỗi tải ODT!")
        Return -1
    EndIf

    ; Giải nén trực tiếp vào thư mục tool
    GUICtrlSetData($LBLSTATUS, "Đang giải nén ODT...")
    If Not _EXTRACTODT($SODT_EXE, @ScriptDir) Then
        GUICtrlSetData($LBLSTATUS, "❌ Lỗi giải nén ODT!")
        Return -2
    EndIf

    Local $SETUP = @ScriptDir & "\setup.exe"
    If Not FileExists($SETUP) Then
        GUICtrlSetData($LBLSTATUS, "❌ Không tìm thấy setup.exe!")
        Return -3
    EndIf

    ; Cài đặt
    GUICtrlSetData($LBLSTATUS, "Đang cài đặt, vui lòng đợi...")
    Local $IRET = RunWait('"' & $SETUP & '" /configure "' & $SCONFIGFILE & '"', "", @SW_HIDE)
    ; Sau khi cài đặt xong thì xoá config.xml
    _REMOVEXML()
    Return $IRET
EndFunc

; ==================== TẠO FILE CONFIG ====================
Func _MAKECONFIG($OPT2019, $OPT2021, $OPT2024, $OPT365, $OPT64BIT, $CMBLANGUAGE, $ACHKAPP, $CHKPROJECT, $CHKVISIO)
    Local $ARCH = GUICtrlRead($OPT64BIT) = $GUI_CHECKED ? "64" : "32"
    Local $LANG = GUICtrlRead($CMBLANGUAGE)

    Local $PRODUCT = "ProPlus2021Retail"
    If GUICtrlRead($OPT2019) = $GUI_CHECKED Then $PRODUCT = "ProPlus2019Retail"
    If GUICtrlRead($OPT2024) = $GUI_CHECKED Then $PRODUCT = "ProPlus2024Retail"
    If GUICtrlRead($OPT365)  = $GUI_CHECKED Then $PRODUCT = "O365ProPlusRetail"

    Local $XML = '<?xml version="1.0" encoding="UTF-8"?>' & @CRLF & _
                 '<Configuration>' & @CRLF & _
                 '  <Add OfficeClientEdition="' & $ARCH & '" Channel="Current">' & @CRLF & _
                 '    <Product ID="' & $PRODUCT & '">' & @CRLF & _
                 '      <Language ID="' & $LANG & '" />' & @CRLF

    For $I = 0 To UBound($AOFFICEAPPS) - 1
        If GUICtrlRead($ACHKAPP[$I]) <> $GUI_CHECKED Then
            $XML &= '      <ExcludeApp ID="' & $AOFFICEAPPS[$I] & '" />' & @CRLF
        EndIf
    Next

    $XML &= '    </Product>' & @CRLF

    If GUICtrlRead($CHKPROJECT) = $GUI_CHECKED Then
        $XML &= '    <Product ID="ProjectPro2019Retail"><Language ID="' & $LANG & '"/></Product>' & @CRLF
    EndIf
    If GUICtrlRead($CHKVISIO) = $GUI_CHECKED Then
        $XML &= '    <Product ID="VisioPro2019Retail"><Language ID="' & $LANG & '"/></Product>' & @CRLF
    EndIf

    $XML &= '  </Add>' & @CRLF & '</Configuration>'

    ; ==== Ghi config.xml ngay tại thư mục Tool ====
    Local $SCONFIG = @ScriptDir & "\config.xml"
    Local $H = FileOpen($SCONFIG, 2)
    FileWrite($H, $XML)
    FileClose($H)

    Return $SCONFIG
EndFunc

; ==== Kiểm tra phiên bản Windows ====
Func _CHECK_WINDOWS()
    Local $OSVERSION = @OSVersion
    Local $OSBUILD   = @OSBuild

    ; Chỉ cho phép chạy trên Windows 10 hoặc Windows 11
    If Not StringInStr($OSVERSION, "WIN_10") And Not StringInStr($OSVERSION, "WIN_11") Then
        MsgBox(16, "⚠️ Không hỗ trợ", _
            "CẢNH BÁO: Công cụ này chỉ hỗ trợ cài Office Online trên Windows 10 và Windows 11." & @CRLF & _
            "Bạn đang dùng: " & $OSVERSION & " (Build " & $OSBUILD & ")." & @CRLF & @CRLF & _
            "👉 Vui lòng tự tải và cài bản Office phù hợp với Windows của bạn.")
        Exit
    EndIf
EndFunc

; ==== Hàm tạo shortcut ====
Func _CREATESHORTCUTS()
    Local $SCMD = 'cmd /c COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\*.lnk" "%AllUsersProfile%\Desktop"'
    RunWait($SCMD, "", @SW_HIDE)
EndFunc

Func _REMOVEXML()
	; Xóa file .xml trong thư mục Tool
	Local $sFile
	; Bắt đầu tìm kiếm file .xml trong thư mục script
	$sSearch = FileFindFirstFile(@ScriptDir & "\*.xml")
	; Kiểm tra xem có tìm thấy file nào không
	If $sSearch = -1 Then
		; Không tìm thấy
		FileClose($sSearch)
	Else
		; Vòng lặp để duyệt và xóa từng file
		Do
			$sFile = FileFindNextFile($sSearch)
			If @error Then ExitLoop ; Thoát vòng lặp khi không còn file nào nữa
			FileDelete(@ScriptDir & "\" & $sFile)
			ConsoleWrite("Đã xóa: " & $sFile & @CRLF)
		Until 0
		; Đóng handle tìm kiếm
		FileClose($sSearch)
	EndIf
EndFunc
