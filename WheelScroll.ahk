;
; Redirect Scrool Function  スクロール制御
;   ・加速対応
;   ・Word / Excel / VBE / 秀丸等の分割ペインも互換スクロールで操作可能
;
;   単体 / 組込み両対応  2008/05/25 (AutoHotkey 1.0.47.06)
;   組込み時 
;     #Include WheelScroll.ahk
;     Gosub,WheelInit             ;初期化 :AutoExecuteセクションに記述
;---------------------------------------------------------------------------
;   2009.06.12  マルチディスプレイ対策 (Thanks IKKIさん)
;   2009.07.22  秀丸v8β1   超暫定対応
;               IKKI氏の WheelAccel.ahk の加速モードを入れ込み
;               Excelスクロール時にアクティブにならないようにした
;               (とりあえず TrackWheel.ahkの旧バージョンから貰ってきた)


;+++++++++++++++++++++++++++++++++++++++++++++++++++++++
;   単体起動用
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++
#NoEnv                  ; 変数名を解釈するとき、環境変数を無視する
#SingleInstance FORCE   ; 複数プロセスで実行の禁止
#WinActivateForce       ; タスクバーアイコン点滅防止
#MaxHotkeysPerInterval 80 ;キーリピートエラー防止用

WheelAutoExecute:
    SendMode Input              ; 送信中にユーザー操作を後回しにする。
    Gosub,WheelInit
    Hotkey,^ESC, WheelExit     ;終了： [Ctrl]+[ESC]
return
WheelExit:
    exitapp
return


;+++++++++++++++++++++++++++++++++++++++++++++++++++++++
;   単体/組込み両用
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++
WheelInit:
;-------------------------------------------------------
;   初期化
;-------------------------------------------------------
    ;--- オプション ---
    DefaultScrollMode = 0           ;基本動作モード  0:WHELL 1:互換SCROLL

    ; IKKI氏の WheelAccel.ahk入れ込み 超暫定対応     2009.07.22
    ; (従来モードは需要なかったら削除するかも)
    AcclMode          = 1           ; 0:従来の加速モード 1:WheelAccel.ahk の加速モード

    ; AcclMode = 0 オプション
    AcclSpeed         = 1           ;加速時の倍率(0で加速OFF)
    AcclTOut          = 300         ;加速タイムアウト値(ms)
    ScrlCount         = 2           ;互換スクロール行数

    ; AcclMode = 1 オプション
	; ホイール加速◆改造版
	minThrottle      := 2           ; 最小加速率
	maxThrottle      := 7           ; 最大加速率
	minWheelSpeed    := 5           ; 最小加速率になるホイール回転速度 (ノッチ/秒)
	maxWheelSpeed    := 30          ; 最大加速率になるホイール回転速度 (ノッチ/秒)
	WA_Debug         := false       ; true にすると加速率とホイール回転速度が表示される

    ;ホイールで動かすコントロールのクラスリスト
    MouseWhellList =MozillaWindowClass

    ;互換モードで動かすコントロールのクラスリスト
    VScroolList =  MdiClient            ;MDI親 (MS-Accessなど)
                  ,VbaWindow            ;VisualBasicEditor
                  ,_WwB                 ;MS-Word(編集領域全体)
                  ,Excel7               ;MS-Excel
;;;;;                  ,OModule                ;MS-Access97   2008.05.20

    ;MDI事前アクティブ化リスト (ｱｸﾃｨﾌﾞ子ｳｨﾝﾄﾞｳのみﾊﾞｰがあるｱﾌﾟﾘなど)
    MdiActivateList = Excel7            ;MS-Excel

    ;--- 互換モード カスタム動作 ---
    ;無視リスト(バイパスして親コントロールを制御する)
    BypassCtlList =   ScrollBar         ;スクロールバー本体
                    , _WwG              ;MS-Word分割ペイン(一つ上の_WwBで制御)
                    , Static            ;秀丸v8β 暫定  2009.07.22

    ;兄弟スクロールバー : ｽｸﾛｰﾙﾊﾞｰが配下ではなく同列にあるｱﾌﾟﾘ
    BrotherScroolBarList = TkfInnerView.UnicodeClass    ;萌ディタ

    ;禁止リスト：ｽｸﾛｰﾙﾊﾝﾄﾞﾙが取れない時は、互換モードを使用しない
    NullShwndTabooList = Excel7         ;MS-Excel(クラッシュ対策)


    ;---- 横スクロール カスタム動作 ---
    ;横スクロール除外リスト
    HDisavledList = 

return

;==============================================
;     Hotkeys
;==============================================
WheelDown::     WheelRedirect()
WheelUp::       WheelRedirect()


;Shiftホイールで横スクロール
+WheelDown::    WheelRedirect(1)
+WheelUp::      WheelRedirect(1)

/* ※※※※※ Logicoolマウス用 設定サンプル ※※※※※※※※※※※※※※

;Logicoolマウスチルト1(uberOptionsで 左:F13 右:F14が割り当てられていると仮定)
; ただし押下解除情報は正しく取れないのでuser.xmlを手動にて編集し
; キーリピートを発生させる必要あり
F13::   WheelRedirect(1,0)
F14::   WheelRedirect(1,1)

;Logicoolマウスチルト2(SetPointで 左:F11 右:F12に割り当てられていると仮定)
F11::       SetTimer,TiltRepeatL,80
F11 up::    SetTimer,TiltRepeatL,OFF
F12::       SetTimer,TiltRepeatR,80
F12 up::    SetTimer,TiltRepeatR,OFF
TiltRepeatL:
    WheelRedirect(1,0)
return
TiltRepeatR:
    WheelRedirect(1,1)
return
;※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※
*/

;

;==============================================
;     Functions
;==============================================
WheelRedirect(mode=0,dir="")
;-------------------------------------------------------------
;   ホイールリダイレクト
;   mode 0:縦スクロール  1:横スクロール (省略時:縦)
;   dir  0:UP(LEFT)      1:DOWN(RIGHT)  (省略時:ホイール準拠)
;-------------------------------------------------------------
{
    global  DefaultScrollMode, AcclSpeed, AcclTOut, ScrlCount
           ,MouseWhellList, VScroolList, MdiActivateList
           ,BypassCtlList, NullShwndTabooList, HDisavledList

    CoordMode,Mouse,Screen
    MouseGetPos,mx,my,hwnd,ctrl,3
    WinGetClass,wcls, ahk_id %hwnd%
    SendMessage,0x84,0,% (my<<16)|mx,,ahk_id %ctrl% ;WM_NCHITTEST
    If (ErrorLevel = 0xFFFFFFFF)
        MouseGetPos,,,,ctrl,2
    ifEqual,ctrl,,  SetEnv,ctrl,%hwnd%              ;2008.05.25
    WinGetClass,ccls,ahk_id %ctrl%
    mccls := ccls                                   ;2009.07.22    秀丸v8β 対応

    ;無視リストチェック：1階層上のコントロールを制御対象とする
    ifInString, BypassCtlList, %ccls%
    {
        ctrl := DllCall("GetParent",UInt,ctrl, UInt)
        WinGetClass,ccls,ahk_id %ctrl%
    }

    ;MDI事前アクティブ化リストチェック : 非ｱｸﾃｨﾌﾞ子ｳｨﾝﾄﾞｳをｱｸﾃｨﾌﾞ化
    if ccls in %MdiActivateList%
    {
        MdiClient := DllCall("GetParent",UInt,ctrl, UInt)
        SendMessage, 0x229, 0,0,,ahk_id %MdiClient% ;WM_MDIGETACTIVE
        if (ctrl != ErrorLevel) {
            if(ccls = "Excel7")                    ;Excelカスタム
			        ControlClick,,ahk_id %ctrl%     ; (改)MID小窓をクリックして前面にならないようにした 2009.07.22
            Else    PostMessage,0x222, %ctrl%,0,,ahk_id %MdiClient%
        }
    }
    scnt := GetScrollBarHwnd(shwnd,mx,my,ctrl,mode,mccls) ;ｽｸﾛｰﾙﾊﾝﾄﾞﾙ取得 2009.07.22

    ;スクロール動作指定
    scmode := DefaultScrollMode<<1 | mode
    if ccls in %HDisavledList%          ;横スクロール禁止
        scmode &= 0x10
    if ccls in %MouseWhellList%         ;ホイールモード
        scmode &= 0x01
    if ccls in %VScroolList%            ;互換モード
        scmode |= 0x10
    if (!shwnd) {                       ;互換モード禁止リスト
        if ccls in %NullShwndTabooList%
            scmode  = 0
    }

    if (!scmode)
            MOUSEWHELL(ctrl,mx,my,dir,AcclSpeed,AcclTOut)
    Else    SCROLL(ctrl,mode,shwnd,dir,ScrlCount,AcclSpeed,AcclTOut)
}

GetScrollBarHwnd(byref shwnd, mx,my,Cntlhwnd,mode=0,mccls="")
;----------------------------------------------------------------------------
; 該当コントロールのスクロールハンドルを返す
;   戻り値 指定方向のスクロールオブジェクト数
;   out    shwnd       スクロールハンドル格納先
;   in     mx,my       マウス位置
;          Cntlhwnd    対象コントロールのハンドル
;          mode        0:VSCROLL(縦) 1:HSCROLL(横)
;          mccls       マウス直下のコントロール名称
;----------------------------------------------------------------------------
{
    global BrotherScroolBarList

    shwnd = 0
    WinGet,lst,ControlList,ahk_id %Cntlhwnd%
    WinGetClass,pcls, ahk_id %Cntlhwnd%

    ;配下にスクロールバーなし
    ifNotInString, lst, ScrollBar
    {    ;兄弟指定がある場合は、自分と同列のスクロールバーを探す
        if pcls in %BrotherScroolBarList%
        {
            Cntlhwnd := DllCall("GetParent",UInt,Cntlhwnd, UInt)
            WinGet,lst,ControlList,ahk_id %Cntlhwnd%
            WinGetClass,pcls, ahk_id %Cntlhwnd%
        }
        else return 0
    }

    ;スクロールバーコントロールの抽出
    vcnt = 0
    hcnt = 0
    Loop,Parse,lst,`n
    {
        ifNotInstring A_LoopField , ScrollBar
            Continue
        ControlGet,hwnd, Hwnd,,%A_LoopField%,ahk_id %Cntlhwnd%
        WinGetpos, sx,sy,sw,sh, ahk_id %hwnd%

        if (sw < sh)    {   ;縦スクロール
            vcnt++
            WinGetpos, vx%vcnt%,vy%vcnt%,vw%vcnt%,vh%vcnt%, ahk_id %hwnd%
            if (vi = "")
            || ((vy%vi%!=sy)&&((sy<my)&&(vy%vi%<sy))||((vy%vi%>my)&&(vy%vi%>sy))) ;上下分割
            || ((vx%vi%!=sx)&&((sx>mx)&&(vx%vi%>sx))||((vx%vi%<mx)&&(vx%vi%<sx))) ;左右分割
            {
                vi := vcnt
                if (mode = 0)   {
                    ret   := vcnt
                    shwnd := hwnd
                }
            }
        }
        if (sw > sh)    {   ;横スクロール
            hcnt++
            WinGetpos, hx%hcnt%,hy%hcnt%,hw%hcnt%,hh%hcnt%, ahk_id %hwnd%
            if (hi = "")
            || ((hx%hi%!=sx)&&((sx<mx)&&(hx%hi%<sx))||((hx%hi%>mx)&&(hx%hi%>sx)))           ;左右(Excel型)
            || ((hy%hi%!=sy)&&((sy+sh>my)&&(hy%hi%>sy))||((hy%hi%+hh%hi%<my)&&(hy%hi%<sy))) ;上下(Word型)
            {
                hi := hcnt
                if (mode = 1)   {
                    ret   := hcnt
                    shwnd := hwnd
                }
            }
        }
    }

    ; 2009.07.22 秀丸8β1 超暫定対応
    ;---アクティブペインにしかバーがないアプリ、可能ならペインを切り替える---
    ;[秀丸]用 カスタム：分割ウィンドウ切り替え 
    if  (pcls="HM32CLIENT" && !(vy1<=my && vy1+vh1 >= my))  ;秀丸 v7未満
     || (pcls="Hidemaru32Class" && mccls = "Static")         ;     v8β1
        PostMessage, 0x111, 142,  0, ,ahk_id %Cntlhwnd%   ;WM_COMMAND
    ;------------------------------------------------------------------------

    return ret
}

;------ PostMessage Scrool Control ------------------------------------------

MOUSEWHELL(hwnd,mx,my,dir="", ASpeed=1,ATOut=300)
;----------------------------------------------------------------------------
; WM_MOUSEWHELLによる任意コントロールスクロール
;       hwnd        該当コントロールのウィンドウハンドル
;       mx,my       マウス位置
;       dir         前後方向 0:UP 1:DOWN
;
;       ASpeed       :加速時の倍率(0で加速OFF)
;       ATOut        :加速タイムアウト値(ms)
;----------------------------------------------------------------------------
{
    ; IKKI氏の WheelAccel.ahk入れ込み 超暫定対応     2009.07.22
    global  AcclMode
    if (AcclMode)  {
        delta := 120 * WA_Throttle()
    }
    else {
        static delta

        ;ホイール加速
        If (A_PriorHotkey <> A_ThisHotkey) || (ATOut < A_TimeSincePriorHotkey) 
           || (0 >= ASpeed)
            delta = 120
        Else If (delta < 1000)
            delta += 120 * ASpeed
    }

    ; wParam: Delta(移動量)
    wpalam  :=GetKeyState("LButton")     | GetKeyState("RButton") <<1 
            | GetKeyState("Shift")   <<2 | GetKeyState("Ctrl")    <<3 
            | GetKeyState("MButton") <<4 | GetKeyState("XButton1")<<5
            | GetKeyState("XButton2")<<6

    If (dir = 0)
         wpalam |=   delta << 16        ;up
    Else if (dir = 1)
         wpalam |= -(delta << 16)       ;down
    Else ifInstring A_ThisHotkey, WheelUp
         wpalam |=   delta << 16        ;up
    Else wpalam |= -(delta << 16)       ;down

    ; lParam: XY座標
    my += (my < 0) ? 0xFFFF : 0  ;マルチディスプレイ対策 2009.06.12
    mx += (mx < 0) ? 0xFFFF : 0
    lpalam := (my << 16) | mx

    ;WM_MOUSEWHELL
    PostMessage, 0x20A, %wpalam%, %lpalam%, , ahk_id %hwnd%
}

SCROLL(hwnd,mode=0,shwnd=0,dir="", ScrlCnt=1,ASpeed=1,ATOut=300)
;----------------------------------------------------------
; WM_SCROLLによる任意コントロールスクロール
;       hwnd        該当コントロールのウィンドウハンドル
;       mode        0:VSCROLL(縦) 1:HSCROLL(横)
;       shwnd       スクロールバーのハンドル
;       dir         前後方向 0:SB_LINEUP/LEFT 1:SB_LINEDOWN/RIGHT
;
;       ScrlCnt      :スクロール行数
;       ASpeed       :加速時の倍率(0で加速OFF)
;       ATOut        :加速タイムアウト値(ms)
;----------------------------------------------------------
{
    static ACount

    ;加速
    If (A_PriorHotkey <> A_ThisHotkey) || (ATOut < A_TimeSincePriorHotkey) 
       || (0 >= ASpeed)
        ACount := ScrlCnt
    Else
        ACount += ScrlCnt * ASpeed

    ;wParam: 方向
    if (dir = "")
    {
        ifInstring A_ThisHotkey, WheelUp
             dir = 0                        ;SB_LINEUP   / SB_LINELEFT
        Else dir = 1                        ;SB_LINEDOWN / SB_LINERIGHT
    }
    
    ;0x114:WM_HSCROLL  0x115:WM_VSCROLL
    msg := 0x115 - mode

    Loop, %ACount%
        PostMessage, %msg%, %dir%, %shwnd%, , ahk_id %hwnd%
    PostMessage, %msg%, 8, %shwnd%, , ahk_id %hwnd% ;SB_ENDSCROLL
}

WA_Throttle() {
;----------------------------------------------------------
; 加速率を線形補間で計算する
; 	minThrottle   = 最小加速率
; 	maxThrottle   = 最大加速率
; 	minWheelSpeed = 最小加速率になるホイール回転速度 (ノッチ/秒)
; 	maxWheelSpeed = 最大加速率になるホイール回転速度 (ノッチ/秒)
; 	WA_Debug      = デバッグモード
;----------------------------------------------------------
	global minThrottle, maxThrottle, minWheelSpeed, maxWheelSpeed, WA_Debug, tooltiptext
	static prevspd := 0
	if (A_PriorHotkey <> A_ThisHotkey || A_TimeSincePriorHotkey <= 0) {
		gosub WA_EraseToolTip
		prevspd := 0
		nextspd := 0
	} else {
		nextspd := 1000 / A_TimeSincePriorHotkey ; 現在のホイール回転速度 (ノッチ/秒)
	}
	spd := (prevspd + nextspd) / 2 ; 直近 2 ノッチの平均回転速度 (ノッチ/秒)
	if (spd < minWheelSpeed) {
		thr := 1
	} else {
		thr := floor((spd - minWheelSpeed) * (maxThrottle - minThrottle) / (maxWheelSpeed - minWheelSpeed) + minThrottle)
	}
	if (thr > maxThrottle) {
		thr := maxThrottle
	}
	
	if (WA_Debug) {
		tooltiptext .= "x" . thr . " (" . round(spd, 1)
; 		tooltiptext .= " = avg(" . round(nextspd, 1) . " + " . round(prevspd, 1) . ")"
		tooltiptext .= " notch/s)`n"
		tooltip % tooltiptext
		settimer WA_EraseToolTip, 10000
	}
	prevspd := nextspd
	return thr
}

WA_EraseToolTip:
;----------------------------------------------------------
; ツールチップを消す
;----------------------------------------------------------
	tooltiptext := ""
	tooltip
	settimer WA_EraseToolTip, off
	return

;----------------------------------------------------------
; <参考> ホイール加速の別実装
; http://f57.aaa.livedoor.jp/~atechs/index.php?plugin=attach&pcmd=open&file=AutoHotKey%20Thread.htm&refer=Download
; 538 ：233：2005/05/09(月) 01:41:23 ID:zU71pxGA
;     WheelUp::
;     WheelDown::
;     　MouseGetPos,x,y,hwnd,cls
;     　MouseGetPos,,,,cls2,1
;     　if(cls != cls2)
;     　　cls := cls2
;     　accel := (A_PriorHotkey == A_ThisHotkey && A_TimeSincePriorHotkey < 80) + (A_PriorHotkey == A_ThisHotkey && A_TimeSincePriorHotkey < 250) + 1
;     　wParam := 0x780000 * accel * (1 - 2 *(A_ThisHotkey = "WheelDown"))
;     　lParam := x + y*0x10000
;     　PostMessage,0x20A, %wParam%,%lParam%, %cls%, ahk_id %hwnd%
;     　return
;     ホイールリダイレクト。例によって加速付き。
;     だいぶ短くなった。今のところMDIを含め殆ど動ようになった。
;     W2kSP4, AHK1.0.32.00
;----------------------------------------------------------
