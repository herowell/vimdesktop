/*----------------------------------------------------------------------------
  IME 制御用 Functuion (コピペ用)
   動作確認環境  2005/09/12
     WinNT4 SP6 / WinXP SP2 (98系でもおそらくはOK)
    IME_Func.ahk からよく使う IME状態取得/セット部のみ抽出
    各関数間に依存関係はありません。
    関数間の依存関係は無いので
    必要な関数単体で
    2008.07.11 v1.0.47以降の 関数ライブラリスクリプト対応用にファイル名を修正
-----------------------------------------------------------------------------
*/

IME_GET(WinTitle="")
;-----------------------------------------------------------
; IMEの状態の取得
;    対象： AHK v1.0.34以降
;   WinTitle : 対象Window (省略時:アクティブウィンドウ)
;   戻り値  1:ON 0:OFF
;-----------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_GETOPENSTATUS
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x005,0,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}

IME_SET(setSts, WinTitle="")
;-----------------------------------------------------------
; IMEの状態をセット
;    対象： AHK v1.0.34以降
;   SetSts  : 1:ON 0:OFF
;   WinTitle: 対象Window (省略時:アクティブウィンドウ)
;   戻り値  1:ON 0:OFF
;-----------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_SETOPENSTATUS
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x006,setSts,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}

;===========================================================================
; IME 入力モード 取得 / セット
;
;    0000xxxx    かな入力
;    0001xxxx    ローマ字入力
;    xxxx0xxx    半角
;    xxxx1xxx    全角
;    xxxxx000    英数
;    xxxxx001    ひらがな
;    xxxxx011    ｶﾅ/カナ
;
;     0 (0x00  0000 0000) かな    半英数
;     3 (0x03  0000 0011)         半ｶﾅ
;     8 (0x08  0000 1000)         全英数
;     9 (0x09  0000 1001)         ひらがな
;    11 (0x0B  0000 1011)         全カタカナ
;    16 (0x10  0001 0000) ローマ字半英数
;    19 (0x13  0001 0011)         半ｶﾅ
;    24 (0x18  0001 1000)         全英数
;    25 (0x19  0001 1001)         ひらがな
;    27 (0x1B  0001 1011)         全カタカナ

IME_GetConvMode(WinTitle="")
;-------------------------------------------------------
; IME 入力モード取得
;   in  WinTitle    (省略時アクティブウィンドウ)
;   戻り値:  入力モード
;--------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_GETCONVERSIONMODE
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x001,0,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}

IME_SetConvMode(ConvMode,WinTitle="")
;-------------------------------------------------------
; IME 入力モードセット
;   in  ConvMode    入力モード
;       WinTitle    (省略時アクティブウィンドウ)
;--------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_SETCONVERSIONMODE
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x002,ConvMode,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}

;===========================================================================
; IME 変換モード
;    0:無変換
;    1:人名/地名
;    8:一般
;   16:話し言葉優先

IME_GetSentenceMode(WinTitle="")
;-------------------------------------------------------
; IME 変換モード取得
;   in  WinTitle    (省略時アクティブウィンドウ)
;   戻り値:  0:無変換 1:人名/地名 8:一般 16:話し言葉優先
;--------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_GETSENTENCEMODE
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x003,0,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}

IME_SetSentenceMode(SentenceMode,WinTitle="")
;-------------------------------------------------------
; IME 変換モードセット
;   in  SentenceMode 0:無変換 1:人名/地名 8:一般 16:話し言葉優先
;       WinTitle     (省略時アクティブウィンドウ)
;--------------------------------------------------------
{
    ifEqual WinTitle,,  SetEnv,WinTitle,A
    WinGet,hWnd,ID,%WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint)

    ;Message : WM_IME_CONTROL  wParam:IMC_SETSENTENCEMODE
    DetectSave := A_DetectHiddenWindows
    DetectHiddenWindows,ON
    SendMessage 0x283, 0x004,SentenceMode,,ahk_id %DefaultIMEWnd%
    DetectHiddenWindows,%DetectSave%
    Return ErrorLevel
}
