;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun


AdobeAcrobat:
    AdobeAcrobat := "AdobeAcrobat"
    AdobeAcrobat_Cls_Name := "AcrobatSDIWindow"
    cur_view := 0

    vim.SetWin(AdobeAcrobat, AdobeAcrobat_Cls_Name)

    vim.mode("normal", AdobeAcrobat)

    vim.map("h", "<Adobe_HomePage>", AdobeAcrobat)
    vim.map("j", "<Adobe_Down>", AdobeAcrobat)
    vim.map("k", "<Adobe_Up>", AdobeAcrobat)
    vim.map("l", "<Adobe_LastPage>", AdobeAcrobat)

    vim.map("m", "<Adobe_MaximizeWin>", AdobeAcrobat)
    vim.map("M", "<Adobe_RestoreWin>", AdobeAcrobat)

    vim.map("t", "<Adobe_ToggleToolsPane>", AdobeAcrobat)

    vim.map("v", "<Adobe_ToggleView>", AdobeAcrobat)
return

<Adobe_NormalMode>:
    vim.mode("normal", AdobeAcrobat)
return

<Adobe_HomePage>:
    Send, {Home}
return

<Adobe_Up>:
    Send, {Up}
return

<Adobe_Down>:
    Send, {Down}
return

<Adobe_LastPage>:
    Send, {End}
return

<Adobe_MaximizeWin>:
    PostMessage, 0x112, 0xF030,,, A,  ; 0x112 = WM_SYSCOMMAND, 0xF030 = SC_MAXIMIZE;for active window
return

<Adobe_RestoreWin>:
    PostMessage, 0x112, 0xF120,,, A,  ; 0x112 = WM_SYSCOMMAND, 0xF030 = SC_MAXIMIZE;for active window
return

<Adobe_ToggleToolsPane>:
    Send, +{F4}
return

<Adobe_ToggleView>:
    global cur_view
    ;static cur_view := 0
    IfEqual, cur_view, 0, Send, ^0
    IfEqual, cur_view, 1, Send, ^1
    IfEqual, cur_view, 2, Send, ^2
    cur_view += 1
    if (cur_view = 3){
        cur_view := 0
    }
return

