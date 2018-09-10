;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun

#Include *i %A_ScriptDir%\lib\IME.ahk

MicrosoftOutlook:
    MSOutlook := "MicrosoftOutlook"
    MSOutlook_Cls_Name := "rctrl_renwnd32"

    vim.SetWin(MSOutlook, MSOutlook_Cls_Name)

    vim.comment("<MSO_SortBySender>", "Sort emails by sender")
    vim.comment("<MSO_SortByRecipient>", "Sort emails by recipient")
    vim.comment("<MSO_SortBySubject>", "Sort emails by subject")
    vim.comment("<MSO_SortByDate>", "Sort emails by date")
    vim.comment("<MSO_PasteFromClipboard>", "Paste from clipboard")
    vim.comment("<MSO_Forward>", "Forward selected mail")
    vim.comment("<MSO_Send>", "Send composed mail")
    vim.comment("<MSO_CopySelectedEmailFromMainOutlookWindow>", "Copy selected email in main Outlook window")

    vim.mode("insert", MSOutlook)

    vim.map("<esc>", "<MSO_NormalMode>", MSOutlook)
    vim.map("^[", "<MSO_NormalMode>", msoutlook)

    vim.mode("normal", MSOutlook)

    vim.map("i", "<MSO_InsertMode>", MSOutlook)
    vim.map("sa", "<MSO_SortBySender>", MSOutlook)
    vim.map("sr", "<MSO_SortByRecipient>", MSOutlook)
    vim.map("ss", "<MSO_SortBySubject>", MSOutlook)
    vim.map("sd", "<MSO_SortByDate>", MSOutlook)
    vim.map("se", "<MSO_Send>", MSOutlook)

    vim.map("h", "<Mso_FirstMailOrMoveLeft>", MSOutlook)
    vim.map("j", "<MSO_Down>", MSOutlook)
    vim.map("k", "<MSO_Up>", MSOutlook)
    vim.map("l", "<MSO_LastMailOrMoveRight>", MSOutlook)
    vim.map("[", "<MSO_NextItem>", MSOutlook)
    vim.map("]", "<MSO_PreviousItem>", MSOutlook)
    
    vim.map("o", "<MSO_Open>", MSOutlook)
    vim.map("r", "<MSO_Reply>", MSOutlook)
    vim.map("v", "<MSO_ReplyToAll>", MSOutlook)
    vim.map("w", "<MSO_Forward>", MSOutlook)
    vim.map("n", "<MSO_New>", MSOutlook)

    vim.map(".", "<MSO_FocusSearchBox>", MSOutlook)

    vim.map("t", "<MSO_ToggleFlag>", MSOutlook)
    vim.map("x", "<MSO_Delete>", MSOutlook)
    vim.map("X", "<MSO_PermanentDelete>", MSOutlook)

    vim.map("r", "<MSO_MarkUnread>", MSOutlook)
    vim.map("R", "<MSO_MarkRead>", MSOutlook)

    vim.map("m", "<MSO_MaximizeWin>", MSOutlook)
    vim.map("M", "<MSO_RestoreWin>", MSOutlook)

    ;Using fv when composing new email will paste from clipboard
    ;It would be useful to use fv in main Outlook window if you have already copied some attachments into clipboard
    ;This action will:
    ;   1. create a new email and will paste the attachment
    ;   2. set email subject to the file name
    vim.map("fv", "<MSO_PasteFromClipboard>", MSOutlook)

    ;Force insert mode shall be disabled in order to use "fv" key binding
    ;Otherwise you can not return back to normal mode due to <esc> will close current email window by default
    ;vim.BeforeActionDo("MSO_ForceInsertMode", MSOutlook)

    vim.map("ff", "<MSO_CopySelectedEmailFromMainOutlookWindow>", MSOutlook)

    vim.map("``", "<MSO_ToggleShowInfo>", MSOutlook)

    vim.BeforeActionDo("MSO_BeforeActionDo", MSOutlook)
return

MSO_BeforeActionDo()
{
    ;MSO_IsEmailOpen()
}

MSO_ChangeIMEToEn()
{
    ;Facilitate searching using EN instead CHN
    if (IME_GetConvMode() = 1025) ;tested with Baidu Wubi
    {
        Send, {Shift}
    }
}

MSO_ForceInsertMode()
{
    ControlGetFocus, ctrl, AHK_CLASS rctrl_renwnd32
    if RegExMatch(ctrl, "_WwG1")
        return true
    
    return false
}

MSO_IsEmailOpen()
{
    ControlGetFocus, ctrl, AHK_CLASS rctrl_renwnd32
    if RegExMatch(ctrl, "_WwG1") 
        return true
    if RegExMatch(ctrl, "RichEdit20WPT2") 
        return true
    if RegExMatch(ctrl, "RichEdit20WPT3") 
        return true
    if RegExMatch(ctrl, "RichEdit20WPT5") 
        return true
    return false
}

<MSO_SortByDate>:
    Send, !vabd
    MSO_ChangeIMEToEn()
Return

<MSO_SortBySender>:
    Send, !vabf
    MSO_ChangeIMEToEn()
    vim.mode("insert", MSOutlook)
Return

<MSO_SortByRecipient>:
    Send, !vabt
    MSO_ChangeIMEToEn()
    vim.mode("insert", MSOutlook)
Return

<MSO_SortBySubject>:
    Send, !vabj
    MSO_ChangeIMEToEn()
    vim.mode("insert", MSOutlook)
Return

<MSO_Send>:
    if MSO_IsEmailOpen()
        Send, !s
Return

<MSO_NormalMode>:
    vim.mode("normal", MSOutlook)
return

<MSO_PasteFromClipboard>:
    send, ^v
return

<MSO_InsertMode>:
    vim.mode("insert", MSOutlook)
return

<Mso_FirstMailOrMoveLeft>:
    if MSO_IsEmailOpen(){
        Send, {Left}
    }
    else{
        Send, {Home}
    }
return

<MSO_Up>:
    Send, {Up}
return

<MSO_Down>:
    Send, {Down}
return

<MSO_LastMailOrMoveRight>:
    if MSO_IsEmailOpen(){
        Send, {Right}
    }
    else{
        Send, {End}
    }
return

<MSO_NextItem>:
    if MSO_IsEmailOpen()
        Send, ^>
return

<MSO_PreviousItem>:
    if MSO_IsEmailOpen()
        Send, ^<
return

<MSO_Open>:
    Send, ^o
return

<MSO_Reply>:
    Send, ^r
return

<MSO_ReplyToAll>:
    Send, ^+r
return

<MSO_Forward>:
    Send, ^f 
return

<MSO_New>:
    Send, ^n 
return

<MSO_FocusSearchBox>:
    Send, ^e 
return

<MSO_ToggleFlag>:
    Send, {Insert} 
return

<MSO_Delete>:
    MsgBox, 49, VIMDesktop-Outlook Confirm Dialog, Do you want to continue? 
    IfMsgBox Cancel 
        return
    Send, {Delete} 
return

<MSO_PermanentDelete>:
    Send, +{Delete} 
return

<MSO_MarkUnread>:
    Send, ^u 
return

<MSO_MarkRead>:
    Send, ^q 
return

<MSO_MaximizeWin>:
    PostMessage, 0x112, 0xF030,,, A,  ; 0x112 = WM_SYSCOMMAND, 0xF030 = SC_MAXIMIZE;for active window
return

<MSO_RestoreWin>:
    PostMessage, 0x112, 0xF120,,, A,  ; 0x112 = WM_SYSCOMMAND, 0xF030 = SC_MAXIMIZE;for active window
return

<MSO_CopySelectedEmailFromMainOutlookWindow>:
    send, ^c
return

<MSO_ToggleShowInfo>:
    vim.GetWin(MSOutlook).SetInfo(!vim.GetWin(MSOutlook).info)
return