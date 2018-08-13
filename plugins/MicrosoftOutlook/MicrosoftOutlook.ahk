;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun


MicrosoftOutlook:
    MSOutlook := "MicrosoftOutlook"
    MSOutlook_Cls_Name := "rctrl_renwnd32"

    vim.SetWin(MSOutlook, MSOutlook_Cls_Name)

    vim.comment("<MSO_Sort_By_Date>", "Sort emails by date")
    vim.comment("<MSO_Sort_By_Sender>", "Sort emails by sender")
    vim.comment("<MSO_Sort_By_Subject>", "Sort emails by subject")
    vim.comment("<MSO_PasteFromClipboard>", "Paste from clipboard")
    vim.comment("<MSO_Forward>", "Forward selected mail")
    vim.comment("<MSO_Send>", "Send composing mail")

    vim.mode("insert", MSOutlook)
    vim.map("<esc>", "<MSO_NormalMode>", MSOutlook)

    vim.mode("normal", MSOutlook)

    vim.map("i", "<MSO_InsertMode>", MSOutlook)
    vim.map("sa", "<MSO_Sort_By_Sender>", MSOutlook)
    vim.map("ss", "<MSO_Sort_By_Subject>", MSOutlook)
    vim.map("sd", "<MSO_Sort_By_Date>", MSOutlook)
    vim.map("se", "<MSO_Send>", MSOutlook)

    vim.map("h", "<MSO_FirstMail>", MSOutlook)
    vim.map("j", "<MSO_Down>", MSOutlook)
    vim.map("k", "<MSO_Up>", MSOutlook)
    vim.map("l", "<MSO_LastMail>", MSOutlook)
    vim.map("[", "<MSO_NextItem>", MSOutlook)
    vim.map("]", "<MSO_PreviousItem>", MSOutlook)
    
    vim.map("o", "<MSO_Open>", MSOutlook)
    vim.map("r", "<MSO_Reply>", MSOutlook)
    vim.map("v", "<MSO_ReplyToAll>", MSOutlook)
    vim.map("f", "<MSO_Forward>", MSOutlook)
    vim.map("n", "<MSO_New>", MSOutlook)

    vim.map(".", "<MSO_FocusSearchBox>", MSOutlook)

    vim.map("t", "<MSO_ToggleFlag>", MSOutlook)

    ;Using fv when composing new email will paste from clipboard
    ;It would be useful to use fv in main Outlook window if you have already copied some attachments into clipboard
    ;This action will:
    ;   1. create a new email and will paste the attachment
    ;   2. set email subject to the file name
    vim.map("fv", "<MSO_PasteFromClipboard>", MSOutlook)

    ;Force insert mode shall be disabled in order to use "fv" key binding
    ;Otherwise you can not return back to normal mode due to <esc> will close current email window by default
    ;vim.BeforeActionDo("MSOutlook_Force_Insert_Mode", MSOutlook)
return

MSOutlook_Force_Insert_Mode()
{
    ControlGetFocus, ctrl, AHK_CLASS rctrl_renwnd32
    if RegExMatch(ctrl, "_WwG1")
        return true
    return false
}

MSO_Is_Email_Open()
{
    ControlGetFocus, ctrl, AHK_CLASS rctrl_renwnd32
    if RegExMatch(ctrl, "_WwG1")
        return true
    return false
}

<MSO_Sort_By_Date>:
    Send, !vabd
Return

<MSO_Sort_By_Sender>:
    Send, !vabf
Return

<MSO_Sort_By_Subject>:
    Send, !vabj
Return

<MSO_Send>:
    if MSO_Is_Email_Open()
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

<MSO_FirstMail>:
    Send, {Home}
return

<MSO_Up>:
    Send, {Up}
return

<MSO_Down>:
    Send, {Down}
return

<MSO_LastMail>:
    Send, {End}
return

<MSO_NextItem>:
    if MSO_Is_Email_Open()
        Send, ^.
return

<MSO_PreviousItem>:
    if MSO_Is_Email_Open()
        Send, ^,
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
