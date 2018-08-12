;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun


MicrosoftOutlook:
    MSOutlook := "MicrosoftOutlook"

    vim.SetWin(MSOutlook, "rctrl_renwnd32")

    vim.comment("<MSO_Sort_By_Date>", "Sort emails by date")
    vim.comment("<MSO_Sort_By_Sender>", "Sort emails by sender")
    vim.comment("<MSO_Sort_By_Subject>", "Sort emails by subject")

    vim.mode("insert", MSOutlook)
    vim.map("<esc>", "<MSO_NormalMode>", MSOutlook)
    vim.mode("normal", MSOutlook)

    vim.map("i", "<MSO_InsertMode>", MSOutlook)
    vim.map("a", "<MSO_Sort_By_Sender>", MSOutlook)
    vim.map("s", "<MSO_Sort_By_Subject>", MSOutlook)
    vim.map("d", "<MSO_Sort_By_Date>", MSOutlook)

    vim.map("h", "<MSO_FirstMail>", MSOutlook)
    vim.map("j", "<MSO_Down>", MSOutlook)
    vim.map("k", "<MSO_Up>", MSOutlook)
    vim.map("l", "<MSO_LastMail>", MSOutlook)
    
    vim.map("o", "<MSO_Open>", MSOutlook)
    vim.map("r", "<MSO_Reply>", MSOutlook)
    vim.map("v", "<MSO_ReplyToAll>", MSOutlook)
    vim.map("f", "<MSO_Forward>", MSOutlook)
    vim.map("n", "<MSO_New>", MSOutlook)

    vim.map(".", "<MSO_FocusSearchBox>", MSOutlook)

    vim.map("t", "<MSO_ToggleFlag>", MSOutlook)
return

<MSO_Sort_By_Date>:
    Send, !vabd
Return

<MSO_Sort_By_Sender>:
    Send, !vabf
Return

<MSO_Sort_By_Subject>:
    Send, !vabj
Return

<MSO_NormalMode>:
    vim.mode("normal", MSOutlook)
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
