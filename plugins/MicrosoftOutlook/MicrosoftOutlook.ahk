;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun

global MSO := "Microsoftoutlook"

MicrosoftOutlook:
    vim.SetWin(MSO, "rctrl_renwnd32")

    vim.comment("<MSO_Sort_By_Date>", "Sort emails by date")
    vim.comment("<MSO_Sort_By_Sender>", "Sort emails by sender")
    vim.comment("<MSO_Sort_By_Subject>", "Sort emails by subject")

    vim.mode("insert", MSO)
    vim.map("<esc>", "<MSO_NormalMode>", MSO)
    vim.mode("normal", MSO)

    vim.map("i", "<MSO_InsertMode>", MSO)
    vim.map("a", "<MSO_Sort_By_Sender>", MSO)
    vim.map("s", "<MSO_Sort_By_Subject>", MSO)
    vim.map("d", "<MSO_Sort_By_Date>", MSO)

    vim.map("h", "<MSO_FirstMail>", MSO)
    vim.map("j", "<MSO_Down>", MSO)
    vim.map("k", "<MSO_Up>", MSO)
    vim.map("l", "<MSO_LastMail>", MSO)
    
    vim.map("o", "<MSO_Open>", MSO)
    vim.map("r", "<MSO_Reply>", MSO)
    vim.map("v", "<MSO_ReplyToAll>", MSO)
    vim.map("f", "<MSO_Forward>", MSO)
    vim.map("n", "<MSO_New>", MSO)

    vim.map(".", "<MSO_FocusSearchBox>", MSO)

    vim.map("t", "<MSO_ToggleFlag>", MSO)
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
    vim.mode("normal", MSO)
return

<MSO_InsertMode>:
    vim.mode("insert", MSO)
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
