;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun

global MSO := "Microsoftoutlook"

MicrosoftOutlook:
    vim.SetWin(MSO, "rctrl_renwnd32")

    vim.comment("<Mso_Sort_By_Date>", "Sort emails by date")
    vim.comment("<Mso_Sort_By_Sender>", "Sort emails by sender")
    vim.comment("<Mso_Sort_By_Subject>", "Sort emails by subject")

    vim.mode("insert", MSO)
    vim.map("<esc>", "<Mso_NormalMode>", MSO)
    vim.mode("normal", MSO)

    vim.map("i", "<Mso_InsertMode>", MSO)
    vim.map("a", "<Mso_Sort_By_Sender>", MSO)
    vim.map("s", "<Mso_Sort_By_Subject>", MSO)
    vim.map("d", "<Mso_Sort_By_Date>", MSO)

    vim.map("h", "<Mso_FirstMail>", MSO)
    vim.map("j", "<Mso_Down>", MSO)
    vim.map("k", "<Mso_Up>", MSO)
    vim.map("l", "<Mso_LastMail>", MSO)
    
    vim.map("o", "<Mso_Open>", MSO)
    vim.map("r", "<Mso_Reply>", MSO)
    vim.map("v", "<Mso_ReplyToAll>", MSO)
    vim.map("f", "<Mso_Forward>", MSO)
    vim.map("n", "<Mso_New>", MSO)

    vim.map(".", "<Mso_FocusSearchBox>", MSO)

    vim.map("t", "<Mso_ToggleFlag>", MSO)
return

<Mso_Sort_By_Date>:
    Send, !vabd
Return

<Mso_Sort_By_Sender>:
    Send, !vabf
Return

<Mso_Sort_By_Subject>:
    Send, !vabj
Return

<Mso_NormalMode>:
    vim.mode("normal", MSO)
return

<Mso_InsertMode>:
    vim.mode("insert", MSO)
return

<Mso_FirstMail>:
    Send, {Home}
return

<Mso_Up>:
    Send, {Up}
return

<Mso_Down>:
    Send, {Down}
return

<Mso_LastMail>:
    Send, {End}
return

<Mso_Open>:
    Send, ^o
return

<Mso_Reply>:
    Send, ^r
return

<Mso_ReplyToAll>:
    Send, ^+r
return

<Mso_Forward>:
    Send, ^f 
return

<Mso_New>:
    Send, ^n 
return

<Mso_FocusSearchBox>:
    Send, ^e 
return

<Mso_ToggleFlag>:
    Send, {Insert} 
return
