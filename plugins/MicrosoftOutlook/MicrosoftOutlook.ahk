;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun

MicrosoftOutlook:
    vim.SetWin("MicrosoftOutlook", "rctrl_renwnd32")

    vim.comment("<Mso_Sort_By_Date>", "Sorte emails ordered by date")
    vim.comment("<Mso_Sort_By_Sender>", "Sorte emails ordered by sender")
    vim.comment("<Mso_Sort_By_Subject>", "Sorte emails ordered by subject")

    vim.mode("insert", "MicrosoftOutlook")
    vim.map("<esc>", "<Mso_NormalMode>", "MicrosoftOutlook")
    vim.mode("normal", "MicrosoftOutlook")
    vim.map("i", "<Mso_InsertMode>", "MicrosoftOutlook")
    vim.map("a", "<Mso_Sort_By_Sender>", "MicrosoftOutlook")
    vim.map("s", "<Mso_Sort_By_Subject>", "MicrosoftOutlook")
    vim.map("d", "<Mso_Sort_By_Date>", "MicrosoftOutlook")

    vim.map("h", "<Mso_FirstMail>", "MicrosoftOutlook")
    vim.map("j", "<Mso_Down>", "MicrosoftOutlook")
    vim.map("k", "<Mso_Up>", "MicrosoftOutlook")
    vim.map("l", "<Mso_LastMail>", "MicrosoftOutlook")
    
    vim.map("o", "<Mso_Open>", "MicrosoftOutlook")
    vim.map("r", "<Mso_Reply>", "MicrosoftOutlook")
    vim.map("v", "<Mso_ReplyToAll>", "MicrosoftOutlook")
    vim.map("f", "<Mso_Forward>", "MicrosoftOutlook")
    vim.map("n", "<Mso_New>", "MicrosoftOutlook")

    vim.map(".", "<Mso_FocusSearchBox>", "MicrosoftOutlook")

    vim.map("t", "<Mso_ToggleFlag>", "MicrosoftOutlook")
Return

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
    vim.mode("normal", "MicrosoftOutlook")
return

<Mso_InsertMode>:
    vim.mode("insert", "MicrosoftOutlook")
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
