;Tested for Outlook 2016
;Referenced Gmail keys for Outlook 2016 version 4.0 by Myrick
;By Lu Da Jun

global Win_Name := "microsoftoutlooks"

MicrosoftOutlook:
    vim.SetWin(Win_Name, "rctrl_renwnd32")

    vim.comment("<Mso_Sort_By_Date>", "Sort emails by date")
    vim.comment("<Mso_Sort_By_Sender>", "Sort emails by sender")
    vim.comment("<Mso_Sort_By_Subject>", "Sort emails by subject")

    vim.mode("insert", Win_Name)
    vim.map("<esc>", "<Mso_NormalMode>", Win_Name)
    vim.mode("normal", Win_Name)

    vim.map("i", "<Mso_InsertMode>", Win_Name)
    vim.map("a", "<Mso_Sort_By_Sender>", Win_Name)
    vim.map("s", "<Mso_Sort_By_Subject>", Win_Name)
    vim.map("d", "<Mso_Sort_By_Date>", Win_Name)

    vim.map("h", "<Mso_FirstMail>", Win_Name)
    vim.map("j", "<Mso_Down>", Win_Name)
    vim.map("k", "<Mso_Up>", Win_Name)
    vim.map("l", "<Mso_LastMail>", Win_Name)
    
    vim.map("o", "<Mso_Open>", Win_Name)
    vim.map("r", "<Mso_Reply>", Win_Name)
    vim.map("v", "<Mso_ReplyToAll>", Win_Name)
    vim.map("f", "<Mso_Forward>", Win_Name)
    vim.map("n", "<Mso_New>", Win_Name)

    vim.map(".", "<Mso_FocusSearchBox>", Win_Name)

    vim.map("t", "<Mso_ToggleFlag>", Win_Name)
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
    vim.mode("normal", Win_Name)
return

<Mso_InsertMode>:
    vim.mode("insert", Win_Name)
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
