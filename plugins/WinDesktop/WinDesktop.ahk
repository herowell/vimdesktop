WinDesktop:
    WDTP := "WinDesktop"
    WDTP_Cls_Name := "WorkerW"

    vim.SetWin(WDTP, WDTP_Cls_Name)

    vim.mode("insert", WDTP)
    vim.map("<esc>", "<WDTP_NormalMode>", WDTP)

    vim.mode("normal", WDTP)
    vim.map("i", "<WDTP_InsertMode>", WDTP)
    vim.map("ff", "<WDTP_Copy>", WDTP)
    vim.map("fx", "<WDTP_Move>", WDTP)
    vim.map("fv", "<WDTP_Paste>", WDTP)
    vim.map("r", "<WDTP_Rename>", WDTP)
    vim.map("x", "<WDTP_Delete>", WDTP)

Return

<WDTP_Copy>:
    Send, ^c
Return

<WDTP_Move>:
    Send, ^x
Return

<WDTP_Paste>:
    Send, ^v
Return

<WDTP_Rename>:
    Send, {F2} 
Return

<WDTP_Delete>:
    Send, {Delete} 
Return

<WDTP_NormalMode>:
    vim.mode("normal", WDTP)
Return

<WDTP_InsertMode>:
    vim.mode("insert", WDTP)
Return