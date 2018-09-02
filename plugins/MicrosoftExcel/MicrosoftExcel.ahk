MicrosoftExcel:
    global Workbook
    global excel
    global Sheet
    global Cell
    global Selection
    global lLastRow ;整个表的最末尾行
    global lLastColumn ;整个表最末尾列
    global SelectionFirstRow ;当前选择内容首行
    global SelectionFirstColumn ;当前选择内容首列
    global SelectionLastColumn ;当前选择内容末列
    global SelectionLastRow ;当前选择内容末行
    global SelectionType ; 当前选择单元格类型 1=A1  2=A1:B1 4=A1:A2 16=A1:B2  18=A1:B1 A1:B2 20=A1:A2 A1:B2
    global FontColor := -4165632  ;填充字体颜色-默认蓝色
    global CellColor := -16711681 ;填充表格颜色-默认黄色

    MSExcel := "MircosoftExcel" 

    vim.SetWin(MSExcel, "XLMAIN")

    vim.comment("<Insert_Mode_MicrosoftExcel>", "insert模式")
    vim.comment("<Normal_Mode_MicrosoftExcel>", "normal模式")
    vim.comment("<MSE_SheetReName>", "重命名当前工作表名称")
    vim.comment("<MSE_GoTo>", "跳转到指定行列值的表格")
    vim.comment("<MSE_SaveAndExit>", "保存并退出")
    vim.comment("<MSE_DiscardAndExit>", "放弃修改并退出")
    vim.comment("<MSE_Undo>", "撤销")
    vim.comment("<MSE_Redo>", "重做")
    vim.comment("<MSE_SaveAndExit>", "保存后退出")
    vim.comment("<MSE_DiscardAndExit>", "不保存退出")
    vim.comment("<MSE_Color_Font>", "设置选中区域字体为上次颜色")
    vim.comment("<MSE_Color_Cell>", "填充选中表格背景为上次颜色")
    vim.comment("<MSE_Color_All>", "同时应用字体颜色、背景颜色")
    vim.comment("<MSE_Color_Menu_Font>", "设置选中区域字体颜色")
    vim.comment("<MSE_Color_Menu_Cell>", "填充选中表格背景颜色")
    vim.comment("<MSE_FocusHome>", "定位到工作表开头")
    vim.comment("<MSE_FocusEnd>", "定位到工作表最后一个单元格")
    vim.comment("<MSE_FocusRowHome>", "定位到当前列首行")
    vim.comment("<MSE_FocusRowEnd>", "定位到当前列尾行")
    vim.comment("<MSE_FocusColHome>", "定位到当前行首列")
    vim.comment("<MSE_FocusColEnd>", "定位到当前行尾列")
    vim.comment("<MSE_FocusAreaLeft>", "定位到当前区域边缘-左")
    vim.comment("<MSE_FocusAreaRight>", "定位到当前区域边缘-右")
    vim.comment("<MSE_FocusAreaUp>", "定位到当前区域边缘-上")
    vim.comment("<MSE_FocusAreaDown>", "定位到当前区域边缘-下")
    vim.comment("<MSE_SelectToAreaLeft>", "选择到当前区域边缘-左")
    vim.comment("<MSE_SelectToAreaRight>", "选择到当前区域边缘-右")
    vim.comment("<MSE_SelectToAreaUp>", "选择到当前区域边缘-上")
    vim.comment("<MSE_SelectToAreaDown>", "选择到当前区域边缘-下")
    vim.comment("<MSE_Delete>", "删除（=Delete键）")
    vim.comment("<MSE_SelectAll>", "选择全部=^a")
    vim.comment("<MSE_Paste_Value>", "粘贴数值")
    vim.comment("<MSE_PageUp>", "向上翻页")
    vim.comment("<MSE_PageDown>", "向下翻页")
    vim.comment("<MSE_Cut>", "剪切")
    vim.comment("<MSE_Replace>", "替换")
    vim.comment("<MSE_Find>", "查找")
    vim.comment("<Alt_Mode_MicrosoftExcel>", "alt命令模式")

    ;insert模式及快捷键
    vim.mode("insert", MSExcel)
    vim.map("<esc>", "<Normal_Mode_MicrosoftExcel>", MSExcel)

    ;normal模式及快捷键
    vim.mode("normal", MSExcel)
    vim.map("i", "<Insert_Mode_MicrosoftExcel>", MSExcel)
    vim.map("<esc>", "<Normal_Mode_MicrosoftExcel>", MSExcel)
    vim.map("I", "<Alt_Mode_MicrosoftExcel>", MSExcel)

    ;数字计数
    vim.map("1", "<1>", MSExcel)
    vim.map("2", "<2>", MSExcel)
    vim.map("3", "<3>", MSExcel)
    vim.map("4", "<4>", MSExcel)
    vim.map("5", "<5>", MSExcel)
    vim.map("6", "<6>", MSExcel)
    vim.map("7", "<7>", MSExcel)
    vim.map("8", "<8>", MSExcel)
    vim.map("9", "<9>", MSExcel)

    ;撤销与重复
    vim.map("u", "<MSE_Undo>", MSExcel)
    vim.map("<c-r>", "<MSE_Redo>", MSExcel)

    ;Z保存与退出
    vim.map("ZZ", "<MSE_SaveAndExit>", MSExcel)
    vim.map("ZQ", "<MSE_DiscardAndExit>", MSExcel)

    ;颜色
    vim.map("""", "<MSE_Color_All>", MSExcel)
    vim.map("'", "<MSE_Color_Menu_Font>", MSExcel)
    vim.map(";", "<MSE_Color_Menu_Cell>", MSExcel)

    ;d删除
    vim.map("dd", "<MSE_Delete>", MSExcel)
    vim.map("D", "<MSE_Delete>", MSExcel)
    vim.map("dr", "<MSE_删除选择行>", MSExcel)
    vim.map("dc", "<MSE_删除选择列>", MSExcel)
    vim.map("dw", "<MSE_工作表删除当前>", MSExcel)

    ;o插入/O插入在右
    vim.map("or", "<MSE_编辑插入新行在前>", MSExcel)
    vim.map("oc", "<MSE_编辑插入新列在左>", MSExcel)
    vim.map("Or", "<MSE_编辑插入新行在后>", MSExcel)
    vim.map("Oc", "<MSE_编辑插入新列在右>", MSExcel)
    vim.map("ow", "<MSE_工作表新建>", MSExcel)

    ;s选择
    vim.map("sk", "<MSE_SelectToAreaUp>", MSExcel)
    vim.map("sj", "<MSE_SelectToAreaDown>", MSExcel)
    vim.map("sh", "<MSE_SelectToAreaLeft>", MSExcel)
    vim.map("sl", "<MSE_SelectToAreaRight>", MSExcel)
    vim.map("sr", "<MSE_选择整行>", MSExcel)
    vim.map("sc", "<MSE_选择整列>", MSExcel)
    vim.map("sa", "<MSE_SelectAll>", MSExcel)

    ;f过滤命令
    vim.map("ff", "<MSE_自动过滤开启>", MSExcel)
    vim.map("fl", "<MSE_过滤当前列下拉菜单>", MSExcel)
    vim.map("fd", "<MSE_过滤打开筛选对话框>", MSExcel)
    vim.map("fo", "<MSE_过滤大于等于当前单元格>", MSExcel)
    vim.map("fu", "<MSE_过滤小于等于当前单元格>", MSExcel)
    vim.map("f.", "<MSE_过滤非空单元格>", MSExcel)
    vim.map("fb", "<MSE_过滤空单元格>", MSExcel)

    ;因不区分数值型与文本型以及日期型的问题，以下过滤功能暂不完整
    vim.map("fB", "<MSE_过滤开头包含当前单元格>", MSExcel)
    vim.map("fE", "<MSE_过滤末尾包含当前单元格>", MSExcel)
    vim.map("fs", "<MSE_过滤等于当前单元格>", MSExcel)
    vim.map("f<", "<MSE_过滤小于当前单元格>", MSExcel)
    vim.map("f>", "<MSE_过滤大于当前单元格>", MSExcel)
    vim.map("fi", "<MSE_过滤包含当前单元格>", MSExcel)
    vim.map("fe", "<MSE_过滤不包含当前单元格>", MSExcel)

    ;以下过滤功能2013版测试无效
    vim.map("fa", "<MSE_过滤取消当前列>", MSExcel)
    vim.map("fA", "<MSE_过滤取消所有列>", MSExcel)

    ;p粘贴
    vim.map("p", "<MSE_Paste>", MSExcel)
    vim.map("P", "<MSE_Paste_Select>", MSExcel)

    ;pv希望以后用代码做，快捷键做会闪一下
    ;vim.map("v", "<MSE_Paste_Value>", MSExcel)

    ;space翻页（PageUp）Shiht-space（PageDown）
    vim.map("<space>", "<MSE_PageDown>", MSExcel)
    vim.map("<S-space>", "<MSE_PageUp>", MSExcel)

    ;x剪切
    vim.map("x", "<MSE_Cut>", MSExcel)

    ;y复制
    vim.map("yy", "<MSE_Copy_Selection>", MSExcel)
    vim.map("Y", "<MSE_Copy_Selection>", MSExcel)
    vim.map("yr", "<MSE_Copy_Row>", MSExcel)
    vim.map("yc", "<MSE_Copy_Col>", MSExcel)
    vim.map("yh", "<MSE_编辑自左侧复制>", MSExcel)
    vim.map("yl", "<MSE_编辑自右侧复制>", MSExcel)
    vim.map("yk", "<MSE_编辑自上侧复制>", MSExcel)
    vim.map("yj", "<MSE_编辑自下侧复制>", MSExcel)
    vim.map("myl", "<MSE_逐行编辑自左侧复制>", MSExcel)
    vim.map("myr", "<MSE_逐行编辑自右侧复制>", MSExcel)
    vim.map("yw", "<MSE_CopyCurrentSheet>", MSExcel)
    vim.map("yW", "<MSE_工作表复制对话框>", MSExcel)

    ;上下左右映射
    vim.map("h", "<left>", MSExcel)
    vim.map("l", "<right>", MSExcel)
    vim.map("k", "<up>", MSExcel)
    vim.map("j", "<down>", MSExcel)

    ;上下左右选择映射
    vim.map("H", "<MSE_向左选择>", MSExcel)
    vim.map("L", "<MSE_向右选择>", MSExcel)
    vim.map("K", "<MSE_向上选择>", MSExcel)
    vim.map("J", "<MSE_向下选择>", MSExcel)

    ;g位置跳转
    vim.map("gg", "<MSE_FocusHome>", MSExcel)
    vim.map("G", "<MSE_FocusEnd>", MSExcel)
    vim.map("grh", "<MSE_FocusRowHome>", MSExcel)
    vim.map("gre", "<MSE_FocusRowEnd>", MSExcel)
    vim.map("gch", "<MSE_FocusColHome>", MSExcel)
    vim.map("gce", "<MSE_FocusColEnd>", MSExcel)
    vim.map("gk", "<MSE_FocusAreaUp>", MSExcel)
    vim.map("gj", "<MSE_FocusAreaDown>", MSExcel)
    vim.map("gh", "<MSE_FocusAreaLeft>", MSExcel)
    vim.map("gl", "<MSE_FocusAreaRight>", MSExcel)
    vim.map("gH", "<MSE_FirstSheet>", MSExcel)
    vim.map("gL", "<MSE_LastSheet>", MSExcel)
    vim.map("gt", "<MSE_NextSheet>", MSExcel)
    vim.map("gT", "<MSE_PreviousSheet>", MSExcel)
    vim.map("go", "<MSE_GoTo>", MSExcel)

    ;F填充
    vim.map("Fk", "<MSE_填充向上>", MSExcel)
    vim.map("Fj", "<MSE_填充向下>", MSExcel)
    vim.map("Fh", "<MSE_填充向左>", MSExcel)
    vim.map("Fl", "<MSE_填充向右>", MSExcel)

    ;r重命名/替换
    vim.map("rr", "<MSE_Replace>", MSExcel)
    vim.map("R", "<MSE_Replace>", MSExcel)
    vim.map("rw", "<MSE_SheetReName>", MSExcel)

    ;/查找
    vim.map("/", "<MSE_Find>", MSExcel)

    ;w宽高/W指定值
    vim.map("wr", "<MSE_自适应宽度选择行>", MSExcel)
    vim.map("wc", "<MSE_自适应宽度选择列>", MSExcel)
    vim.map("Wr", "<MSE_编辑行宽指定值>", MSExcel)
    vim.map("Wc", "<MSE_编辑列宽指定值>", MSExcel)

    ;工作表

    vim.map(">w", "<MSE_工作表移动向后>", MSExcel)
    vim.map("<w", "<MSE_工作表移动向前>", MSExcel)

    ;:字体颜色命令

    ;;单元格颜色命令

    ;%页面设置命令

    ;^设置格式命令

    ;@视图指令

    ;-横向线颜色命令

    ;|纵向ActiveSheet.线颜色指令

    ;`字体命令
    vim.map("<S-,>", "<XLmain_字体放大>", MSExcel)
    vim.map("<S-.>", "<XLmain_字体缩小>", MSExcel)

    ;(名称
    vim.map("<S-9>n", "<MSE_名称工作簿定义>", MSExcel)
    vim.map("<S-9>N", "<MSE_名称当前工作表定义>", MSExcel)

    ;编辑

    ;行指令
    ;vim.map("rh", "<MSE_隐藏选择行>", MSExcel)
    ;vim.map("rH", "<MSE_隐藏选择行取消>", MSExcel)

    ;行填充作用不明显
    ;vim.map("rf", "<MSE_行填充>", MSExcel)

    ;列指令
    ;vim.map("ch", "<MSE_隐藏选择列>", MSExcel)
    ;vim.map("cH", "<MSE_隐藏选择列取消>", MSExcel)

    ;vim.map("e", "<MSE_编辑行宽变窄>", MSExcel)
    ;vim.map("E", "<MSE_编辑行宽变宽>", MSExcel)
    ;vim.map("q", "<MSE_编辑列宽变窄>", MSExcel)
    ;vim.map("Q", "<MSE_编辑列宽变宽>", MSExcel)

    ;m多区域逐行处理
    ;vim.map("mr", "<MSE_逐行合并>", MSExcel)
    ;vim.map("mbd", "<MSE_逐行边框下框线>", MSExcel)
    ;vim.map("mbu", "<MSE_逐行边框上框线>", MSExcel)
    ;vim.map("mbs", "<MSE_逐行边框外侧框线>", MSExcel)
    ;vim.map("mbt", "<MSE_逐行边框粗匣框线>", MSExcel)
    ;vim.map("mR", "<MSE_取消逐行合并>", MSExcel)

    ;测试
    ;vim.map("t5", "<XLMIAN_获取活动工作表边界>", MSExcel)
    vim.map("t1", "<LastRow>", MSExcel)
    vim.map("t2", "<LastColumn>", MSExcel)

    vim.BeforeActionDo("MSE_BeforeActionDo",  MSExcel)
return

;Action 如要跳转，请使用查找功能/

MSE_BeforeActionDo()
{
    ControlGetFocus, ctrl, AHK_CLASS XLMAIN
    ;Excel61 is active when editing
    If RegExMatch(ctrl, "EXCEL61")
        Return True
    return False
}

<Normal_Mode_MicrosoftExcel>:
    Send, {esc}
    vim.Mode("normal", MSExcel)
    getExcel().Application.StatusBar := "Normal Mode"
    
return

<Insert_Mode_MicrosoftExcel>:
    vim.Mode("insert", MSExcel)

    ;插入模式下使用由Excel接管状态栏
    getExcel().Application.StatusBar := blank
return

<Alt_Mode_MicrosoftExcel>:
    vim.Mode("insert", MSExcel)

    ;插入模式下使用由Excel接管状态栏
    getExcel().Application.StatusBar := blank
    {
        send {alt}
        return
    }
return

<MSE_Undo>:
{
    send ^z
    return
}

<MSE_Redo>:
{
    send ^y
    return
}

<MSE_Delete>:
{
    send,{Del}
    return
}

;by dlt:改用快捷键方式，可被撤销
<MSE_删除选择行>:
{
    send ^-
    send !r
    send {Enter}
    return
}

<MSE_删除选择列>:
{
    Excel_Selection()
    Selection.EntireColumn.Delete
    objrelease(excel)
    return
}

<MSE_工作表删除当前>:
{
    Excel_ActiveSheet()
    excel.ActiveWindow.SelectedSheets.delete
    ;objRelease(excel)
    return
}

;o插入
<MSE_编辑插入新行在前>:
{
    send,{AppsKey}
    send,i
    send,{enter}
    sleep,5
    send,r
    send,{enter}
    return
}

<MSE_编辑插入新列在左>:
{
    send,{AppsKey}
    send,i
    send,{enter}
    sleep,5
    send,c
    send,{enter}
    return
}

<MSE_工作表新建>:
{
    Excel_ActiveSheet()
    getExcel().ActiveWorkbook.Sheets.Add
    ;objRelease(excel)
    return
}


;s选择
<MSE_SelectToAreaUp>:
{
    send,^+{Up}
    return
}

<MSE_SelectToAreaDown>:
{
    send,^+{Down}
    return
}

<MSE_SelectToAreaLeft>:
{
    send,^+{Left}
    return
}

<MSE_SelectToAreaRight>:
{
    send,^+{Right}
    return
}

<MSE_选择整行>:
{
    Excel_Selection()
    Selection.EntireRow.Select
    objrelease(excel)
    return
}

<MSE_选择整列>:
{
    Excel_Selection()
    Selection.EntireColumn.Select
    objrelease(excel)
    return
}

<MSE_SelectAll>:
{
    send,^a
    return
}

;space翻页
<MSE_PageDown>:
{
    send,{PgDn}
    return
}

<MSE_PageUp>:
{
    send,{PgUp}
    return
}

;x剪切
<MSE_Cut>:
{
    send,^x
    return
}

;r置换
<MSE_Replace>:
{
    send,^h
    return
}

;/查找
<MSE_Find>:
{
    send,^f
    return
}

;控制
<MSE_向左选择>:
{
    send,+{left}
    return
}


<MSE_向右选择>:
{
    send,+{right}
    return
}

<MSE_向上选择>:
{
    send,+{up}
    return
}

<MSE_向下选择>:
{
    send,+{down}
    return
}

<MSE_名称工作簿定义>:
{
    Excel_Selection()
    InputBox, OutputVar ,输入名称
    If ErrorLevel
        Return
    inputbox, comments ,输入注释
    If ErrorLevel
        Return
    address:=Selection.address
    Name:=OutputVar
    RefersToR1C1:=address
    excel.ActiveWorkbook.Names.Add(Name,RefersToR1C1)
    ActiveWorkbook.Names(OutputVar).Comment := "comments"
    ;objRelease(excel)
    return
}

<MSE_名称当前工作表定义>:
{
    Excel_Selection()
    InputBox, OutputVar ,输入名称
    If ErrorLevel
        Return
    inputbox, comments ,输入注释
    If ErrorLevel
        Return
    address:=Selection.address
    Name:=OutputVar
    RefersToR1C1:=address
    excel.ActiveSheet.Names.Add(Name,RefersToR1C1)
    ActiveSheet.Names(OutputVar).Comment := "comments"
    ;objRelease(excel)
    return
}

<MSE_定位空单元格>:
{
    Excel_Selection()
    MSE_定位对象(4)
    ;objRelease(excel)
    return
}

<MSE_定位任意格式>:
{
    Excel_Selection()
    MSE_定位对象(-4172)
    ;objRelease(excel)
    return
}

<MSE_定位验证条件全部>:
{
    Excel_Selection()
    MSE_定位对象(-4174)
    ;objRelease(excel)
    return
}

<MSE_定位注释>:
{
    Excel_Selection()
    MSE_定位对象(-4144)
    ;objRelease(excel)
    return
}

<MSE_定位常量全部>:
{
    Excel_Selection()
    MSE_定位公式变量(2,23)
    ;objRelease(excel)
    return
}

<MSE_定位公式全部>:
{
    Excel_Selection()
    MSE_定位公式变量(-4123,23)
    ;objRelease(excel)
    return
}

<MSE_定位已用区域最末单元格>:
{
    Excel_Selection()
    MSE_定位对象(11)
    ;objRelease(excel)
    return
}

<MSE_定位相同格式>:
{
    Excel_Selection()
    MSE_定位对象(-4173)
    ;objRelease(excel)
    return
}

<MSE_定位验证条件相同>:
{
    Excel_Selection()
    MSE_定位对象(-4175)
    ;objRelease(excel)
    return
}

<MSE_定位可见>:
{
    Excel_Selection()
    MSE_定位对象(12)
    ;objRelease(excel)
    return
}

MSE_定位对象(value)
{
    Selection.SpecialCells(value).Select
    return
}

MSE_定位公式变量(value,indicate)
{
    Selection.SpecialCells(value,indicate).Select
    return
}

;过滤
<MSE_自动过滤开启>:
{
    Excel_ActiveSheet()
    If excel.ActiveSheet.AutoFilterMode
        excel.ActiveSheet.AutoFilterMode := False
    Else
        excel.Selection.AutoFilter
    ;XLMIAN_获取活动工作表边界()
    ;excel.ActiveSheet.Range("A1" , MSE_ColToChar(lLastColumn) . "1").Select
    ;msgbox,%range%
    ;excel.Application.Dialogs(447).Show(fid,excel.ActiveCell.Value)
    objrelease(excel)
    return
}

<MSE_过滤打开筛选对话框>:
{
    Excel_ActiveSheet()
    address:=excel.ActiveSheet.AutoFilter.Range.Address
    StringReplace, address, address, $,,All
    FoundPosSeperate := RegExMatch(address,":")
    StringLeft, parta, address, FoundPosSeperate-1
    StringMid, partb, address, FoundPosSeperate+1 , 50
    RegExMatch(parta,"[A-Z]+",ColumnLeftName)
    RegExMatch(parta,"[0-9]+",RowUp)
    fid_first_column:=excel.ActiveSheet.Range(ColumnLeftName "1:" ColumnLeftName "1").Column
    fid:=excel.ActiveCell.Column - fid_first_column + 1
    value:=excel.ActiveCell.Value
    excel.Application.Dialogs(447).Show(fid, value)
    objrelease(excel)
    return
}

<MSE_过滤等于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("=",value)
    objrelease(excel)
    return
}

<MSE_过滤小于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("<",value)
    objrelease(excel)
    return
}


<MSE_过滤大于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter(">",value)
    objrelease(excel)
    return
}

<MSE_过滤大于等于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter(">=",value)
    objrelease(excel)
    return
}

<MSE_过滤小于等于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("<=",value)
    objrelease(excel)
    return
}

<MSE_过滤不等于当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("<>",value)
    objrelease(excel)
    return
}

<MSE_过滤非空单元格>:
{
    Excel_ActiveSheet()
    ;value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("<>","")
    objrelease(excel)
    return
}

<MSE_过滤空单元格>:
{
    Excel_ActiveSheet()
    ;value:=excel.ActiveCell.Value
    ;msgbox,%value%
    MSE_CustomAutoFilter("=","")
    objrelease(excel)
    return
}

<MSE_过滤包含当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    value=%value%*
    msgbox,%value%
    MSE_CustomAutoFilter("=*",valve)
    objrelease(excel)
    return
}

<MSE_过滤不包含当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    value=%value%*
    MSE_CustomAutoFilter("<>*",valve)
    objrelease(excel)
    return
}

<MSE_过滤开头包含当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    value=*%value%
    MSE_CustomAutoFilter("=",valve)
    objrelease(excel)
    return
}

<MSE_过滤末尾包含当前单元格>:
{
    Excel_ActiveSheet()
    value:=excel.ActiveCell.Value
    ;msgbox,%value%
    value=%value%*
    MSE_CustomAutoFilter("=",valve)
    objrelease(excel)
    return
}

<MSE_过滤当前列下拉菜单>:
{
    Excel_ActiveSheet()
    ;msgbox,ArithmeticOpr %ArithmeticOpr%
    ;msgbox,CurrentValue %CurrentValue%
    ;msgbox,CriteriaValue %CriteriaValue%
    address:=excel.ActiveSheet.AutoFilter.Range.Address
    StringReplace, address, address, $,,All
    FoundPosSeperate := RegExMatch(address,":")
    StringLeft, parta, address, FoundPosSeperate-1
    StringMid, partb, address, FoundPosSeperate+1 , 50
    ;msgbox,%parta%
    ;msgbox,%partb%
    RegExMatch(parta,"[A-Z]+",ColumnLeftName)
    RegExMatch(parta,"[0-9]+",RowUp)
    ;msgbox,%RowUp%
    column:=excel.ActiveCell.Column
    ;msgbox,%column%
    excel.ActiveSheet.Cells( RowUp , column ).Activate
    send,!{down}
    objrelease(excel)
    return
}

<MSE_过滤取消当前列>:
{
    Excel_Selection()
    value:=excel.ActiveCell.Value
    MSE_CustomAutoFilter("",valve)
    objrelease(excel)
    return
}

<MSE_过滤取消所有列>:
{
    Excel_Selection()
    If excel.ActiveSheet.FilterMode = True
        excel.ActiveSheet.ShowAllData
    objrelease(excel)
    return
}


MSE_CustomAutoFilter(ArithmeticOpr,CurrentValue)
{
    ;msgbox,ArithmeticOpr %ArithmeticOpr%
    ;msgbox,CurrentValue %CurrentValue%
    CriteriaValue = %ArithmeticOpr%%CurrentValue%
    ;msgbox,CriteriaValue %CriteriaValue%
    address:=excel.ActiveSheet.AutoFilter.Range.Address
    ;msgbox,address %address%
    XLMIAN_获取Range边界(address)
    fid_first_column:=excel.ActiveSheet.Range(ColumnLeftName "1:" ColumnLeftName "1").Column
    fid:=excel.ActiveCell.Column - fid_first_column + 1
    Field:=fid
    ;msgbox,Field %Field%
    Criteria1:=CriteriaValue
    ;msgbox,Criterial %Criterial%
    excel.ActiveSheet.Range("A1").CurrentRegion.AutoFilter(Field,Criteria1)
    return
}

;行指令
<MSE_取消逐行合并>:
{
    Excel_ActiveCell()
    if excel.Selection.Columns.Count > 1
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).unmerge
        }
    }
    objrelease(excel)
    return
}

;by dlt:快捷键实现，可被撤销
<MSE_隐藏选择行>:
{
    send ^9
    return
}

;by dlt:Ctrl+Shift+(
<MSE_隐藏选择行取消>:
{
    send ^+(
    return
}

<MSE_自适应宽度选择行>:
{
    Excel_Selection()
    Selection.EntireRow.AutoFit
    objrelease(excel)
    return
}

<MSE_编辑行宽指定值>:
{
    Excel_Selection()
    Default:=Selection.RowHeight
    InputBox, inputvar,输入行宽,,,,,,,,,%Default%
    If ErrorLevel
        Return
    ;InputBox, OutputVar Title  ,,,,,,,,,
    Selection.RowHeight:=inputvar
    tooltip,%inputvar%
    ;objRelease(excel)
    sleep,500
    tooltip,
    return
}

;列指令

<MSE_自适应宽度选择列>:
{
    Excel_Selection()
    Selection.EntireColumn.AutoFit
    objrelease(excel)
    return
}

<MSE_隐藏选择列>:
{
    Excel_Selection()
    Selection.EntireColumn.Hidden := True
    objrelease(excel)
    return
}

<MSE_隐藏选择列取消>:
{
    Excel_Selection()
    Selection.EntireColumn.Hidden := False
    objrelease(excel)
    return
}



<MSE_编辑列宽指定值>:
{
    Excel_Selection()
    Default:=Selection.ColumnWidth
    InputBox, inputvar,输入列宽,,,,,,,,,%Default%
    If ErrorLevel
        Return
    Selection.ColumnWidth:=inputvar
    tooltip,%inputvar%
    ;objRelease(excel)
    sleep,500
    tooltip,
    return
}



;多行指令
<MSE_逐行合并>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            return
        }
    else if SelectionType=2
        {
            excel.Selection.merge
        }
    else if SelectionType=4 ;A1:A4
        {
            Return
        }
    else if SelectionType=16
        {
            rowcount:=excel.selection.rows.count
            Loop, %rowcount%
            {
                address:=excel.selection.rows(A_Index).address
                StringReplace, address, address, $,,All
                ;msgbox,%address%
                FoundPosSeperate := RegExMatch(address,":")
                StringLeft, parta, address, FoundPosSeperate-1
                StringMid, partb, address, FoundPosSeperate+1 , 50
                excel.range(parta ":" partb).merge
            }
        }
    objrelease(excel)
    return


    ; Excel_ActiveCell()
    ; if excel.Selection.Columns.Count > 1
    ; {
    ;     rowcount:=excel.selection.rows.count
    ;     Loop, %rowcount%
    ;     {
    ;         address:=excel.selection.rows(A_Index).address
    ;         StringReplace, address, address, $,,All
    ;         ;msgbox,%address%
    ;         FoundPosSeperate := RegExMatch(address,":")
    ;         StringLeft, parta, address, FoundPosSeperate-1
    ;         StringMid, partb, address, FoundPosSeperate+1 , 50
    ;         excel.range(parta ":" partb).merge
    ;     }
    ; }
    ; objrelease(excel)
    ; return
}

;边框

<MSE_边框下框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(9).Weight := 2
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(9).Weight := 2
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(9).LineStyle := 1
            excel.range(address).Borders(9).Weight := 2
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).Borders(9).LineStyle := 1
            excel.range(parta ":" partb).Borders(9).Weight := 2
        }
    }
    objrelease(excel)
    return
}


<MSE_边框上框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(8).Weight := 2
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(8).Weight := 2
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(8).LineStyle := 1
            excel.range(address).Borders(8).Weight := 2
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).Borders(8).LineStyle := 1
            excel.range(parta ":" partb).Borders(8).Weight := 2
        }
    }
    objrelease(excel)
    return
}


<MSE_边框左框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(7).Weight := 2
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(7).Weight := 2
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(7).LineStyle := 1
            excel.range(address).Borders(7).Weight := 2
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).Borders(7).LineStyle := 1
            excel.range(parta ":" partb).Borders(7).Weight := 2
        }
    }
    objrelease(excel)
    return
}

<MSE_边框右框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(10).LineStyle := 1
            excel.Selection.Borders(10).Weight := 2
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(10).LineStyle := 1
            excel.Selection.Borders(10).Weight := 2
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(10).LineStyle := 1
            excel.range(address).Borders(10).Weight := 2
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).Borders(10).LineStyle := 1
            excel.range(parta ":" partb).Borders(10).Weight := 2
        }
    }
    objrelease(excel)
    return
}

<MSE_边框无框线>:
{
    Excel_ActiveCell()
    excel.Selection.Borders(7).LineStyle := -4142
    excel.Selection.Borders(8).LineStyle := -4142
    excel.Selection.Borders(9).LineStyle := -4142
    excel.Selection.Borders(10).LineStyle := -4142
    excel.Selection.Borders(11).LineStyle :=-4142
    excel.Selection.Borders(12).LineStyle :=-4142
    objrelease(excel)
    return
}

<MSE_边框所有框线>:
{
    Excel_ActiveCell()
    excel.Selection.Borders(7).LineStyle := 1
    excel.Selection.Borders(8).LineStyle := 1
    excel.Selection.Borders(9).LineStyle := 1
    excel.Selection.Borders(10).LineStyle :=1
    excel.Selection.Borders(11).LineStyle :=1
    excel.Selection.Borders(12).LineStyle :=1
    excel.Selection.Borders(7).Weight := 2
    excel.Selection.Borders(8).Weight := 2
    excel.Selection.Borders(9).Weight := 2
    excel.Selection.Borders(10).Weight :=2
    excel.Selection.Borders(11).Weight := 2
    excel.Selection.Borders(12).Weight :=2
    objrelease(excel)
    return
}

<MSE_边框四边框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(10).LineStyle :=1
            excel.Selection.Borders(7).Weight := 2
            excel.Selection.Borders(8).Weight := 2
            excel.Selection.Borders(9).Weight := 2
            excel.Selection.Borders(10).Weight :=2
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(10).LineStyle :=1
            excel.Selection.Borders(7).Weight := 2
            excel.Selection.Borders(8).Weight := 2
            excel.Selection.Borders(9).Weight := 2
            excel.Selection.Borders(10).Weight :=2
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(7).LineStyle := 1
            excel.range(address).Borders(8).LineStyle := 1
            excel.range(address).Borders(9).LineStyle := 1
            excel.range(address).Borders(10).LineStyle :=1
            excel.range(address).Borders(7).Weight := 2
            excel.range(address).Borders(8).Weight := 2
            excel.range(address).Borders(9).Weight := 2
            excel.range(address).Borders(10).Weight :=2
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50
            excel.range(parta ":" partb).Borders(7).LineStyle := 1
            excel.range(parta ":" partb).Borders(8).LineStyle := 1
            excel.range(parta ":" partb).Borders(9).LineStyle := 1
            excel.range(parta ":" partb).Borders(10).LineStyle :=1
            excel.range(parta ":" partb).Borders(7).Weight := 2
            excel.range(parta ":" partb).Borders(8).Weight := 2
            excel.range(parta ":" partb).Borders(9).Weight := 2
            excel.range(parta ":" partb).Borders(10).Weight :=2
        }
    }
    objrelease(excel)
    return
}

<MSE_边框四边粗匣框线>:
{
    Excel_ActiveCell()
    MSE_GetSelectionType()
    MSE_GetSelectionInfo()
    ;msgbox,%SelectionType%
    if SelectionType=1
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(10).LineStyle :=1
            excel.Selection.Borders(7).Weight := -4138
            excel.Selection.Borders(8).Weight := -4138
            excel.Selection.Borders(9).Weight := -4138
            excel.Selection.Borders(10).Weight :=-4138
        }
    else if SelectionType=2
        {
            excel.Selection.Borders(7).LineStyle := 1
            excel.Selection.Borders(8).LineStyle := 1
            excel.Selection.Borders(9).LineStyle := 1
            excel.Selection.Borders(10).LineStyle :=1
            excel.Selection.Borders(7).Weight := -4138
            excel.Selection.Borders(8).Weight := -4138
            excel.Selection.Borders(9).Weight := -4138
            excel.Selection.Borders(10).Weight :=-4138
        }
    else if SelectionType=4 ;A1:A4
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            excel.range(address).Borders(7).LineStyle := 1
            excel.range(address).Borders(8).LineStyle := 1
            excel.range(address).Borders(9).LineStyle := 1
            excel.range(address).Borders(10).LineStyle :=1
            excel.range(address).Borders(7).Weight := -4138
            excel.range(address).Borders(8).Weight := -4138
            excel.range(address).Borders(9).Weight := -4138
            excel.range(address).Borders(10).Weight :=-4138
        }
    }
    else if SelectionType=16
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            address:=excel.selection.rows(A_Index).address
            StringReplace, address, address, $,,All
            ;msgbox,%address%
            FoundPosSeperate := RegExMatch(address,":")
            StringLeft, parta, address, FoundPosSeperate-1
            StringMid, partb, address, FoundPosSeperate+1 , 50

            excel.range(parta ":" partb).Borders(7).LineStyle := 1
            excel.range(parta ":" partb).Borders(8).LineStyle := 1
            excel.range(parta ":" partb).Borders(9).LineStyle := 1
            excel.range(parta ":" partb).Borders(10).LineStyle :=1
            excel.range(parta ":" partb).Borders(7).Weight := -4138
            excel.range(parta ":" partb).Borders(8).Weight := -4138
            excel.range(parta ":" partb).Borders(9).Weight := -4138
            excel.range(parta ":" partb).Borders(10).Weight :=-4138
        }
    }
    objrelease(excel)
    return
}

<MSE_边框粗匣框线>:
{
    Excel_ActiveCell()
    excel.Selection.Borders(7).LineStyle := 1
    excel.Selection.Borders(8).LineStyle := 1
    excel.Selection.Borders(9).LineStyle := 1
    excel.Selection.Borders(10).LineStyle :=1
    excel.Selection.Borders(7).Weight := -4138
    excel.Selection.Borders(8).Weight := -4138
    excel.Selection.Borders(9).Weight := -4138
    excel.Selection.Borders(10).Weight :=-4138
    objrelease(excel)
    return
}

<MSE_边框上下框线>:
{
    Excel_ActiveCell()
    excel.Selection.Borders(5).LineStyle := -4142
    excel.Selection.Borders(6).LineStyle := -4142
    excel.Selection.Borders(7).LineStyle := -4142

    excel.Selection.Borders(8).LineStyle := 1
    excel.Selection.Borders(8).ColorIndex := 0
    excel.Selection.Borders(8).TintAndShade := 0
    excel.Selection.Borders(8).Weight := 2

    excel.Selection.Borders(9).LineStyle := 1
    excel.Selection.Borders(9).ColorIndex := 0
    excel.Selection.Borders(9).TintAndShade := 0
    excel.Selection.Borders(9).Weight := 2

    excel.Selection.Borders(10).LineStyle := -4142
    excel.Selection.Borders(11).LineStyle := -4142
    excel.Selection.Borders(12).LineStyle := -4142
    objrelease(excel)
    return
}

;编辑

<MSE_编辑插入新行在后>:
{
    send,{down}
    send,{AppsKey}
    send,i
    send,{enter}
    sleep,5
    send,r
    send,{enter}
    return
}

<MSE_编辑插入新列在右>:
{
    send,{right}
    send,{AppsKey}
    send,i
    send,{enter}
    sleep,5
    send,c
    send,{enter}
    return
}

<MSE_Copy_Selection>:
{
    send ^c
    return
}

<MSE_Copy_Row>:
{
    Excel_Selection()
    Selection.EntireRow.Select
    objrelease(excel)
    send ^c
    return
}

<MSE_Copy_Col>:
{
    Excel_Selection()
    Selection.EntireColumn.Select
    objrelease(excel)
    send ^c
    return
}

<MSE_Paste>:
{
    send ^v
    return
}

<MSE_Paste_Select>:
{
    send ^!v
    return
}

<MSE_Paste_Value>:
{
    send ^!v!v{enter}
    return
}

<MSE_Color_Font>:
{
    getExcel().Selection.Font.Color := FontColor
    return
}



<MSE_Color_Cell>:
{
    getExcel().Selection.Interior.Color := CellColor
    return
}

<MSE_Color_All>:
{
    getExcel().Selection.Font.Color := FontColor
    getExcel().Selection.Interior.Color := CellColor
    return
}

<MSE_Color_Menu_Font>:
{
    InputColor(color)

    if color = null
        return
    if color = Transparent
    {
        MsgBox 字体颜色不支持透明色
        return
    }

    FontColor := ToBGR(color)
    getExcel().Selection.Font.Color := FontColor
    return
}

<MSE_Color_Menu_Cell>:
{
    InputColor(color)

    if color = null
        return

    if color = Transparent
    {
        getExcel().Selection.Interior.Pattern := -4142
        return
    }

    CellColor := ToBGR(color)
    getExcel().Selection.Interior.Color := CellColor
    return
}

;编辑复制
<MSE_编辑自左侧复制>:
{
    send,{left}
    send,^c
    send,{right}
    send,^v
    return
}

<MSE_逐行编辑自左侧复制>:
{
    Excel_ActiveCell()
    if excel.Selection.Columns.Count = 1
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            send,{left}
            send,^c
            send,{right}
            send,^v
            sleep,500
            send,{down}
        }
    }
    objrelease(excel)
    return
}


<MSE_编辑自右侧复制>:
{
    send,{right}
    send,^c
    send,{left}
    send,^v
    return
}

<MSE_逐行编辑自右侧复制>:
{
    Excel_ActiveCell()
    if excel.Selection.Columns.Count = 1
    {
        rowcount:=excel.selection.rows.count
        Loop, %rowcount%
        {
            send,{right}
            send,^c
            send,{left}
            send,^v
            sleep,500
            send,{down}
        }
    }
    objrelease(excel)
    return
}


<MSE_编辑自上侧复制>:
{
    send,{up}
    send,^c
    send,{down}
    send,^v
    return
}

<MSE_编辑自下侧复制>:
{
    send,{down}
    send,^c
    send,{up}
    send,^v
    return
}

;定位
<MSE_FocusHome>:
{
    send,^{Home}
    return
}

<MSE_FocusEnd>:
{
    send ^{End}
    return
}

; 模拟输入9个Ctrl+Up,差不多能到行首了
<MSE_FocusRowHome>:
{
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    send ^{Up}
    return
}

; 模拟输入9个Ctrl+Down，差不多能到尾行了
<MSE_FocusRowEnd>:
{
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    send ^{Down}
    return
}

; 快捷键Home可直接定位到首列
<MSE_FocusColHome>:
{
    send,{Home}
    return
}

; 貌似没有快捷键直接定位到尾列--同时该功能貌似没什么作用...
<MSE_FocusColEnd>:
{
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    send,^{Right}
    return
}

<MSE_FocusAreaUp>:
{
    send,^{Up}
    return
}

<MSE_FocusAreaDown>:
{
    send,^{Down}
    return
}

<MSE_FocusAreaLeft>:
{
    send,^{Left}
    return
}

<MSE_FocusAreaRight>:
{
    send,^{Right}
    return
}

;对齐
<MSE_对齐左>:
{
    Excel_Selection()
    Selection.HorizontalAlignment := -4131
    ;objRelease(excel)
    return
}

<MSE_对齐水平中间>:
{
    Excel_Selection()
    Selection.HorizontalAlignment := -4108
    ;objRelease(excel)
    return
}

<MSE_对齐右>:
{
    Excel_Selection()
    Selection.HorizontalAlignment := -4152
    ;objRelease(excel)
    return
}

<MSE_对齐顶>:
{
    Excel_Selection()
    Selection.VerticalAlignment := -4160
    ;objRelease(excel)
    return
}

<MSE_对齐垂直中间>:
{
    Excel_Selection()
    Selection.VerticalAlignment := -4108
    ;objRelease(excel)
    return
}

<MSE_对齐底>:
{
    Excel_Selection()
    Selection.VerticalAlignment := -4107
    ;objRelease(excel)
    return
}

;单元格颜色
<MSE_单元格颜色黑>:
{
    Excel_Selection()
    Selection.Interior.color:= 0x000000
    ;objRelease(excel)
    return
}

;字体命令
<XLmain_字体缩小>:
{
    Excel_Selection()
    currentFontSize := Selection.Font.Size
    Selection.Font.Size := currentFontSize - 1
    ;objRelease(excel)
    return
}

<XLmain_字体放大>:
{
    Excel_Selection()
    currentFontSize := Selection.Font.Size
    Selection.Font.Size := currentFontSize + 1
    ;objRelease(excel)
    return
}

<excel_find>:
{
    GUI,XLFind:Destroy
    GUI,XLFind:Add,Edit,w200 h20 gXLFind
    GUI,XLFind:Add,Button,w50 x160 center Default,确定
    GUI,XLFind:Show
    return
}

;工作表
<MSE_SheetReName>:
{
    InputBox, NewSheetName , Please input the new name for active sheet
    If ErrorLevel
        Return
    if StrLen(NewSheetName) > 0
    {
        getExcel().ActiveSheet.Name := NewSheetName
    }
    return
}

<MSE_CopyCurrentSheet>:
{
    After:=getExcel().ActiveSheet
    getExcel().ActiveSheet.Copy(After)
    return
}

<MSE_FirstSheet>:
{
    getExcel().Worksheets(1).Select
    return
}

<MSE_LastSheet>:
{
    getExcel().Worksheets(getExcel().Worksheets.Count).Select
    return
}

<MSE_工作表复制对话框>:
{
    Excel_ActiveSheet()
    excel.Application.Dialogs(283).Show
    ;objRelease(excel)
    return
}

<MSE_NextSheet>:
{
    xls := getExcel()
    If xls.ActiveSheet.index = xls.Worksheets.Count
        xls.Worksheets(1).Select
    Else
        xls.ActiveSheet.Next.Select
    return
}

<MSE_PreviousSheet>:
{
    xls := getExcel()
    If xls.ActiveSheet.index =1
        xls.Worksheets(xls.Worksheets.Count).Select
    Else
        xls.ActiveSheet.Previous.Select
    return
}

<MSE_GoTo>:
{
    Excel_ActiveSheet()
    InputBox, Reference , 输入跳转到的位置，如B5/b5：第二列，第5行
    If ErrorLevel
        Return
    excel.ActiveSheet.Range(Reference).Select
    ;objRelease(excel)
    return
}

<MSE_SaveAndExit>:
{
    send ^s
    send !{F4}
    return
}

<MSE_DiscardAndExit>:
{
    getExcel().ActiveWorkbook.Saved := true
    getExcel().Quit
    return
}

<MSE_工作表移动向后>:
{
    If excel.ActiveSheet.index < excel.Worksheets.Count - 1
    {
        After :=excel.Sheets(excel.ActiveSheet.index + 2)
        getExcel().ActiveSheet.Move(After)
    }
    Else
    {
        getExcel().Sheets(excel.Worksheets.Count).Move(excel.ActiveSheet)
        getExcel().Sheets(excel.Worksheets.Count).select
    }
    ;objRelease(excel)
    return
}

<MSE_工作表移动向前>:
{
    excel := getExcel()
    getExcel().ActiveSheet.Select
    count:=getExcel().Worksheets.Count
    If getExcel().ActiveSheet.index = count
    {
        Before:=getExcel().Sheets(1)
        getExcel().ActiveSheet.Move(Before)
    }
    Else
    {
        index:=getExcel().ActiveSheet.index + 1
        After:=getExcel().Sheets(index)
        getExcel().ActiveSheet.Move(After)
    }
    ;objRelease(excel)
    return
}

;填充

<MSE_填充向下>:
{
    Excel_Selection()
    Selection.FillDown
    ;objRelease(excel)
    return
}

<MSE_填充向上>:
{
    Excel_Selection()
    Selection.FillUp
    ;objRelease(excel)
    return
}

<MSE_填充向左>:
{
    Excel_Selection()
    Selection.FillLeft
    ;objRelease(excel)
    return
}

<MSE_填充向右>:
{
    Excel_Selection()
    Selection.FillRight
    ;objRelease(excel)
    return
}

;===================================================================
;直接获取Excel
getExcel()
{
    try xls := ComObjActive("Excel.Application") ;handle to running application
    Catch {
        MsgBox % "no existing Excl ojbect:  Need to create one"
        xls := ComObjCreate("Excel.Application")
    }
    return xls
}

Excel_ActiveSheet()
{
    ;objRelease(excel)
    Sheet := getExcel().ActiveSheet ; 当前工作表
    return 
}

Excel_ActiveCell()
{
    ;objRelease(excel)
    ;excel := ComObjCreate("Excel.Application") ; 创建Excel对象
    Cell := getExcel().ActiveCell ; 当前单元格
    return
}

Excel_Selection()
{
    ;objRelease(excel)
    ;excel := ComObjCreate("Excel.Application") ; 创建Excel对象
    Selection:=getExcel().Selection ;选择对象
    return
}

Excel_Direction(x=0, y=0)
{
    objExcel := Excel_GetObj()
    app  := ObjExcel.Application
    cell := app.ActiveCell
    addr := Cell_Address(cell)
    x_new := CharCalc(Addr["x"],x)
    y_new := addr["y"] + y
    If y_new < 1
        y_new := 1
    new := "$" x_new "$" y_new
    app.range(new).Activate
}

Excel_CellActivate(Location){
    objExcel := Excel_GetObj()
    app  := ObjExcel.Application
    app.range(Location).Activate
}

CharCalc(char,count)
{
    StringUpper,Char,Char
    SingleChars := []
    NumberChars := []
    ReturnChars := []
    MaxBit := Strlen(char)
    SingleChars[0] := MaxBit
    Loop,Parse,Char
    {
        Pos := MaxBit - A_Index + 1
        SingleChars[Pos] := Asc(A_LoopField)-64
    }
    ; 800 => abcd
    idx := 26
    Loop
    {
        If count >= %idx%
            idx := idx * idx
        Else {
            MaxBit := A_Index
            Break
        }
    }
    NumberChars[0] := MaxBit
    Loop % MaxBit
    {
        Pos := MaxBit - A_index
        NumberChars[Pos+1] := Floor(Count/26**Pos)
        count := Mod(count,26**Pos)
    }
    s := SingleChars[0]
    n := NumberChars[0]
    If s > %n%
        MaxBit := s
    Else
        MaxBit := n
    Pos := 1
    Add := 0
    Loop,% MaxBit
    {
        s := SingleChars[Pos]
        If not strlen(s)
            s := 0
        n := NumberChars[Pos]
        If not strlen(n)
            n := 0
        r := ReturnChars[Pos]
        If not strlen(r)
            r := 0
        sum := s + n + r
        If sum > 26
        {
            sum := sum - 26
            ReturnChars[Pos+1] := 1
        }
        ReturnChars[Pos] := sum
        Pos++
    }
    msg := ""
    For  i , k in ReturnChars
        msg := Chr(k+64) msg
    return msg
}

Cell_Address(cell)
{
    addr := []
    OldAddr := Cell.Address()
    Loop,Parse,OldAddr,$
    {
        If not Strlen(A_LoopField)
            Continue
        If Strlen(addr["x"])
            addr["y"] := A_LoopField
        Else
            addr["x"] := A_LoopField
    }
    return addr
}

Excel_GetObj()
{
    ControlGet, hwnd, hwnd, , Excel71, ahk_class MicrosoftExcel
    ObjExcel := Excel[hwnd]
    If IsObject(ObjExcel)
        return ObjExcel
    ObjExcel := Acc_ObjectFromWindow(hwnd, -16)
    Excel[hwnd] := ObjExcel

    return ObjExcel
}

MSE_获取Range地址(address)
{
    StringReplace, address, address, $,,All
    return
}

XLMIAN_获取活动工作表边界()
{
    lLastRow := excel.Cells(excel.Rows.Count, 1).End(-4162).Row
    lLastColumnAddress := excel.Cells(1, excel.Columns.Count).End(-4159).Address
    StringReplace, lLastColumnAddress, lLastColumnAddress, $,,All
    RegExMatch(lLastColumnAddress,"[A-Z]+",lLastColumn)
    return
}


<LastColumn>:
{
    Excel_Selection()
    XLMIAN_获取活动工作表边界()
    msgbox,lLastRow %lLastRow% lLastColumn %lLastColumn%
    MSE_GetSelectionType()

    if SelectionType = 1
    {
        rng:=excel.Selection
        ;msgbox,rng %rng%
        MSE_GetSelectionInfo()
        ;msgbox,%SelectFirstColumn%%SelectFirstRow% %SelectionLastColumn%%SelectionLastRow%
        ;msgbox,%SelectFirstColumn%%SelectFirstRow%:%SelectFirstColumn%%lLastRow%
        excel.range(SelectionFirstColumn SelectionFirstRow ":" SelectionFirstColumn lLastRow ).Select ;填充列
        rng.AutoFill(excel.selection,9)
        ;objRelease(excel)
    }
    else if SelectionType = 4
    {
        rng:=excel.Selection
        MSE_GetSelectionInfo()

        excel.range(SelectionFirstColumn SelectionFirstRow ":" SelectionFirstColumn lLastRow).Select
        rng.AutoFill(excel.selection,9)
        ;objRelease(excel)
    }
    Return
}

<LastRow>:
{
    Excel_Selection()

    MSE_GetSelectionType()
    XLMIAN_获取活动工作表边界()
    if SelectionType=1
    {
        rng:=excel.Selection
        MSE_GetSelectionInfo()
        excel.range(SelectionFirstColumn SelectionFirstRow ":" lLastColumn SelectionFirstRow).select
        rng.AutoFill(excel.selection,9)

        ;objRelease(excel)
    }
    else if SelectionType=2
    {
        rng:=excel.Selection
        MSE_GetSelectionInfo()
        excel.range(SelectionFirstColumn SelectionFirstRow ":" lLastColumn SelectionFirstRow).select
        rng.AutoFill(excel.selection,9)

        ;objRelease(excel)
    }
    else
        ;objRelease(excel)
        Return
}

MSE_ColToChar(index)
{
    If(index <= 26)
    {
        return Chr(64+index)
    }
    Else If (index > 26)
    {
        return Chr((index-1)/26+64) . Chr(mod((index - 1),26)+65)
    }
}

XLMIAN_获取Range边界(address)
{
    StringReplace, address, address, $,,All
    FoundPosSeperate := RegExMatch(address,":")
    StringLeft, parta, address, FoundPosSeperate-1
    StringMid, partb, address, FoundPosSeperate+1 , 50
    RegExMatch(parta,"[A-Z]+",ColumnLeftName)
    RegExMatch(parta,"[0-9]+",RowUp)
    ;msgbox,ColumnLeftName%ColumnLeftName% RowUp %RowUp%

    RegExMatch(partb,"[A-Z]+",ColumnRightName)
    RegExMatch(partb,"[0-9]+",RowDown)
    return
}

MSE_GetSelectionType()
{
    if excel.Selection.Columns.Count =1 And excel.Selection.Rows.Count =1 ;A1
    {
        SelectionType:=1
    }
    else if excel.Selection.Columns.Count >1 And excel.Selection.Rows.Count =1 ;A1:B1
    {
        SelectionType:=2
    }
    else if excel.Selection.Columns.Count =1 And excel.Selection.Rows.Count >1 ;A1:A2
    {
        SelectionType:=4
    }
    else if excel.Selection.Columns.Count >1 And excel.Selection.Rows.Count >1  ;A1:B2
    {
        SelectionType:=16
    }
    else
        return
}

MSE_GetSelectionInfo()
{
    address:=excel.Selection.Address
    ;msgbox,address %address%
    StringReplace, address, address, $,,All
    ;msgbox,address %address%

    if SelectionType = 1 ;A1
    {
        ;msgbox,SelectionType %SelectionType%
        RegExMatch(address,"[A-Z]+",SelectionFirstColumn)
        RegExMatch(address,"[0-9]+",SelectionFirstRow)
        ;msgbox,SelectionFirstColumn %SelectionFirstColumn%  SelectionFirstRow %SelectionFirstRow%
        SelectionLastColumn:=SelectionFirstColumn
        SelectionLastRow:=SelectionFirstRow

        return
    }
    else if SelectionType = 2 ;A1:B1
    {
        FoundPosSeperate := RegExMatch(address,":")
        StringLeft, parta, address, FoundPosSeperate-1
        StringMid, partb, address, FoundPosSeperate+1 , 50

        RegExMatch(parta,"[A-Z]+",SelectionFirstColumn)
        RegExMatch(parta,"[0-9]+",SelectionFirstRow)


        RegExMatch(partb,"[A-Z]+",SelectionLastColumn)
        SelectionLastRow:=SelectionFirstRow


    }
    else if SelectionType = 4 ;A1:A2
    {
        FoundPosSeperate := RegExMatch(address,":")
        StringLeft, parta, address, FoundPosSeperate-1
        StringMid, partb, address, FoundPosSeperate+1 , 50

        RegExMatch(parta,"[A-Z]+",SelectionFirstColumn)
        RegExMatch(parta,"[0-9]+",SelectionFirstRow)

        SelectionLastColumn:=SelectionFirstColumn
        RegExMatch(partb,"[0-9]+",SelectionLastRow)
    }
    else if SelectionType = 16  ;A1:B2
    {
        FoundPosSeperate := RegExMatch(address,":")
        StringLeft, parta, address, FoundPosSeperate-1
        StringMid, partb, address, FoundPosSeperate+1 , 50
        RegExMatch(parta,"[A-Z]+",SelectionFirstColumn)
        RegExMatch(parta,"[0-9]+",SelectionFirstRow)
        ;msgbox,ColumnLeftName%ColumnLeftName% RowUp %RowUp%

        RegExMatch(partb,"[A-Z]+",SelectionLastColumn)
        RegExMatch(partb,"[0-9]+",SelectionLastRow)
    }
    else
        return
}

XLFind:
    XLFind()
return

XLFind()
{
    ControlGetText, findstring, Edit1, A
    If not Strlen(findstring)
        return
    If RegExMatch(findstring,"^[a-zA-Z]*$")
        Excel_CellActivate(findstring "1")
    Else
        Excel_CellActivate(findstring)
}

<excel_null>:
return

<excel_replace>:
{
    ClipSaved := ClipboardAll

    Send,^c
    ClipWait
    String := Clipboard

    Clipboard := ClipSaved
    ClipSaved =

    GUI,Replace:Destroy
    GUI,Replace:Add,Edit,w400 h300 ,%String%
    GUI,Replace:show

    return
}

#Include %A_ScriptDir%\plugins\MicrosoftExcel\InputColor.ahk
