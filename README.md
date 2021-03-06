## ![icon](doc/images/vimdesktop_32.jpg) VimDesktop


Update History by Lu Da Jun

- Sep. 09 2018: Adding basic support for Windows Desktop
- Sep. 02 2018: Pending issue that can not input upper letters when CapsLock is on and using Microsoft Wubi 
- Aug. 22 2018: Simple Adobe Acrobat Reader DC support added
- Aug. 20 2018: Maximize & Restore window for Outlook added
- Aug. 17 2018: 
    - MSO_CopySelectedEmailFromMainOutlookWindow added
    - Modify the function of key h, j to support move left and right when composing email
- Aug. 13 2018: 
    - Add fv key binding to paste from clipboard
        - Using fv when composing new email will paste from clipboard

        - It would be useful to using fv in main Outlook window if you have already copied some attachments into clipboard. This action will:
            1. create a new email and will paste the attachment
            2. set email subject to the file name
    - More functions added or adjusted...
- Aug. 11 2018: Add support for Outlook

Todo List:
- enable ff&fv when reply/forward/new email instead of showing a find/replace messagebox when f key triggered

### Below information referenced from original developer:
---
让所有 Windows 桌面程序拥有 Vim 操作风格的辅助工具。

### 链接

[详细介绍](https://github.com/goreliu/vimdesktop/wiki)

[Introduction](https://github.com/goreliu/vimdesktop/wiki/VimDesktop-Introduction)

[下载地址](https://github.com/goreliu/vimdesktop/releases/latest)

[文档首页](https://github.com/goreliu/vimdesktop/wiki)

[TC 快捷键列表](https://github.com/goreliu/vimdesktop/wiki/TC%E5%BF%AB%E6%8D%B7%E9%94%AE%E5%88%97%E8%A1%A8)

[API 列表](https://github.com/goreliu/vimdesktop/wiki/API%E5%88%97%E8%A1%A8)

[更新历史](https://github.com/goreliu/vimdesktop/wiki/HISTORY)

### VimDesktop 历史

VimDesktop 的前身是 [linxinhong](http://git.oschina.net/linxinhong) 的 ViATc （现已停止更新）：[github地址](https://github.com/linxinhong/ViATc) [sourceforge地址](https://sourceforge.net/p/viatc/home/%E4%B8%BB%E9%A1%B5/)

之后 [linxinhong](http://git.oschina.net/linxinhong) 将其升级为 [VimDesktop 1](https://github.com/victorwoo/vimdesktop)，该版本由 [linxinhong](http://git.oschina.net/linxinhong)、[victorwoo](https://github.com/victorwoo)、[wideweide](https://github.com/wideweide) 等人协同开发。

再之后 [linxinhong](http://git.oschina.net/linxinhong) 又将其升级为 [VimDesktop 2](http://git.oschina.net/linxinhong/VimDesktop)，2版本的核心文件和1版本相比改动较大，配置文件格式也有所不同。

目前 [linxinhong](http://git.oschina.net/linxinhong) 将主要精力放在 [QuickZ](http://git.oschina.net/linxinhong/QuickZ) 上， [VimDesktop 2](http://git.oschina.net/linxinhong/VimDesktop) 已有较长时间没有更新。

### 此版本 VimDesktop 的历史和缘由

我2016年1月接触到的 VimDesktop，当时在网上搜到了两个版本的 VimDesktop，如上所述。我分别试用后感觉 [VimDesktop 1](https://github.com/victorwoo/vimdesktop) 版本的好用些，主要是 `TotalCommander_Dialog` 插件很有用，配置文件也更方便些，当时并未考虑过修改代码。但使用过程中慢慢发现一些问题，或者有功能缺失，改了很多代码。修改过程中，发现这个版本的核心文件 `lib/vimcore.ahk` 缺陷比较多，功能也相对薄弱，而 [VimDesktop 2](http://git.oschina.net/linxinhong/VimDesktop) 的`lib/class_vim.ahk` 是 `lib/vimcore.ahk` 的升级版，功能更强大，缺陷也少些。于是改用 [VimDesktop 2](http://git.oschina.net/linxinhong/VimDesktop) 的部分核心文件，而插件部分还是沿用之前的代码，除了必要的兼容性改动。

此版本包含 [VimDesktop 1](https://github.com/victorwoo/vimdesktop) 和 [VimDesktop 2](http://git.oschina.net/linxinhong/VimDesktop) 的全部功能。

