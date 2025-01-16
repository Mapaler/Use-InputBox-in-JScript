# 在 WindowsScriptHost 的 JScript 内使用 prompt/InputBox<br>Use prompt/InputBox in JScript on WindowsScriptHost
> 注意：[JScript](https://learn.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/hbxc2t98(v=vs.84)) 不是现在浏览器或 Node.js 里的 JavaScript 语言。  
> Notice: [JScript](https://learn.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/hbxc2t98(v=vs.84)) is not current JavaScript language in Browser or Node.js.
>
> 和 VBS 一样，你可以双击使用 WScript 运行，或使用 CScript 在命令行内运行。  
> As with VBS, you can double-click to run with WScript, or use CScript to run from within the command line.

## 背景<br>Background
WindowsScriptHost 执行速度比 PowerShell 快得多，并且兼容性也更广，从 Win XP 到 Win 11 都能支持。所以我写的很多脚本是用 WindowsScriptHost 运行的。  
WindowsScriptHost executes much faster than PowerShell and is more compatible, from Win XP to Win 11. So many of the scripts I write run in WindowsScriptHost.

WindowsScriptHost 支持 VBScript 和 JScript 两种语言，基于 ECMAScript3 的 JScript 语法写起来肯定是比 VBScript 要简单快捷的，还可以支持一些新对象的 polyfill。  
WindowsScriptHost supports VBScript and JScript, and the JScript syntax based on ECMAScript3 is certainly easier and faster to write than VBScript, and it can also support polyfills for some new objects.

但是 JScript 不支持 VBS 的 InputBox （甚至 C# 都不支持，也得调用 VB.Net），所以我收集和书写了几种实现方案来分享。这些原始代码基本都是 2015 年收集和书写的，但是我今天将他们进行了优化。  
However, JScript doesn't support VBS's InputBox (and even C# doesn't support it, so you have to call VB.Net), so I've collected and written a few implementations to share. Most of the original code was collected and written in 2015, but I've optimized it today.

## 解决方案<br>Solution
### Type-A: MSScriptControl.ScriptControl
Code File: [Type-A](Type-A.js)

原理：用 MSScriptControl.ScriptControl 对象运行 VBS 代码  
Principle: Run VBS code by MSScriptControl.ScriptControl Object

原始代码来自于 [Demon's Blog](https://demon.tw/programming/javascript-vbs-inpubox-msgbox.html) ，我 2015 年发现这个代码无法在 64 位 Windows 运行，经检查发现 MSScriptControl.ScriptControl 只有 32 位的文件，因此需要用 32 位 WScript 运行才能正常使用。我在这个示例代码的开头部分加入了 64 位系统重定向到 32 位 WScript 运行的代码。  
The original code comes from [Demon's Blog](https://demon.tw/programming/javascript-vbs-inpubox-msgbox.html), I found out in 2015 that this code does not work on 64-bit Windows, and upon inspection found that MSScriptControl.ScriptControl only has a 32-bit file, so I need to use 32-bit WScript Run to work properly. I've added code at the beginning of this sample code that redirects a 64-bit system to a 32-bit WScript run.

### Type-B: Another VBS
Code File: [Type-B](Type-B.js)

原理：在系统临时文件夹生成一个 VBS 文件用来执行 InputBox，通过标准输入流输入函数参数和标准输出流读取返回值。  
Principle: Generate a VBS file in the temporary folder of the system to execute the InputBox, and read the return value through the input function parameters of the standard input stream and the standard output stream.

这完全是我自己独立想出来的方法，用起来有以下缺点：  
This is a completely independent method that I came up with, and it has the following disadvantages when used to use it:
* 用行来分割读取参数内容，因此 InputBox 无法输入多行提示文本。  
Lines are used to split the content of the read parameter, so the InputBox cannot enter multiple lines of prompt text.
* 启动一个另外的进程执行 VBS ，InputBox 弹出框后的 MsgBox 不在前台。  
Start another process to execute VBS, and the MsgBox after the InputBox pop-up box is not in the foreground.

### Type-C: WSF file
Code File: [Type-C](Type-C.wsf)

原理：WSF 文件支持不同代码混写，可以使用其他代码的函数。所以在 WSF 文件里运行 VBS，来调用 InputBox  
Principle: WSF files can be mixed with different codes, and functions from other codes can be used. So run VBS in a WSF file to call InputBox

这是我发现 WSF 这种格式后想出来的方法，最大的缺点这种格式就是没有语法高亮支持，写东西很麻烦。  
This is the method I came up with after discovering the WSF format, and the biggest drawback of this format is that it doesn't have syntax highlighting support, and it's cumbersome to write.

## 官方文档链接<br>Link to official documentation
* [wscript.exe](https://learn.microsoft.com/windows-server/administration/windows-commands/wscript)
* [cscript.exe](https://learn.microsoft.com/windows-server/administration/windows-commands/cscript)
* [JScript and VBScript](https://learn.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1et7k7c(v=vs.84))
* [WScript Object](https://learn.microsoft.com/previous-versions//at5ydy31(v=vs.85))
* [Windows Script Host](https://learn.microsoft.com/previous-versions/windows/it-pro/windows-server-2003/cc784547(v=ws.10))