/**
 * 缺点：MSScriptControl.ScriptControl 只在 x86 环境下有，所以我在前面加入了切换到 x86 模式的代码。
 * Disadvantages: MSScriptControl.ScriptControl is only available in x86 environments, so I added the code to switch to x86 mode in the front.
 */
var wsh = new ActiveXObject('WScript.Shell');
function isWin64(){
	var procEnv = wsh.Environment("Process");
	var architecture = procEnv("PROCESSOR_ARCHITECTURE");
	if (/AMD64|X64/i.test(architecture)) return true;
	if (/x86/i.test(architecture)) return !Boolean(procEnv("PROCESSOR_ARCHITEW6432"));
	return true;
}
if (isWin64() && /System32/i.test(WScript.FullName)) {
	var x86WScriptExe = WScript.FullName.replace(/\\System32\\/i,"\\SysWOW64\\");
	var command = '"' + x86WScriptExe + '" "' + WScript.ScriptFullName + '"';
	if (WScript.Arguments.length) command += ' ' + WScript.Arguments.join(' ');
    wsh.Run(command);
	WScript.Quit();
}

/**
 * 以下代码来自：
 * The following code is from:
 * https://demon.tw/programming/javascript-vbs-inpubox-msgbox.html
 */

function MsgBox(prompt, buttons, title)
{
	if(buttons == undefined) buttons = 0;
	if(title == undefined) title = "";

	var sc = new ActiveXObject("MSScriptControl.ScriptControl");
	var code = 'MsgBox("'+ prompt + '","' + buttons + '","' + title + '")';

	sc.Language = "VBScript";
	return sc.Eval(code);
}

function InputBox(prompt, title, def)
{
	if(title == undefined) title = "";
	if(def == undefined) def = "";
	
	var sc = new ActiveXObject("MSScriptControl.ScriptControl");
	var code = 'Inputbox("'+ prompt + '","' + title + '","' + def + '")';
	
	sc.Language = "VBScript";
	result = sc.Eval(code);
	
	if(result != undefined) return result;
	else return "";
}

/**
 * 测试代码
 * Test Code
 */
var str = InputBox("Enter your name");
if(str != "")
{
	var res = MsgBox("Are you " + str + "?", 4);
	if(res == 6)
		MsgBox("Yes,you are " + str);
	else
		MsgBox("No,you are not " + str);
}