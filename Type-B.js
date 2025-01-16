/**
 * 缺点：InputBox无法输入多行提示，InputBox弹出框确定后的MsgBox不在前台。
 * Disadvantages: InputBox cannot use multi-line prompts, and the MsgBox after the InputBox pop-up box is confirmed is not in the foreground.
 */
(function() {
	var wsh = new ActiveXObject('WScript.Shell');
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	this.VbMsgBoxStyle = {
		"vbOKOnly":0,
		"vbOKCancel":1,
		"vbAbortRetryIgnore":2,
		"vbYesNoCancel":3,
		"vbYesNo":4,
		"vbRetryCancel":5,
		"vbCritical":16,
		"vbQuestion":32,
		"vbExclamation":48,
		"vbInformation":64,
		"vbDefaultButton1":0,
		"vbDefaultButton2":256,
		"vbDefaultButton3":512,
		"vbDefaultButton4":768,
		"vbApplicationModal":0,
		"vbSystemModal":4096,
		"vbMsgBoxHelpButton":16384,
		"vbMsgBoxSetForeground":65536,
		"vbMsgBoxRight":524288,
		"vbMsgBoxRtlReading":1048576
	};
	this.VbMsgBoxResult = {
		"vbOK":1,
		"vbCancel":2,
		"vbAbort":3,
		"vbRetry":4,
		"vbIgnore":5,
		"vbYes":6,
		"vbNo":7
	};
	
	this.MsgBox = function(strText,nType,strTitle,nSecondsToWait) {
		nSecondsToWait = nSecondsToWait ? nSecondsToWait : 0;
		strTitle = strTitle || "";
		nType = nType || 0;
		return wsh.Popup(strText, nSecondsToWait, strTitle, nType);
	};
	this.alert = this.MsgBox;
	this.confirm = function(strText) {
		return MsgBox(strText, VbMsgBoxStyle.vbQuestion | VbMsgBoxStyle.vbOKCancel) == VbMsgBoxResult.vbOK;
	}
	
	this.InputBox = function(sPrompt,sTitle,sDefault) {
		sPrompt = sPrompt || "";
		sTitle = sTitle || "";
		sDefault = sDefault || "";
		var iptvbsCode = 
			[
				"Dim prompt,title,default" ,
				"prompt = WScript.StdIn.ReadLine()" ,
				"title = WScript.StdIn.ReadLine()" ,
				"default = WScript.StdIn.ReadLine()" ,
				"Dim input" ,
				"input = InputBox(prompt,title,default)" ,
				"If input = False Then" ,
				"	WScript.StdOut.WriteLine(0)" ,
				"Else" ,
				"	WScript.StdOut.WriteLine(-1)" ,
				"End If" ,
				"WScript.StdOut.WriteLine(input)"
			].join("\r\n");
        //var debug = 1;
        var iptvbsPath;
        iptvbsPath = (typeof(debug) != "undefined") ? WScript.ScriptName + ".vbs" : wsh.Environment("Process").Item("temp") + "\\" + WScript.ScriptName + ".vbs";
        var iptvbs = fso.OpenTextFile(iptvbsPath,2,true,-1);
		iptvbs.Write(iptvbsCode);
		iptvbs.Close();
		var oExec = wsh.Exec("wscript.exe \"" + iptvbsPath + "\"");
		oExec.StdIn.WriteLine(sPrompt.replace("\n","")) ;
		oExec.StdIn.WriteLine(sTitle.replace("\n","")) ;
		oExec.StdIn.WriteLine(sDefault.replace("\n","")) ;
		var noNull = parseInt(oExec.StdOut.ReadLine());
		var input = oExec.StdOut.ReadLine();
		if (!noNull) return false;
		return input;
	};
	this.prompt = this.InputBox;
})();

/**
* 测试代码
* Test Code
*/

var str = prompt("Enter your name");
if(str != "")
{
	var res = confirm("Are you " + str + "?");
	if(res)
		alert("Yes,you are " + str);
	else
		alert("No,you are not " + str);
}