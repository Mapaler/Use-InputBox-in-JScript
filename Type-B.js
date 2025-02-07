/**
 * 缺点：InputBox弹出框后的MsgBox不在前台。
 * Disadvantages: the MsgBox after the InputBox pop-up box is not in the foreground.
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
	this.OpenTextFileIOMode = {
		ForReading: 1,
		ForWriting: 2,
		ForAppending: 8
	}
	this.OpenTextFileFormat = {
		TristateUseDefault: -2,
		TristateTrue: -1,
		TristateFalse: 0
	}
	
	this.MsgBox = function(strText,nType,strTitle,nSecondsToWait) {
		strText = strText || ""
		nType = nType || 0;
		strTitle = strTitle || "";
		nSecondsToWait = nSecondsToWait || 0;
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
				"Function GetNamedArgument(name, default)" ,
				"	If WScript.Arguments.Named.Exists(name) Then" ,
				"		GetNamedArgument = WScript.Arguments.Named.Item(name)" ,
				"	Else" ,
				"		GetNamedArgument = default" ,
				"	End If" ,
				"End Function" ,
				"Dim promptl,titlel,defaultl" ,
				"promptl = CInt(GetNamedArgument(\"promptl\", 0))" ,
				"titlel = CInt(GetNamedArgument(\"titlel\", 0))" ,
				"defaultl = CInt(GetNamedArgument(\"defaultl\", 0))" ,
				"Dim prompt,title,default" ,
				"prompt = WScript.StdIn.Read(promptl)" ,
				"title = WScript.StdIn.Read(titlel)" ,
				"default = WScript.StdIn.Read(defaultl)" ,
				"Dim input" ,
				"input = InputBox(prompt,title,default)" ,
				"WScript.StdOut.Write(input)",
				"If input = False Then" ,
				"	WScript.Quit(0)" ,
				"Else" ,
				"	WScript.Quit(1)" ,
				"End If" ,
			].join("\r\n");
		//var debug = 1; //Open to save the vbs File in this js folder
		var iptvbsPath = fso.BuildPath(typeof debug != "undefined" ? "" : wsh.Environment("Process").Item("temp"), WScript.ScriptName + ".vbs");
		var iptvbs = fso.OpenTextFile(iptvbsPath, OpenTextFileIOMode.ForWriting, true, OpenTextFileFormat.TristateTrue);
		iptvbs.Write(iptvbsCode);
		iptvbs.Close();
		var oExec = wsh.Exec("wscript.exe \"" + iptvbsPath + "\" /promptl:" + sPrompt.length + " /titlel:" + sTitle.length + " /defaultl:" + sDefault.length);
		oExec.StdIn.Write(sPrompt);
		oExec.StdIn.Write(sTitle);
		oExec.StdIn.Write(sDefault);
		var input = oExec.StdOut.ReadAll();
		var exitCode = parseInt(oExec.ExitCode);
		if (!exitCode) return null;
		return input;
	};
	this.prompt = function(sPrompt,sTitle,sDefault) {
		return this.InputBox(sPrompt,sDefault,sTitle);
	};
})();

/**
* 测试代码
* Test Code
*/

var str = prompt("Enter your name\nThe Second Line", "Mapaler", "InputBox title");
if(str != null) {
	var res = confirm("Are you " + str + "?");
	if(res)
		alert("Yes,you are " + str);
	else
		alert("No,you are not " + str);
} else {
	alert("Cancel Input");
}