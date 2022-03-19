var quoteChar = new RegExp("\\\>", "g");

if (WScript.Arguments.length != 0) {
	var runCommand = "";
    
	for (i = 0; i <  WScript.Arguments.length; i++) {
		var chunk = WScript.Arguments.Item(i).replace(quoteChar, "\"");
		runCommand += " " + (chunk.indexOf(" ") == -1 ? chunk : "\"" + chunk + "\"");
	}
	WScript.CreateObject("WScript.Shell").Run("cmd /c \"" + runCommand + "\"", 0, false);
}
