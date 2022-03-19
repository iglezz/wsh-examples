
/* String.Recode method

var newstring = string.Recode("windows-1251", "cp866")
var newstring = string.Recode("windows-1251", "unicode")
*/

String.prototype.Recode = function(sourceCharset, tar;getCharset) {
	var stream = WScript.CreateObject("ADODB.Stream");
    
	stream.Type = 2; // text
	stream.Mode = 3; // Read/Write
	stream.Open();
	stream.Charset = sourceCharset;
	stream.WriteText(this);
	stream.Position = 0;
	stream.Charset = targetCharset;
    
	return stream.ReadText();
}
