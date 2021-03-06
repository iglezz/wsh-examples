/*******************************************************************************
Name: NewFileFolderEx.js
Author: iglezz (iglezz@gmail.com)

Create files/folders with placeholders for dynamic name formatting.

TODO:
* placeholder to get special folder
* elevate if not enough permissions for file/folder creation
*******************************************************************************/

var RET_OK = 1;
var ERR_EMPTYNAME    = -2;
var ERR_EXISTS       = -3;
var ERR_CREATEPARENT = -4;
var ERR_CREATEFILE   = -5;
var ERR_CREATEFOLDER = -6;

var ERR_INT_UNDEFCNT = -20;
var ERR_INT_NOCOUNT  = -26;

var MSG_HELP               = 1;
var MSG_HELP_ASKREADME     = 2;
var MSG_PARAM_ERR_HEAD     = 3;
var MSG_PARAM_ERR_NONAME   = 4;
var MSG_PARAM_ERR_UNKNOW   = 5;
var MSG_ERR_EMPTYNAME      = 6;
var MSG_ERR_EXISTS         = 7;
var MSG_FILE               = 8;
var MSG_FOLDER             = 9;
var MSG_ERR_CREATEFILE     = 10;
var MSG_ERR_CREATEFOLDER   = 11;
var MSG_FOR                = 12;
var MSG_ERR_INTERNAL       = 13;
var MSG_WITHENCODING       = 14;
var MSG_EXEC               = 15;
var MSG_SEND               = 16;
var MSG_DOIT_ALWAYS        = 17;
var MSG_DOIT_SUCCESS       = 18;
var MSG_SEND_TITLE         = 19;
var MSG_SEND_ACTIVE        = 20;
var MSG_SEND_DELAY         = 21;
var MSG_CREATE             = 22;
var MSG_ASKCONTINUE        = 23;

var ENC_ANSI               = 0
var ENC_UTF8_BOM           = 1
var ENC_UTF16LE_BOM        = 2
var ENC_UTF16BE_BOM        = 3

var ConsoleMode = WScript.FullName.slice(WScript.FullName.lastIndexOf("\\") + 1).toLowerCase() == "cscript.exe" ? true : false;

var Config = {
	// locale for script and <D> placeholder:
	// 0 == english
	// 1 == русский
	locale     : 1,
	// new file encoding:
	encoding   : ENC_ANSI,
	// default <C> counter placeholder
	counter    : "<C (?)>",
	// default string for empty `/exec` command
	// "=" == ShellExec with default verb
	execString : "=",
	// default string for empty `/send` command
	sendString : "Total Commander :{F16}^r"
}

var App = new function() {
	this.name = "NewFileFolderEx";
	this.date = "2019.01.08";
	this.version = "0.6.8";
	this.title = this.name + " " + this.version;
	
	this.encodingsList = [];
	this.encodingsList[ENC_ANSI]        = "ANSI";
	this.encodingsList[ENC_UTF8_BOM]    = "UTF-8 BOM";
	this.encodingsList[ENC_UTF16LE_BOM] = "UTF-16 LE BOM";
	this.encodingsList[ENC_UTF16BE_BOM] = "UTF-16 BE BOM";
	
	this.localesList = ["en", "ru"];
	
	// adding locales
	var messages = [];
	
	var mstrings = [];
	mstrings.push("english");
	mstrings[MSG_HELP            ] = "Parameters: \n\n/file FILE   \t- create file\n/folder FOLDER\t- create folder\n/nc\t\t- disable counter\n\n/ansi, \t\t- set file encoding\n/utf8, \t\t  (with BOM for multibyte)\n/utf16be,\n/utf16le or /unicode\n\n/exec [COMMAND]\t- execute command\n/send [SEND]\t- send keystrokes\n\n/test\t\t- test run\n/quiet\t\t- supress output\n/eng\t\t- use English locale";
	mstrings[MSG_HELP_ASKREADME  ] = "Open readme file?";
	mstrings[MSG_PARAM_ERR_HEAD  ] = "Error:";
	mstrings[MSG_PARAM_ERR_NONAME] = "File/folder name is not set";
	mstrings[MSG_PARAM_ERR_UNKNOW] = "Unknown parameters:";
	mstrings[MSG_ERR_EMPTYNAME   ] = "File/folder name is not set";
	mstrings[MSG_ERR_EXISTS      ] = "File/folder already exist with this name:";
	mstrings[MSG_FILE            ] = "Create new file ";
	mstrings[MSG_FOLDER          ] = "Create new folder:";
	mstrings[MSG_ERR_CREATEFILE  ] = "Can't create file:";
	mstrings[MSG_ERR_CREATEFOLDER] = "Can't create folder:";
	mstrings[MSG_FOR             ] = "for:";
	mstrings[MSG_ERR_INTERNAL    ] = "Internal error:";
	mstrings[MSG_WITHENCODING    ] = "Using encoding: ";
	mstrings[MSG_EXEC            ] = "Execute command";
	mstrings[MSG_SEND            ] = "Send keystrokes";
	mstrings[MSG_DOIT_ALWAYS     ] = " in any case:";
	mstrings[MSG_DOIT_SUCCESS    ] = " on successful creation:";
	mstrings[MSG_SEND_TITLE      ] = "To window with title:";
	mstrings[MSG_SEND_ACTIVE     ] = "To the active window";
	mstrings[MSG_SEND_DELAY      ] = "Delay %% ms before send";
	mstrings[MSG_ASKCONTINUE     ] = "Continue?";
	
	messages.push(mstrings);
	
	var mstrings = [];
	mstrings.push("русский");
	mstrings[MSG_HELP            ] = "Параметры: \n\n/file FILE   \t- создать файл\n/folder FOLDER\t- создать папку\n/nc\t\t- отключить счётчик\n\n/ansi, \t\t- выбор кодировки для файла\n/utf8, \t\t  (с BOM для многобайтных)\n/utf16be,\n/utf16le или /unicode\n\n/exec [COMMAND]\t- выполнить команду\n/send [SEND]\t- передать клавиши\n\n/test\t\t- тестовый запуск\n/quiet\t\t- подавление вывода сообщений\n/eng\t\t- use English locale";
	mstrings[MSG_HELP_ASKREADME  ] = "Открыть подробную справку?";
	mstrings[MSG_PARAM_ERR_HEAD  ] = "Ошибка:";
	mstrings[MSG_PARAM_ERR_NONAME] = "Не задано имя файла/папки";
	mstrings[MSG_PARAM_ERR_UNKNOW] = "Неизвестные параметры:";
	mstrings[MSG_ERR_EMPTYNAME   ] = "Не задано имя файла/папки";
	mstrings[MSG_ERR_EXISTS      ] = "Файл/папка таким именем существует:";
	mstrings[MSG_CREATE          ] = "Создать ";
	mstrings[MSG_FILE            ] = "Файл:";
	mstrings[MSG_FOLDER          ] = "Папка:";
	mstrings[MSG_ERR_CREATEFILE  ] = "Ошибка создания файла:";
	mstrings[MSG_ERR_CREATEFOLDER] = "Ошибка создания папки:";
	mstrings[MSG_FOR             ] = "для:";
	mstrings[MSG_ERR_INTERNAL    ] = "Внутренняя ошибка:";
	mstrings[MSG_WITHENCODING    ] = "В кодировке: ";
	mstrings[MSG_EXEC            ] = "Выполнить команду";
	mstrings[MSG_SEND            ] = "Отправить ";
	mstrings[MSG_DOIT_ALWAYS     ] = " (в любом случае): ";
	mstrings[MSG_DOIT_SUCCESS    ] = " (если успешно создан): ";
	mstrings[MSG_SEND_TITLE      ] = "В окно с заголовком: ";
	mstrings[MSG_SEND_ACTIVE     ] = "В активное окно";
	mstrings[MSG_SEND_DELAY      ] = "С задержкой %% мс";
	mstrings[MSG_ASKCONTINUE     ] = "Продолжить?";
	
	messages.push(mstrings);
	
	this.MSG = function(id) {
		return messages[Config.locale][id];
	}
};

var AppParam = new AppParameters(WScript.Arguments);

switch (AppParam.errorCode) {
	case 1:
		ShowHelpAndExit(AppParam.quiet, 0);
	
	case 2:
	case 10:
	case 11:
		ShowHelpAndExit(AppParam.quiet, 1, AppParam.GetErrorMessage());
}

var newff = new NewFileFolder();
var newName = newff.GetName(AppParam.name, AppParam.isFolder, AppParam.useCounter);
var exec = new Execute(AppParam.execString);
var send = new SendKeys(AppParam.sendString);

if (AppParam.testMode) {
	var wsh = WScript.CreateObject("WScript.Shell");
	
	var testMessage = App.MSG(AppParam.isFolder ? MSG_FOLDER : MSG_FILE);
	testMessage += "\n>  " + newName;
	if (!AppParam.isFolder)
		testMessage += "\n" + App.MSG(MSG_WITHENCODING) + "\n> " + App.encodingsList[AppParam.encoding];
	
	if (exec.execAvailable) {
		testMessage += "\n\n" + App.MSG(MSG_EXEC) + 
				       App.MSG(exec.execAlways ? MSG_DOIT_ALWAYS : MSG_DOIT_SUCCESS) +
					   "\n>  " + exec.GetC(newName) + 
					   "\n>  " + exec.GetP(newName) + 
					   "\n>  execMode: " + exec.execMode;
	}
	
	if (send.sendAvailable) {
		testMessage += "\n\n" + App.MSG(MSG_SEND) + 
				       App.MSG(send.sendAlways ? MSG_DOIT_ALWAYS : MSG_DOIT_SUCCESS) +
					   "\n>  " + send.keys + "\n" + 
					   (send.active ? App.MSG(MSG_SEND_ACTIVE) : (App.MSG(MSG_SEND_TITLE) + "\n>  " + send.title)) + 
					   (send.delay > 0 ? "\n" + App.MSG(MSG_SEND_DELAY).replace("%%", send.delay) : "");
	}
	
	if (ConsoleMode) {
		WScript.Echo(testMessage);
		WScript.Quit(0);
	} else {
		testMessage += "\n\n\n" + App.MSG(MSG_ASKCONTINUE);
		if (wsh.Popup(testMessage, 0, App.title, 4) != 6 )
			WScript.Quit(0);
	}
}

switch (newName) {
	case ERR_EMPTYNAME:
		ShowMessageAndExit(AppParam.quiet, -newName, App.MSG(MSG_ERR_EMPTYNAME));
		
	case ERR_INT_NOCOUNT:
	case ERR_INT_UNDEFCNT:
		ShowMessageAndExit(AppParam.quiet, -newName, App.MSG(MSG_ERR_INTERNAL) + "\n\"" + AppParam.name) + "\"";
		
	case ERR_EXISTS:
		if (exec.execAvailable && exec.execAlways && !AppParam.testMode)
			exec.Exec(newff.lastName);
		if (send.sendAvailable && send.sendAlways && !AppParam.testMode)
			send.Send();
		ShowMessageAndExit(AppParam.quiet, -newName, App.MSG(MSG_ERR_EXISTS) + "\n\"" + newff.lastName + "\"");
}

var createResult = newff.Create(newName, AppParam.isFolder, AppParam.encoding);

switch (createResult) {
	case ERR_CREATEPARENT:
		ShowWarning(AppParam.quiet, App.MSG(MSG_ERR_CREATEFOLDER) + "\n" + newff.lastName + "\n" + App.MSG(MSG_FOR) + "\n" + newName + "\n\n" + newff.errorText);
		break;
		
	case ERR_CREATEFILE:
		ShowWarning(AppParam.quiet, App.MSG(MSG_ERR_CREATEFILE) + "\n" + newName + "\n\n" + newff.errorText);
		break;
	
	case ERR_CREATEFOLDER:
		ShowWarning(AppParam.quiet, App.MSG(MSG_ERR_CREATEFOLDER) + "\n" + newName + "\n\n" + newff.errorText);
		break;
}

if (exec.execAvailable && (createResult == RET_OK || exec.execAlways))
	exec.Exec(newName);
	
if (send.sendAvailable && (createResult == RET_OK || send.sendAlways))
	send.Send();
	
WScript.Quit(createResult == RET_OK ? RET_OK : -createResult);


function NewFileFolder() {
	// Object Properties:
	
	this.errorCode = 0;
	this.errorText = "";
	this.lastName = "";
	
	// Object Methods:
	
	// GetName(name, type, count) 
	// return success: "full file/folder name" 
	// return error: -# errorcode (ERR_xxxxx values)
	this.GetName = function(pName, pFolder, pCount, pDefCounter) {
		if (pName == undefined || pName == "")
			return ERR_EMPTYNAME;
		
		var isFolder = (pFolder == undefined ? false : (pFolder == true ? true : false));
		var useCounter = (pCount == undefined ? true : (pCount == true ? true : false));
		var defaultCounter = pDefCounter == undefined ? Config.counter : pDefCounter;
		if (defaultCounter == undefined)
			return ERR_INT_UNDEFCNT;
		
		var baseName = GetBaseName(pName, defaultCounter);
		baseName = nffFSO.GetAbsolutePathName(baseName);
		
		if (!useCounter) {
			var testName = baseName.replace(/\<C[^>]+\>/g, "");
			if (CheckExists(testName, isFolder)) {
				this.lastName = testName;
				return ERR_EXISTS;
			} else
				return testName;
		}
		
		var counters = baseName.match(/\<C[^>]+\>/g);
		
		if (counters == null)
			return ERR_INT_NOCOUNT;
		
		var previousName = "";
		var testName = "";
		
		do {
			testName = baseName;
			
			for (var i = 0; i < counters.length; i++) {
				testName = testName.replace(counters[i], GetFinalCounter(counters[i], counterValue, bPrintCounter));
			}
			
			counterValue++;
			
			if (!bPrintCounter)
				bPrintCounter = true;
			
			if (previousName == testName) {
				this.lastName = testName;
				return ERR_EXISTS;
			} else
				previousName = testName;
			
		} while ( CheckExists(testName, isFolder) );
		
		return testName;
	}
	
	// Create(name, type)
	// return success: 0
	// return error: -# errorcode (ERR_xxxxx values)
	this.Create = function(fullname, isFolder, encoding) {
		var parentfolders = CreateParentFolders(fullname);
		if (typeof(parentfolders) == "object") {
			this.errorCode = parentfolders.code;
			this.errorText = parentfolders.text;
			this.lastName = parentfolders.name;
			return ERR_CREATEPARENT;
		}
		
		if (isFolder) {
			try {
				nffFSO.CreateFolder(fullname);
			} catch(error) {
				this.errorCode = error;
				this.errorText = error.description;
				return ERR_CREATEFOLDER;
			}
		} else {
			try {
				switch (encoding) {
					case ENC_ANSI:
						var file = nffFSO.CreateTextFile(fullname, true, false);
						file.Close();
						break;
						
					case ENC_UTF8_BOM:
						var stream = new ActiveXObject('ADODB.Stream');
						stream.Type = 2;
						stream.Charset = "utf-8";
						stream.Open();
						stream.WriteText("");
						stream.SaveToFile(fullname, 2);
						stream.Close();   
						break;
						
					case ENC_UTF16LE_BOM:
						var file = nffFSO.CreateTextFile(fullname, true, true);
						file.Close();
						break;
						
					case ENC_UTF16BE_BOM:
						var stream = new ActiveXObject('ADODB.Stream');
						stream.Type = 2;
						stream.Charset = "unicodeFEFF";
						stream.Open();
						stream.WriteText("");
						stream.SaveToFile(fullname, 2);
						stream.Close();   
						break;
						
				}
			} catch(error) {
				this.errorCode = error;
				this.errorText = error.description;
				return ERR_CREATEFILE;
			}
		}
		
		return RET_OK;
	}
	
	// Internal variable and functions
	var nffWSH = WScript.CreateObject("WScript.Shell");
	var nffFSO = new ActiveXObject("Scripting.FileSystemObject");
	var bPrintCounter  = false;
	var counterValue = 1;
	
	
	function GetBaseName(name, defaultCounterString) {
		var parser = new StringParser(Config.locale);
		// parser.SetDate(TodayDate);
		
		var baseName = parser.GetParsed(nffWSH.ExpandEnvironmentStrings(name));
		
		var bCounterExist = false;
		
		var matches = baseName.match(/\<C[^>]+\>/g);
		if (matches != null) {
			for (var i = 0; i < matches.length; i++) {
				var counterString = matches[i];
				
				var delimiter = counterString.indexOf(":");
				if (delimiter > 2) {
					counterString = "<C" + matches[i].slice(delimiter + 1);
					
					var counterOptions = matches[i].slice(2, delimiter);
					var option = counterOptions.match(/\d+/);
					if (option != null) {
						bPrintCounter = true;
						counterValue = option[0];
					}
				}
				
				if (counterString.indexOf("?") != -1)
					bCounterExist = true;
				
				if (counterString == "<C>")
					counterString = "<C?>";
				
				if (counterString != matches[i])
					baseName = baseName.replace(matches[i], counterString);
		   }
		}
		
		if (!bCounterExist) {
			var slashPos = baseName.lastIndexOf("\\");
			var dotPos = baseName.lastIndexOf(".");
			
			// leading dot in '.name' != 'name.ext' delimiter
			if ((slashPos + 1) < dotPos)
				baseName = baseName.slice(0, dotPos) + defaultCounterString + baseName.slice(dotPos);
			else
				baseName = baseName + defaultCounterString;
		}
		
		return baseName;
	}
	
	function GetFinalCounter(counterString, counterNumber, returnNotEmpty) {
		if (!returnNotEmpty) return "";
		
		var outputString = counterString.slice(2,-1);
		
		var matches = counterString.match(/\?+/g);
		if (matches == null) return outputString;
		
		var counterString = "";
		
		for (var i = 0; i < matches.length; i++) {
			counterString = "";
			for (var j = 0; j < (matches[i].length - counterNumber.toString().length); j++) {
				counterString += "0";
			}
			counterString += counterNumber.toString();
			
			outputString = outputString.replace(/\?+/, counterString);
		}
		
		return outputString;
	}
	
	function CheckExists(fullname, isFolder) {
		return (!isFolder && nffFSO.FileExists(fullname)) ||
			   ( isFolder && nffFSO.FolderExists(fullname));
	}
	
	function GetFilteredName(fullname) {
		var slashPos = fullname.lastIndexOf("\\");
		
		var path = fullname.slice(0, slashPos + 1);
		var name = fullname.slice(slashPos + 1).replace(/[:\*\|\<\>\?\/]/g, "");
		
		return path + name;
	}
	
	function CreateParentFolders(fullname) {
		// TODO::version with recursion (??)
		var folders = fullname.split("\\");
		var path = folders[0];
		
		for (var i = 1; i < (folders.length - 1); i++) {
			path += "\\" + folders[i];
			
			if (!nffFSO.FolderExists(path)) {
				try {
					nffFSO.CreateFolder(path);
				} catch(error) {
					return {
						code: error.number,
						text: error.description,
						name: path
					}
				}
			}
		}
		
		return true;
	}
	
}	

function StringParser(locale_index) {
	var now = new Date();
	var localeIndex = locale_index == undefined ? 0 : locale_index;
	
	var localesList = ["en", "ru"];
	var monthsNames = [
		[ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ],
		[ "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" ]
		];
	var weekdayNames = [
		[ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ],
		[ "Воскресенье", "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" ]
		];
	var weekdayNames3 = [
		[ "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" ],
		[ "Вск", "Пнд", "Втр", "Срд", "Чтв", "Птн", "Сбт" ]
		];
	var weekdayNames2 = [
		[ "Su", "Mo", "Tu", "We", "Th", "Fr", "Sa" ],
		[ "Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб" ]
		];
	var dateMatches = [
		[ "YYYY", "now.getFullYear()" ],
		[ "YYY",  "now.getFullYear()" ], // workaround for GetParsedString2() v1.0 (it handles YYYY as two YY)
		[ "YY"  , "now.getFullYear().toString().slice(2)" ],
		[ "MMMM", "monthsNames[localeIndex][ now.getMonth() ]" ],
		[ "MMM" , "monthsNames[localeIndex][ now.getMonth() ].slice(0, 3)" ],
		[ "MM"  , "(now.getMonth() < 9 ? '0' : '') + (now.getMonth() + 1)" ],
		[ "M"   , "(now.getMonth() + 1)" ],
		[ "DDD" , "var d = now.getDate(); d + ( d == 1 ? 'st' : ( d == 2 ? 'nd'  : ( d == 3 ? 'rd' : 'th' ) ) ); " ],
		[ "DD"  , "(now.getDate() < 10 ? '0' : '') + now.getDate()" ],
		[ "D"   , "now.getDate()" ],
		[ "WWWW", "weekdayNames[localeIndex][ now.getDay() ]" ],
		[ "WWW" , "weekdayNames3[localeIndex][ now.getDay() ]" ],
		[ "WW"  , "weekdayNames2[localeIndex][ now.getDay() ]" ],
		[ "WS"  , "(now.getDay() + 1)" ], // Sunday is first
		[ "W"   , "now.getDay() == 0 ? 7 : now.getDay()" ], // Monday is first
		[ "hh"  , "(now.getHours() < 10 ? '0' : '') + now.getHours()" ],
		[ "h"   , "now.getHours()" ],
		[ "mm"  , "(now.getMinutes() < 10 ? '0' : '') + now.getMinutes()" ],
		[ "m"   , "now.getMinutes()" ],
		[ "ss"  , "(now.getSeconds() < 10 ? '0' : '') + now.getSeconds()" ],
		[ "s"   , "now.getSeconds()" ]
		];

	
	this.SetDate = function (dateObject) {
		now = dateObject;
	}
	
	this.GetParsed = function (inputString) {
		var matches = inputString.match(/\<[^>]+\>/g);
		if (matches == null) return inputString;
		
		var outputSring = inputString;
		var matchReplace;
		for (var i = 0; i < matches.length; i++) {
			switch (matches[i].charAt(1)) {
				case "D": // Date
					matchReplace = GetParsedDate(matches[i]);
					break;
				
				default:
					matchReplace = undefined;
			}
			
			if (matchReplace != undefined)
				outputSring = outputSring.replace(matches[i], matchReplace);
		}
		
		return outputSring;
	}
	
	function GetParsedDate(placeholder) {
		var dateString = placeholder.slice(2, -1);
		
		var locale = localesList[Config.locale];
		var caseNum = 0;
		
		var delimiterPosition = dateString.indexOf(":");
		if (delimiterPosition != -1) {
			var params = dateString.slice(0, delimiterPosition);
			
			var caseMatch = params.match(/[+-]/);
			if (caseMatch != null) {
				caseNum = caseMatch[0] == "+" ? 2 : 1;
				params = params.replace(caseMatch[0], "");
			}
			
			locale = params;
			
			dateString = dateString.slice(delimiterPosition + 1)
	   }
		
		for (var i = 0; i < localesList.length; i++ ) {
			if (locale == localesList[i]) {
				localeIndex = i;
				break;
			}
		}
		
		return GetParsedString2(dateString, dateMatches, caseNum);
	}

	function GetParsedString2(inputString, matchesArray, stringCase) {
		// GetParsedString2() parser version 1.0
		var source = inputString + "\t"; 
		var result = "";
		var chunk = "";
		var symb = "";
		var chunkMatchIndex = -1;
		var newMatchIndex;
		var scase = stringCase == undefined ? 0 : stringCase;
		
		function GetMatchIndex(matchString) {
			for (var i = 0; i < matchesArray.length; i++) {
				if (matchString == matchesArray[i][0]) return i;
			}
			return -1;
		}
		
		for (var i = 0; i < source.length; i++) {
			symb = source.charAt(i);
			newMatchIndex = GetMatchIndex(chunk + symb);
			
			if (newMatchIndex == -1) {
				if (chunkMatchIndex == -1) {
					result += chunk;
				} else {
					var str = eval(matchesArray[chunkMatchIndex][1]);
					switch (stringCase) {
						case 1:
							str = str.toString().toLowerCase();
							break;
						case 2:
							str = str.toString().toUpperCase();
					}
					result += str;
				}
				chunk = symb;
				chunkMatchIndex = GetMatchIndex(chunk);
			} else {
				chunk += symb;
				chunkMatchIndex = newMatchIndex;
			}
		}
		
		return result;
	}
	
}

function SendKeys(_send) {
	this.title = "";
	this.keys = "";
	this.sendAlways = false;
	this.sendAvailable = false;
	this.delay = 0;
	this.active = false;
	
	this.Send = function() {
		if (!this.sendAvailable) 
			return 0;
		
		var send = false;
		
		if (this.active) {
			send = true;
		} else {
			if (this.title != "")
				send = wsh.AppActivate(this.title);
		}
		
		if (this.delay > 0)
			WScript.Sleep(this.delay);
		
		if (send) {
			var wsh = WScript.CreateObject("WScript.Shell");
			wsh.SendKeys(this.keys); 
		}
		
		return 1;
	}
	
	if (this._send != "") {
		this.sendAlways = _send.charAt(0) == "!" ? true : false;
		var inputString = _send.slice(this.sendAlways);
		
		var options = "";
		var delimiterPosition = inputString.indexOf("*");
		
		if (delimiterPosition != -1) {
			options = inputString.slice(0, delimiterPosition);
			inputString = inputString.slice(delimiterPosition + 1);
		}
		
		if (options != "") {
			var op = options.match(/d\d+/);
			if (op != null) {
				this.delay = parseInt(op[0].slice(1));
				options = options.replace(op[0], "");
			}
			
			var op = options.match(/a/);
			if (op != null) {
				this.active = true;
				options = options.replace(op[0], "");
			}
		}
		
		if (this.active) {
			this.keys = inputString;
			this.sendAvailable = true;
		}
		
		var delimiterPosition = inputString.indexOf(":");
		if (delimiterPosition > 0 && !this.active) {
			this.title = inputString.slice(0, delimiterPosition);
			this.keys = inputString.slice(delimiterPosition + 1);
			if (this.title.replace(/\s+/g, "") != "")
				this.sendAvailable = true;
		}
	}
}

function Execute(_run) {
	this.command = "";
	this.parameters = "";
	this.execAlways = false;
	this.execAvailable = false;
	this.quoteChar = "";
	this.execMode = 10;
	this.ShellExec = false;
	
	this.Exec = function(filename) {
		var params = this.parameters.replace(/\?/g, filename);
		if (this.command.charAt(0) == "=") {
			// ShellExecute method: https://docs.microsoft.com/en-us/windows/win32/shell/shell-shellexecute
			var verb = this.command.slice(1);
			var sha = new ActiveXObject("shell.application");
			try {
				sha.ShellExecute(filename, params, "", verb, this.execMode);
			} catch(error) {}

		} else {
			// Run method: https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/d5fk67ky%28v%3dvs.84%29
			var wsh = WScript.CreateObject("WScript.Shell");
			try {
				wsh.Run("\"" + this.command + "\" " + params, this.execMode, false);
			} catch(error) {}
		}
		
		return 1;
	}
	
	this.GetC = function(filename) {
		var cmd;
		if (this.command.charAt(0) == "=") {
			var verb = this.command.slice(1);
			cmd = "ShellExecute " + (verb == "" ? "default" : verb) + ", " + filename;
		} else {
			cmd = this.command;
		}
		return cmd;
	}
	
	this.GetP = function(filename) {
		return this.parameters.replace(/\?/g, filename);
	}
	
	if (_run != "") {
		this.execAvailable = true;
		this.execAlways = _run.charAt(0) == "!" ? true : false;
		
		var inputString = _run.slice(this.execAlways);
		var options = "";
		var delimiterPosition = inputString.indexOf("*");
		
		if (delimiterPosition != -1) {
			options = inputString.slice(0, delimiterPosition);
			inputString = inputString.slice(delimiterPosition + 1);
		}
		
		if (options != "") {
			var op = options.match(/q.{1}/);
			if (op != null) {
				this.quoteChar = op[0].charAt(1);
				options = options.replace(op[0], "");
			}
			
			var op = options.match(/m\d{1}/);
			if (op != null) {
				this.execMode = parseInt(op[0].charAt(1));
				options = options.replace(op[0], "");
			}
		}
		
		var delimiters = inputString.match(/\:/g);
		if (delimiters == null) {
			if (inputString.indexOf("?") == -1) {
				this.command = inputString;
			}
			else {
				this.parameters = inputString;
			}
		} else {
			if (inputString.search(/[a-zA-Z]\:\\/) == 0) {
				// PROG = "x:\path\app...", slice by second ":" char (if present)
				var pos = inputString.slice(3).indexOf(":");
				if (pos == -1) {
					// only one ":"
					this.command = inputString;
				} else {
					this.command = inputString.slice(0, pos + 3);
					this.parameters = inputString.slice(pos + 4);
				}
			} else {
				var pos = inputString.indexOf(":");
				this.command = inputString.slice(0, pos);
				this.parameters = inputString.slice(pos + 1);
			}
		}
	}
	
	if (this.command == "")
		this.command = "=";
	
	if (this.quoteChar != "")
		try {
			var quote = new RegExp("\\" + this.quoteChar, "g");
			this.parameters = this.parameters.replace(quote, '"');
		} catch(error) {}
	
	if ((this.parameters == "" || this.parameters == "?") && this.command.charAt(0) != "=")
		this.parameters = '"?"';
	
	
}

function AppParameters(ArgsObj) {
	this.name = "";
	this.isFolder = false;
	this.useCounter = true;
	this.encoding = Config.encoding;
	
	this.execString = "";
	this.sendString = "";
	
	this.testMode = false;
	this.quiet = false;
	
	// var error = new AppErrors();
	// 0  = ok 
	// 1  = no params / show help
	// 2  = unknown params
	// 10 = file not set
	// 11 = folder  not set
	this.errorCode = 0; 
	this.unknownParams = "";
	
	this.GetErrorMessage = function() {
		var message = App.MSG(MSG_PARAM_ERR_HEAD) + "\n";
		
		switch (this.errorCode) {
			case 10:
			case 11:
				message += "* " + App.MSG(MSG_PARAM_ERR_NONAME) + "\n";
			
			case 2:
				message += "* " + App.MSG(MSG_PARAM_ERR_UNKNOW) + "\n" + this.unknownParams;
		}
		
		return message;
	}
	
	if (ArgsObj.length == 0) {
		this.errorCode = 1;
		return this;
	} 
	
	for (i = 0; i <  ArgsObj.length; i++) {
		
		switch (ArgsObj.Item(i)) {
			case "/f":
			case "/file":
				++i;
				if (i < ArgsObj.length) {
					this.name = ArgsObj.Item(i);
					this.isFolder = false;
				} else {
					this.errorCode = 10;
				}
				break;
			
			case "/ansi":
				this.encoding = ENC_ANSI;
				break;
			
			case "/utf8":
				this.encoding = ENC_UTF8_BOM;
				break;
			
			case "/utf16le":
			case "/unicode":
				this.encoding = ENC_UTF16LE_BOM;
				break;
			
			case "/utf16be":
				this.encoding = ENC_UTF16BE_BOM;
				break;
			
			case "/d":
			case "/dir":
			case "/folder":
				++i;
				if (i < ArgsObj.length) {
					this.name = ArgsObj.Item(i);
					this.isFolder = true;
				} else {
					this.errorCode = 11;
				}
				break;
			
			case "/e":
			case "/exec":
				if ((i + 1) < ArgsObj.length && ArgsObj.Item(i + 1).charAt(0) != "/") 
					this.execString = ArgsObj.Item(++i);
				else
					this.execString = Config.execString;
				break;
			
			case "/s":
			case "/send":
				if ((i + 1) < ArgsObj.length && ArgsObj.Item(i + 1).charAt(0) != "/")
					this.sendString = ArgsObj.Item(++i);
				else
					this.sendString = Config.sendString;
				break;
			
			case "/nc":
				this.useCounter = false;
				break;
			
			case "/t":
			case "/test":
				this.testMode = true;
				break;
			
			case "/q":
			case "/quiet":
				this.quiet = true;
				break;
			
			case "/eng":
				Config.locale = 0;
				break;
			
			case "/?":
				this.errorCode = 1;
				break;
			
			default:
				this.errorCode = 2;
				this.unknownParams += " " + ArgsObj.Item(i) + "\n";
		}
	}
	
	if (this.name == "" && this.errorCode > 1 && this.errorCode < 10)
		this.errorCode = 10;
}

function ShowHelpAndExit(quietMode, exitCode, prefixMessage) {
	if (quietMode)
		WScript.Quit(exitCode);
	
	var popupMessage = (prefixMessage == undefined ? "" : prefixMessage + "\n\n") ;
	popupMessage += App.MSG(MSG_HELP);
	
	if (ConsoleMode) {
		WScript.Echo(popupMessage);
		WScript.Quit(exitCode);
	}
	
	popupMessage += "\n\n" + App.MSG(MSG_HELP_ASKREADME);
	var wsh = WScript.CreateObject("WScript.Shell");
	
	if (wsh.Popup(popupMessage, 0, App.title, 4) == 6)
		OpenReadmeFile();
	
	WScript.Quit(exitCode);
}

function ShowMessageAndExit(quietMode, exitCode, messageText) {
	if (!quietMode)
		WScript.Echo(messageText);
	WScript.Quit(exitCode);
}

function ShowWarning(quietMode, messageText) {
	if (!quietMode)
		WScript.Echo(messageText);
}

function OpenReadmeFile() {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var fullnameWithoutExt = WScript.ScriptFullName.slice(0, WScript.ScriptFullName.lastIndexOf("."));
	
	var readmeFile = fullnameWithoutExt + "." + App.localesList[Config.locale] + ".txt";
	
	if (!fso.FileExists(readmeFile)) {
		readmeFile = fullnameWithoutExt + ".txt";
		if (!fso.FileExists(readmeFile)) {
			WScript.Echo("Error:\nReadme file not found!");
			return 0
		}
	}
	var sha = new ActiveXObject("shell.application");
	sha.ShellExecute(readmeFile);
	return 1;
}

function dbg(s){WScript.Echo(s);}