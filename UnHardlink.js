function UnHardlink() {
    var app = WScript.CreateObject("Shell.Application");
    var fso = WScript.CreateObject("Scripting.FileSystemObject");	
    var wsh = WScript.CreateObject("WScript.Shell");
    
    String.prototype.GetLongFileName = function() {
        var fileobject = fso.GetFile(this);
        
        return app.NameSpace(fileobject.ParentFolder.Path).ParseName(fileobject.Name).Path;
    }

    function GetTempFullName(path) {
        var delimiter = fso.FileExists(path) ? "" : "\\";
        
        do {
            tempfile = path + delimiter + fso.GetTempName();
        }
        while(fso.FileExists(tempfile));
            
        return tempfile;
    }
    
	var fsutil = fso.GetSpecialFolder(1) + "\\fsutil.exe";
    
    // Properties:
    this.fsutil = fso.FileExists(fsutil) ? fsutil : "";
    
    // Methods:
    this.ProcessFile = function(dirtyfilename) {
        if (this.fsutil == "")
            return { message: "Error: fsutil not found", returncode: 2}
        
        var targetfile = fso.GetFile(dirtyfilename.GetLongFileName());
        var targetfullname = targetfile.Path;
        
        var execobj = wsh.Exec(fsutil + " hardlink list " + "\"" + targetfullname + "\"");
        
        var output = "";
        while (execobj.Status == 0) {
            if (!execobj.StdOut.AtEndOfStream)
                output += execobj.StdOut.ReadAll();
            WScript.Sleep(100);
        }

        var re = new RegExp("\r\n", "g");
        var matches = output.match(re);
        
        if (matches.length == 1 )
            return { message: "Error: Not hardlink", returncode: 3}

        // TODO:: Check free space more carefully (cluster size, compressed files, ..)
        var filesize = targetfile.Size;
        var freespace = fso.GetDrive(fso.GetDriveName(targetfullname));
        
        if (filesize > freespace)
            return { message: "Error: Not enough free space", returncode: 10}
        
        tempfullname = GetTempFullName(targetfullname);
        
        try {
            fso.CopyFile(targetfullname, tempfullname)
        } catch(error) {
            return { message: "Error: (" + error + ") " + error.description, returncode: 11}
        }
        
        try {
            fso.DeleteFile(targetfullname, true)
        } catch(error) {
            return { message: "Error: (" + error + ") " + error.description, returncode: 13}
        }
        
        try {
            fso.MoveFile(tempfullname, targetfullname)
        } catch(error) {
            return { message: "Error: (" + error + ") " + error.description, returncode: 11}
        }
        
        return { message: "Success: " + targetfullname, returncode: 0}
   }

}


function QuitWithMessage(message, exitcode) {
	WScript.Echo(message);
	WScript.Quit(exitcode == undefined ? 0 : exitcode);
}


function msg(s) {WScript.Echo(s)}


function main(args) {
    var fso = WScript.CreateObject("Scripting.FileSystemObject");	

    if (args.length == 0)
        QuitWithMessage("Usage: " + fso.GetFileName(WScript.ScriptFullName) + " FILE", 0);
        
    var file = args.Item(0);

    if (!fso.FileExists(file))
        QuitWithMessage("Error: file not found:\n " + file, 1);

    var u = new UnHardlink();

    if (u.fsutil == "")
        QuitWithMessage("Error: fsutil not found" , 2);

    var ret = u.ProcessFile(file);
    
    WScript.Quit(ret.returncode);
}

main(WScript.Arguments);
