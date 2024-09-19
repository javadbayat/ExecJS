String.prototype.startsWith = new Function("str", "return(this.substr(0,str.length) === str);");
String.prototype.endsWith = new Function("suffix", "return(this.substr(this.length - suffix.length) === suffix);");
beep.toString = new Function("return String.fromCharCode(7);");
Number.prototype.toHRESULT = function () {
    var hr = (0x100000000 + this).toString(16).toUpperCase();
    if (hr.length > 8)
        throw new Error("Number too large for HRESULT.");
    while (hr.length != 8) hr = "0" + hr;
    return ("0x" + hr);
};
Number.prototype.toFileSize = function () {
    var divideby;
    var tagg;
    if (this / 1024 < 1) {
        divideby = 1;
        tagg = "bytes";
    }
    else if (this / (1024 * 1024) < 1) {
        divideby = 1024;
        tagg = "KB";
    }
    else if (this / (1024 * 1024 * 1024) < 1) {
        divideby = 1024 * 1024;
        tagg = "MB";
    }
    else if (this / (1024 * 1024 * 1024 * 1024) < 1) {
        divideby = 1024 * 1024 * 1024;
        tagg = "GB";
    }
    var i = this / divideby;
    var x = Number(Math.floor(i * 100) / 100);
    return x.toFixed(2) + " " + tagg;
};
Number.prototype.toTime = function () {
    var period = this;
    
    with (Math) {
        var hours = floor(period / 3600000);
        period %= 3600000;
        
        var minutes = floor(period / 60000);
        period %= 60000;
        
        var seconds = floor(period / 1000);
        period %= 1000;
    }
    
    return hours + ":" + minutes + ":" + seconds + ":" + period;
};

var wshShell = WScript.CreateObject("WScript.Shell");
var FSO = WScript.CreateObject("Scripting.FileSystemObject");

FatalError.prototype = Error;
//Predefined variables
var ejs, _internal, _echo = true, _lastres = "", 
_lastiniterr = null, _lasterr = null, _ = this, 
_target = _, _cmdexit = "", _cmdcancel = "#", 
_void = {toString:undefined}, _version = "1.0",
_args = col2arr(WScript.Arguments), _cmd = "",
_timing = false, _sdate = null, _edate = null,
_cmdsetup = "setup", _cmdhelp = "help", _ncmds = 0,
_sysarch = wshShell.Environment.Item("PROCESSOR_ARCHITECTURE");

//Initialization function
(function() {    
    if (FSO.GetBaseName(WScript.FullName).toLowerCase() != "cscript") {
        var cmdLine = 'cmd /c color f0 & cscript "' + WScript.ScriptFullName + '" //D';
        for (var i = 0;i < _args.length;i++)
            if (_args[i].indexOf(" ") == -1)
                cmdLine += " " + _args[i];
            else
                cmdLine += ' "' + _args[i] + '"';
        wshShell.Run(cmdLine);
        WScript.Quit();
    }
    
    var argIndex = findArg("/a", "-a");
    if (argIndex != -1) {
        var objref;
        if (objref = _args[argIndex + 1]) {
            try {
                attach(GetObject(objref).value);
            }
            catch (e) {
                alert("Unable to attach to the object specified in command line.\n" + e.message);
                _target = _;
            }
        }
    }
    
    argIndex = findArg("/v", "-v");
    if (argIndex != -1) {
        if (_sysarch != "x86") {
            WSH.StdErr.WriteLine("Error: Virtualization is only supported on 32-bit Windows.");
            WSH.Quit();
        }
        
        _ = _target = new ActiveXObject("MSScriptControl.ScriptControl");
        _.Language = "JScript";
        _.CodeObject.WSH = _.CodeObject.WScript = WScript;
    }
    
    argIndex = findArg("/t", "-t");
    if (argIndex != -1)
        _timing = true;
    
    argIndex = findArg("/c", "-c");
    if (argIndex != -1) {
        _cmd = _args[argIndex + 1];
        if (_ instanceof ActiveXObject) {
            if (_timing)
                WScript.Echo("# Exec started at " + ((_sdate = new Date).toLocaleTimeString()));
            _lastres = _.eval(_cmd) + "";
            WScript.Echo(_lastres);
            if (_timing && (!!_sdate)) {
                _edate = new Date;
                var period = (_edate.valueOf()) - (_sdate.valueOf());
                WScript.Echo("# Exec lasted for " + (period.toTime()));
            }
            WScript.Quit();
        }
    }
    
    function findArg(arg1, arg2) {
        for (var i = 0;i < _args.length;i++) {
            var arg = _args[i].toLowerCase();
            if ((arg1 == arg) || (arg2 == arg))
                return i;
        }
        return -1;
    }
})();

var shell = WScript.CreateObject("Shell.Application");
var voice = WScript.CreateObject("SAPI.SpVoice");
var locator = WScript.CreateObject("WbemScripting.SWbemLocator");
var wshNetwork = WScript.CreateObject("WScript.Network");

if (_sysarch == "x86") {
    var VBS = WScript.CreateObject("MSScriptControl.ScriptControl");
    VBS.Language = "VBScript";
    VBS.AddObject("WScript", WScript);
    VBS.AddObject("WSH", WSH);

    VBS.AddCode("Function ConvertArray(jsa)\nDim vba()\nReDim vba(jsa.length - 1)\nDim i, e\ni = 0\nFor Each e In jsa\nIf IsObject(e) Then\nSet vba(i) = e\nElse\nvba(i) = e\nEnd If\ni = i + 1\nNext\nConvertArray = vba\nEnd Function");
    var getVBArray = VBS.Eval('GetRef("ConvertArray")');
}
else
    var VBS = null;

var errInfo = {
    hr: function (errno) {
        return system("certutil -error " + errno.toHRESULT());
    }, win32: function (errno) {
        return system("net helpmsg " + errno);
    }
};

try {
    var cmdlib = WScript.CreateObject("Microsoft.CmdLib");
    cmdlib.ScriptingHost = WScript.Application;
}
catch (_lastiniterr) { }
try {
    var cui = WScript.CreateObject("CompatUI.Util");
}
catch (_lastiniterr) { }
try {
    var scriptex = WScript.CreateObject("Scriptex.Util");
}
catch (_lastiniterr) { }

_internal = {
    ShellInput: function () {
        WScript.StdOut.Write((_target === _) ? ">>> " : "==> ");
        var cmd = WScript.StdIn.ReadLine();
        if (cmd == _cmdexit)
            throw null;
        if (cmd == _cmdsetup) {
            _sdate = null;
            return "setup()";
        }
        if (cmd == _cmdhelp) {
            _sdate = null;
            return "help()";
        }
        if (cmd.charAt(0) == " ")
            do {
                WScript.StdOut.Write((_target === _) ? ">>>  " : "==>  ");
                cmd += "\n" + WScript.StdIn.ReadLine();
                var lc = cmd.charAt(cmd.length - 1);
                if (lc == _cmdcancel)
                    return "";
            } while (lc != " ")
        if (_timing)
            WScript.StdOut.WriteLine("# Exec started at " + 
            ((_sdate = new Date).toLocaleTimeString()));
        return cmd;
    }, ShellOutput: function () {
        WScript.StdOut.Write(_echo ? _lastres : "");
        if (_timing && (!!_sdate)) {
            _edate = new Date;
            var period = (_edate.valueOf()) - (_sdate.valueOf());
            WScript.StdOut.WriteLine("# Exec lasted for " + (period.toTime()));
        }
    }, ShellExceptionHandler: function () {
        if (!_lasterr) {
            if (_target === _)
                return true;
            else
                attach();
        }
        else if (_lasterr instanceof FatalError)
            _lasterr.handle();
        else if (_lasterr.message)
            WScript.StdErr.WriteLine("Error: " + _lasterr.message);
        else if (_lasterr.number)
            WScript.StdErr.WriteLine("Error Code: " + _lasterr.number.toHRESULT());
        return false;
    }, vie:1, vhta:1
};

if (_cmd) {
    if (_timing)
        WScript.Echo("# Exec started at " + ((_sdate = new Date).toLocaleTimeString()));
    _lastres = _target.eval(_cmd) + "";
    WScript.Echo(_lastres);
    if (_timing && (!!_sdate)) {
        _edate = new Date;
        var period = (_edate.valueOf()) - (_sdate.valueOf());
        WScript.Echo("# Exec lasted for " + (period.toTime()));
    }
    WScript.Quit();
}

main:
    while (1) {
        try {
            if ((typeof _target != "object") || (!_target))
                throw new FatalError;
            _lastres = _target.eval(_internal.ShellInput()) + "\n";
            _internal.ShellOutput();
            _ncmds++;
        }
        catch (_lasterr) {
            if (_internal.ShellExceptionHandler())
                break main;
        }
    }

function toString() {
    return "[object ExecJS Root]";
}

//Utility Functions

function FatalError() {
    this.handle = function () {
        if (wshShell.Popup("A Fatal Error caused ExecJS to stop working correctly, and it will quit. Do you wish to debug the program?", 0, "ExecJS Fatal Error", 17) == 1)
            throw this;
        else
            WScript.Quit();
    };
    this.toString = function () {
        return "[object FatalError]";
    };
}

function newDictionary() {
    return WScript.CreateObject("Scripting.Dictionary");
}

function attach(scriptRoot) {
    try {
        _target.ejs = undefined;
    }
    catch (e) {}
    if (scriptRoot)
        (_target = scriptRoot).ejs = _;
    else
        _target = _;
    return _target;
}

function Thing(initialVal, noobjref) {
    this.interval = 100;
    this.value = initialVal;
    this.objref = noobjref ? "" : getref(this);
    
    this.toString = function() {
        return this.objref ? this.objref : "?";
    };
    
    this.acquire = function(maxTime) {
        var oldValue = this.value;
        var foo = 0;
        
        if (!maxTime)
            maxTime = Infinity;
        
        while (this.value === oldValue) {
            WScript.Sleep(this.interval);
            
            if ((foo += this.interval) >= maxTime)
                break;
        }
        
        return this.value;
    };
}

function acquireSomething(maxTime) {
    var something = new Thing;
    WScript.Echo("Thing: " + something);
    return something.acquire(maxTime);
}

function attachSomething(maxTime) {
    var root = acquireSomething(maxTime);
    if (root)
        return attach(root);
}

function getref(obj) {
    try {
        return scriptex.GenerateOBJREF(obj ? obj : _);
    }
    catch (e) {
        return "";
    }
}

function system(cmd) {
    return wshShell.Exec("cmd /c " + cmd).StdOut.ReadAll();
}

function print(txt) {
    WScript.Echo(txt);
    return _void;
}

function sleep(t) {
    if (t != Infinity)
        WScript.Sleep(t);
    else
        while (1)
            WScript.Sleep(60000);
}

function beep(repeat, delay) {
    if (!repeat)
        repeat = 1;
    
    if (delay) {
        for (var i = 0;i < repeat;i++) {
            if (i)
                sleep(delay);
            
            WScript.StdOut.Write(String.fromCharCode(7));
        }
    }
    else {
        var s = "";
        for (var i = 0;i < repeat;i++)
            s += String.fromCharCode(7);
        
        WScript.StdOut.Write(s);
    }
    
    return _void;
}

function obj_dump(obj) {
    var e;
    if (!obj)
        obj = _;
    
    for (var m in obj) {
        try {
            if (obj[m] instanceof Function)
                WScript.Echo("Method:\t" + m);
            else
                WScript.Echo("Property:\t" + m);
        }
        catch (e) {
            WScript.Echo("Unknown:\t" + m);
        }
    }
    
    return _void;
}

function col2arr(col) {
    if (typeof col == "object") {
        var arr = [];
        for (var ce = new Enumerator(col);!ce.atEnd();ce.moveNext())
            arr.push(ce.item());
        return arr;
    }
    else if (typeof col == "string")
        return col.split("");
    else
        return (new VBArray(col)).toArray();
}

function alert(msg) {
    if (typeof msg == "object") {
        try {
            msg = String(msg);
        }
        catch (e) {
            msg = "[object]";
        }
    }
    
    wshShell.Popup(msg, 0, "ExecJS", 48);
}

function confirm(msg) {
    if (typeof msg == "object") {
        try {
            msg = String(msg);
        }
        catch (e) {
            msg = "[object]";
        }
    }
    
    return wshShell.Popup(msg, 0, "ExecJS", 33) == 1;
}

function prompt(msg, def) {
    if (!VBS)
        throw new Error("This feature is not currently supported on 64-bit Windows.");
    
    if (typeof msg == "object") {
        try {
            msg = String(msg);
        }
        catch (e) {
            msg = "[object]";
        }
    }
    
    if (typeof def == "object") {
        try {
            def = String(def);
        }
        catch (e) {
            def = "[object]";
        }
    }
    
    VBS.CodeObject.M = msg;
    VBS.CodeObject.D = def;
    
    var res = VBS.Eval('InputBox(M, "ExecJS", D)');
    
    VBS.CodeObject.M = VBS.CodeObject.D = undefined;
    return (res === undefined) ? null : res;
}

function child() {
    wshShell.Run("ExecJS.js //D");
}

function virtualIE(title, dontAttach) {
    var IE = WScript.CreateObject("InternetExplorer.Application");
    IE.Navigate("about:blank");
    while (IE.ReadyState != 4)
        WScript.Sleep(1000);
    IE.Visible = true;
    var win = IE.Document.parentWindow;
    win.name = "VirtualIE" + (_internal.vie++);
    if (title)
        win.document.title = title;
    if (win.execScript)
        win.execScript(" ");
    if (!dontAttach)
        attach(win);
    return IE;
}

function virtualJS(dontAttach, dontAddWSH) {
    if (_sysarch != "x86")
        throw new Error("Virtualization is only supported on 32-bit Windows.");
    
    var sc = WScript.CreateObject("MSScriptControl.ScriptControl");
    sc.Language = "JScript";
    var root = sc.CodeObject;
    if (!dontAddWSH)
        root.WSH = root.WScript = WScript;
    if (!dontAttach)
        attach(root);
    return sc;
}

function virtualHTA(title, dontAttach) {
    if (scriptex) {
        var th = new Thing;
        wshShell.Run("mshta javascript:void(GetObject('" + th + "').value=this);");
        var hta = th.acquire(10000);
        if (!hta)
            throw new Error("Unable to acquire the HTA.");
        hta.name = "VirtualHTA" + (_internal.vhta++);
        hta.document.title = title ? title : hta.name;
        return dontAttach ? hta : attach(hta);
    }
    
    return virtualHTA2(title, dontAttach);
}

function virtualHTA2(title, dontAttach) {
    var id = _internal.vhta++;
    wshShell.Run('mshta "javascript:wbc=document.appendChild(document.createElement(\'object\'));wbc.classid=\'clsid:8856F961-340A-11D0-A96B-00C04FD705A2\';wbc.RegisterAsBrowser=true;wbc.PutProperty(\'VHTA' + id + '\', window);"');
    
    for (var i = 0;i < 10;i++) {
        var wins = shell.Windows();
        for (var j = 0;j < wins.Count;j++) {
            var win = wins.Item(j);
            if (win) {
                var hta = win.GetProperty("VHTA" + id);
                if (hta) {
                    hta.name = "VirtualHTA" + id;
                    hta.document.title = title ? title : hta.name;
                    hta.document.firstChild.removeNode(true);
                    return dontAttach ? hta : attach(hta);
                }
            }
        }
        
        sleep(500);
    }
    
    throw new Error("Unable to acquire the HTA.");
}

function setup(undo) {
    if (undo) {
        wshShell.RegDelete("HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\ExecJS.js\\");
    }
    else {
        if (!scriptex)
            print("Note: Seems like you don't have the Scriptex Library installed! Scriptex is required for some of ExecJS's features to work, such as Virtual HTAs, object browsing, OBJREF generation, .etc");
        
        var sfn = WScript.ScriptFullName;
        wshShell.RegWrite("HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\ExecJS.js\\Path", FSO.GetParentFolderName(sfn));
        wshShell.RegWrite("HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\ExecJS.js\\", sfn);
    }
    
    print("Setup done!");
    return _void;
}

function help() {
    return print(
        "Welcome to ExecJS!\n" +
        " ExecJS is a JScript command console that is useful and easy to use, as\n" +
        " well as advanced and perfect. It lets you seamlessly execute JScript\n" +
        " commands - whether they are single-line or multiline - and then view\n" +
        " their results in the blink of an eye! It can also run JScript code in\n" +
        " the context of other applications - including HTAs and webpages in IE\n" +
        " - whether they are on the local computer, or a remote machine on the\n" +
        " network. Plus, it can open a blank page on the IE browser or a virtual\n" +
        " HTA window, allowing you to execute commands in the context of it for\n" +
        " the sake of a coding/scientific test. Therefore, ExecJS is useful to\n" +
        " debug an HTA, do calculations, test properties and methods of various\n" +
        " Automation Objects, cheat at games and webpages, transfer data over the\n" +
        " network, .etc.\n" +
        "\nHow to install:\n" +
        " If this is the first time you're running this app, please execute\n" +
        " 'setup'. Then, getting the message 'Setup done!' indicates success of\n" +
        " the installation.\n" +
        "\nHow to use:\n" +
        " To execute a single-line command, just type it and press Enter. To\n" +
        " execute a multiline command, type a space, followed by the first line\n" +
        " of code. Press Enter, and type the other lines of code. When you reached\n" +
        " the last line, append a space character to it, an press Enter to begin\n" +
        " the execution. Note: if you've written part of a multiline command and\n" +
        " you wish to not write the rest and just cancel its execution, type an\n" +
        " octothorp (#), and press Enter.\n" +
        " To execute commands in the context of an IE browser or a virtual HTA\n" +
        " window, try the virtualIE() and virtualHTA() functions, respectively.\n" +
        " You can also specify a title for the window as a parameter to these\n" +
        " functions.\n" +
        " To exit the app, leave the line blank and press Enter. Do NOT use the\n" +
        " close button in the title bar of the console window to exit.\n" +
        "\nAbout:\n" +
        " This app is made by Javad Bayat. Version: 1.0\n" +
        " Hope the documentation of ExecJS will be published soon, so that you get\n" +
        " to know ExecJS better."
    );
}