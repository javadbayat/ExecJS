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

FatalError.prototype = Error;
//Predefined variables
var ejs, _internal, _echo = true, _lastres = "", 
_lastiniterr = null, _lasterr = null, _ = this, 
_target = _, _cmdexit = "", _cmdcancel = "#", 
_void = {toString:undefined}, _version = "1.0",
_args = col2arr(WScript.Arguments), _cmd = "",
_timing = false, _sdate = null, _edate = null,
_cmdsetup = "setup", _cmdhelp = "help", _ncmds = 0,
_vd_entry_limit = 512;

var wshShell = WScript.CreateObject("WScript.Shell");
var FSO = WScript.CreateObject("Scripting.FileSystemObject");

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

var VBS = WScript.CreateObject("MSScriptControl.ScriptControl");
VBS.Language = "VBScript";
VBS.AddObject("WScript", WScript);
VBS.AddObject("WSH", WSH);

VBS.AddCode("Function ConvertArray(jsa)\nDim vba()\nReDim vba(jsa.length - 1)\nDim i, e\ni = 0\nFor Each e In jsa\nIf IsObject(e) Then\nSet vba(i) = e\nElse\nvba(i) = e\nEnd If\ni = i + 1\nNext\nConvertArray = vba\nEnd Function");
var getVBArray = VBS.Eval('GetRef("ConvertArray")');

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

function var_dump() {
    var vdui = virtualHTA("Variant Dumper", true);
    
    vdui.ejs = _;
    var document = vdui.document;
    document.charset = "UTF-8";
    document.body.scroll = "auto";
    
    var styleSheet = document.createStyleSheet();
    with (styleSheet) {
        addRule("#btnDebug", "position:absolute; right:7px; bottom:7px;");
        addRule("#btnDebug button", "color:red; zoom:125%;");
        addRule(".variantEntry", "width:100%; height:auto;");
        addRule(".variantIcon", "cursor:hand; font-size:large; margin-right:4px;");
        addRule(".booleanValue", "color:blue;");
        addRule(".decimalValue", "color:blue;");
        addRule(".hexadecimalValue", "color:red;");
        addRule(".octalValue", "color:green;");
        addRule(".doubleValue", "color:orange;");
        addRule(".nanValue", "border:1px solid red; color:gray; padding:1px;");
        addRule(".infinityValue", "border:1px solid green; color:gray; padding:1px;");
        addRule(".unsupportedType", "font-size:large; color:red; font-weight:bold;");
        addRule(".stringInfo", "display:inline; height:7em; width:300px; border:1px solid black; overflow:auto; vertical-align:middle;");
        addRule(".stringAttributes", "color:green;");
        addRule("button", "margin:4px;");
        addRule(".subset", "width:100%; height:auto; margin-left:40px; border-left:1px solid black; padding:4px;");
        addRule(".dateInfo h1", "font-size:large;");
        addRule(".dispex", "width:300px; display:block;");
        addRule(".collectionContent .indicator", "font-size:larger;");
        addRule(".collectionContent .indicator span", "color:red;");
        addRule(".functionInfo h1", "font-size:large; margin:5px;");
        addRule(".functionInfo code", "background-color:#c0c0c0; display:block; overflow:auto;");
        addRule(".objectContent h1", "font-size:large; margin:5px;");
        addRule(".objrefBox", "overflow:scroll; width:50%; vertical-align:middle");
        addRule(".arrayBounds span", "color:blue;");
    }
    
    vdui.findElemByName = vdui.Function('c', 'n',
            'for (var i = 0;i < c.length;i++) {' +
            'if (c[i].name == n) return c[i];' +
            '} ' +
            'return null;'
    );
    
    var funcStore = [
        vdui.Function( //0
            'if (!ejs.scriptex) {' +
            'alert("This feature requires Scriptex to be installed on the system.");' +
            'return;' +
            '}' +
            'var iid = prompt("Enter the IID of the interface to check its existence on the object.", "");' +
            'if (!iid) return;' +
            'try {' +
            'var interfaceName = ejs.wshShell.RegRead("HKCR\\\\Interface\\\\" + iid + "\\\\");' +
            '} catch (e) {' +
            'alert("Unable to find information about the interface in the registry:\\n" + e.message);' +
            'return;' +
            '}' +
            'var obj = this.parentElement.parentElement.variantValue;' +
            'var eid = this.parentElement.parentElement.entryID;' +
            'ejs.print("V" + eid + "# Checking the existence of the interface \'" + interfaceName + "\'...");' +
            'ejs.print("V" + eid + "# IID: " + iid);' +
            'var result = ejs.scriptex.SupportsInterface(obj, iid);' +
            'ejs.print("V" + eid + "# " + (result ? "Interface exists on the object." : "Interface does not exist on the object.") + ejs.beep);'
        ),
        vdui.Function( //1
            'var entry = this.parentElement.parentElement;' +
            'if (entry.variantOBJREF) {' +
            'ejs.scriptex.Clipboard = entry.variantOBJREF;' +
            '} else {' +
            'var objref = ejs.getref(entry.variantValue);' +
            'if (!objref) {' +
            'alert("Unable to generate an OBJREF Moniker for the object.");' +
            'return;' +
            '}' +
            'ejs.print("V" + (entry.entryID) + "# " + objref + ejs.beep);' +
            'entry.variantOBJREF = objref;' +
            'this.innerText = "Copy OBJREF";' +
            '}'
        ),
        vdui.Function( //2
            'if (event.shiftKey) {' +
            'var btn = findElemByName(this.children, "btnMakeDate");' +
            'if (btn) btn.style.display = "";' +
            '}'
        ),
        vdui.Function( //3
            'var btn = findElemByName(this.children, "btnMakeDate");' +
            'if (btn) btn.style.display = "none";'
        ),
        vdui.Function('this.style.color = "red";'), //4
        vdui.Function('this.style.color = "";'), //5
        vdui.Function( //6
            'var val = this.parentElement.variantValue;' +
            'var varName = prompt("Please enter the variable name to put the value into.", "");' +
            'if (varName) ejs[varName] = val;'
        ),
        vdui.Function( //7
            'if ((event.keyCode == 13) && (document.selection.type == "Text")) {' +
            'var range = document.selection.createRange();' +
            'if (range.parentElement() !== this.children[0]) return;' +
            'var txt = range.text;' +
            'range.moveStart("character", -1);' +
            'if (range.text.indexOf("\\\\") != -1) return;' +
            'var eid = this.parentElement.entryID;' +
            'var result = "V" + eid + "# Char codes: ";' + 
            'for (var i = 0;i < txt.length;i++) {' +
            'if (i) result += ", ";' +
            'result += txt.charCodeAt(i);' +
            '} ' +
            'ejs.print(result);' +
            '}'
        ),
        vdui.Function( //8
            'if (this.style.backgroundColor == "#9bbb59") {' +
            'this.advancedSection.style.display = "none";' +
            'this.style.backgroundColor = "#c0504d";' +
            '} else {' +
            'this.advancedSection.style.display = "";' +
            'this.style.backgroundColor = "#9bbb59";' +
            '}'
        ),
        vdui.Function( //9
            'alert("Value: " + this.parentElement.variantValue);'
        ),
        vdui.Function( //10
            'var entry = this.parentNode.parentNode.parentNode;' +
            'ejs.print("V" + (entry.entryID) + "# Dumping user-defined members of the object...");' +
            'ejs.obj_dump(entry.variantValue);' +
            'ejs.beep();'
        ),
        vdui.Function( //11
            "var r = document.body.createTextRange();" +
            "r.moveToElementText(this);" +
            "r.select();"
        ),
        vdui.Function( //12
            "this.parentNode.lastChild.removeNode(true);" +
            "de(this.th.value, this.parentNode);"
        ),
        vdui.Function( //13
            'try {' +
            'var d = new Date(this.parentElement.variantValue);' +
            'if (isNaN(d.valueOf())) throw new Error("Invalid date!");' +
            '} catch(e) {' +
            'alert("Unable to make a date object from this variant:\\n" + e.message);' +
            'return;' +
            '} ' +
            'this.dumpDate(d, this.parentElement);' +
            'this.removeNode(true);'
        )
    ];
    
    vdui.de = dumpEntry;
    
    var baseEntryID = 1;
    var exception_limitReached = 1;
    try {
        for (var i = 0;i < arguments.length;i++)
            dumpEntry(arguments[i], document.body);
    }
    catch (exception) {
        if (exception === exception_limitReached)
            print(baseEntryID + " variant entries processed so far.\nAborting dump operation due to entry limit...");
        else
            throw exception;
    }
    
    insertDebugButton();
    
    var exitEvent = new Thing(false, true);
    var exitFunc = vdui.eval("onunload = function() { arguments.callee.ee.value = true; }");
    exitFunc.ee = exitEvent;
    exitEvent.acquire();
    
    return _void;
    
    function dumpEntry(v, container, vk) {
        if (baseEntryID >= _vd_entry_limit)
            throw exception_limitReached;
        
        var arrow = String.fromCharCode(8594, 32);
        
        var entry = document.createElement("div");
        entry.className = "variantEntry";
        entry.variantValue = v;
        entry.variantKey = vk;
        entry.entryID = baseEntryID++;
        
        entry.onmouseover = funcStore[2];
        entry.onmouseout = funcStore[3];
        
        if (vk !== undefined) {
            var spnKey = document.createElement("span");
            spnKey.className = "variantKey";
            spnKey.innerText = formatKey(vk);
            entry.appendChild(spnKey);
            
            addText(entry, arrow);
        }
        
        var spnIcon = document.createElement("span");
        spnIcon.className = "variantIcon";
        spnIcon.innerText = String.fromCharCode(9784);
        spnIcon.title = "V" + entry.entryID;
        spnIcon.onmouseover = funcStore[4];
        spnIcon.onmouseout = funcStore[5];
        spnIcon.onclick = funcStore[6];
        entry.appendChild(spnIcon);
        
        if (scriptex)
            var vt = scriptex.GetVarType(v);
        else {
            vdui.foo = v;
            vdui.execScript("If IsObject(foo) Then foo = 9 Else foo = VarType(foo) End If", "VBScript");
            var vt = vdui.foo;
        }
        
        var spnType = document.createElement("span");
        spnType.className = "variantType";
        
        if (scriptex)
            spnType.innerText = scriptex.GetVarTypeString(v);
        else {
            switch (vt) {
            case 0 : //VT_EMPTY
                spnType.innerText = "VT_EMPTY";
                break;
            case 1 : //VT_NULL
                spnType.innerText = "VT_NULL";
                break;
            case 11 : //VT_BOOL
                spnType.innerText = "VT_BOOL";
                break;
            case 3 : //VT_I4
                spnType.innerText = "VT_I4";
                break;
            case 5 : //VT_R8
                spnType.innerText = "VT_R8";
                break;
            case 8 : //VT_BSTR
                spnType.innerText = "VT_BSTR";
                break;
            case 7 : //VT_DATE
                spnType.innerText = "VT_DATE";
                break;
            case 13 : //VT_UNKNOWN
                spnType.innerText = "VT_UNKNOWN";
                break;
            case 9 : //VT_DISPATCH
                spnType.innerText = "VT_DISPATCH";
                break;
            default :
                spnType.innerText = "?";
            }
        }
        
        spnType.title = "Type code: " + vt + "\nJS type: " + (typeof v);
        
        switch (typeof v) {
        case "undefined" :
            spnType.style.color = "black";
            break;
        case "boolean" :
            spnType.style.color = "rgb(192, 0, 0)";
            break;
        case "number" :
            spnType.style.color = "rgb(255, 102, 0)";
            break;
        case "string" :
            spnType.style.color = "rgb(0, 102, 0)";
            break;
        case "object" :
            spnType.style.color = "rgb(0, 102, 255)";
            break;
        case "function" :
            spnType.style.color = "rgb(153, 0, 204)";
            break;
        case "date" :
            spnType.style.color = "rgb(255, 0, 255)";
            break;
        case "unknown" :
            spnType.style.color = "rgb(153, 102, 51)";
            break;
        }
        
        entry.appendChild(spnType);
                
        if ((vt != 0) && (vt != 1)) {
            addText(entry, String.fromCharCode(32, 32, 8594, 32, 32));
            
            switch (vt) {
            case 11 : //VT_BOOL
                var spnValue = document.createElement("span");
                spnValue.className = "booleanValue";
                spnValue.innerText = v;
                entry.appendChild(spnValue);
                break;
            case 3 : //VT_I4
                var spnValue = document.createElement("span");
                
                var spnDecimal = document.createElement("span");
                spnDecimal.className = "decimalValue";
                spnDecimal.innerText = v;
                spnDecimal.title = "Locale: " + v.toLocaleString();
                spnValue.appendChild(spnDecimal);
                
                addText(spnValue, " (HEX: ");
                
                var spnHex = document.createElement("span");
                spnHex.className = "hexadecimalValue";
                spnHex.innerText = v.toString(16).toUpperCase();
                spnValue.appendChild(spnHex);
                
                addText(spnValue, ", OCT: ");
                
                var spnOct = document.createElement("span");
                spnOct.className = "octalValue";
                spnOct.innerText = v.toString(8);
                spnValue.appendChild(spnOct);
                
                addText(spnValue, ")");
                
                entry.appendChild(spnValue);
                
                addDateCreationButton();
                break;
            case 5 : //VT_R8
                var spnValue = document.createElement("span");
                spnValue.innerText = v;
                if (isNaN(v))
                    spnValue.className = "nanValue";
                else if ((v == Infinity) || (v == -Infinity))
                    spnValue.className = "infinityValue";
                else {
                    spnValue.className = "doubleValue";
                    spnValue.title = "Locale: " + v.toLocaleString();
                }
                
                entry.appendChild(spnValue);
                
                addDateCreationButton();
                break;
            case 8 : //VT_BSTR
                var divInfo = document.createElement("div");
                divInfo.className = "stringInfo";
                
                addText(divInfo, '"');
                var spnValue = document.createElement("span");
                spnValue.className = "stringValue";
                spnValue.innerText = codeForm(v);
                divInfo.appendChild(spnValue);
                addText(divInfo, '" ');
                
                divInfo.onkeyup = funcStore[7];
                
                var spnAttr = document.createElement("span");
                spnAttr.className = "stringAttributes";
                spnAttr.innerText = "(length: " + (v.length) + ")";
                divInfo.appendChild(spnAttr);
                
                entry.appendChild(divInfo);
                
                addDateCreationButton();
                break;
            case 7 : //VT_DATE
                var spnValue = document.createElement("span");
                spnValue.className = "dateValue";
                spnValue.innerHTML = formatDate(v);
                entry.appendChild(spnValue);
                
                dumpDate(new Date(v), entry, true);
                break;
            case 13 : //VT_UNKNOWN
                var btnAdvanced = document.createElement("button");
                btnAdvanced.innerText = "Advanced";
                btnAdvanced.style.backgroundColor = "#c0504d";
                
                btnAdvanced.onclick = funcStore[8];
                
                entry.appendChild(btnAdvanced);
                
                var as = document.createElement("div");
                as.className = "subset";
                as.style.display = "none";
                
                addCommonButtonsForObjects();
                
                btnAdvanced.advancedSection = entry.appendChild(as);
                break;
            case 9 : //VT_DISPATCH
                var btnValue = document.createElement("button");
                btnValue.innerText = "Value";
                btnValue.onclick = funcStore[9];
                entry.appendChild(btnValue);
                
                var btnAdvanced = document.createElement("button");
                btnAdvanced.innerText = "Advanced";
                btnAdvanced.style.backgroundColor = "#c0504d";
                
                btnAdvanced.onclick = funcStore[8];
                
                entry.appendChild(btnAdvanced);
                
                var as = document.createElement("div");
                as.className = "subset";
                as.style.display = "none";
                
                addCommonButtonsForObjects();
                
                var btnBrowse = document.createElement("button");
                btnBrowse.innerText = "Browse";
                as.appendChild(btnBrowse);
                
                if ((!!scriptex) && (scriptex.SupportsInterface(v, "{A6EF9860-C720-11D0-9337-00A0C90DCAA9}"))) {
                    var fs = document.createElement("fieldset");
                    fs.className = "dispex";
                    
                    var legend = document.createElement("legend");
                    legend.innerText = "DispatchEx";
                    fs.appendChild(legend);
                    
                    var btnDump = document.createElement("button");
                    btnDump.innerText = "Dump";
                    btnDump.onclick = funcStore[10];
                    fs.appendChild(btnDump);
                    
                    as.appendChild(fs);
                }
                
                btnAdvanced.advancedSection = entry.appendChild(as);
                
                var isDictionary = scriptex ? scriptex.SupportsInterface(v, "{42C642C1-97E1-11CF-978F-00A02463E06F}") : false;
                try {
                    if (v instanceof Enumerator) {
                        var ce = v;
                        ce.moveFirst();
                    }
                    else
                        var ce = new Enumerator(v);
                    
                    var cc = document.createElement("div");
                    cc.className = "subset collectionContent";
                    
                    var indicator = document.createElement("span");
                    indicator.className = "indicator";
                    addText(indicator, "Collection(");
                    
                    var spnNumItems = document.createElement("span");
                    indicator.appendChild(spnNumItems);
                    
                    addText(indicator, ")");
                    cc.appendChild(indicator);
                    
                    addText(cc, "\n{");
                    
                    var numItems = 0;
                    for (;!ce.atEnd();ce.moveNext()) {
                        var i = ce.item();
                        if (isDictionary)
                            dumpEntry(v(i), cc, i);
                        else
                            dumpEntry(i, cc, numItems);
                        
                        numItems++;
                    }
                    
                    addText(cc, "}");
                    
                    spnNumItems.innerText = numItems;
                    entry.appendChild(cc);
                }
                catch (e) {
                    if (e === exception_limitReached)
                        throw e;
                    // if (e.number != -2146827837)
                        // throw e;
                }
                
                if (v instanceof Date)
                    dumpDate(v, entry);
                else if (v instanceof Function) {
                    var functionInfo = document.createElement("div");
                    functionInfo.className = "subset functionInfo";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "Function object";
                    functionInfo.appendChild(heading);
                    
                    var font = document.createElement("font");
                    font.color = "blue";
                    font.innerText = "Number of parameters ";
                    functionInfo.appendChild(font);
                    
                    addText(functionInfo, arrow);
                    
                    var nParameters = document.createElement("span");
                    nParameters.innerText = v.length;
                    functionInfo.appendChild(nParameters);
                    
                    lineBreak(functionInfo);
                    
                    font = document.createElement("font");
                    font.color = "blue";
                    font.innerText = "Function code:";
                    functionInfo.appendChild(font);
                    
                    var functionCode = document.createElement("code");
                    functionCode.innerText = v.toString();
                    functionInfo.appendChild(functionCode);
                    
                    entry.appendChild(functionInfo);
                }
                else if (v instanceof Error) {
                    var errorInfo = document.createElement("div");
                    errorInfo.className = "subset objectContent";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "Error info " + arrow + v.name;
                    errorInfo.appendChild(heading);
                    
                    var list = document.createElement("ul");
                    
                    var itemDescription = document.createElement("li");
                    itemDescription.innerText = "Description " + arrow + v.description;
                    list.appendChild(itemDescription);
                    
                    var itemNumber = document.createElement("li");
                    itemNumber.innerText = "Number " + arrow + (v.number) + 
                            " ( HRESULT " + arrow + (v.number.toHRESULT()) + " )";
                    list.appendChild(itemNumber);
                    
                    errorInfo.appendChild(list);
                    entry.appendChild(errorInfo);
                }
                else if (v instanceof Boolean) {
                    var boolInfo = document.createElement("div");
                    boolInfo.className = "subset objectContent";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "Boolean object";
                    boolInfo.appendChild(heading);
                    
                    var V = v.valueOf();
                    dumpEntry(V, boolInfo);
                    
                    entry.appendChild(boolInfo);
                }
                else if (v instanceof Number) {
                    var numberInfo = document.createElement("div");
                    numberInfo.className = "subset objectContent";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "Number object";
                    numberInfo.appendChild(heading);
                    
                    var V = v.valueOf();
                    dumpEntry(V, numberInfo);
                    
                    entry.appendChild(numberInfo);
                }
                else if (v instanceof String) {
                    var stringInfo = document.createElement("div");
                    stringInfo.className = "subset objectContent";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "String object";
                    stringInfo.appendChild(heading);
                    
                    var V = v.valueOf();
                    dumpEntry(V, stringInfo);
                    
                    entry.appendChild(stringInfo);
                }
                else if (v instanceof Thing) {
                    var thingInfo = document.createElement("div");
                    thingInfo.className = "subset thingInfo";
                    
                    var heading = document.createElement("h1");
                    heading.innerText = "Thing object";
                    thingInfo.appendChild(heading);
                    
                    var list = document.createElement("ul");
                    
                    var itemInterval = document.createElement("li");
                    itemInterval.innerText = "Interval " + arrow + v.interval;
                    list.appendChild(itemInterval);
                    
                    if (v.objref) {
                        var itemOBJREF = document.createElement("li");
                        itemOBJREF.innerText = v.objref;
                        itemOBJREF.className = "objrefBox";
                        itemOBJREF.ondblclick = funcStore[11];
                        list.appendChild(itemOBJREF);
                    }
                    
                    var itemValue = document.createElement("li");
                    itemValue.innerText = "Value" + String.fromCharCode(8595, 32);
                    
                    var btnRefresh = document.createElement("button");
                    btnRefresh.innerText = "Refresh";
                    btnRefresh.th = v;
                    btnRefresh.onclick = funcStore[12];
                    itemValue.appendChild(btnRefresh);
                    
                    dumpEntry(v.value, itemValue);
                    
                    list.appendChild(itemValue);
                    
                    thingInfo.appendChild(list);
                    entry.appendChild(thingInfo);
                }
                else if (v instanceof VBArray) {
                    var lb = v.lbound();
                    var ub = v.ubound();
                    var arr = v.toArray();
                    
                    var vbac = document.createElement("div");
                    vbac.className = "subset vbarrayContent";
                    
                    var spnBounds = document.createElement("span");
                    spnBounds.className = "arrayBounds";
                    addText(spnBounds, "VBArray(");
                    
                    var spnLBound = document.createElement("span");
                    spnLBound.innerText = lb;
                    spnBounds.appendChild(spnLBound);
                    
                    addText(spnBounds, " To ");
                    
                    var spnUBound = document.createElement("span");
                    spnUBound.innerText = ub;
                    spnBounds.appendChild(spnUBound);
                    
                    addText(spnBounds, ")");
                    vbac.appendChild(spnBounds);
                    
                    addText(vbac, "\n{");
                    
                    for (var i = 0;i < arr.length;i++)
                        dumpEntry(arr[i], vbac, lb + i);
                    
                    addText(vbac, "}");
                    
                    entry.appendChild(vbac);
                }
                break;
            default :
                if (vt > 8192) {
                    //Dump the safearray
                    var lb, ub, arr;
                    
                    var spnBounds = document.createElement("span");
                    spnBounds.className = "arrayBounds";
                    addText(spnBounds, "SafeArray(");
                    
                    var spnLBound = document.createElement("span");
                    vdui.foo = v;
                    vdui.execScript("foo = LBound(foo)", "VBScript");
                    spnLBound.innerText = lb = vdui.foo;
                    spnBounds.appendChild(spnLBound);
                    
                    addText(spnBounds, " To ");
                    
                    var spnUBound = document.createElement("span");
                    vdui.foo = v;
                    vdui.execScript("foo = UBound(foo)", "VBScript");
                    spnUBound.innerText = ub = vdui.foo;
                    spnBounds.appendChild(spnUBound);
                    
                    addText(spnBounds, ")");
                    entry.appendChild(spnBounds);
                    
                    if (scriptex) {
                        arr = scriptex.ConvertArray(v, new Array);
                        
                        var ac = document.createElement("div");
                        ac.className = "subset arrayContent";
                        
                        addText(ac, "{");
                        
                        for (var i = 0;i < arr.length;i++)
                            dumpEntry(arr[i], ac, lb + i);
                        
                        addText(ac, "}");
                        
                        entry.appendChild(ac);
                    }
                }
                else {
                    var spn = document.createElement("span");
                    spn.className = "unsupportedType";
                    spn.innerText = String.fromCharCode(9587);
                    spn.title = "This variant type is not supported.";
                    entry.appendChild(spn);
                }
            }
        }
        
        container.appendChild(entry);
        
        if (functionCode) {
            if (functionCode.offsetHeight > 350)
                functionCode.style.height = 350;
        }
        
        function addText(elem, txt) {
            return elem.appendChild(document.createTextNode(txt));
        }
        
        function lineBreak(elem) {
            elem.appendChild(document.createElement("br"));
        }
        
        function codeForm(str) {
            var result = '';
            for (var i = 0;i < str.length;i++) {
                switch (str.charAt(i)) {
                case "\0" :
                    result += "\\0";
                    break;
                case "\t" :
                    result += "\\t";
                    break;
                case "\r" :
                    result += "\\r";
                    break;
                case "\n" :
                    result += "\\n";
                    break;
                case "\\" :
                    result += "\\\\";
                    break;
                case "\"" :
                    result += "\\\"";
                    break;
                case "\'" :
                    result += "\\\'";
                    break;
                default :
                    result += str.charAt(i);
                }
            }
            
            return result;
        }
        
        function formatKey(key) {
            var result = "[";
            
            if (typeof key == "number")
                result += key;
            else {
                try {
                    key = String(key);
                }
                catch (e) {
                    key = "?";
                }
                
                result += '"' + codeForm(key) + '"';
            }
            
            result += "]";
            return result;
        }
        
        function addDateCreationButton() {
            var btn = document.createElement("button");
            btn.name = "btnMakeDate";
            btn.innerText = "Make Date";
            
            btn.onclick = funcStore[13];
            btn.dumpDate = dumpDate;
            
            btn.style.display = "none";
            entry.appendChild(btn);
        }
        
        function dumpDate(d, container, noHeading) {
            var dateInfo = document.createElement("div");
            dateInfo.className = "subset dateInfo";
            dateInfo.dateObject = d;
            
            if (!noHeading) {
                var heading = document.createElement("h1");
                heading.innerHTML = "Date info " + arrow + (formatDate(d));
                dateInfo.appendChild(heading);
            }
            
            var list = document.createElement("ul");
            
            var itemValue = document.createElement("li");
            itemValue.innerText = "Value " + arrow + (d.valueOf());
            list.appendChild(itemValue);
            
            var itemLocaleDate = document.createElement("li");
            itemLocaleDate.innerText = "Locale date " + arrow + (d.toLocaleDateString());
            list.appendChild(itemLocaleDate);

            var itemLocaleTime = document.createElement("li");
            itemLocaleTime.innerText = "Locale time " + arrow + (d.toLocaleTimeString());
            list.appendChild(itemLocaleTime);
            
            var itemUTC = document.createElement("li");
            itemUTC.innerText = "UTC " + arrow + (d.toUTCString());
            list.appendChild(itemUTC);
            
            var itemDay = document.createElement("li");
            itemDay.style.color = "#008000";
            var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
            itemDay.innerText = "Day of week " + arrow + (days[d.getDay()]);
            itemDay.title = "Day of week in UTC " + arrow + (days[d.getUTCDay()]);
            list.appendChild(itemDay);

            var itemYear = document.createElement("li");
            itemYear.style.color = "#ff8040";
            itemYear.innerText = "Year " + arrow + (d.getFullYear());
            itemYear.title = "Year in UTC " + arrow + (d.getUTCFullYear());
            list.appendChild(itemYear);
            
            var itemMonth = document.createElement("li");
            itemMonth.style.color = "#ff0000";
            itemMonth.innerText = "Month " + arrow + (d.getMonth() + 1);
            itemMonth.title = "Month in UTC " + arrow + (d.getUTCMonth() + 1);
            list.appendChild(itemMonth);
            
            var itemDate = document.createElement("li");
            itemDate.style.color = "#0000a0";
            itemDate.innerText = "Day " + arrow + (d.getDate());
            itemDate.title = "Day in UTC " + arrow + (d.getUTCDate());
            list.appendChild(itemDate);
            
            var itemHours = document.createElement("li");
            itemHours.style.color = "#ff00ff";
            itemHours.innerText = "Hours " + arrow + (d.getHours());
            itemHours.title = "Hours in UTC " + arrow + (d.getUTCHours());
            list.appendChild(itemHours);
            
            var itemMinutes = document.createElement("li");
            itemMinutes.style.color = "#5c6600";
            itemMinutes.innerText = "Minutes " + arrow + (d.getMinutes());
            itemMinutes.title = "Minutes in UTC " + arrow + (d.getUTCMinutes());
            list.appendChild(itemMinutes);
            
            var itemSeconds = document.createElement("li");
            itemSeconds.style.color = "#663300";
            itemSeconds.innerText = "Seconds " + arrow + (d.getSeconds());
            itemSeconds.title = "Seconds in UTC " + arrow + (d.getUTCSeconds());
            list.appendChild(itemSeconds);
            
            var itemMilliseconds = document.createElement("li");
            itemMilliseconds.innerText = "Milliseconds " + arrow + (d.getMilliseconds());
            itemMilliseconds.title = "Milliseconds in UTC " + arrow + (d.getUTCMilliseconds());
            list.appendChild(itemMilliseconds);
            
            var itemTimezoneOffset = document.createElement("li");
            itemTimezoneOffset.innerText = "Time zone offset " + arrow + (d.getTimezoneOffset()) + " minutes";
            list.appendChild(itemTimezoneOffset);
            
            dateInfo.appendChild(list);

            container.appendChild(dateInfo);
        }
        
        function formatDate(strDate) {
            if (typeof strDate != "string")
                strDate += "";
            
            try {
                var result = strDate.split(" ");
                result[0] = result[0].fontcolor("#008000");
                result[5] = result[5].fontcolor("#ff8040");
                result[1] = result[1].fontcolor("#ff0000");
                result[2] = result[2].fontcolor("#0000a0");
                
                result[3] = result[3].split(":");
                result[3][0] = result[3][0].fontcolor("#ff00ff");
                result[3][1] = result[3][1].fontcolor("#5c6600");
                result[3][2] = result[3][2].fontcolor("#663300");
                
                result[3] = result[3].join(":");
                result = result.join(" ");
                return result;
            }
            catch (e) {
                return strDate;
            }
        }
        
        function addCommonButtonsForObjects() {
            var btnInterfaceExists = document.createElement("button");
            btnInterfaceExists.innerText = "Interface exists?";
            btnInterfaceExists.onclick = funcStore[0];
            as.appendChild(btnInterfaceExists);
            
            var btnOBJREF = document.createElement("button");
            btnOBJREF.innerText = "Generate OBJREF";
            btnOBJREF.onclick = funcStore[1];
            as.appendChild(btnOBJREF);
        }
    }

    function insertDebugButton() {
        var debugButton = document.createElement("button");
        debugButton.innerText = "Debug!";
        debugButton.style.visibility = "hidden";
        
        debugButton.onclick = vdui.Function("arguments.callee.exposeWindow(window);");
        debugButton.onclick.exposeWindow = function(vdui) {
            if (!scriptex) {
                WScript.StdErr.WriteLine("VD Error: Debugging the Variant Dumper User Interface requires Scriptex to be installed on the system.");
                return;
            }
            
            var th = new Thing(vdui);
            scriptex.Clipboard = th.objref;
            print(beep + "The OBJREF Moniker for the Variant Dumper User Interface has been placed on clipboard.");
        };
        
        var debugButtonContainer = document.createElement("span");
        debugButtonContainer.id = "btnDebug";
        debugButtonContainer.onmouseover = vdui.Function('if (event.shiftKey) this.firstChild.style.visibility = "visible";');
        debugButtonContainer.onmouseout = vdui.Function('this.firstChild.style.visibility = "hidden";');
        debugButtonContainer.appendChild(debugButton);
        document.body.appendChild(debugButtonContainer);
    }
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