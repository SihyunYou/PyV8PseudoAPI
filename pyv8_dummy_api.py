#
# -- PyV8 Window Script Host Dummy API Module --
#
# Standalone JScript file to be executed by Windows Script Host
# (This file is not going to be used as a web application)

import PyV8
import win32com.client
import os
import sys
import platform
import _winreg
import time

def regkey_value(path, name="", start_key = None):
    if isinstance(path, str):
        path = path.split("\\")
    if start_key is None:
        start_key = getattr(winreg, path[0])
        return regkey_value(path[1:], name, start_key)
    else:
        subkey = path.pop(0)
    with winreg.OpenKey(start_key, subKey) as handle:
        assert handle
        if path:
            return regkey_value(path, name, handle)
        else:
            desc, i = None, 0
            while not desc or desc[0] != name:
                desc = winreg.EnumValue(handle, i)
                i += 1
            return desc[1]
                
#
### WScript(Window Script) class to be default object
#
class MyWScript(PyV8.JSClass) :
    def Sleep(self, x) :
        time.sleep(x/1000.)
        
    def Arguments(self) :
	return len(sys.argv)
	
    def CreateObject(self, progid) :
        print '[*] CreateObject :', progid
        if progid.lower() == 'wscript.shell' :
            return MyWshShell()
        elif progid.lower() == 'msxml2.xmlhttp' :
            return MyXMLHTTP()
        elif progid.lower() == 'adodb.stream' :
            return MyAdodbStream()
	elif progid.lower() == 'scripting.filesystemobject' :
	    return MyFileSystemObject()

    def ActiveXObject(self, progid) :
	print '[*] ActiveXObject :', progid
        if progid.lower() == 'wscript.shell' :
            return MyWshShell()
        elif progid.lower() == 'msxml2.xmlhttp' :
            return MyXMLHTTP()
        elif progid.lower() == 'msxml2.serverxmlhttp.6.0':
	    return MyXML2SERVERHTTP()
        elif progid.lower() == 'adodb.stream' :
            return MyAdodbStream()
	elif progid.lower() == 'scriptlet.typelib' :
	    return MyScriptlet()
	elif progid.lower() == 'scripting.filesystemobject' :
            return MyFileSystemObject()
	elif progid.lower() == 'dynamicwrapperx.2' :
	    return MyDynamicWrapper()

	
    def Enumerator(self, progid) :
	return MyEnumerator()

#
### The type or class of the object to create.
#
class MyWshShell(PyV8.JSClass) : 
    def __init__(self) :
       self.Type = 0
       self.Charset = 0

    def ExpandEnvironmentStrings(self, x) :
        s = _winreg.ExpandEnvironmentStrings(unicode(x))
        print '[*] WshShell.ExpandEnvironmentStrings :', x
        print '    [-] :', s
        return s
    def Environment(self, x) :
       print '[*] WshShell.Environment :', x
       if x.lower() == 'process':
            return MyWshProcEnv()
        
    def Run(self, x) :
        print '[*] WshShell.Run :'
        print '    [-] :', x      
    def Run(self, x, y=0, z=1) :
        print '[*] WshShell.Run :'
        print '    [-] :', x   

    def RegRead(self, RegPath) :
        print '[*] WshShell.RegRead :'
        print '    [-] :', RegPath
        return reg_keyvalue(RegPath, RegPath.substr(RegPath.rfind('\\'), RegPath.length))        
    def RegWrite(self, RegPath, x, y) :
        print '[*] WshShell.RegWrite :'
        print '    [-] :', RegPath
        print '    [-] :', x
        print '    [-] :', y
        
    def SpecialFolders(self, x) :
        print '[*] Stream.SpecialFolders :'
        print '    [-] :', x

class MyXMLHTTP(PyV8.JSClass) : 
    def open(self, x, y, z) :
        print '[*] XMLHTTP.open :'
        print '    [-] :', x  
        print '    [-] :', y  
        print '    [-] :', z
    def send(self) :
        pass
        
class MyXML2SERVERHTTP(PyV8.JSClass) :
    def setTimeouts(self, a, b, c, d) :
	print '[*] MSXMLSERVERHTTP.setTimeouts :'
        print '    [-] :', a
        print '    [-] :', b  
        print '    [-] :', c
        print '    [-] :', d

class MyFileSystemObject(PyV8.JSClass) :
    Drives = -1
    def CreateTextFile(self, dir, flags) :
        print '    [-] CreateTextFile :', dir
        return FilePointer()
    def GetFolder(self, path) :
        print '    [-] GetFolder :', path
        return FilePointer()
    def CopyFile(self, exist, new) :
        print '    [-] CopyFile : ', exist, ' , ', new
    def FileExists(self, t) :
	print '    [-] :', t

class MyAdodbStream(PyV8.JSClass) :
    def open(self) :
        pass
    def write(self, x) :
        pass
    def SaveToFile(self, x, y) :
        print '[*] SaveToFile :', x
    def close(self) :
        pass

    def Open(self) :
	pass
    def WriteText(self, x) :
        print '[*] Stream.WriteText :'
        print '    [-] :', x
    def SaveToFile(self, x) :
	print '[*] Stream.SaveToFile :'
	print '    [-] :', x
    def LoadFromFile(self, x) :
	print '[*] Stream.LoadFromFile :'
	print '    [-] :', x
    def ReadText(self) :
	pass
    def Close(self) :
	pass

#
# Sub class
#
class FilePointer(PyV8.JSClass) :
    Files = -1
    def Write(self, txt) :
        print '        [-] Write : ', txt
    def Close(self) :
        pass
    
class MyEnumerator(PyV8.JSClass) :
    def atEnd(self) :
	return 1
    def moveNext(self) :
	return 0
    def item(self) :
	return 0

class MyScriptlet(PyV8.JSClass) :
    GUID = 'aaaa'
    def substr(self, x, y) :
        pass

# Gilles Laurent's DynaWrap ocx a chance
# This kind of dll needs to be registered on the target system like regsvr32 /s DynaWrap.dll

# It is restricted to 32-bit DLLs, and this might be inconvenient for you to use, but it works on a 64bit Windows. You can't access function exported by ordinal number and you can't directly handle 64bit or greater values/pointers.

class MyDynamicWrapper(PyV8.JSClass) :
    def Register(self, a, b, c, d) :
	print '[*] DynamicWrapper.Register :'
	print '    [-] :', a
	print '    [-] :', b
	print '    [-] :', c
	print '    [-] :', d
    def Space(self, x) :
	return ' ' * x
    def NumGet(self, x, y, z) :
        pass
    def SystemFunction036(self, x, y) :
	pass



class MyWshProcEnv(PyV8.JSClass) :
    def Arch(self, x) :
        if x.lower() == 'processor_architecture':
            if os.name == 'nt' and sys.version_info[:2] < (2,7):
                mc = os.environ.get("PROCESSOR_ARCHITEW6432", 
                       os.environ.get('PROCESSOR_ARCHITECTURE', ''))
            else:
                mc = platform.machine()
            
            machine2bits = {'AMD64': 'x64', 'x86_64': 'x64', 'i386': 'x86', 'x86': 'x86'}
            return machine2bits.get(mc, None)
        elif x.lower() == 'processor_architetew6432':
            """Return type of machine."""
            if os.name == 'nt' and sys.version_info[:2] < (2,7):
                return os.environ.get("PROCESSOR_ARCHITEW6432", 
                       os.environ.get('PROCESSOR_ARCHITECTURE', ''))
            else:
                return platform.machine()
        
        




class MyArgs(PyV8.JSClass) :
    def length(self) :
	    return len(sys.argv)
        
class Global(PyV8.JSClass):
    WScript = MyWScript()
    WshShell = MyWshShell()
    args = MyArgs()

    def GetObject(self, name) :
        return win32com.client.GetObject(name)

if len(sys.argv) != 2 :
    print 'Usage : PyV8DummyAPI.py <.js>'
    exit()
    
s = open(sys.argv[1],'rb').read()
ctx = PyV8.JSContext(Global())
ctx.enter()

ret = ctx.eval(s)
print ctx.locals.keys()
