Option Explicit
'Creates a connection to Windows Shell Object.
Set oShell = createObject("wscript.shell")
 
'Launches Notepad Application
oshell.run("Notepad")
wscript.quit
