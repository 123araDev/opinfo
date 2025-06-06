Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Créer une page noire temporaire
blackPath = fso.GetSpecialFolder(2) & "\black.html"
Set file = fso.CreateTextFile(blackPath, True)
file.WriteLine "<html><head><HTA:APPLICATION WINDOWSTATE='maximize' BORDER='none'/><style>body{background-color:black;margin:0;cursor:none;}</style></head><body></body></html>"
file.Close

' Ouvrir la page noire
shell.Run "mshta.exe """ & blackPath & """", 1, False

WScript.Sleep 1500 ' Attendre 1.5 secondes

' Bip système
shell.SendKeys Chr(7)

WScript.Sleep 1000 ' Pause avant BSOD

' Lancer le faux BSOD (assure-toi que le fichier est dans le même dossier)
currentPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
shell.Run """" & currentPath & "\SYSCONFIG.hta" & """"
