#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

menu, tray, icon, tray.ico
critical, on

kw :="$xD$xL$xR$yVolltext ab"
lf :="$xD$xL$xR$4LF"
kwadd:="kostenfrei verfügbar$4KW"
eromm := "4066 XA-DE$cDE-206"
4950 := "4950 http://nbn-resolving.org/"
2050 := "2050 "

^!f::
SetTitleMatchMode 2
MsgBox, 4, LIZENZFREIER Datensatzes ausgewählt?
IfMsgBox, No
{
Return
}
sleep, 333 ;
if WinExist("urn_lf")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, urn_lf.xlsx  ; 
sleep, 333
send, {Ctrl Down}c{Ctrl Up} ; 	kopiert die PPN aus der Liste u. speichert in Zwischenablage
Sleep, 500
ppn := clipboard ; 		speichert die PPN in den Zwischenspeicher
send, {Ctrl down}5{Ctrl up} ; 	streicht die PPN durch
sleep, 333
send, {right} ; 		Springt auf die URN
sleep, 333
WinActivate WinIBW 3.7 ; 	öffnet die WinIBW
sleep, 500
send, f ppn{Space}%ppn% ; 	sucht die PPN
WinWait, Vollanzeige
sleep, 500
clipboard :=  ; 		löscht die Zwischenablage
sleep, 330
if WinExist("urn_lf")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, urn_lf.xlsx ; 
sleep, 333
Send, {Ctrl Down}c{Ctrl Up} ; 	markiert die urn
Sleep, 333
urnlink := clipboard
sleep, 333
send, {left} ; 			springt an den Anfang der Zeile
sleep, 500, 
send, {down} ;			Springt auf die nächste PPN
sleep, 333 ; 



WinActivate WinIBW 3.7 ; 	öffnet die WinIBW
sleep, 500
Send, {Enter}
sleep, 333
SendInput, !c
sleep, 333
send, k
Sleep, 333
Send, {Enter}
WinWait, Titel ändern ; 	wartet bis der Titel im Modus "bearbeiten" ist.
SendInput, {right 7}
sleep, 333
SendInput, {delete}
sendInput, r ; 			Oaa wird zu Oar 
sleep, 333
SendInput, {Enter} ; 		neue Zeile
Sleep, 333
sendInput, ^f
Winwait, Suchen	 ; 		Wartet, bis das Fenster "Suchen" auftaucht
sendInput, 4950	;		Sucht das Feld 4950
sleep, 500	
sendInput, {Enter}
Sleep, 333
sendInput, +{end}
Sleep, 333
SendInput, {delete}
Sleep, 500
Send, %4950% ; 			Resolver und 4950 wird eingefügt
Sleep, 500
send, %urnlink% ; 		link wird in 4950 ergänzt
Sleep, 333
send, {left 2}
sleep, 333 ;
send, %LF% ;			Lizenz wird eingefügt
sleep, 333
SendInput, {Enter}
sleep, 333 ; 
Send, %2050% ;
sleep, 333
send, %urnlink% ;
SendInput, {Enter}
Sleep, 333
send, %eromm% ; 		4066 wird eingefügt
Sleep, 333
clipboard :=  ; 		löscht die Zwischenablage
return





^!l::
SetTitleMatchMode 2
MsgBox, 4, LIZENZPFLICHTIGEN Datensatz ausgewählt?
IfMsgBox, No
{
Return
}
if WinExist("urn_kw.xls")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, urn_kw.xls ; 
Sleep, 333
SendInput, {Ctrl Down}c{Ctrl Up} ; 	kopiert die PPN 
Sleep, 500
ppn := clipboard ; 		speichert die PPN in den Zwischenspeicher
sleep, 333
send, {Ctrl down}5{Ctrl up} ;   streicht die PPN durch
sleep, 330
Send, {right}
sleep, 500
WinActivate WinIBW 3.7 ; 	öffnet die WinIBW
sleep, 500
send, f ppn{Space}%ppn% ; 	sucht die PPN
WinWait, Vollanzeige
clipboard :=  ; 		löscht die Zwischenablage
sleep, 330
if WinExist("urn_kw.xls")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, urn_kw.xls ; 
Sleep, 333
Send, {Ctrl Down}c{Ctrl Up} ; 	markiert die urn
Sleep, 333
urnlink := clipboard
sleep, 333



WinActivate WinIBW 3.7 ; 	öffnet die WinIBW
sleep, 500
SendInput, {Enter}
sleep, 333
SendInput, !c
sleep, 333
send, k
Sleep, 333
SendInput, {Enter}
WinWait, Titel ändern ; wartet bis der Titel im Modus "bearbeiten" ist.
SendInput, {right 7}
sleep, 333
SendInput, {delete}
sendInput, r ; 			Oaa wird zu Oar 
sleep, 333
SendInput, {Enter} ; 		neue Zeile
Sleep, 333
sendInput, ^f
Winwait, Suchen	 ; 		Wartet, bis das Fenster "Suchen" auftaucht
send, 4950	
sleep, 500	
sendInput, {Enter}
Sleep, 333
sendInput, +{end}
Sleep, 333
SendInput, {delete}
Sleep, 333
send, %2050%
sleep, 333
send, %urnlink% ;		URN wird eingefügt
sleep, 333
Send, %4950% ; 			4950 und Resolver wird eingefügt
Sleep, 333
send, %urnlink% ; 		link wird in 4950 ergänzt
sleep, 333
clipboard :=  ; 		löscht die Zwischenablage
sleep, 333 ; 
sendInput, {left 2}
sleep, 333 ;
if WinExist("urn_kw.xls")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, urn_kw.xls ; 
sleep, 500
send, {right}
sleep, 500 ; 
SendInput, {Ctrl Down}c{Ctrl Up} ; 	kopiert die Lizenzinfos aus der Liste u. speichert in Zwischenablage
sleep, 333
jkw := clipboard ; 
send, {left 2} ; 
sleep, 333
send, {down} ;
WinActivate WinIBW 3.7 ; 	öffnet die WinIBW
WinWait, WinIBW
send, %jkw%  ;
clipboard :=  ; 		löscht die Zwischenablage
sleep, 333
send, %eromm% ; 		4066 wird eingefügt
return








