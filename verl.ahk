#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Menu, Tray, Icon, tray.ico

date:= "30.11.2021"
password:="z"
DetectHiddenText, On
DetectHiddenWindows, On

^m::
MouseGetPos, xpos, ypos 
MsgBox, Der Zeiger befindet sich auf X%xpos% Y%ypos%.
return

varMitID:="http://lhzbw.gbv.de:9090/lbs4/general/html/de/lbsMain_1024.htm?SCRIPT="
^r::
SetTitleMatchMode 2
MsgBox, 4,  Ist die richtige Nutzernummer ausgewählt?
IfMsgBox, No
{
Return
}
if WinExist("Nutzer.xlsx")
    WinActivate ; 		Verwendet das von WinExist gefundene Fenster.
else Run, Nutzer.xlsx ; 
sleep, 333
send, {Ctrl Down}c{Ctrl Up} ; 	kopiert die Nutzerummer aus der Liste
sleep, 500
id := clipboard ; 		speichert die Nutzernummer  in den Zwischenspeicher
sleep, 500
WinActivate LBS4 ; 		Verwendet das von WinExist gefundene Fenster.
sleep, 500
send, %id% ; 	fügt die Nummer aus der Liste ein
sleep, 500
Send, {shift down} ;
send, {alt down}5{alt up}
send, {shift up} ;
Sleep, 500 ;
Send, {shift down} ;
send, {alt down}
send, n
send, {shift up}
send, {alt up}
send, {delete} ;
Sleep, 500
send, %date% ;
sleep, 500 
send, {tab 4} ;
return
send, {enter}
clipboard :=  ; 		löscht die Zwischenablage
Sleep, 333
return

