'----------------------------------------------ADVENT ENGINE----------------------------------------------------------------
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objArgs = Wscript.Arguments
If UCase( Right( WScript.FullName, 12 ) ) <> "\CSCRIPT.EXE" Then
	For Each strArg in objArgs
		startwithargs=startwithargs&" "&strArg
	Next
	WshShell.run("cscript //nologo .\"&WScript.ScriptName&" "&startwithargs)
	Set WshShell = Nothing
	wscript.quit
end if
Dim Sound(3),switchtonew,displaymap,gifdebug,tempchar,playsequences,gip,dlsize,updatever,walk,packlist,lockwalk,playerchar,walkdir,urlcount,tempupurl,debugarg,globalvolume,askedForPermission,donotuse,temphpto,temphp,hp,reloadmap,plhealth,hurtmsg,hurt,EnemyFollow,gotenemy,spawned,debugargA,debugargB,debugargC,debugargD,DEBUGA,triggerloadmap,triggerload,plequip,inventoryshown,walkwaittime,invalidpos,oldx,oldy,exval,htmlver,updatecache,upchoice,ypos,xpos,tempy,tempx,verscheck1,verscheck2,verscheck3,saveepisodic,goblinkey,inventory,maploadcount,plname,plage,plmoney,x,evaluated,mappath,episodic,gamever,mapcache,validinput,uinput,mapinfo,opt,debug,row,foundmapend
'IMPORTANT LINE NUMBERS
'335 - PlaySound
'412 - CheckForUpdates
'501 - Enemy Functions
'1344 - EvaluateMapString
'1376 - Enemy Collision checking
'1560 - Player Collision checking
'1700 - Map Trigger Processing
'1794 - Sequence Map Special Characters
'1871 - LoadMap
CONST bytesToKb = 1024
Const ForReading = 1
Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set xHttp = createobject("Microsoft.XMLHTTP")
Set bStrm = createobject("Adodb.Stream")
Set Sound(0) = CreateObject("WMPlayer.OCX")
Set Sound(1) = CreateObject("WMPlayer.OCX")
Set Sound(2) = CreateObject("WMPlayer.OCX")
Set Sound(3) = CreateObject("WMPlayer.OCX")
Set http = CreateObject("Microsoft.XmlHttp")
if objFSO.FolderExists("./base")<>true then
	msgbox "FATAL ERROR!"&vbCr&"The required directory './base' was not found. Please re-install ADVENT.",016,"Fatal Error - ADVENT"
	wscript.quit
end if
switchtonew=0
gifdebug=0
'gip is gif map in progress
gip=0
playsequences=1
updatever=""
lockwalk=0
playerchar="☺"
globalvolume=40
askedForPermission=0
hurt=0
hurtmsg=""
DEBUGA=0
mappath=""
debugarg=0
debug = 0
triggerload=0
gamever="1.7.5"
Function UserInput( myPrompt )
	If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		' If so, use StdIn and StdOut
		WScript.StdOut.Write myPrompt & " "
		UserInput = WScript.StdIn.ReadLine
	Else
		' If not, use InputBox( )
		UserInput = InputBox( myPrompt )
	End If
End Function
function MessageBoxNew(str,message,clr,prompt)
	if clr=1 then
		For i=0 to 28
			wscript.echo ""
		Next
	end if
	if clr>1 then
		'█▒
		'┐┌─├┴
		'max progress is 73
		templatelength=Len(" ┌────────────────────────────────────────────────────────────────────────────┐")
		progress=""
		for i=0 to clr
			progress=progress&"█"
		Next
		progress=" │ "&progress
		Do Until Len(progress)=(templatelength-2)
			progress=progress&" "
		Loop
		progress=progress&" │"
		wscript.echo "v"&gamever
		wscript.echo " "
		wscript.echo " "
		wscript.echo "		 ████   █████   ██   ██  █████  █████   ██████"
		wscript.echo "		██▒▒██  ██▒▒██  ██▒  ██▒ ██▒▒▒▒ ██▒▒██   ▒██▒▒▒"
		wscript.echo "		██████▒ ██▒ ██▒  ██ ██▒▒ ████   ██▒ ██▒   ██▒"
		wscript.echo "		██▒▒██▒ ██▒ ██▒  ██▒██▒  ██▒▒▒  ██▒ ██▒   ██▒"
		wscript.echo "		██▒ ██▒ █████▒▒   ███▒▒  █████  ██▒ ██▒   ██▒"
		wscript.echo "		 ▒▒  ▒▒  ▒▒▒▒▒     ▒▒▒    ▒▒▒▒▒  ▒▒  ▒▒    ▒▒"
		wscript.echo "							UPDATER"
		wscript.echo " "
		wscript.echo " ┌──────────────────┐"
		wscript.echo " │ Overall Progress │"
		wscript.echo " ├──────────────────┴─────────────────────────────────────────────────────────┐"
		wscript.echo progress
		wscript.echo " └────────────────────────────────────────────────────────────────────────────┘"
		wscript.echo " "
		wscript.echo " "
		wscript.echo " "
		wscript.echo " "
		wscript.echo " "
	end if
	wscript.echo ""
	if str="NONE" or clr>1 then
		itemtoecho="╔═════════════════════════════════════════════════════════════════════════════╗"
	else
		itemtoecho="╔════ "&str&" ════════════════════════════════════════════════════════════════╗"
	end if
	wscript.echo itemtoecho
	msgtoecho="║  "&message
	Do While Len(msgtoecho)+1<>Len(itemtoecho)
		msgtoecho=msgtoecho&" "
	Loop
	msgtoecho=msgtoecho&"║"
	wscript.echo msgtoecho
	wscript.stdout.write("╚")
	For i = 0 to Len(itemtoecho)-3
	wscript.stdout.write("═")
	Next
	wscript.stdout.write("╝")
	if prompt=1 and clr<2 then
		WScript.StdIn.ReadLine
	end if
end Function
Function Exec(path)
	exval=""
	if objFso.FileExists(path) and inStr(path,".asc")<>0 then
		Set objExec = objFSO.OpenTextFile(path, ForReading)
		Do While objExec.AtEndOfStream = False
			strLine = objExec.ReadLine
			if exval="" then
				exval=strLine
			else
				exval=exval&" : "&strLine
			end if
		Loop
		objExec.close
		On Error Resume Next
		execute(exval)
		If Err.Number <> 0 Then
			msgbox "Script:	"&path&vbCr&"Function:	Exec()"&vbCr&"Error:	"&Err.Description&vbCr&"Number:	"&Err.Number&vbCr&"Source:	"&Err.Source&vbCr&""&vbCr&"ADVENT will not run this script.",016,"Caught Error - Advent"
			Err.Clear
		End If
	else
		if objFSO.FileExists(path)=0 then
			wscript.echo "Exec() :: File '"&path&"' does not exist."
		else
			wscript.echo "Exec() :: The Exec function can only execute valid '.asc' files ("&path&")."
		end if
	end if
end Function
function WriteConfigFile(def)
	Set objFileCFG=objFSO.CreateTextFile("./config.ini",True)
	if def=0 then
		objFileCFG.Write "[GameStartup]"&vbCrLf
		objFileCFG.Write "checkForUpdates="&checkForUpdates&vbCrLf
		objFileCFG.Write "askedForPermission="&askedForPermission&vbCrLf
		objFileCFG.Write "[Saving]"&vbCrLf
		objFileCFG.Write "savefilePrefix="&savefilePrefix&vbCrLf
		objFileCFG.Write "maximumSaveFiles="&maximumSaveFiles&vbCrLf
		objFileCFG.Write "[Audio]"&vbCrLf
		objFileCFG.Write "globalvolume="&globalvolume&vbCrLf
		objFileCFG.Write "[Gameplay]"&vbCrLf
		if playerchar="☺" then
			objFileCFG.Write "playerchar=1"&vbCrLf
		else
			objFileCFG.Write "playerchar="&playerchar&vbCrLf
		end if
		objFileCFG.Write "[Misc]"&vbCrLf
		objFileCFG.Write "playsequences="&playsequences
	else
		objFileCFG.Write "[GameStartup]"&vbCrLf
		objFileCFG.Write "checkForUpdates=1"&vbCrLf
		objFileCFG.Write "askedForPermission=0"&vbCrLf
		objFileCFG.Write "[Saving]"&vbCrLf
		objFileCFG.Write "savefilePrefix=savegame"&vbCrLf
		objFileCFG.Write "maximumSaveFiles=10"&vbCrLf
		objFileCFG.Write "[Audio]"&vbCrLf
		objFileCFG.Write "globalvolume=40"&vbCrLf
		objFileCFG.Write "[Gameplay]"&vbCrLf
		objFileCFG.Write "playerchar=1"&vbCrLf
		objFileCFG.Write "[Misc]"&vbCrLf
		objFileCFG.Write "playsequences=1"
	end if
	objFileCFG.close
end function
'Below declares the configuration file variables
Dim checkForUpdates,savefilePrefix,maximumSaveFiles
if objFSO.FileExists("./config.ini")<>true then
	WriteConfigFile(1)
	savefilePrefix="savegame"
	maximumSaveFiles=10
	checkForUpdates=1
	askedForPermission=0
	globalvolume=40
	playerchar="☺"
else
	Set objFile = objFSO.OpenTextFile("./config.ini", ForReading)
	row=0
	Do While objFile.AtEndOfStream = False
		strLine = objFile.ReadLine
		row=row+1
		if inStr(strLine,"checkForUpdates=")=1 then
			if NOT isNumeric(mid(strLine,inStr(strLine,"checkForUpdates=")+Len("checkForUpdates="),Len(strLine)-Len("checkForUpdates="))) then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'checkForUpdates' IS NOT A BOOL)"
				wscript.quit
			end if
			checkForUpdates=mid(strLine,inStr(strLine,"checkForUpdates=")+Len("checkForUpdates="),1)
		end if
		if inStr(strLine,"savefilePrefix=")=1 then
			savefilePrefix=mid(strLine,inStr(strLine,"savefilePrefix=")+Len("savefilePrefix="),Len(strLine)-Len("savefilePrefix="))
		end if
		if inStr(strLine,"maximumSaveFiles=")=1 then
			if NOT isNumeric(mid(strLine,inStr(strLine,"maximumSaveFiles=")+Len("maximumSaveFiles="),Len(strLine)-Len("maximumSaveFiles="))) then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'checkForUpdates' IS NOT AN INT)"
				wscript.quit
			end if
			maximumSaveFiles=mid(strLine,inStr(strLine,"maximumSaveFiles=")+Len("maximumSaveFiles="),Len(strLine)-Len("maximumSaveFiles="))
		end if
		if inStr(strLine,"askedForPermission=")=1 then
			if NOT isNumeric(mid(strLine,inStr(strLine,"askedForPermission=")+Len("askedForPermission="),Len(strLine)-Len("askedForPermission="))) then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'askedForPermission' IS NOT A BOOL)"
				wscript.quit
			end if
			askedForPermission=mid(strLine,inStr(strLine,"askedForPermission=")+Len("askedForPermission="),Len(strLine)-Len("askedForPermission="))
		end if
		if inStr(strLine,"globalvolume=")=1 then
			if NOT isNumeric(mid(strLine,inStr(strLine,"globalvolume=")+Len("globalvolume="),1)) then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'globalvolume' IS NOT AN INT)"
				wscript.quit
			end if
			globalvolume=mid(strLine,inStr(strLine,"globalvolume=")+Len("globalvolume="),Len(strLine)-Len("globalvolume="))
		end if
		if inStr(strLine,"playerchar=")=1 then
			if Len(mid(strLine,inStr(strLine,"playerchar=")+Len("playerchar="),Len(strLine)-Len("playerchar=")))<>1 then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'playerchar' LENGTH IS TOO LONG (MAX 1))"
				wscript.quit
			end if
			playerchar=mid(strLine,inStr(strLine,"playerchar=")+Len("playerchar="),Len(strLine)-Len("playerchar="))
			if playerchar="1" then
				playerchar="☺"
			end if
		end if
		if inStr(strLine,"playsequences=")=1 then
			if NOT isNumeric(mid(strLine,inStr(strLine,"playsequences=")+Len("playsequences="),Len(strLine)-Len("playsequences="))) then
				msgbox "There is a problem with your ADVENT CFG file. (VAR 'playsequences' IS NOT A BOOL)"
				wscript.quit
			end if
			playsequences=mid(strLine,inStr(strLine,"playsequences=")+Len("playsequences="),Len(strLine)-Len("playsequences="))
		end if
	Loop
	objFile.close
end if
Sub pauseScript()
      Dim strMessage, Input
      Wscript.StdOut.Write strMessage
      WScript.Echo "Press ENTER to continue."

      Do While Not WScript.StdIn.AtEndOfLine
            Input = WScript.StdIn.Read(1)
      Loop
End Sub
function getHTML(url,plain)
	On Error Resume Next
	http.open "GET", URL, False
	http.send ""
	if err.Number = 0 Then
		if plain=1 then
			getHTML = ConvertHTML2PlainText(http.responseText)
		end if
		if plain=0 then
			getHTML = http.responseText
		end if
		if plain=2 then
			getHTML = Len(http.responseText)
		end if
	Else
		if DEBUGA=1 then
			Wscript.Echo "INTERNET CALL ERROR (" & Err.Number & "): " & Err.Description
		end if
		getHTML = "0"
	End If
end function
function DownloadBinary(url,saveto)
	xHttp.Open "GET", url, False
	xHttp.Send
	with bStrm
		.type = 1
		.open
		.write xHttp.responseBody
		.savetofile saveto, 2
	end with
	bStrm.close
end function
Function ConvertHTML2PlainText(ByVal sText)
    Dim oRegEx
    Set oRegEx = New RegExp
        oRegEx.Pattern = "</?.+?/?>"
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
 
    sText = oRegEx.Replace(sText, "")
 
    oRegEx.Pattern = "&gt;"
    sText = oRegEx.Replace(sText, ">")
 
    oRegEx.Pattern = "&lt;"
    sText = oRegEx.Replace(sText, "<")
 
    oRegEx.Pattern = "&quot;"
    sText = oRegEx.Replace(sText, """")
 
    ConvertHTML2PlainText = sText
End Function
evaluated=0
Sub StopSound(ch)
	Sound(ch).Controls.Stop
End Sub
Sub PlaySound(SoundFile,ch)
	if objFSO.FileExists(SoundFile)=False then
		Exit Sub
	end if
	if globalvolume=0 then
		exit sub
	else
		Sound(ch).URL = SoundFile
		Sound(ch).settings.volume = globalvolume
		Sound(ch).Controls.play
		'do while Sound(ch).currentmedia.duration = 0
		'	wscript.sleep 100
		'loop
		'wscript.sleep int(Sound.currentmedia.duration)+1)*1000
	end if
End Sub
function ShowLoadProgress(name,val,max)
	wscript.echo ""
	prog=""
	for i=1 to val
		prog=prog&"█ "
	Next
	For i=1 to (max-val)
		prog=prog&"∙ "
	Next
	ShowLoadProgress=name&": "&prog
end Function
validinput=0
'-----------------------------------------------START GAME FUNCTIONS---------------------------------------------
'-----------------------------------------------PERMISSION FOR FILES---------------------------------------------
if askedForPermission=0 then
	Do Until lCase(debugargA)="y" or lCase(debugargA)="n"
		wscript.echo " "
		wscript.echo " "
		wscript.echo " "
		wscript.echo "  ┌──────────────────────────────────────────────────────────────────────────┐"
		wscript.echo "  │ ADVENT v"&gamever&" - Permission Check                                         │"
		wscript.echo " ╓┴──────────────────────────────────────────────────────────────────────────┴╖"
		wscript.echo " ║                                                                            ║"
		wscript.echo " ║   ADVENT needs your permission to do the following things:                 ║"
		wscript.echo " ║      ► Modify files in this directory (Including Subfolders)               ║"
		wscript.echo " ║      ► Connect to the internet                                             ║"
		wscript.echo " ║      ► Play audio                                                          ║"
		wscript.echo " ║                                                                            ║"
		wscript.echo " ║                                                                            ║"
		wscript.echo " ║   In order to run Advent, these conditions must be allowed. We will only   ║"
		wscript.echo " ║   attempt actions that the user has permitted.                             ║"
		wscript.echo " ║                                                                            ║"
		wscript.echo " ║   NOTE: If you agree with these permissions, this message will not         ║"
		wscript.echo " ║   appear on the next game startup.                                         ║"
		wscript.echo " ║                                                                            ║"
		wscript.echo " ╚════════════════════════════════════════════════════════════════════════════╝"
		wscript.echo " "
		wscript.echo " "
		wscript.echo " "
		debugargA=UserInput("Do you agree with these permissions (y/n)?>")
	Loop
	if lCase(debugargA)="y" then
		askedForPermission=1
		WriteConfigFile(0)
	else
		wscript.quit
	end if
end if
if objFSO.FolderExists("./addons/")=0 then
	objFSO.CreateFolder("./addons/")
end if
if objFSO.FolderExists("./sounds/")=0 then
	objFSO.CreateFolder("./sounds/")
end if
if objFSO.FileExists("./config.cfg")=true then
	objFSO.DeleteFile("./config.cfg")
end if
'-------------------------------------------------------UPDATE---------------------------------------------------
Dim updatedesc
verscheck1="https://raw.githubusercontent.com/ClayHanson/advent/master/update"
verscheck2="https://raw.githubusercontent.com/ClayHanson/advent/master/updatebackup"
Function CheckForGameUpdate(url)
	htmlver=getHTML(url,1)
	if htmlver<>"0" then
		htmlstuff=htmlver
		htmlver=left(htmlver,inStr(htmlver,"}")-1)
		updatever=htmlver
		if htmlver<>gamever then
			if inStr(htmlstuff,"(")=0 or inStr(htmlstuff,")")=0 then 
				if url=verscheck2 then
					wscript.echo "Error while getting current game version."
					UserInput("Press ENTER to continue.")
					exit function
				end if
				Do Until lCase(choice)="y" or lCase(choice)="n"
					wscript.echo "Error while getting current game version."
					choice=UserInput("Try second update link (y/n)?>")
				Loop
				if lCase(choice)="y" then
					CheckForGameUpdate(verscheck2)
				end if
				Exit Function
			end if
			updatedesc=Mid(htmlstuff,InStr(htmlstuff,"(") + Len("("),inStr(htmlstuff,")")-(InStr(htmlstuff,"(") + Len("(")))
			updatedesc=Replace(updatedesc,"NL","""&vbCr&""")
			'wscript.echo "CHANGELOG:"
			execute("upmsg=""There is a new version of ADVENT (v"&htmlver&"). ""&vbCr&""Changelog:""&vbCr&"""&updatedesc&"""&vbCr&""Do you want to install this update?""")
			upchoice=msgbox(upmsg,3,"Update Available - Advent")
			if upchoice=6 then
				MessageBoxNew "NONE","Preparing for update...",2,0
				intStart=InStr(htmlstuff,"[")
				intStart=intStart + Len("[")
				intEnd=inStr(htmlstuff,"]")
				downloadurl=Mid(htmlstuff,intStart,intEnd-intStart)
				execute("dlsize=getHTML("""&downloadurl&""",2)")
				urlcount=Mid(htmlstuff,InStr(htmlstuff,"%count") + Len("%count"),inStr(htmlstuff,"/count%")-(InStr(htmlstuff,"%count") + Len("%count")))
				MessageBoxNew "NONE",ShowLoadProgress("Downloading file",1,urlcount+1)&" (ADVENT.VBS - "&CInt(dlsize/bytesToKb)&" Kb)",10,0
				dlfilename=objFSO.GetFileName(downloadurl)
				DownloadBinary downloadurl, "./Advent.vbs"
				if inStr(htmlstuff,"%")<>0 then
					g=73-urlcount-2
					for i=1 to urlcount
						g=g+1
						tempupurl=Mid(htmlstuff,InStr(htmlstuff,"@"&i&"@") + Len("@"&i&"@"),inStr(htmlstuff,"@/"&i&"@")-(InStr(htmlstuff,"@"&i&"@") + Len("@"&i&"@")))
						temppath=Mid(htmlstuff,InStr(htmlstuff,"@path"&i&"@") + Len("@path"&i&"@"),inStr(htmlstuff,"@/path"&i&"@")-(InStr(htmlstuff,"@path"&i&"@") + Len("@path"&i&"@")))
						tocheck=Replace(temppath,objFSO.GetFileName(temppath),"")
						if objFSO.FolderExists(tocheck)=False then
							objFSO.CreateFolder(tocheck)
						end if
						execute("dlsize=getHTML("""&tempupurl&""",2)")
						if CInt(dlsize/bytesToKb)<>0 then
							size=CInt(dlsize/bytesToKb)&" Kb"
						else
							size=dlsize&" bytes"
						end if
						showthistext=ShowLoadProgress("Downloading file",i+1,urlcount+1)&" ("&UCase(objFSO.GetFileName(temppath))&" - "&size&")"
						if Len(showthistext)>73 then
							showthistext=ShowLoadProgress("Downloading file",i+1,urlcount+1)
						end if
						MessageBoxNew "NONE",showthistext,g,0
						if objFSO.FileExists(temppath) then
							objFSO.DeleteFile(temppath)
						end if
						DownloadBinary tempupurl,temppath
					Next
				end if
				MessageBoxNew "NONE","Deleting CONFIG.INI",72,0
				if objFSO.FileExists("./config.ini") then
					objFSO.DeleteFile("./config.ini")
				end if
				MessageBoxNew "NONE","Restarting ADVENT...",73,0
				wscript.sleep 500
				For Each strArg in objArgs
					startwithargs=startwithargs&" "&strArg
				Next
				WshShell.run("cscript .\"&WScript.ScriptName&" "&startwithargs)
				wscript.quit
				Exit Function
			end if
			Exit Function
		end if
	else
		Exit Function
	end if
end Function
if checkForUpdates=1 then
	wscript.echo "Checking game version..."
	CheckForGameUpdate verscheck1
end if
'----------------------------------------------------ENEMY FUNCTIONS---------------------------------------------
Dim npcenemy(10),enemycount
enemycount = 0
Function UpdateEnemyPosition(name,eypos,expos)
	if enemycount <> 0 then
		for i = 0 to enemycount
			if inStr(npcenemy(i),"[") <> 0 and inStr(npcenemy(i),"]") <> 0 then
				tempenemyname=Mid(npcenemy(i),inStr(npcenemy(i),"[")+1,inStr(npcenemy(i),"]")-(inStr(npcenemy(i),"[")+1))
				if lCase(tempenemyname) = lCase(name) then
					repxpos=Mid(npcenemy(i),inStr(npcenemy(i),"_x")+2,2)
					repypos=Mid(npcenemy(i),inStr(npcenemy(i),"_y")+2,2)
					'msgbox npcenemy(i)
					if Len(eypos) < 2 then
						Do until Len(eypos)=2
							eypos="0"&eypos
						Loop
					end if
					if Len(expos) < 2 then
						Do until Len(expos)=2
							expos="0"&expos
						Loop
					end if
					for d = 0 to enemycount-1
						'msgbox npcenemy(d)
						nam=Mid(npcenemy(d),inStr(npcenemy(d),"[")+1,inStr(npcenemy(d),"]")-(inStr(npcenemy(d),"[")+1))
						if inStr(npcenemy(d),"_x"&expos) <> 0 and nam <> tempenemyname then
							expos=expos-1
						end if
						if inStr(npcenemy(d),"_y"&eypos) <> 0 and nam <> tempenemyname then
							eypos=eypos-1
						end if
					next
					if eypos < 0 then
						eypos=0
					end if
					if expos < 0 then
						expos=0
					end if
					if Len(eypos) < 2 then
						Do until Len(eypos)=2
							eypos="0"&eypos
						Loop
					end if
					if Len(expos) < 2 then
						Do until Len(expos)=2
							expos="0"&expos
						Loop
					end if
					npcenemy(i)=Replace(npcenemy(i),"_x"&repxpos,"_x"&expos)
					npcenemy(i)=Replace(npcenemy(i),"_y"&repypos,"_y"&eypos)
					Exit Function
				end if
			end if
		next
	end if
end Function
Function GetEnemyName(str)
	GetEnemyName=Mid(str,inStr(str,"[")+1,inStr(str,"]")-(inStr(str,"[")+1))
end Function
Function AddMapEnemy(name,eypos,expos,hp)
	if enemycount >= 10 then
		msgbox "Max map enemy count reached."
		Exit Function
	end if
	if enemycount <> 0 then
		for i = 0 to enemycount
			if inStr(npcenemy(i),"[") <> 0 and inStr(npcenemy(i),"]") <> 0 then
				tempenemyname=Mid(npcenemy(i),inStr(npcenemy(i),"[")+1,inStr(npcenemy(i),"]")-(inStr(npcenemy(i),"[")+1))
				if lCase(tempenemyname) = lCase(name) then
					msgbox "Enemy '"&name&"' already defined. Skipping."
					Exit Function
				end if
			end if
		next
	end if
	if Len(eypos) < 2 then
		Do until Len(eypos)=2
			eypos="0"&eypos
		Loop
	end if
	if Len(expos) < 2 then
		Do until Len(expos)=2
			expos="0"&expos
		Loop
	end if
	if Len(hp) < 3 then
		Do until Len(hp)=3
			hp="0"&hp
		Loop
	end if
	npcenemy(enemycount) = "["&name&"]_x"&expos&"_y"&eypos&"_HP"&hp
	'msgbox "done:"&vbCr&npcenemy(enemycount)
	enemycount=enemycount+1
End Function
'----------------------------------------------------ENEMY AI-------------------------------------------------------
Function CalcEnemyPos()
	if enemycount <> 0 then
		for i = 0 to enemycount
			if inStr(npcenemy(i),"[") <> 0 and inStr(npcenemy(i),"]") <> 0 then
				encxpos=Mid(npcenemy(i),inStr(npcenemy(i),"_x")+2,2)
				encypos=Mid(npcenemy(i),inStr(npcenemy(i),"_y")+2,2)
				tempkeepx = encxpos
				tempkeepy = encypos
				tempchangex = encxpos
				tempchangey = encypos
				notchanged=1
				if CInt(tempkeepx) < CInt(xpos) and notchanged=1 then
					tempchangex=CInt(tempchangex+1)
					'notchanged=0
				end if
				if CInt(tempkeepx) > CInt(xpos) and notchanged=1 then
					tempchangex=CInt(tempchangex-1)
					'notchanged=0
				end if
				if CInt(tempkeepy) < CInt(ypos) and notchanged=1 then
					tempchangey=CInt(tempchangey+1)
					'notchanged=0
				end if
				if CInt(tempkeepy) > CInt(ypos) and notchanged=1 then
					tempchangey=CInt(tempchangey-1)
					'notchanged=0
				end if
				UpdateEnemyPosition Mid(npcenemy(i),inStr(npcenemy(i),"[")+1,inStr(npcenemy(i),"]")-(inStr(npcenemy(i),"[")+1)),tempchangey,tempchangex
			end if
		next
	end if
end Function
'----------------------------------------------------GIF-TYPE STUFF-------------------------------------------------
function PlaySequence(name,frames,delay,loadafter)
	if playsequences=0 then
		Exit Function
	end if
	if objFSO.FileExists(name&"cfg.txt")=False then
		msgbox "ERROR"&vbCr&"PlaySequence: Required file '"&name&"cfg.txt' does not exist. Cannot play sequence.",016,"Error - Advent"
		Exit Function
	end if
	if frames<0 then
		if gifdebug=1 then
			msgbox "PlaySequence :: Frame count not supplied; will get from gif cfg."
		end if
		Set objFileGCFG = objFSO.OpenTextFile(name&"cfg.txt",ForReading)
		Do While objFileGCFG.AtEndOfStream = False
			strLine=lCase(objFileGCFG.readLine)
			if inStr(strLine,"length=")=1 then
				frames=Replace(strLine,"length=","")
			end if
		Loop
		objFileGCFG.close
	end if
	tempmappath=mappath
	gip=1
	tempchar=playerchar
	playerchar=""
	for i=0 to frames
		if objFSO.FileExists(name&i&".txt")=False then
			playerchar=tempchar
			gip=0
			msgbox "ERROR"&vbCr&"PlaySequence: Frame "&i&" of map '"&name&"' does not exist.",016,"Error - Advent"
			Exit Function
		end if
	Next
	for i=0 to frames
		LoadGifMap name&i
		wscript.sleep delay
	Next
	playerchar=tempchar
	gip=0
	if loadafter<>"0" then
		LoadMap loadafter,0
	end if
End Function
function LoadGifMap(path)
	if inStr(path,".txt")<>0 then
		path=Replace(path,".txt","")
	end if
	displaymap=""
	if objFSO.FileExists(path&".txt")=False then
		msgbox "LoadGifMap() :: GifMap '"&path&".txt' does not exist."
		Exit Function
	end if
	Set objFileGIFMap = objFSO.OpenTextFile(path&".txt", ForReading)
	Do While objFileGIFMap.AtEndOfStream = False
		strLine = objFileGIFMap.readLine
		if inStr(strLine,"///")<>1 then
			EvaluateGifString(mid(strLine,4,len(strLine)))
		end if
	Loop
	wscript.echo mid(displaymap,1,Len(displaymap)-1)
	objFileGIFMap.close
end Function
function EvaluateGifString(str)
	i=0
	If Len(str)<30 then
		Do Until Len(str)=30
			str=str&" "
		Loop
	End If
	Do Until Len(str)=i
		i=i+1
		specialchar=0
		if mid(str,i,1)="4" then
			specialchar=1
			MapCacheAdd("█")
		end if
		if mid(str,i,1)="3" then
			specialchar=1
			MapCacheAdd("▓")
		end if
		if mid(str,i,1)="2" then
			specialchar=1
			MapCacheAdd("▒")
		end if
		if mid(str,i,1)="1" then
			specialchar=1
			MapCacheAdd("▒")
		end if
		if mid(str,i,1)="+" then
			specialchar=1
			MapCacheAdd("∙")
		end if
		if mid(str,i,1)="P" then
			specialchar=1
			MapCacheAdd(tempchar)
		end if
		if mid(str,i,1)="N" then
			specialchar=1
			MapCacheAdd("☻")
		end if
		if specialchar=0 then
			MapCacheAdd(mid(str,i,1))
		end if
	Loop
	MapCacheAdd("NL")
end function
'----------------------------------------------------REST OF STUFF--------------------------------------------------
function MapCacheAdd(str)
	if str="NL" then
		displaymap=displaymap&vbCrLf
	else
		displaymap=displaymap&str
	end if
end Function
function HasItem(str)
	toadd=lcase(str)
	toadd=mid(toadd,2,len(str))
	toadd=ucase(mid(str,1,1))&toadd
	if inStr(toadd,"Vfs")<>0 then
		toadd=Replace(toadd,"Vfs","VFS")
	end if
	if inStr(inventory,"{"&toadd&";")<>0 then
		HasItem = True
	else
		HasItem = False
	end if
end Function
function PayForItem(str,desc,showtext,price)
	if plmoney>=price then
		if HasItem(str)=False then
			plmoney=plmoney-price
			AddItem str,desc,showtext
		end if
	else
		hurt=1
		hurtmsg="You don't have enough money to purchase this item. ($"&price&")"
	end if
end Function
function AddItem(str,desc,showtext)
	toadd=lcase(str)
	toadd=mid(toadd,2,len(str))
	toadd=ucase(mid(str,1,1))&toadd
	if inStr(toadd,"Vfs")<>0 then
		toadd=Replace(toadd,"Vfs","VFS")
	end if
	if HasItem(toadd)=False then
		inventory=inventory&"{"&toadd&";"&desc&toadd&"}"
		if inventoryshown="" then
			inventoryshown=toadd
		else
			if inStr(str," ")<>0 then
				inventoryshown=inventoryshown&","&Replace(toadd," ","_")
			else
				inventoryshown=inventoryshown&","&toadd
			end if
		end if
		if showtext=1 then
			wscript.echo ""
			itemtoecho="╔═════════════════════════ You got the "&toadd&"! ═════════════════════════╗"
			'itemtoecho="+============================You got the "&toadd&"!============================+"
			wscript.echo itemtoecho
			desctoecho="║  "&desc
			Do While Len(desctoecho)+1<>Len(itemtoecho)
				desctoecho=desctoecho&" "
			Loop
			desctoecho=desctoecho&"║"
			wscript.echo desctoecho
			wscript.stdout.write("╚")
			For i = 0 to Len(itemtoecho)-3
				wscript.stdout.write("═")
			Next
			wscript.stdout.write("╝")
			wscript.echo ""
			userinput("Press ENTER to continue.")
			ypos=ypos+1
		end if
	end if
end function
function RemoveItem(str,story)
	toadd=lcase(str)
	toadd=mid(toadd,2,len(str))
	toadd=ucase(mid(str,1,1))&toadd
	if inStr(toadd,"Vfs")<>0 then
		toadd=Replace(toadd,"Vfs","VFS")
	end if
	if HasItem(toadd)=True then
		intStart=InStr(lcase(inventory),"{"&lcase(toadd))
		intStart=intStart + Len("{"&lcase(toadd))+1
		intEnd=inStr(lcase(inventory),lcase(toadd)&"}")
		desc=Mid(inventory,intStart,intEnd-intStart)
		if inStr(desc,"Story")<>0 and story=0 then
			wscript.echo "You cannot discard story items."
			UserInput("Press ENTER to continue.")
			LoadMap mappath,0
		else
			inventory=replace(inventory,"{"&toadd&";"&desc&toadd&"}","")
			if inStr(inventoryshown,","&toadd)<>0 then
				inventoryshown=Replace(inventoryshown,","&toadd,"")
			else
				inventoryshown=Replace(inventoryshown,toadd,"")
			end if
		end if
	end if
end function
Function SaveGame()
	slot="-1"
	Do until CInt(slot)<maximumSaveFiles and CInt(slot)>0
		slot=UserInput("Slot (1 - "&maximumSaveFiles&"):")
	Loop
	if CInt(slot)<=10 then
		slot="0"&slot
	end if
	Set objFile = objFSO.CreateTextFile(".\base\"&savefilePrefix&slot&".avs",True)
	'FORMAT:
	'Line 1: PL NAME, Line 2: PL MONEY, Line 3: PL AGE, Line 4: MAP, Line 5: XPOS, Line 6: YPOS
	objFile.Write strReverse(plname)&vbCrLf
	objFile.Write strReverse(plmoney)&vbCrLf
	objFile.Write strReverse(plage)&vbCrLf
	objFile.Write strReverse(mappath)&vbCrLf
	objFile.Write strReverse(inventory)&vbCrLf
	objFile.Write strReverse(xpos)&vbCrLf
	objFile.Write strReverse(ypos)&vbCrLf
	objFile.Write strReverse(plequip)&vbCrLf
	objFile.Write strReverse(inventoryshown)&vbCrLf
	objFile.Write strReverse(plhealth)&vbCrLf
	for j = 0 to 10
		if inStr(npcenemy(j),"[") <> 0 and inStr(npcenemy(j),"]") <> 0 and inStr(npcenemy(j),"_HP") <> 0 and inStr(npcenemy(j),"_x") <> 0 and inStr(npcenemy(j),"_y") <> 0 then
			objFile.Write strReverse(npcenemy(j))&vbCrLf
		else
			'msgbox npcenemy(j)
			objFile.Write " "&vbCrLf
		end if
	Next
	objFile.close
	if objFSO.FileExists(".\base\"&savefilePrefix&slot&".avs") then
		wscript.echo "Game Saved."
	else
		wscript.echo "Failed to save game."
	end if
End Function
Function evaluateInput(uin)
	if uin="" then
		evaluated=1
	end if
	if plhealth <= 0 and plname <> "" then
		wscript.echo "You were slain."
		UserInput("Press ENTER to continue.")
		wscript.quit
	end if
	if (uin="walk" or uin="run" or uin="teleport") and evaluated=0 then
		if uin="run" then
			walkwaittime=50
		end if
		if uin="walk" then
			walkwaittime=100
		end if
		if uin="teleport" then
			if HasItem("teleporter") then
				walkwaittime=0
			else
				walkwaittime=100
				msgbox "You do not have a Teleporter.",0,"Cannot teleport!"
				uin="walk"
			end if
		end if
		Do
			reloadmap=0
			walk=UserInput("Direction (w=Forward, a=Left, d=Right, s=Down, Stop Moving=done):")
			walk=lcase(walk)
			if walk="done" or walk="back" or walk="exit" then
				Exit Do
			end if
			For i = 1 to Len(walk)
				if uin="teleport" then
					EnemyFollow=2
				end if
				if walk="wait" then
					LoadMap mappath,0
				end if
				if mid(walk,i,1)="w" and walk <> "wait" then
					ypos=ypos-1
					if lockwalk=0 then
						walkdir="up"
					end if
					LoadMap mappath,0
				end if
				if mid(walk,i,1)="a" and walk <> "wait" then
					xpos=xpos-1
					if lockwalk=0 then
						walkdir="left"
					end if
					LoadMap mappath,0
				end if
				if mid(walk,i,1)="s" and walk <> "wait" then
					ypos=ypos+1
					if lockwalk=0 then
						walkdir="down"
					end if
					LoadMap mappath,0
				end if
				if mid(walk,i,1)="d" and walk <> "wait" then
					xpos=xpos+1
					if lockwalk=0 then
						walkdir="right"
					end if
					LoadMap mappath,0
				end if
				oldx=xpos
				oldy=ypos
				'ypos=ypos-1
				'LoadMap mappath,0
				wscript.sleep walkwaittime
				if uin="teleport" then
					EnemyFollow=0
				end if
			Next
			'LoadMap mappath,0
		Loop
	end if
	if uin="lockwalk" and evaluated=0 then
		if lockwalk=0 then
			lockwalk=1
			wscript.echo "Locked weapon direction."
		else
			lockwalk=0
			wscript.echo "Unlocked weapon direction."
		end if
	end if
	if uin="stats" and evaluated=0 then
		wscript.echo " "
		wscript.echo "PLAYER NAME: "&plname
		wscript.echo "AGE: "&plage
		wscript.echo "MONEY: $"&plmoney
		wscript.echo "HP: "&plhealth
		wscript.echo " "
	end if
	if uin="inventory" and evaluated=0 then
		'msgbox inventory
		wscript.echo " "
		wscript.echo "INVENTORY:"
		inventoryecho=inventory
		'Do until (inStr(inventoryecho,";") and inStr(inventoryecho,"}"))=0
		'	inventoryecho=Replace(inventoryecho,mid(inventoryecho,inStr(inventoryecho,";"),inStr(inventoryecho,"}")),"")
		'Loop
		'inventoryecho=Replace(inventoryecho,"{","")
		choice=inventoryshown
		if inStr(choice,"_") then
			choice=Replace(inventoryshown,"_"," ")
		end if
		if inStr(lcase(choice),lcase(plequip))<>0 and plequip<>"" then
			wscript.echo Replace(choice,plequip,plequip&" (EQUIPPED)")
		else
			wscript.echo choice
		end if
		wscript.echo " "
	end if
	if inStr(uin,"examine ")=1 and evaluated=0 then
		itemtosearchfor=lcase(Mid(uin,inStr(uin," ")+1,Len(uin)-inStr(uin," ")))
		toadd=mid(itemtosearchfor,2,len(itemtosearchfor))
		itemtosearchfor=ucase(mid(itemtosearchfor,1,1))&toadd
		if HasItem(itemtosearchfor) then
			intStart=InStr(inventory,"{"&itemtosearchfor)
			intStart=intStart + Len("{"&itemtosearchfor)+1
			intEnd=inStr(inventory,itemtosearchfor&"}")
			desc=Mid(inventory,intStart,intEnd-intStart)
			if inStr(desc,"DP") <> 0 then
				tempdp=Mid(desc,inStr(desc,"DP")-3,3)
				execute("donotuse=CInt("&tempdp&")")
				desc=replace(desc,tempdp&"DP",donotuse&" DP (Damage Points)")
			end if
			if inStr(desc,"HP") <> 0 then
				temphp=Mid(desc,inStr(desc,"HP")-3,3)
				execute("donotuse=CInt("&temphp&")")
				desc=replace(desc,temphp&"HP",donotuse&" HP (Health Points)")
			end if
			msgbox desc,0,"Examine "&itemtosearchfor&" - Advent"
		end if
	end if
	if (inStr(uin,"discard ")=1 or inStr(uin,"drop ")=1) and evaluated=0 then
		itemtosearchfor=Mid(uin,inStr(uin," ")+1,Len(uin)-inStr(uin," "))
		if HasItem(itemtosearchfor) then
			Do Until lcase(choice)="n" or lcase(choice)="y"
				choice=UserInput("Are you sure you want to discard '"&UCase(itemtosearchfor)&"' (y/n)?>")
			Loop
			if LCase(choice)="y" then
				RemoveItem itemtosearchfor,0
			end if
		end if
	end if
	if inStr(uin,"use ")=1 and evaluated=0 then
		itemtosearchfor=lcase(Mid(uin,inStr(uin," ")+1,Len(uin)-inStr(uin," ")))
		if HasItem(itemtosearchfor)=True then
			intStart=InStr(lcase(inventory),"{"&itemtosearchfor)
			intStart=intStart + Len("{"&itemtosearchfor)+1
			intEnd=inStr(lcase(inventory),itemtosearchfor&"}")
			desc=Mid(inventory,intStart,intEnd-intStart)
			gotaction=0
			if inStr(lcase(desc),"hp")<>0 and gotaction=0 then
				gotaction=1
				tempaddhp=Mid(desc,inStr(desc,"HP")-3,3)
				execute("spawned=CInt("&tempaddhp&")")
				wscript.echo "Healed "&spawned&" HP."
				plhealth=plhealth+spawned
				if plhealth>100 then
					plhealth=100
				end if
				RemoveItem itemtosearchfor,0
			end if
			if inStr(lcase(desc),"dp")<>0 and gotaction=0 then
				gotaction=1
				choice=""
				if lcase(itemtosearchfor)=lcase(plequip) then
					Do until lcase(choice)="y" or lcase(choice)="n"
						choice=UserInput("Unequip the "&itemtosearchfor&"? (y/n)>")
					Loop
					if lcase(choice)="y" then
						plequip=""
					end if
				else
					Do until lcase(choice)="y" or lcase(choice)="n"
						choice=UserInput("Equip the "&itemtosearchfor&"? (y/n)>")
					Loop
					if lcase(choice)="y" then
						plequip=itemtosearchfor
					end if
				end if
				LoadMap mappath,0
			end if
			if gotaction=0 then
				wscript.echo "It didn't do anything."
				userinput("Press ENTER to continue.")
				LoadMap mappath,0
			end if
		end if
	end if
	if uin="look" and evaluated=0 then
		evaluated=1
		LoadMap mappath,0
		msgbox mapinfo
		evaluated=0
	end if
	if uin="save" and evaluated=0 then
		SaveGame()
	end if
	if uin="exit" and evaluated=0 then
		exitchoice=UserInput("Do you want to save before you quit (y/n)?:")
		if exitchoice="y" then
			SaveGame()
		end if
		evaluated=1
		wscript.quit
	end if
	if uin="help" and evaluated=0 then
		wscript.echo "Commands:"
		wscript.echo "WALK/RUN/TELEPORT		Move around the map. Use the WASD keys."
		wscript.echo "STATS				View your stats."
		wscript.echo "INVENTORY			View the items in your inventory."
		wscript.echo "EXAMINE [ITEM NAME]		Shows the item's description."
		wscript.echo "USE [ITEM NAME]			Uses an item."
		wscript.echo "DISCARD / DROP [ITEM NAME]	Drop an item."
		wscript.echo "LOOK				Look around the map."
		wscript.echo "SAVE				Save the game."
		wscript.echo "EXIT				Exit the game."
		wscript.echo "HELP				Prints out a list of commands."
		wscript.echo "SUICIDE				Game over."
		wscript.echo "LOCKWALK				Lock/unlock your weapon's current direction."
		if debugarg=1 then
			wscript.echo "TEST				Helps with map testing."
			wscript.echo "RELMAP0				Reload the current map & keeps the player's position."
			wscript.echo "RELMAP1				Reload the current map & resets the player's position."
		end if
		Userinput("Press ENTER to continue.")
		LoadMap mappath,0
	end if
	if uin="test" and evaluated=0 and debugarg=1 then
		opt=""
		Do Until opt="1" or opt="2" or opt="3" or opt="4" or opt="5" or opt="6" or opt="7" or opt="8"
			opt=UserInput("1=Change Age,2=Change Money,3=Add Item,4=Variable list,5=AddEnemy,6=ViewEnemies,7=ToggleAI,8=Modify Health>")
			if opt="1" then
				plage=UserInput("Age (Current:"&plage&")>")
				Exit Do
			end if
			if opt="2" then
				plmoney=UserInput("Money (Current: $"&plmoney&")>")
				Exit Do
			end if
			if opt="3" then
				AddItem UserInput("NEW ITEM NAME>"),UserInput("NEW ITEM DESC>"),1
			end if
			if opt="4" then
				msgbox "PLNAME="&plname&vbCr&"PLAGE="&plage&vbCr&"PLMONEY="&plmoney&vbCr&"PLEQUIP="&plequip&vbCr&"INVENTORY="&inventory&vbCr&"INVENTORYSHOWN="&inventoryshown&vbCr&"XPOS AND YPOS: "&xpos&"  "&ypos
			end if
			if opt="5" then
				debugargA=UserInput("Enemy Name>")
				debugargB=UserInput("Enemy YPOS>")
				debugargC=UserInput("Enemy XPOS>")
				debugargD=UserInput("Enemy HP>")
				AddMapEnemy debugargA,debugargB,debugargC,debugargD
			end if
			if opt="6" then
				msgecho = ""
				for i = 0 to enemycount
					msgecho = msgecho&vbCr&npcenemy(i)
				Next
				msgbox msgecho
			end if
			if opt="7" then
				if EnemyFollow <> 2 then
					EnemyFollow = 2
					wscript.echo "AI Disabled."
				else
					wscript.echo "AI Enabled."
					EnemyFollow = 0
				end if
			end if
			if opt="8" then
				plhealth=UserInput("New Health (Current:"&plhealth&")>")
			end if
		Loop
	end if
	if uin="relmap0" and evaluated=0 and debugarg=1 then
		LoadMap mappath,0
	end if
	if uin="relmap1" and evaluated=0 and debugarg=1 then
		maptoload=mappath
		mappath=""
		LoadMap maptoload,1
	end if
	if uin="suicide" and evaluated=0 then
		evaluated=1
		wscript.echo "You decide that the best course of action here is to exit the game. You ascend to heaven."
		pauseScript()
		wscript.echo "GAME OVER!"
		wscript.quit
	end if
	evaluated=0
end function
function ShowStartGameGUI()
	Do
		wscript.echo("  ► Default Story by Clay Hanson (default)")
		For Each objFolder in objFSO.GetFolder(".\addons\").SubFolders
			tempname=objFolder.name
			path=objFolder.path
			valid=1
			if inStr(tempname,"mp_")=1 then
				if objFSO.FileExists(path&"\description.txt")=False then
					wscript.echo "ERROR :: Mappack '"&tempname&"' does not have a 'description.txt'."
					valid=0
				end if
				if objFSO.FileExists(path&"\mapcfg.txt")=False then
					wscript.echo "ERROR :: Mappack '"&tempname&"' does not have a 'config.ini'."
					valid=0
				end if	
				if valid=1 then
					set objFileDesc = objFSO.OpenTextFile(path&"\description.txt", ForReading)
					if objFileDesc.AtEndOfStream = False then
						title=objFileDesc.readLine
					else
						title=tempname
					end if
					if objFileDesc.AtEndOfStream = False then
						author=objFileDesc.readLine
					else
						author="N/A"
					end if
					objFileDesc.close
					wscript.echo("  ► "&title&" by "&author&" ("&tempname&")")
				end if
			end if
		Next
		choice=UserInput("Mappack Filename (Example: mp_example)>")
		if objFSO.FolderExists("./addons/"&choice&"/") and choice<>"default" and choice<>"" then
			plname=""
			Do Until plname<>""
				plname=userinput("Player Name>")
			Loop
			pathtemp="./addons/"&choice&"/"
			if objFSO.FileExists(pathtemp&"mapcfg.txt") then
				Set objMapINI = objFSO.OpenTextFile(pathtemp&"mapcfg.txt", ForReading)
				Do While objMapINI.AtEndOfStream = False
					strLine=objMapINI.readLine
					if inStr(lCase(strLine),"map=")=1 then
						mappath=pathtemp&Replace(lCase(strLine),"map=","")
					end if
					if inStr(lCase(strLine),"age=")=1 then
						plage=Replace(strLine,"age=","")
					end if
					if inStr(lCase(strLine),"money=")=1 then
						plmoney=Replace(strLine,"money=","")
					end if
					if inStr(lCase(strLine),"itemname=")=1 then
						itemname=Replace(strLine,"itemname=","")
					end if
					if inStr(lCase(strLine),"itemdesc=")=1 then
						itemdesc=Replace(strLine,"itemdesc=","")
					end if
					if inStr(lCase(strLine),"script=")=1 then
						if objFSO.FileExists(Replace(strLine,"script=","")) then
							Exec(Replace(strLine,"script=",""))
						end if
					end if
				Loop
				objMapINI.close
				plhealth="100"
				if itemname<>"" and itemdesc<>"" then
					AddItem itemname,itemdesc,0
					plequip=itemname
				end if
				LoadMap mappath,1
				Exit Do
			end if
		else
			if choice="default" and choice<>"" then
				StartGame(0)
			else
				wscript.echo "Map pack does not exist."
				UserInput("Press ENTER to continue.")
				LoadMap "./base/t_2016",1
			end if
		end if
	Loop
end function
function StartGame(et)
	if et=1 then
		LoadMap mappath,2
	else
		plname=""
		Do Until plname<>""
			plname=userinput("Player Name>")
		Loop
		plage="15"
		plequip="Stick"
		plmoney="50"
		plhealth="100"
		AddItem "Stick","A stick. Deals 005DP.",0
		'PlaySequence "./maps/sequences/intro/intro",10,40,"0"
		LoadMap "maps\start",1
	end if
end function
if objFSO.FileExists("./base/t_2016.txt")<>true or objFSO.FileExists("./base/t_2016.ch")<>true or objFSO.FileExists("./base/t_2016.desc")<>true then
	msgbox "FATAL ERROR!"&vbCr&"The required map './base/t_2016' was not found. Please re-install ADVENT.",016,"Fatal Error - ADVENT"
	wscript.quit
end if
episodic=1
'DO NOT USE, BROKEN FEATURE. KEEP VALUE AT 1.
if episodic=2 then
	wscript.quit
	If objFSO.FolderExists(".\EP2")<>true Then
		wscript.echo "Error loading episodic game data."
		pauseScript()
		wscript.quit
	end if
	if objFSO.FileExists(".\EP2\gameinfo.txt")<>true then
		wscript.echo "Could not find 'gameinfo.txt'."
		pauseScript()
		wscript.quit
	end if
	loadMap "./EP2/maps/t_2016ep2",1
else
	'WScript.Echo objArgs.Count
	foundmap=0
	foundage=0
	foundmoney=0
	foundname=0
	For Each strArg in objArgs
		if lastargstr="-map" then
			foundmap=1
			maptoload=strArg
		end if
		if lastargstr="-age" and foundmap=1 then
			foundage=1
			plage=strArg
		end if
		if lastargstr="-money" and foundmap=1 then
			foundmoney=1
			plmoney=strArg
		end if
		if lastargstr="-name" and foundmap=1 then
			foundname=1
			plname=strArg
		end if
		if lastargstr="-xpos" and foundmap=1 then
			foundxpos=1
			xpos=strArg
		end if
		if lastargstr="-ypos" and foundmap=1 then
			foundypos=1
			ypos=strArg
		end if
		if lastargstr="-debug" then
			debugarg=strArg
		end if
		lastargstr=strArg
	Next
	if foundmap=1 then
		if foundage=0 then
			plage=0
		end if
		if foundmoney=0 then
			plmoney=15
		end if
		if foundname=0 then
			plname="DEBUG"
		end if
		plhealth="100"
		mappath=""
		if foundxpos=0 and foundypos=0 then
			loadmap maptoload,1
		else
			loadmap maptoload,0
		end if
	else
		PlaySequence "./base/intro/intro",3,70,"0"
		LoadMap "./base/t_2016",1
	end if
end if
Function EvaluateMapString(str)
	i=0
	If Len(str)<30 then
		Do Until Len(str)=30
			str=str&" "
		Loop
	End If
	tempx=0
	invalidpos=0
	Do Until Len(str)=i
		i=i+1
		specialchar=0
		if DEBUGA=1 then
			'wscript.echo tempy&", "&ypos&"; "&tempx&", "&xpos
		end if
		drewsword=0
		gotenemy=0
		if enemycount <> 0 then
			for g = 0 to enemycount
				if inStr(npcenemy(g),"[") <> 0 and inStr(npcenemy(g),"]") <> 0 then
					hp=Mid(npcenemy(g),inStr(npcenemy(g),"_HP")+3,3)
					if hp <> "000" then
						tempenemyypos=Mid(npcenemy(g),inStr(npcenemy(g),"_y")+2,2)
						tempenemyxpos=Mid(npcenemy(g),inStr(npcenemy(g),"_x")+2,2)
						tempname=Mid(npcenemy(g),inStr(npcenemy(g),"[")+1,inStr(npcenemy(g),"]")-(inStr(npcenemy(g),"[")+1))
						'msgbox tempenemyypos&" | "&tempy&vbCr&tempenemyxpos&" | "&tempx&vbCr&spawned&vbCr&tempname
						if CInt(tempenemyypos)=CInt(tempy) and CInt(tempenemyxpos)=CInt(tempx) and inStr(mappath,"./base/t_")=0 and inStr(spawned,"_"&tempname&"_")=0 then
							if inStr("0123456789",mid(str,i,1)) <> 0 then
								specialchar=1
								temphp=Mid(npcenemy(g),inStr(npcenemy(g),"_HP")+3,3)
								npcenemy(g)=replace(npcenemy(g),"_HP"&temphp,"_HP000")
							end if
						'	if mid(str,i,1)="v" then
						'		specialchar=1
						'		UpdateEnemyPosition tempname,tempenemyypos+1,tempenemyxpos
						'	end if
						'	if mid(str,i,1)=">" then
						'		specialchar=1
						'		UpdateEnemyPosition tempname,tempenemyypos+1,tempenemyxpos+1
						'	end if
						'	if mid(str,i,1)="<" then
						'		specialchar=1
						'		UpdateEnemyPosition tempname,tempenemyypos-1,tempenemyxpos-1
						'	end if
						'	if mid(str,i,1)="^" then
						'		specialchar=1
						'		UpdateEnemyPosition tempname,tempenemyypos-1,tempenemyxpos
						'	end if
							if mid(str,i,1)="X" then
								specialchar=1
								UpdateEnemyPosition tempname,tempenemyypos-1,tempenemyxpos-1
							end if
							execute("temphp=CInt("&hp&")")
							if specialchar=0 and temphp<100 then
								MapCacheAdd(UCase(Mid(tempname,1,1)))
							else
								MapCacheAdd(mid(str,i,1))
							end if
							'spawned=spawned&"_"&tempname&"_"
							gotenemy=1
							if walkdir="right" and CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos+1) and inStr(mappath,"./base/t_")=0 then
								if plequip<>"" then
									intStart=InStr(lcase(inventory),"{"&lcase(plequip))
									intStart=intStart + Len("{"&plequip)+1
									intEnd=inStr(lcase(inventory),lcase(plequip)&"}")
									desc=Mid(inventory,intStart,intEnd-intStart)
									if inStr(desc,"DP") <> 0 then
										tempdp=Mid(desc,inStr(desc,"DP")-3,3)
										hurt=1
										execute("spawned=CInt("&tempdp&")")
										hurtmsg="Dealt "&spawned&" DP to "&UCase(tempname)&"."
										temphpto=hp
										temphpto=hp-spawned
										if temphpto < 0 then
											temphpto=0
										end if
										execute("temphpto="""&temphpto&"""")
										Do until Len(temphpto)=3
											temphpto="0"&temphpto
										Loop
										execute("hp="""&hp&"""")
										Do until Len(hp)=3
											hp="0"&hp
										Loop
										npcenemy(g)=replace(npcenemy(g),"_HP"&hp,"_HP"&temphpto)
									else
										msgbox "PLEQUIP ITEM "&plequip&" does not have DP in it's description. Correct format:"&vbCr&"015DP"
									end if
								end if
							end if
							if walkdir="left" and CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos-1) and inStr(mappath,"./base/t_")=0 then
								if plequip<>"" then
									intStart=InStr(lcase(inventory),"{"&lcase(plequip))
									intStart=intStart + Len("{"&plequip)+1
									intEnd=inStr(lcase(inventory),lcase(plequip)&"}")
									desc=Mid(inventory,intStart,intEnd-intStart)
									if inStr(desc,"DP") <> 0 then
										tempdp=Mid(desc,inStr(desc,"DP")-3,3)
										hurt=1
										execute("spawned=CInt("&tempdp&")")
										hurtmsg="Dealt "&spawned&" DP to "&UCase(tempname)&"."
										temphpto=hp
										temphpto=hp-spawned
										if temphpto < 0 then
											temphpto=0
										end if
										execute("temphpto="""&temphpto&"""")
										Do until Len(temphpto)=3
											temphpto="0"&temphpto
										Loop
										execute("hp="""&hp&"""")
										Do until Len(hp)=3
											hp="0"&hp
										Loop
										npcenemy(g)=replace(npcenemy(g),"_HP"&hp,"_HP"&temphpto)
									else
										msgbox "PLEQUIP ITEM "&plequip&" does not have DP in it's description. Correct format:"&vbCr&"015DP"
									end if
								end if
							end if
							if walkdir="up" and CInt(tempy)=CInt(ypos-1) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 then
								if plequip<>"" then
									intStart=InStr(lcase(inventory),"{"&lcase(plequip))
									intStart=intStart + Len("{"&plequip)+1
									intEnd=inStr(lcase(inventory),lcase(plequip)&"}")
									desc=Mid(inventory,intStart,intEnd-intStart)
									if inStr(desc,"DP") <> 0 then
										tempdp=Mid(desc,inStr(desc,"DP")-3,3)
										hurt=1
										execute("spawned=CInt("&tempdp&")")
										hurtmsg="Dealt "&spawned&" DP to "&UCase(tempname)&"."
										temphpto=hp
										temphpto=hp-spawned
										if temphpto < 0 then
											temphpto=0
										end if
										execute("temphpto="""&temphpto&"""")
										Do until Len(temphpto)=3
											temphpto="0"&temphpto
										Loop
										execute("hp="""&hp&"""")
										Do until Len(hp)=3
											hp="0"&hp
										Loop
										npcenemy(g)=replace(npcenemy(g),"_HP"&hp,"_HP"&temphpto)
									else
										msgbox "PLEQUIP ITEM "&plequip&" does not have DP in it's description. Correct format:"&vbCr&"015DP"
									end if
								end if
							end if
							if walkdir="down" and CInt(tempy)=CInt(ypos+1) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 then
								if plequip<>"" then
									intStart=InStr(lcase(inventory),"{"&lcase(plequip))
									intStart=intStart + Len("{"&plequip)+1
									intEnd=inStr(lcase(inventory),lcase(plequip)&"}")
									desc=Mid(inventory,intStart,intEnd-intStart)
									if inStr(desc,"DP") <> 0 then
										tempdp=Mid(desc,inStr(desc,"DP")-3,3)
										hurt=1
										execute("spawned=CInt("&tempdp&")")
										hurtmsg="Dealt "&spawned&" DP to "&UCase(tempname)&"."
										temphpto=hp
										temphpto=hp-spawned
										if temphpto < 0 then
											temphpto=0
										end if
										execute("temphpto="""&temphpto&"""")
										Do until Len(temphpto)=3
											temphpto="0"&temphpto
										Loop
										execute("hp="""&hp&"""")
										Do until Len(hp)=3
											hp="0"&hp
										Loop
										npcenemy(g)=replace(npcenemy(g),"_HP"&hp,"_HP"&temphpto)
									else
										msgbox "PLEQUIP ITEM "&plequip&" does not have DP in it's description. Correct format:"&vbCr&"015DP"
									end if
								end if
							end if
							if CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 then
								plhealth=plhealth-5
								hurt=1
								hurtmsg="Recieved 5 DP(s) from "&UCase(tempname)&"."
							end if
						end if
					end if
				end if
			next
		end if
		if walkdir="right" and CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos+1) and inStr(mappath,"./base/t_")=0 and gotenemy=0 then
			if plequip<>"" then
				MapCacheAdd("/")
				drewsword=1
			end if
		end if
		if walkdir="down" and CInt(tempy)=CInt(ypos+1) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 and gotenemy=0 then
			if plequip<>"" then
				MapCacheAdd("|")
				drewsword=1
			end if
		end if
		if walkdir="left" and CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos-1) and inStr(mappath,"./base/t_")=0 and gotenemy=0 then
			if plequip<>"" then
				MapCacheAdd("\")
				drewsword=1
			end if
		end if
		if walkdir="up" and CInt(tempy)=CInt(ypos-1) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 and gotenemy=0 then
			if plequip<>"" then
				MapCacheAdd("|")
				drewsword=1
			end if
		end if
		if CInt(tempy)=CInt(ypos) and CInt(tempx)=CInt(xpos) and inStr(mappath,"./base/t_")=0 and gotenemy=0 then
			'OLD COLLISION
'			if mid(str,i,1)="v" then
'				specialchar=1
'				ypos=ypos+1
'			end if
'			if mid(str,i,1)=">" then
'				specialchar=1
'				xpos=xpos+1
'			end if
'			if mid(str,i,1)="<" then
'				specialchar=1
'				xpos=xpos-1
'				reloadmap=1
'			end if
'			if mid(str,i,1)="^" then
'				specialchar=1
'				ypos=ypos-1
'				reloadmap=1
'			end if
			if mid(str,i,1)="X" and switchtonew=0 then
				specialchar=1
				'ypos=ypos-1
				ypos=oldy
				xpos=oldx
				reloadmap=1
			end if
			if switchtonew=1 then
				switchtonew=0
			end if
			'NPC CHARACTERS: àáâãäabcdefghijklmnopqrstuwxyz
			if inStr("àáâãäabcdefghijklmnopqrstuwxyz",mid(str,i,1)) <> 0 then
				'Talk to NPC
				specialchar=1
				Set objFileNPC = objFSO.OpenTextFile(mappath&".ch", ForReading)
				found=0
				Do While objFileNPC.AtEndOfStream = False
					strLine = objFileNPC.ReadLine
					if inStr(strLine,mid(str,i,1)&"]")=1 then
						found=1
						strLine=replace(strLine,mid(str,i,1)&"]","")
						msg=strLine
					end if
				Loop
				objFileNPC.close
				gaveitem=0
				if found=1 then
					msgecho=msg
					if inStr(msgecho,"%NAME%") then
						msgecho=Replace(msgecho,"%NAME%",plname)
					end if
					if inStr(msgecho,"%MONEY%") then
						msgecho=Replace(msgecho,"%MONEY%",plmoney)
					end if
					if inStr(msgecho,"%AGE%") then
						msgecho=Replace(msgecho,"%AGE%",plage)
					end if
					if inStr(msgecho,"\n") then
						msgecho=Replace(msgecho,"\n","""&vbCr&""")
					end if
					if inStr(msg,"{")<>0 and inStr(msg,"}")<>0 and inStr(msg,"[")<>0 and inStr(msg,"]")<>0 then
						msgecho=replace(msgecho,mid(msgecho,inStr(msgecho,"{"),Len(msgecho)-(inStr(msgecho,"{")-1)),"")
					end if
					toexecute="msgbox """&msgecho&""""
					execute(toexecute)
					if inStr(msg,"{")<>0 and inStr(msg,"}")<>0 and inStr(msg,"[")<>0 and inStr(msg,"]")<>0 then
						gaveitem=1
						cannotbuy=0
						intStart=InStr(msg,"{")
						intStart=intStart + Len("{")
						intEnd=inStr(msg,"}")
						tempitemdesc=Mid(msg,inStr(msg,"[")+1,inStr(msg,"]")-(inStr(msg,"[")+1))
						tempitem=Mid(msg,intStart,intEnd-intStart)
						choice=0
						if inStr(tempitem,"%")<>0 then
							reqage=mid(tempitem,inStr(tempitem,"%")+1,3)
							if CInt(reqage)>CInt(plage) then
								hurt=1
								hurtmsg="You are too young to buy this item. (Required age: "&CInt(reqage)&"+)"
								cannotbuy=1
							end if
							tempitem=replace(tempitem,"%"&reqage,"")
						end if
						if inStr(tempitem,"$")<>0 and cannotbuy=0 then
							pricetopay=mid(tempitem,inStr(tempitem,"$")+1,4)
							Do until choice=7 or choice=6
								if inStr("aeiou",lcase(mid(tempitem,1,1))) then
									choice=MsgBox("Do you want to pay $"&CInt(pricetopay)&" for an '"&Replace(tempitem,"$"&pricetopay,"")&"'?",3)
								else
									choice=MsgBox("Do you want to pay $"&CInt(pricetopay)&" for a '"&Replace(tempitem,"$"&pricetopay,"")&"'?",3)
								end if
							Loop
							if choice=6 then
								if isNumeric(pricetopay) then
									PayForItem Replace(tempitem,"$"&pricetopay,""),tempitemdesc,1,CInt(pricetopay)
								else
									msgbox "PRICE PAY ERROR: Price is not stated properly. (Proper format: $0051, Improper format: $51)"
								end if
							end if
						else
							if cannotbuy=0 then
								AddItem tempitem,tempitemdesc,1
							end if
						end if
					end if
				else
					'wscript.stdout.write(mid(str,i,1))
					specialchar=0
				end if
			end if
			if mid(str,i,1)="$" then
				'Pay money
				specialchar=1
				Set objFileCH = objFSO.OpenTextFile(mappath&".ch", ForReading)
				found=0
				Do While objFileCH.AtEndOfStream = False
					strLine = objFileCH.ReadLine
					if inStr(strLine,"$]")=1 then
						found=1
						strLine=replace(strLine,"$]","")
						money=CInt(strLine)
					end if
				Loop
				objFileCH.close
				if found=1 then
					if plmoney>=money then
						plmoney=plmoney-money
					else
						hurt=1
						hurtmsg="You don't have enough money."
						ypos=ypos-1
					end if
				else
					if plmoney>=5 then
						plmoney=plmoney-5
					else
						hurt=1
						hurtmsg="You don't have enough money. (Could not find pay amount in map's .CH file - Defaulting to $5.)"
						ypos=ypos-1
					end if
				end if
			end if
			if inStr("0123456789",mid(str,i,1))<>0 and gip=0 then
				'Trigger
				specialchar=1
				Set objFileTRIG = objFSO.OpenTextFile(mappath&".ch", ForReading)
				found=0
				Do While objFileTRIG.AtEndOfStream = False
					strLine = objFileTRIG.ReadLine
					if inStr(strLine,mid(str,i,1)&"E]")=1 then
						strLine=replace(strLine,mid(str,i,1)&"E]","")
						execute(strLine)
					end if
					if inStr(strLine,mid(str,i,1)&"]")=1 then
						found=1
						strLine=replace(strLine,mid(str,i,1)&"]","")
						mappathtemp=strLine
					end if
				Loop
				objFileTRIG.close
				if found=1 then
					hastheitem=0
					if inStr(mappathtemp,"[")<>0 and inStr(mappathtemp,"]")<>0 then
						intStart=InStr(mappathtemp,"[")
						intStart=intStart + Len("[")
						intEnd=inStr(mappathtemp,"]")
						itemtohave=Mid(mappathtemp,intStart,intEnd-intStart)
						mappathtemp=replace(mappathtemp,"["&itemtohave&"]","")
						if HasItem(itemtohave)=False then
							hurt=1
							if inStr("aeiou",mid(lCase(itemtohave),1,1))<>0 then
								hurtmsg="You need an '"&UCase(Itemtohave)&"' to pass through here."
							else
								hurtmsg="You need a '"&UCase(Itemtohave)&"' to pass through here."
							end if
							hastheitem=1
						else
							hastheitem=2
						end if
					end if
					if inStr(mappathtemp,"{")<>0 and inStr(mappathtemp,"}")<>0 and inStr(mappathtemp,"_x")<>0 and inStr(mappathtemp,"_y")<>0 and (hastheitem=0 or hastheitem=2) then
						intStart=InStr(mappathtemp,"{")
						intStart=intStart + Len("{")
						intEnd=inStr(mappathtemp,"}")
						postoset=Mid(mappathtemp,intStart,intEnd-intStart)
						ypos=mid(postoset,inStr(postoset,"_y")+2,2)
						xpos=mid(postoset,inStr(postoset,"_x")+2,2)
						mappathtemp=Replace(mappathtemp,"_x"&xpos,"")
						mappathtemp=Replace(mappathtemp,"_y"&ypos,"")
						mappathtemp=Replace(mappathtemp,"{","")
						mappathtemp=Replace(mappathtemp,"}","")
						triggerload=1
						triggerloadmap=mappathtemp
						switchtonew=1
						'loadMap mappath,0
					else
						if hastheitem<>1 then
							triggerload=2
							triggerloadmap=mappathtemp
							'loadmap mappathtemp,1
						end if
					end if
				end if
			end if
			if specialchar=0 then
				'Player
				if specialchar=0 then
					if playerchar="" then
						MapCacheAdd(mid(str,i,1))
					else
						MapCacheAdd(playerchar)
					end if
				end if
			else
				if mid(str,i,1)<>"*" then
					MapCacheAdd("░")
				else
					if inStr("àáâãää",mid(str,i,1)) then
						MapCacheAdd("☻")
					else
						MapCacheAdd("█")
					end if
				end if
			end if
		else
			if inStr(mappath,"./base/t_")=0 and drewsword=0 and gotenemy=0 then
				if inStr("X",mid(str,i,1)) then
					specialchar=1
					MapCacheAdd("█")
				end if
				if inStr("0123456789",mid(str,i,1)) <> 0 and gip=0 then
					'changelevel Trigger
					specialchar=1
					MapCacheAdd("░")
				end if
				if gip=1 then
					if mid(str,i,1)="4" then
						specialchar=1
						MapCacheAdd("█")
					end if
					if mid(str,i,1)="3" then
						specialchar=1
						MapCacheAdd("▓")
					end if
					if mid(str,i,1)="2" then
						specialchar=1
						MapCacheAdd("▒")
					end if
					if mid(str,i,1)="1" then
						specialchar=1
						MapCacheAdd("▒")
					end if
					if mid(str,i,1)="P" then
						specialchar=1
						MapCacheAdd(tempchar)
					end if
					if mid(str,i,1)="N" then
						specialchar=1
						MapCacheAdd("☻")
					end if
				end if
				if inStr("àáâãääabcdefghijklmnopqrstuwxyz",mid(str,i,1)) <> 0 then
					'NPC
					specialchar=1
					found=0
					Set objFileNPC = objFSO.OpenTextFile(mappath&".ch", ForReading)
					Do While objFileNPC.AtEndOfStream = False
						strLine = objFileNPC.ReadLine
						if inStr(strLine,mid(str,i,1)&"]")=1 then
							found=1
						end if
					Loop
					if found=1 then
						MapCacheAdd("☻")
					else
						MapCacheAdd(mid(str,i,1))
					end if
					objFileNPC.close
				end if
				if mid(str,i,1)="|" and gip=0 then
					specialchar=1
					MapCacheAdd("│")
				end if
				if mid(str,i,1)="_" and gip=0 then
					specialchar=1
					MapCacheAdd("─")
				end if
				if specialchar=0 then
					MapCacheAdd(mid(str,i,1))
				end if
			else
				if drewsword=0 and gotenemy=0 then
					if inStr(mappath,"./base/t_")=1 then
						if mid(str,i,1)="&" then
							specialchar=1
							MapCacheAdd("█")
						end if
						if mid(str,i,1)="*" then
							specialchar=1
							MapCacheAdd("▒")
						end if
					end if
					if specialchar=0 then
						MapCacheAdd(mid(str,i,1))
					end if
				end if
			end if
		end if
		tempx=tempx+1
	Loop
	tempy=tempy+1
	MapCacheAdd("NL")
end Function
function LoadMap(path, resetpos)
	If objFSO.FileExists(path&".txt")=False then
		wscript.echo "Error loading map '"&path&".txt' (Map does not exist)."
		Exit Function
	end if
	If objFSO.FileExists(path&".ch")=False and gip=0 then
		wscript.echo "Error loading map '"&path&".txt' (Map is missing .CH file)."
		Exit Function
	end if
	If objFSO.FileExists(path&".desc")=False and gip=0 then
		wscript.echo "Error loading map '"&path&".txt' (Map is missing .DESC file)."
		Exit Function
	end if
	spawned=""
	if DEBUGA=1 then
		wscript.echo "DEBUG :: Attempting to load map '"&path&"'..."
	end if
	if DEBUGA=1 then
		wscript.echo "DEBUG :: Global Variable 'mappath' equals """&mappath&"""."
		wscript.echo "DEBUG :: Local Variable 'path' equals """&path&"""."
		MsgBox "Pause.",0,"DEBUG"
	end if
	displaymap=""
	if path <> mappath then
		enemycount=0
		npcenemy(0)=""
		npcenemy(1)=""
		npcenemy(2)=""
		npcenemy(3)=""
		npcenemy(4)=""
		npcenemy(5)=""
		npcenemy(6)=""
		npcenemy(7)=""
		npcenemy(8)=""
		npcenemy(9)=""
		npcenemy(10)=""
	end if
	tempmappath = mappath
	mappath = path
	if DEBUGA=1 then
		wscript.echo "DEBUG :: Global Variable 'mappath' equals """&mappath&"""."
		wscript.echo "DEBUG :: Local Variable 'path' equals """&path&"""."
		MsgBox "Pause.",0,"DEBUG"
	end if
	if DEBUGA=1 then
		wscript.echo "DEBUG :: objFSO.FileExists '"&path&"' equals "&objFSO.FileExists(path&".txt")&"."
		MsgBox "Pause.",0,"DEBUG"
	end if
	Set objFile = objFSO.OpenTextFile(path&".txt", ForReading)
	row=0
	foundmapend = 0
	Do While objFile.AtEndOfStream = False
		strLine = objFile.ReadLine
		if strLine="///" then
			foundmapend = 1
		end if
		if foundmapend<>1 and inStr(strLine,"%%%%")=0 then
			row=row+1
		end if
	Loop
	if debug=1 then
		wscript.echo "Got "&row&" lines of text."
	end if
	if row>=26 then
		wscript.echo "FATAL MAP ERROR :: The size of the map in '"&path&".txt' is more than or equal to 26."
		msgbox "FATAL MAP ERROR :: The size of the map in '"&path&".txt' is more than or equal to 26."
		wscript.quit
	end if
	objFile.close
	foundmapend = 0
	mapcache=""
	if resetpos=1 then
		xpos=0
		ypos=0
	end if
	if gip=0 then
		Set objFileEN = objFSO.OpenTextFile(path&".ch", ForReading)
		found=0
		Do While objFileEN.AtEndOfStream = False
			strLine = objFileEN.ReadLine
			if inStr(lCase(strLine),"start]")=1 and resetpos = 1 then
				strLine=replace(strLine,Mid(strLine,1,6),"")
				execute(strLine)
			end if
			if tempmappath <> path and inStr(lCase(strLine),"enemy]") <> 0 and inStr(strLine,"_HP") <> 0 and inStr(strLine,"{") <> 0 and inStr(strLine,"}") <> 0 and inStr(strLine,"_x") <> 0 and inStr(strLine,"_y") <> 0 then
				tempname=Mid(strLine,inStr(strLine,"{")+1,inStr(strLine,"}")-(inStr(strLine,"{")+1))
				tempypos=Mid(strLine,inStr(strLine,"_y")+2,2)
				tempxpos=Mid(strLine,inStr(strLine,"_x")+2,2)
				temphp=Mid(strLine,inStr(strLine,"_HP")+3,3)
				AddMapEnemy tempname,tempypos,tempxpos,temphp
			end if
		Loop
		objFileEN.close
		if resetpos <> 1 then
			if EnemyFollow=1 then
				EnemyFollow=0
				Call CalcEnemyPos()
			else
				if EnemyFollow<>2 then
					EnemyFollow=1
				end if
			end if
		end if
	end if
	Set objFile = objFSO.OpenTextFile(path&".txt", ForReading)
	tempx=0
	tempy=0
	Do While objFile.AtEndOfStream = False
		strLine = objFile.ReadLine
		if inStr(strLine,"%%%%")=1 and resetpos=1 then
			ypos=mid(strLine,inStr(strLine,"_y")+2,2)
			xpos=mid(strLine,inStr(strLine,"_x")+2,2)
		end if
		if strLine="///" then
			foundmapend = 1
		end if
		if foundmapend<>1 and inStr(strLine,"%%%%")=0 then
			if mapcache="" then
				mapcache=mapcache&mid(strLine,4,len(strLine))
			else
				mapcache=mapcache&vbLf&mid(strLine,4,len(strLine))
			end if
			EvaluateMapString(mid(strLine,4,len(strLine)))
		end if
	Loop
	objFile.close
	wscript.echo mid(displaymap,1,Len(displaymap)-1)
	mapinfo = ""
	if gip=1 then
		Exit Function
	end if
	Set objFile = objFSO.OpenTextFile(path&".desc", ForReading)
	Do While objFile.AtEndOfStream = False
		strLine = objFile.ReadLine
		if mapinfo = "" then
			mapinfo=mapinfo&strLine
		else
			mapinfo=mapinfo&vbCr&strLine
		end if
	Loop
	objFile.close
	if hurt=1 then
		hurt=0
		wscript.echo hurtmsg
	end if
	uinput=""
	validinput=0
	if triggerload=1 then
		triggerload=0
		LoadMap triggerloadmap,0
	end if
	if triggerload=2 then
		triggerload=0
		LoadMap triggerloadmap,1
	end if
	if reloadmap=1 then
		reloadmap=0
		LoadMap mappath,0
		Exit Function
	end if
	if resetpos=0 then
		Exit Function
	end if
	if inStr(path,"./base/t_")=0 then
		if DEBUGA=1 then
			wscript.echo "DEBUG :: Map name includes no special string(s) - Initiating default menu."
		end if
		Do until validinput = 1
			uinput=UserInput("What to do?")
			evaluateInput(uinput)
		Loop
	end if
	if inStr(path,"./base/t_")=1 then
		if DEBUGA=1 then
			wscript.echo "DEBUG :: Map name includes special string './base/t_' - Initiating menu input."
		end if
		Do
			uinput=UserInput("Option: ")
			if uinput<>"" then
				if uinput="1" then
					ShowStartGameGUI()
					Exit Do
				end if
				if uinput="2" then
					slot="-1"
					Do until CInt(slot)<CInt(maximumSaveFiles) and CInt(slot)>0
						slot=UserInput("Slot (1 - "&maximumSaveFiles&"):")
					Loop
					if CInt(slot)<=10 then
						slot="0"&slot
					end if
					if objFSO.FileExists("./base/"&savefilePrefix&slot&".avs")<>true then
						wscript.echo "Save file does not exist."
					else
						'FORMAT:
						'Line 1: PL NAME, Line 2: PL MONEY, Line 3: PL AGE, Line 4: MAP, Line 5: Has key to goblins' place
						Set objFile = objFSO.OpenTextFile("./base/"&savefilePrefix&slot&".avs", ForReading)
						maploadcount=0
						if objFile.AtEndOfStream = False then
							plname = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got PLNAME."
							end if
						end if
						if objFile.AtEndOfStream = False then
							plmoney = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got PLMONEY."
							end if
						end if
						if objFile.AtEndOfStream = False then
							plage = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got PLAGE."
							end if
						end if
						if objFile.AtEndOfStream = False then
							mappath = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got MAPPATH."
							end if
						end if
						if objFile.AtEndOfStream = False then
							inventory = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got INVENTORY."
							end if
						end if
						if objFile.AtEndOfStream = False then
							xpos = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got XPOS"
							end if
						end if
						if objFile.AtEndOfStream = False then
							ypos = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got YPOS."
							end if
						end if
						if objFile.AtEndOfStream = False then
							plequip = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got PLEQUIP."
							end if
						end if
						if objFile.AtEndOfStream = False then
							inventoryshown = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got INVENTORYSHOWN."
							end if
						end if
						if objFile.AtEndOfStream = False then
							plhealth = strReverse(objFile.ReadLine)
							maploadcount=maploadcount+1
							if DEBUGA=1 then
								wscript.echo "DEBUG :: Got PLHEALTH."
							end if
						end if
						enemycount=0
						for f = 0 to 10
							if objFile.AtEndOfStream = False then
								templine = strReverse(objFile.ReadLine)
								if inStr(templine,"[") <> 0 and inStr(templine,"]") <> 0 and inStr(templine,"_HP") <> 0 and inStr(templine,"_x") <> 0 and inStr(templine,"_y") <> 0 then
									npcenemy(f) = templine
									enemycount=enemycount+1
									if DEBUGA=1 then
										wscript.echo "DEBUG :: Got NPCENEMY("&lsen&")."
									end if
								end if
								maploadcount=maploadcount+1
							end if
						next
						if DEBUGA=1 then
							wscript.echo "DEBUG :: "&maploadcount
						end if
						if maploadcount=21 and isNumeric(goblinkey) and isNumeric(plmoney) and isNumeric(plage) then
							StartGame(1)
						else
							wscript.echo "Save file is corrupted and cannot be used."
							mappath=""
							plage=""
							plmoney=""
							plname=""
						end if
						objFile.close
					end if
				end if
				if uinput="3" then
					Do
						editopt=InputBox("Change an option:"&vbCr&"1) Check for updates on game start ("&checkForUpdates&")"&vbCr&"2) Save file prefix ("&savefilePrefix&")"&vbCr&"3) Maximum save file count ("&maximumsavefiles&")"&vbCr&"4) Volume ("&globalvolume&")"&vbCr&"5) Play map sequences (Kind of like a GIF) ("&playsequences&")"&vbCr&"6) Save and exit options menu","Change options - ADVENT")
						if editopt="1" then
							if checkForUpdates=1 then
								checkForUpdates=0
							else
								checkForUpdates=1
							end if
						end if
						if editopt="2" then
							Do
								savefilePrefix=InputBox("Save File Prefix (A-Z ONLY)>")
								if inStr("1234567890!@#$%^&*()_+-={}|[]\:"",./<>?l;'",savefilePrefix)<>0 then
									msgbox "Invalid characters. You can only enter A-Z."
								else
									Exit Do
								end if
							Loop
						end if
						if editopt="3" then
							Do
								maximumSaveFiles=InputBox("Maximum Save Files>")
								if isNumeric(maximumSaveFiles)<>True then
									msgbox "Invalid characters. You can only enter 0-9."
								else
									if maximumSaveFiles<1 then
										msgbox "Invalid number. Must be larger than 0."
									else
										Exit Do
									end if
								end if
							Loop
						end if
						if editopt="4" then
							Do
								globalvolume=InputBox("Volume (Lower than 100)>")
								if isNumeric(globalvolume)<>True then
									msgbox "Invalid characters. You can only enter 0-9."
								else
									if globalvolume>100 or globalvolume<-1 then
										msgbox "Invalid number. Must be lower than 100 and higher than 0."
									else
										Exit Do
									end if
								end if
							Loop
						end if
						if editopt="5" then
							if playsequences=1 then
								playsequences=0
							else
								playsequences=1
							end if
						end if
						if editopt="6" then
							WriteConfigFile(0)
							Exit Do
						end if
					Loop
				end if
				if uinput="4" then
					wscript.echo "Created By: Clay Hanson"
					wscript.echo "Helpers: Michael Hart (Bitl)"
					wscript.echo "Game Version: v"&gamever
					wscript.echo " "
					wscript.echo "Please give credit if you make modifications of this game."
					pausescript()
					LoadMap mappath,0
				end if
				if uinput="5" then
					wscript.quit
				end if
				if uinput="debug" then
					wscript.echo "ENTERING DEBUG STATE, SETTING VARIABLES."
					plname="DEBUG"
					plage=UserInput("Age>")
					plmoney=UserInput("Money>")
					plhealth=100
					validinput=1
					Exit Do
				end if
			end if
		Loop
	end if
End Function
'error handling - go here if you somehow bypass the normal DO loops.
Do
	x=userinput("DEBUG OPTION (0=Load Map, 1=Execute File)>")
	if x=0 then
		x=userinput("MAP PATH>")
		LoadMap x,1
	else
		x=userinput("FILE PATH>")
		Exec(x)
	end if
Loop

'9:36 AM 5/20/2016
'http://creativecommons.org/licenses/by-nc-sa/4.0/