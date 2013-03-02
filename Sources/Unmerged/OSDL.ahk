#NoEnv
#include httpRequest.ahk	;http://www.autohotkey.com/board/topic/67989-func-httprequest-for-web-apis-ahk-b-ahk-lunicodex64/
/*
OpenSubtitles Downloader

Sites:
https://github.com/submersion/opensubtitles-downloader
http://www.autohotkey.com/board/forum/49-scripts/
http://forum.opensubtitles.org/viewtopic.php?f=11&t=14045

License:
GNU GPL v3

Changes in v1.0A5
- Added support for search for subtitles in multiple languages
- Minor GUI and functionality fixes
Changes in v1.0A4
- Send To menu item now works with 
	- multiple files selected
	- folders selected - search does not include subfolders
- Changed subtitle file name indexing from "index" to "(index)" when there are multiple subtitle files with the same name 
- Fixed results regex matching bug
- Traytip notifications now include the name of the movie file, and the name of downloaded subtitle file
Changes in v1.0A3:
- Tray menu
Changes in v1.0A2:
- Control flow
- Parsing different result page for single result search in searchOS()
- Replaced opensubtitles_hash() with GetOpenSubtitlesHash() for faster speed
- Languages in alphabetical order in OSDLGUI DDL
- Added OSDL GUI functionality: 
	- save preferred language on OSDL DDL change, 
	- show send to menu item state
	- traytip notifications + option to disable them
*/

;// SETTINGS
saveSettings	= 1 ;1 - yes; 0 - no
trayTipOn		= 1 ; 1 - traytip notifications on (1) or off (0)		
scriptName		= OpenSubtitles Downloader
scriptVersion 	= v1.0A5

;// SCRIPT
Menu, Tray, NoStandard
Menu, Tray, Add, E&xit, GuiClose
;// Languages
SplitPath, A_ScriptFullPath,,iniDir,,iniNameNoExt
iniPath := iniDir "\" iniNameNoExt ".ini"
IniRead, defLanguage, %iniPath% ,Default,language
if (!defLanguage) ;no language settings
	defLanguage = All Languages
DDLCont = All Languages|Albanian|Arabic|Armenian|Basque|Bengali|Bosnian|Brazilian|Breton|Bulgarian|Catalan|Chinese|Croatian|Czech|Danish|Dutch|English|Esperanto|Estonian|Finnish|French|Georgian|German|Galician|Greek|Hebrew|Hindi|Hungarian|Icelandic|Indonesian|Italian|Japanese|Kazakh|Khmer|Korean|Latvian|Lithuanian|Luxembourgish|Macedonian|Malay|Norwegian|Occitan|Persian|Polish|Portuguese|Romanian|Russian|Serbian|Sinhalese|Slovak|Slovenian|Spanish|Swahili|Swedish|Syriac|Tagalog|Thai|Turkish|Ukrainian|Urdu|Vietnamese
Loop, parse, defLanguage, |
	StringReplace, DDLCont, DDLCont, %A_LoopField%, %A_LoopField%|
if (!InStr(DDLCont,"||")){ ;invalid language settings
	StringReplace, DDLCont, DDLCont, All Languages|, All Languages||
	defLanguage = All Languages
}
StringReplace, defLanguageForDisplay, defLanguage, |,`,%a_space% , All

clNr = %0%
if (!clNr) ;// no command line parameters
	Gosub, OSDLGUI
else 
	Gosub, CommandLine
return 

OSDLGUI:
	Gosub, OSDLSendToVarsInit
	if (FileExist(sendToItemPath)){
		FileGetShortcut, %sendToItemPath%,outTarget,,currentMenuItemLanguage
		needle = "([a-zA-Z|]*)"
		RegExMatch(currentMenuItemLanguage,needle,out,1)
		StringReplace, currentMenuItemLanguage, out1, |,`,%a_space% , All
	}		
	if (!FileExist(sendToItemPath) || !currentMenuItemLanguage)
		currentMenuItemLanguage = No menu item
 	
	gui, 1:default
	gui, +LabelOSDLGUI
	gui, add, tab2, w500 h270, General|Right click menu item|About
	gui, tab, 1
	gui, add, edit,w250 h25 readonly vOSDLGUIEdit,
	gui, add, button,yp+0 xp+250 w50 gOSDLBrowseFile, &Browse
	;gui, add, ddl,vOSDLGUIDDL yp+0 xp+60 w100 gOSDLChangeDDL,%DDLCont%
	gui, add, edit, yp+0 xp+60 w160 h20 ReadOnly vOSDLGUILangEdit, %defLanguageForDisplay%
	gui, add, listbox, vOSDLGUILB gOSDLChangeDDL Multi yp+30 xp+0 w160 h100, %DDLCont%
	gui, add, text,, (Ctrl+Click to select more than one)
	gui, add, button,yp-100 x20 h25 w150 gOSDLSearch,&Search and download!
	
	gui, tab, 2
	gui, add, text, h20, Right click menu item (in the Send To menu)
	gui, add, button, x20 h30 yp+25 w150 gOSDLSendto, Add/refresh menu item
	gui, add, button, x20 h30 yp+40 w150 gOSDLSendto, Remove menu item
	gui, add, text, xp160 yp-40 w70,for language:
	;gui, add, DDL, xp80 w100 vOSDLSendToLang gOSDLChangeDDL, %DDLCont%
	gui, add, edit, xp+80 w160 ReadOnly vOSDLGUISendToLangEdit, %defLanguageForDisplay%
	gui, add, listbox, vOSDLGUISendToLB gOSDLChangeDDL Multi yp+30 xp+0 w160 h100, %DDLCont%
	
	
	gui, add, text, x20 yp+120 h20, (Use right click on movie file -> Send To -> Download Subtitle.)
	gui, add, text, w50 h20, Currently: 
	gui, add, edit, ReadOnly xp60 h20 w400 vOSDLMenuItemLangTxt, % (currentMenuItemLanguage != "No menu item" ?  "Menu item added for " : "") currentMenuItemLanguage
	gui, tab, 3
	gui, add, text,, %scriptName%
	gui, add, edit, w470 h200 ReadOnly, Version: %scriptVersion%`nAuthor: nordan / gahks`nHomepage: Look for %scriptName% at:`nhttp://www.autohotkey.com/board/forum/49-scripts/`nhttps://github.com/submersion/opensubtitles-downloader`nLicense: GNU GPL v3. The included libraries and functions might have different licenses, check em out.`nThanks to: shajul and Sean for the Unz(), VxE for httpRequest(), just me for GetOpenSubtitlesHash(). Cheers to Delusion for his subdownloader script, which was the inspiration for this one.
	gui, add, text,, 
	gui, add, statusbar,,
	gui, show,, %scriptName% - %scriptVersion%
return

OSDLSendToVarsInit:
	sendToPath := a_appdata "\Microsoft\Windows\SendTo"
	sendToItemPath := sendToPath "\Download Subtitle.lnk" 
return

OSDLSendto:
	GuiControlGet, sendToLang,, OSDLGUISendToLB
	StringReplace, sendToLangForDisplay, sendToLang, |,`,%a_space% , All}
	Gosub, OSDLSendToVarsInit
	if (a_guicontrol == "Add/refresh menu item"){
		if(FileExist(sendToItemPath))
			FileDelete, %sendToItemPath%
		if (A_IsCompiled){
			sendToTargetPath := A_ScriptFullPath
			sendToCL = "%sendToLang%" 
		} else {
			sendToTargetPath := A_AhkPath
			sendToCL = "%A_ScriptFullPath%" "%sendToLang%" 
		}
		FileCreateShortcut, %sendToTargetPath%, %sendToItemPath%, %A_ScriptDir%, %sendToCL%, Download Subtitle
		if (!Errorlevel){
			SB_SetText("Shortcut created/refreshed (for " . sendToLangForDisplay .  ")")
			GuiControl,,OSDLMenuItemLangTxt, Menu item added for %sendToLangForDisplay%
		}
	} else if (a_guicontrol == "Remove menu item"){
		FileDelete, %sendToItemPath%
		if (!ErrorLevel){
			SB_SetText("Shortcut removed.")
			GuiControl,,OSDLMenuItemLangTxt, No menu item
		}
	}
return


OSDLBrowseFile:
	gui, +owndialogs
	gui, +disabled
	FileSelectFile, filePath, 1,,Select a movie file!, Video files (*.*)
	if (!ErrorLevel){
		GuiControl,, OSDLGUIEdit, %filePath%
	}
	gui, -disabled
return

OSDLChangeDDL:
if a_guicontrol in OSDLGUILB,OSDLGUISendToLB
{	
		GuiControlGet, defLanguage2,, %A_GuiControl%
		if (defLanguage != defLanguage2)
			defLanguage := defLanguage2
		else 
			return
		StringReplace, DDLCont, DDLCont, ||, |, All
		Loop, parse, defLanguage2, |
				StringReplace, DDLCont, DDLCont, %A_LoopField%, %A_LoopField%|
		if (a_guicontrol == "OSDLGUISendToLB")
			GuiControl,,OSDLGUILB, |%DDLCont%
		else if (a_guicontrol == "OSDLGUILB")
			GuiControl,,OSDLGUISendToLB, |%DDLCont%
		StringReplace, defLanguageForDisplay, defLanguage, |,`,%a_space% , All
		GuiControl,,OSDLGUILangEdit, %defLanguageForDisplay%
		GuiControl,,OSDLGUISendToLangEdit, %defLanguageForDisplay%
		if (saveSettings){
			IniWrite, %defLanguage%, %iniPath%, Default, language
			if (!ErrorLevel)
				SB_SetText("Preferred language(s) changed to: " . defLanguageForDisplay)
		}
}
return

OSDLSearch:
	gui, +owndialogs
	gui, +disabled
	GuiControlGet,defLanguage,,OSDLGUILB
	GuiControlGet, filePath,,OSDLGUIEdit
	if (saveSettings){
		IniWrite, %defLanguage%, %iniPath%, Default, language
		;if (!ErrorLevel && !filePath)
		;	SB_SetText("Preferred language(s) changed to " . defLanguage)
	}
	defLanguage:=lang(defLanguage)
	if (!filePath) {
		gui, -disabled
		return
	}
	SplitPath, filePath,,,outExt 
	if outExt not in 3g2,3gp,3gp2,3gpp,60d,ajp,asf,asx,avchd,avi,bik,bix,box,cam,dat,divx,dmf,dv,dvr-ms,evo,flc,fli,flic,flv,flx,gvi,gvp,h264,m1v,m2p,m2ts,m2v,m4e,m4v,mjp,mjpeg,mjpg,mkv,moov,mov,movhd,movie,movx,mp4,mpe,mpeg,mpg,mpv,mpv2,mxf,nsv,nut,ogg,ogm,omf,ps,qt,ram,rm,rmvb,swf,ts,vfw,vid,video,viv,vivo,vob,vro,wm,wmv,wmx,wrap,wvx,wx,x264,xvid
	{
		Msgbox, 4,Are you sure?,This file doesn't seem like a video file. Are you sure you would like to continue?
		IfMsgBox, No
		{
			gui, -disabled
			return
		}
	}
	Gui, Destroy
	Gosub, SearchOS
	if (!WinExist("ahk_id " CSGUIHwnd))
		Gosub, GuiClose
return

CommandLine:
	if (clNr == 1) 
	{
		filePath = %1% 
		defLanguage := lang()
		Gosub, SearchOS
	} else {
		clLang = %1%
		defLanguage := lang(clLang)
		filePath = %2%
		Loop, % clNr 
		{
			if (a_index==1)
				continue
			filePath := %a_index%
			FileGetAttrib, fileAttr, %filePath%
			if (fileAttr == "D"){
				Loop, %filePath%\*.*, 0, 0
				{
					if A_LoopFileExt not in 3g2,3gp,3gp2,3gpp,60d,ajp,asf,asx,avchd,avi,bik,bix,box,cam,dat,divx,dmf,dv,dvr-ms,evo,flc,fli,flic,flv,flx,gvi,gvp,h264,m1v,m2p,m2ts,m2v,m4e,m4v,mjp,mjpeg,mjpg,mkv,moov,mov,movhd,movie,movx,mp4,mpe,mpeg,mpg,mpv,mpv2,mxf,nsv,nut,ogg,ogm,omf,ps,qt,ram,rm,rmvb,swf,ts,vfw,vid,video,viv,vivo,vob,vro,wm,wmv,wmx,wrap,wvx,wx,x264,xvid
						continue
					filePath := A_LoopFileLongPath
					Gosub, SearchOS
				}
			} else {
				filePath := %a_index%
				Gosub, SearchOS
			}
		}
	}
	if (!WinExist("ahk_id " CSGUIHwnd))
		Gosub, GuiClose
return

SearchOS:
	SplitPath, filePath, trayFileName
	if (trayTipOn)
		TrayTip, %scriptName%,%trayFileName%:`nHashing...,,1
	hash := GetOpenSubtitlesHash(filePath)
	if (trayTipOn)
		TrayTip, %scriptName%,%trayFileName%:`nSearching for available subtitles...,,1
	subtitles := searchOS(hash,defLanguage)
	if (!subtitles)
		if (trayTipOn){
			TrayTip, %scriptName%,%trayFileName%:`nNo subtitles found...,,1
			Sleep, 1500
		}
	if (defLanguage == "all" || InStr(defLanguage,",")){
		if (trayTipOn)
			TrayTip,
		Gosub, ChooseSubtitleGUI
	} else {
		Gosub, ChooseMostDownloadedSubtitle
		Gosub, DLSubtitle
	}
return
	
ChooseSubtitleGUI:
	gui, 2:default
	gui +LabelCSGUI 
	gui, add, text,, Subtitles for: 
	gui, add, edit, ReadOnly w500 h20, %filePath%
	gui, add, ListView,w500 h100 vLV gLV -LV0x10, Language| |x downloaded|Release|Download Link
	loop, parse, subtitles, csv
	{
		StringSplit,tmp,a_loopfield,|
		LV_Add("",tmp4, tmp3, tmp2, tmp5, tmp1)
		ix := a_index
	}
	sysget, mon, MonitorWorkArea
	newHeight := (ix * 30 > (monBottom-200) ? monBottom-200 : ix*30 )
	guicontrol, move, LV, h%newHeight%
	types=Text,Text,Integer,Text
	StringSplit,types,types,`,
	loop, 4 
		LV_ModifyCol(a_index,types%a_index%)
	LV_ModifyCol(2,"Auto")
	LV_ModifyCol(4,"Auto")
	LV_ModifyCol(5,0)
	LV_ModifyCol(1,"Sort")
	gui, add, button, w80 gLV yp+%newHeight%,&Download!
	gui, add, statusbar
	gui, show,,%scriptName% - %scriptVersion%
	gui, +lastfound
	CSGUIHwnd := WinExist()
return

LV:
	gui, 2:default
	if (a_guievent == "DoubleClick" || A_GuiControl == "&Download!"){
		if (!LV_GetNext(0))
		{
			SB_SetText("No item selected. Please select an item from the list!")
			return
		}
		 LV_GetText(bestRated1, LV_GetNext(0),5)
	 }
	else 
		return
	Gui, Destroy
	Gosub, DLSubtitle
	Gosub, CSGUIClose
return

GuiClose:
OSDLGUIClose:
CSGUIClose:
	ExitApp
return

ChooseMostDownloadedSubtitle:
	bestRated1 := ""
	bestRated2 = 0
	Loop, parse, subtitles, csv
	{
		StringSplit,tmp, a_loopfield, |
		if (tmp2>=bestRated2) {
			bestRated1 := tmp1 
			bestRated2 := tmp2
		}
	}
return

DLSubtitle:
	if (trayTipOn)
		TrayTip, %scriptName%,%trayFileName%:`nDownloading subtitle file...,,1
	SplitPath, filePath,outName, outDir,,outNameNoExt
	dlPath := a_temp . "\OSDL.zip"
	FileDelete, %dlPath%
	URLDownloadToFile, %bestRated1%, %dlPath%
	if (!FileExist(dlPath))
		return
	unzipPath := a_temp . "\OSDLunzip"
	FileRemoveDir, %unzipPath%, 1
	UnZ(dlPath,unzipPath)
	Loop, %unzipPath%\*.*
	{
		if (a_loopfileext != "nfo") {
			index := ""
			if (fExist:=FileExist(outDir . "\" . outNameNoExt . "." . A_LoopFileExt)){
				while(fExist){
					index++
					fExist := FileExist(outDir . "\" . outNameNoExt . "(" . index ")." . A_LoopFileExt)
				}
			}
			if (index)
				index := "(" . index . ")"
			outPath := outDir . "\" .  outNameNoExt . index "." . A_LoopFileExt
			FileCopy, %A_LoopFileFullPath%, %outPath%
		}
	}
	if (trayTipOn)
		TrayTip, %scriptName%,%outPath%:`nDownload finished.,,1
return

;// Searches Opensubtitles using file hash and a language code
;// fileHash - os hash of the movie to search subtitles for
;// languageCode - http://www.opensubtitles.org/addons/export_languages.php
;// return - CSV with the values in the following format: "subtitleDownloadPath|timesDownloaded|format|language|subReleaseName"
;// eg.: http://dl.opensubtitles.org/download/sub/4132885|2933|srt|English|
searchOS(fileHash,languageCode){
	searchUrl := "http://www.opensubtitles.org/search/sublanguageid-" . languageCode . "/moviehash-" . fileHash . "/rss_2_00"
	ret := HTTPRequest(searchUrl,data:="")
	if (!ret)
		return ret
	needle = <item>(?:\s)*<title>(?:.*?) - ([a-zA-Z]*)(?:.*?) - (?:[a-zA-Z]*?)<\/title>(?:\s*)<link>http://www.opensubtitles.org/(?:.*?)/subtitles/([0-9]*)/(?:.*?)<\/link>(?:\s*)<description>(?:\s*?)(?:Released as: (.*?)`;(?:\s*?))?Format: (.*?)`;(?:\s*?)Uploaded at (?:.*?)(?:\s*?)Download: http://www.opensubtitles.org - ([0-9]*)x(?:\s*)<\/description>
	pos:=1	
	while (pos:=RegExMatch(data, needle, out, pos+1)){
		list .= "http://dl.opensubtitles.org/download/sub/" out2 "|" out5 "|" out4 "|" out1 "|" out3 "`,"
	}
	StringTrimRight, list, list, 1
	if (!list){
		needle = <h1>(?:\s?)<a title="Download" href="/en/subtitleserve/sub/([0-9]*)">(?:.*?) ([a-zA-Z]*) ([a-z]*) subtitles</a>
		pos:=1
		if (RegexMatch(data, needle, out, pos)){
			list := "http://dl.opensubtitles.org/download/sub/" out1 "|" 0 "|" out3 "|" out2 "|"
		}
	}
	return list
}

;//Converts language strings to language codes
lang(inlang="All Languages"){
	var = all|All Languages,alb|Albanian,ara|Arabic,arm|Armenian,baq|Basque,ben|Bengali,bos|Bosnian,bre|Breton,bul|Bulgarian,cat|Catalan,chi|Chinese,cze|Czech,dan|Danish,dut|Dutch,eng|English,epo|Esperanto,est|Estonian,fin|Finnish,fre|French,geo|Georgian,ger|German,glg|Galician,ell|Greek,heb|Hebrew,hin|Hindi,hrv|Croatian,hun|Hungarian,ice|Icelandic,ind|Indonesian,ita|Italian,jpn|Japanese,kaz|Kazakh,khm|Khmer,kor|Korean,lav|Latvian,lit|Lithuanian,ltz|Luxembourgish,mac|Macedonian,may|Malay,nor|Norwegian,oci|Occitan,per|Persian,pol|Polish,por|Portuguese,rus|Russian,scc|Serbian,sin|Sinhalese,slo|Slovak,slv|Slovenian,spa|Spanish,swa|Swahili,swe|Swedish,syr|Syriac,tgl|Tagalog,tha|Thai,tur|Turkish,ukr|Ukrainian,urd|Urdu,vie|Vietnamese,rum|Romanian,pob|Brazilian
	ret:=""
	Loop, Parse, inlang, |
	{
		cinlang := a_loopfield
		Loop, Parse, var, CSV 
		{
			StringSplit, langs, A_LoopField, |
			if (langs2 == cinlang) 
				ret := (ret ? ret "," : "") langs1
		}
	}
	if (!ret)
		return "all"
	return ret
}

/*
Zip/Unzip file(s)/folder(s)/wildcard pattern files
Requires: Autohotkey_L, Windows > XP
URL: http://www.autohotkey.com/forum/viewtopic.php?t=65401
Credits: Sean for original idea
*/
Unz(sZip, sUnz)
{
    fso := ComObjCreate("Scripting.FileSystemObject")
    If Not fso.FolderExists(sUnz)  ;http://www.autohotkey.com/forum/viewtopic.php?p=402574
       fso.CreateFolder(sUnz)
    psh  := ComObjCreate("Shell.Application")
    zippedItems := psh.Namespace( sZip ).items().count
    psh.Namespace( sUnz ).CopyHere( psh.Namespace( sZip ).items, 4|16 )
    Loop {
        sleep 50
        unzippedItems := psh.Namespace( sUnz ).items().count
        ToolTip Unzipping in progress..
        IfEqual,zippedItems,%unzippedItems%
            break
    }
    ToolTip
}

/*
GetOpenSubtitlesHash(FilePath) by just me
http://www.autohotkey.com/board/topic/89313-subtitles-downloader-opensubtitlesorg/?p=565995
*/
GetOpenSubtitlesHash(FilePath) {
   ;http://trac.opensubtitles.org/projects/opensubtitles/wiki/HashSourceCodes
   Static X := { 0: "0",  1: "1",  2: "2",  3: "3",  4: "4",  5: "5",  6: "6",  7: "7"
              ,  8: "8",  9: "9", 10: "A", 11: "B", 12: "C", 13: "D", 14: "E", 15: "F"}
   ; Check the file size ---------------------------------------------------------------------------
   ; 9000000000 > $moviebytesize >= 131072 bytes (changed > to  >= for the lower limit)
   FileGetSize, FileSize, %FilePath%
   If (FileSize < 131072) || (FileSize >= 9000000000)
      Return ""
   ; Read the first and last 64 KB -----------------------------------------------------------------
   VarSetCapacity(FileParts, 131072)         ; allocate sufficient memory
   File := FileOpen(FilePath, "r")           ; open the file
   File.Seek(0, 0)                           ; set the file pointer (just for balance)
   File.RawRead(FileParts, 65536)            ; read the first 64 KB
   File.Seek(-65536, 2)                      ; set the file pointer for the last 64 KB
   File.RawRead(&FileParts + 65536, 65536)   ; read the last 64 KB
   File.Close()                              ; got all we need, so the file can be closed
   ; Now calculate the hash using two UINTs for the low- and high-order parts of an UINT64 ---------
   LoUINT := FileSize & 0xFFFFFFFF           ; store low-order UINT of file size
   HiUINT := FileSize >> 32                  ; store high-order UINT of file size
   Offset := -4                              ; to allow adding 4 on first iteration
   Loop, 16384 {                             ; 131072 / 8
      LoUINT += NumGet(FileParts, Offset += 4, "UInt") ; add first UINT value to low-order UINT
      HiUINT += NumGet(FileParts, Offset += 4, "UInt") ; add second UINT value to high-order UINT
   }
   ; Adjust the probable overflow of the low-order UINT
   HiUINT += LoUINT >> 32                    ; add the overflow to the high-order UINT
   LoUINT &= 0xFFFFFFFF                      ; remove the overflow from the low-order UINT
   ; Now get the hex string, i.e. the hash ---------------------------------------------------------
   Hash := ""
   VarSetCapacity(UINT64, 8, 0)
   NumPut((HiUINT << 32) | LoUINT, UINT64, 0, "UInt64")
   Loop, 8
      Hash .= X[(Byte := NumGet(UINT64, 8 - A_Index, "UChar")) >> 4] . X[Byte & 0x0F]
   Return Hash
}
; ==================================================================================================