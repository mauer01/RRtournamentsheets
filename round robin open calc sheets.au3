#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Program Files (x86)\AutoIt3\Icons\au3.ico
#AutoIt3Wrapper_Outfile=round robin open calc sheets.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         Mauer01

 Script Function:
	Making Open Calc Sheets for 2 game Round Robin Tournaments

	Textfile needs to be formated like this, for now.

	:*tournament1*
	*player1*
	*player2*
	*player3*
	...
	:*tournament2*
	*player1*
	*player2*
	*player3*
	...
	:...

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include "OOoCalc.au3"
#include "Array.au3"
#include "MsgBoxConstants.au3"

dim $playernames[0]



if $cmdline[0] = 1 Then
	$file = FileReadToArray($cmdline[1])
	If @error then Exit
Else
	$file = FileReadToArray(FileOpenDialog("choose a textfile",@ScriptDir,"Text (*.txt)"))
	If @error then
		$file = ClipGet()
		if stringleft($file,1) <> ":" then
			$file = ":" & InputBox("Choose Tournament Name","If you dont hit cancel names will get copied out of your clipboard","Standard") & @CRLF & $file
			if @error then Exit

		EndIf
		$file = StringSplit($file,@CRLF,3)
		_ArrayDisplay($file)
	EndIf
EndIf


$filedir = @ScriptDir & "\"
$ods = $filedir & @hour & @min & @sec & "tournament.ods"
$HNDL_ODS = _OOoCalc_BookNew(True)
OnAutoItExitRegister("_exit")

_OOoCalc_SheetDelete($HNDL_ODS,1)
_OOoCalc_SheetDelete($HNDL_ODS,1)
dim $tournaments[1][3] = [[0,0,0]]
$linemarker = 0
If StringLeft($file[$linemarker],1) <> ":" Then
	MsgBox(0,"","file not correctly setuped")
	Exit
EndIf


While $linemarker < UBound($file)
	$tournament_name = StringTrimLeft($file[$linemarker],1)
	$linemarker += 1
	dim $playernames[0], $players[1] = [0]
	Do
		_ArrayAdd($playernames,$file[$linemarker])
		$linemarker += 1
	Until  ($linemarker = UBound($file) or StringLeft($file[$linemarker],1) = ":" )
	Dim $players[1] = [0]
	for $i = 0 to UBound($playernames) -1
		_ArrayAdd($players,"A" & 4+$i)
		$players[0] += 1
	Next
	dim $dummyarray[1][3] = [[$tournament_name,$players,$playernames]]

	_ArrayAdd($tournaments,$dummyarray)
	$tournaments[0][0] += 1
WEnd

For $index = 1 to $tournaments[0][0]
	$tournament_name = $tournaments[$index][0]
	$players = $tournaments[$index][1]
	$playernames = $tournaments[$index][2]
	constructsheet($HNDL_ODS,$tournament_name,$players,$playernames)

Next


_OOoCalc_SheetDelete($HNDL_ODS,1)
_OOoCalc_BookSaveAs($HNDL_ODS,$ods)
_OOoCalc_BookClose($HNDL_ODS)
OnAutoItExitUnRegister("_exit")
Exit



Func constructsheet($HNDL_ODS,$tournament_name,$players,$playernames)
	_OOoCalc_SheetAddNew($HNDL_ODS,$tournament_name)
	_OOoCalc_SheetActivate($HNDL_ODS,$tournament_name)
	_OOoCalc_WriteCell($HNDL_ODS,$tournament_name,"A1")
	_OOoCalc_WriteCell($HNDL_ODS,"Players","A3")
	_OOoCalc_WriteCell($HNDL_ODS,"White","G2")
	_OOoCalc_WriteCell($HNDL_ODS,"Black","I2")
	_OOoCalc_WriteCell($HNDL_ODS,"Result:","H3")
	_OOoCalc_WriteCell($HNDL_ODS,"Any links (YT/Twitch with when you play/chessin5d)","F2")
	dim $row = ["Ranking:","Name","Points","Sonneborn-Berger","","","","","","Score","Name","Points","SB","Tied points","Black games"]
	_OOoCalc_WriteRowFromArray($row,10,2)
	dim $row = ["'1.","=AB4","=AC4","=AD4","","=VLOOKUP(C4;$U$4:$V$" & $players[0]+3 &";2;0)*I4/2","","=VLOOKUP(E4;$U$4:$V$" & $players[0]+3 &";2;0)*G4/2","","=V4*1000000+W4*10000+1/ROW()","=A4","=SUMIF(C$1:C$1048576;U4;G$1:G$1048576)+SUMIF(E$1:E$1048576;U4;I$1:I$1048576)","=SUMIF(C$1:C$1048576;U4;R$1:R$1048576)+SUMIF(E$1:E$1048576;U4;P$1:P$1048576)","","=SUMIF(E$1:E$1048576;U4;I$1:I$1048576)","","=LARGE($T$4:$T$" & $players[0]+3 &";ROW()-3)","=VLOOKUP($AA4;$T$4:$Y$" & $players[0]+3 &";COLUMN()-26;0)","=VLOOKUP($AA4;$T$4:$Y$" & $players[0]+3 &";COLUMN()-26;0)","=VLOOKUP($AA4;$T$4:$Y$" & $players[0]+3 &";COLUMN()-26;0)","=VLOOKUP($AA4;$T$4:$Y$" & $players[0]+3 &";COLUMN()-26;0)"]
	_OOoCalc_WriteRowFromArray($row,10,3)
	_OOoCalc_WriteCell($HNDL_ODS,'=IF(INT(AA5)=INT(AA4);K4;ROW()-3&".")',"K5")
	_OOoCalc_DraggingDown("K5:K5",UBound($playernames)-2)
	For $i = 1 to 17
	_OOoCalc_ColumnSetProperties($HNDL_ODS,13+$i,2275,False,False)
	Next
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,2,0,3,9000)
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,2,0,6,9000,8)
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,2,0,10,9000)
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,3,0,12,9000,13)
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,1,"N3")
	_OOoCalc_HorizontalAlignSet($HNDL_ODS,1,"M3")
	_OOoCalc_FontSetProperties($HNDL_ODS,"A1",-1,-1,-1,-1,True,False,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"F2",-1,-1,-1,-1,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"H3",-1,-1,-1,-1,True,False,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"K3",-1,-1,-1,-1,True,False,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"L3",-1,-1,-1,-1,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"M3",-1,-1,-1,-1,True)
	_OOoCalc_FontSetProperties($HNDL_ODS,"N3",-1,-1,-1,-1,True)
	_OOoCalc_ColumnSetProperties($HNDL_ODS,5,1000,True,True)
	_OOoCalc_ColumnSetProperties($HNDL_ODS,13,1000,True,True)

	GenerateRoundRobinSchedule($players)
	_ArrayShuffle($playernames)

	For $i = 0 to $players[0]-1
		_OOoCalc_WriteCell($HNDL_ODS,$playernames[$i],"A" & 4 + $i)
	Next

EndFunc

Func _OOoCalc_WriteRowFromArray($row,$x,$y)
	For $i = 0 to UBound($row)-1
		_OOoCalc_WriteCell($HNDL_ODS,$row[$i],$y,$x+$i)
	Next
EndFunc

Func _OOoCalc_WriteRowFromArray2($row,$x,$y)
	For $i = 0 to UBound($row)-1
		_OOoCalc_WriteFormula($HNDL_ODS,$row[$i],$y,$x+$i)
	Next
EndFunc



Func _OOoCalc_DraggingDown($range,$count)
	local $row,$startcol,$endcol,$rangewidth,$i = 0,$colmarker
	$copyrange = _ArrayFromString($range,"")
	while $copyrange[$i] = 0
		$startcol &= $copyrange[$i]
		$i += 1
	WEnd
	do
		$row &= $copyrange[$i]
		$i += 1
	Until $copyrange[$i] = ":"
	$i += 1
	while $copyrange[$i] = 0

		$endcol &= $copyrange[$i]
		$i += 1
	WEnd
	$startcoln = Twentysixintodec($startcol)
	$endcoln = Twentysixintodec($endcol)
	$rangewidth = $endcoln - $startcoln + 1
	For $y = 1 to $count
		$colmarker = $startcol
		For $x = 1 to $rangewidth
			_OOoCalc_RangeMoveOrCopy($HNDL_ODS,$colmarker & $row,$colmarker & $row+$y, 1)
			$colmarker = Alphabeticalcountup($colmarker)
		Next
	Next
EndFunc

Func Alphabeticalcountup($letters)
	$array = _ArrayFromString($letters,"")
	$c = UBound($array)-1
	If String($array[$c]) <> "Z" Then
		$array[$c] = AlphabeticalNumberToLetter(LetterToAlphabeticalNumber($array[$c])+1)
	ElseIf $array[$c] = "Z" Then
		_ArrayDelete($array,$c)
		if UBound($array) = 0 Then
			Return "AA"
		EndIf
		$array = _ArrayFromString(Alphabeticalcountup(_ArrayToString($array,""))&"A","")
	EndIf
	Return _ArrayToString($array,"")
EndFunc

Func GenerateRoundRobinSchedule($teams)
    Local $numTeams = $teams[0]
	_ArrayDelete($teams,0)
    Local $schedule = []
    ; If odd number of teams, add a dummy team to make it even
    If Mod($numTeams, 2) <> 0 Then
        _ArrayAdd($teams, "BYE")
        $numTeams += 1
    EndIf

    ; Number of rounds
    Local $numRounds = $numTeams - 1
	$marker = 2
    For $round = 0 To $numRounds - 1
        _OOoCalc_WriteCell($HNDL_ODS,"Round" & $round+1,"C" & $marker)
		_OOoCalc_FontSetProperties($HNDL_ODS,"C"&$marker,-1,-1,-1,-1,True,False,True)
		$marker += 2
        For $match = 0 To ($numTeams / 2) - 1
            Local $home = Mod(($round + $match),($numTeams - 1))
            Local $away = Mod(($numTeams - 1 - $match + $round),($numTeams - 1))
            ; Fix the position of the last team
            If $match = 0 Then
                $away = $numTeams - 1
            EndIf
			If $teams[$away] <> "bye" Then
				_OOoCalc_WriteFormula($HNDL_ODS,"=" & $teams[$home], "C" & $marker)
				_OOoCalc_WriteCell($HNDL_ODS,"vs.", "D" & $marker)
				_OOoCalc_WriteFormula($HNDL_ODS,"=" & $teams[$away], "E" & $marker)
				_OOoCalc_WriteCell($HNDL_ODS,":", "H" & $marker)
				$marker += 1
				_OOoCalc_WriteFormula($HNDL_ODS,"=" & $teams[$away], "C" & $marker)
				_OOoCalc_WriteCell($HNDL_ODS,"vs.", "D" & $marker)
				_OOoCalc_WriteFormula($HNDL_ODS,"=" & $teams[$home], "E" & $marker)
				_OOoCalc_WriteCell($HNDL_ODS,":", "H" & $marker)
				$marker += 2
			EndIf
        Next
    Next
	_OOoCalc_DraggingDown("L4:AE4",UBound($playernames)-1)
	_OOoCalc_DraggingDown("P4:R4",$marker-4)
EndFunc

func _exit()
	_OOoCalc_BookSaveAs($HNDL_ODS,$ods)
	_OOoCalc_BookClose($HNDL_ODS)
endfunc

Func AlphabeticalNumberToLetter($number)
	If $number < 0 then
		; Invalid input, not a valid alphabetical number
		Return ""
	EndIf
	local $returnstring
	If $number > 26 Then
		$newnumber = Int($number/26)

	EndIf
	$returnstring &= Chr($number + Asc("A") - 1)
	Return $returnstring
EndFunc


func reducenumber($number,$c)
	local $count

	While $number > $c
		$count += 1
		$number -= $c
	WEnd
	If $count = 0 Then
		Return $number
	ElseIf $count < $c Then
		Return $count & $number
	ElseIf $count > $c Then
		Return reducenumber($count,$c) & $number
	EndIf
EndFunc

Func LetterToAlphabeticalNumber($letter)
	$letter = StringUpper($letter) ; Convert to uppercase to handle lowercase letters

	If StringRegExp($letter, "[A-Z]") = 0 Then
		; Invalid input, not a letter
		Return 0
	EndIf

	Return Asc($letter) - Asc("A") + 1
EndFunc


Func Twentysixintodec($number)
	local $returnstring = 0,$i = 1
	$len = StringLen($number)
	For $i = 1 to $len

		$returnstring += LetterToAlphabeticalNumber(StringMid($number,$i,1))*26^($len-$i)

	Next
	Return $returnstring
EndFunc
