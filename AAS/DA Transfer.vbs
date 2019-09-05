Dim Resulname, ResulpoMsg, Syear, Smonth, Sday
Syear = Year(Now)
Smonth = Month(Now)
Sday = Day(Now)
If Sday < 10 then
	Sday= 0 & Sday
End if
If Smonth < 10 then
	Smonth = 0 & Smonth
End if 
mydate= Syear & Smonth & Sday
Resulname= inputbox("Input a file name for the Transfer Results", "Input filename",mydate) 'Input a file name for the Results

Do while Resulname =""  'Check if the file name has been input
 ResulpoMsg=MsgBox("The Results' filename is empty!", "53", "Error")
 If ResulpoMsg=4 then
  Resulname= inputbox("Please input a file name for the Results", "Input filename")
 Elseif ResulpoMsg=2 then
  WScript.Echo "Script is End!"
  WScript.quit
 end if
Loop

Function SelectFile( ) ' A function for Selecting files
    Dim objExec, strMSHTA, wshShell
    SelectFile = ""  
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();ResulizeTo(0,0);" & "<" & "/script>"""
    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )
    SelectFile = objExec.StdOut.ReadLine( )
    Set objExec = Nothing
    Set wshShell = Nothing
End Function

Dim dataFile, Txtfile, Dat, calnum
Dim ADat, Mydate, Mytime
ADat="Solution Label      	Type	Element             	Flags   	Soln Conc 	Units   	Corr Con  	Units   	Int     	Date    	Time     "
Mydate= Date
Mytime= Time
M=Msgbox("Select Result Files","0", "SelectFiles")
dataFile = SelectFile( )
If dataFile = "" Then 
    WScript.Echo "No file selected. Script is End!"
	Wscript.quit
Else
    Call ReadTextFile
	' Open A txt file for input the Results
	Dim fso, reFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set reFile = fso.OpenTextFile("D:\" & Resulname & ".txt", ForWriting, True)
	'For i= 0 to calnum
    reFile.Write ADat
    'Next
	reFile.Close
	WScript.Echo "The Results have been output.  Script is Finished!"
End If

Const ForReading = 1, ForWriting = 2
Function ReadTextFile ' Load the original data
   Dim fso, MyFile,reg
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set MyFile = fso.OpenTextFile(dataFile, ForReading)
   Do While MyFile.AtEndOfStream <> True
      Txtfile= MyFile.readline
	  reg="(Pb|Cd|Ni)((?:\w+\s+)?\d+/[\d\s-=]*\w?\d)\s+(-?\d+\.\d+)\s+mg/L"
	  Call RegExpTest(reg, Txtfile) 
   Loop 
End Function

Function RegExpTest(patrn, strng)  ' use the regular expression to match the target information 
   Dim regEx, Match, Matches  
   Set regEx = New RegExp  
   regEx.Pattern = patrn   
   regEx.IgnoreCase = True   
   regEx.Global = True   
   If regEx.Test(strng) = True Then
    Set Matches = regEx.Execute(strng)   
    For Each Match in Matches
	'Set Match = Matches(0)
	Call TransferForm(Match.SubMatches(0), Match.SubMatches(1), Match.SubMatches(2))
	
	ADat =ADat & vbCrlf & Dat 
	Next
   End if
End Function
  
Function TransferForm(Elem, Projt, Resul)
  
 If Elem = "Pb" then
     Projt=Replace(Projt, "=", "+")
	 Projt=Replace(Projt, "-", "+")
	 Dat= Projt & vbTab & vbTab & "Pb 220.353" & vbTab & vbTab & Resul & vbTab & "mg/L" & vbTab & vbTab & vbTab & vbTab& "10/16/2014" & vbTab & "11:26:05"
 elseif Elem = "Cd" then
     Projt=Replace(Projt, "=", "+")
	 Projt=Replace(Projt, "-", "+")
	 Dat= Projt & vbTab & vbTab & "Cd 228.802" & vbTab &  vbTab  & Resul & vbTab & "mg/L" & vbTab & vbTab & vbTab & vbTab & "10/16/2014" & vbTab & "11:26:05"
 elseif Elem = "Ni" then
     Projt=Replace(Projt, "=", "+")
	 Projt=Replace(Projt, "-", "+")
	 Dat= Projt & vbTab & vbTab & "Ni 216.555" & vbTab &  vbTab  & Resul & vbTab & "mg/L"& vbTab & vbTab & vbTab & vbTab & "10/16/2014" & vbTab & "11:26:05"
 End if
End Function
' Update1 change the regex to match Project NO. like MM  12345/1+2+3
' 20160414 Update2 change the regex to match Project NO. like 12345/1+2+3 
' Any question, pls. contact Justin, thanks~