Dim sFilename, sMsg, sDate, sYear, sMonth, sDay 
sYear = Year(Now)
sMonth = Month(Now)
sDay = Day(Now)
If sDay < 10 then
	sDay= 0 & sDay
End if
If sMonth < 10 then
	sMonth = 0 & sMonth
End if
sDate= sYear & sMonth & sDay
sFilename= inputbox("Input a file name for the results", "Input filename","Batch " & sDate & "-1") 'Input a file name for the results

Do while sFilename =""  'Check if the file name has been input
	sMsg= MsgBox ("The results' filename is empty!", "53", "Error")
	If sMsg=4 then
		sFilename= inputbox("Please input a file name for the results", "Input filename")
	Elseif sMsg=2 then
		WScript.Echo "Script is End!"
		WScript.quit
	End if
Loop

Function SelectFile( ) ' A function for Selecting files
    Dim objExec, strMSHTA, wshShell
    SelectFile = ""  
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )
    SelectFile = objExec.StdOut.ReadLine( )
    Set objExec = Nothing
    Set wshShell = Nothing
End Function

Function ReadTextFile ' Load the original data
	Dim fso, oFile, sReg
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = fso.OpenTextFile(sNewPath, ForReading)
	'Do While oFile.AtEndOfStream <> True
		Txtfile= oFile.readall	
		sReg="(?:AM|MM|ECO|EN) \d+/[\d\+]*\w+"
		Call RegExpTest(sReg, Txtfile) 
	'Loop
	set fso = nothing
	set oFile = nothing   
End Function

Function RegExpTest(patrn, str)  ' use the regular expression to match the target information and calculate 
	Dim regEx, Match, Matches, label  
	Set regEx = New RegExp  
	regEx.Pattern = patrn   
	regEx.IgnoreCase = True   
	regEx.Global = True   
	If regEx.Test(str) = True Then
		Set Matches = regEx.Execute(str)   
		For Each Match in Matches   
			label = Match
			regEx.Pattern = "^aging"		
			If regEx.Test(str) = True Then
				sReslut = sReslut & label & "A" & vbcrlf & label & "A+D" & vbcrlf
			Else
				sReslut = sReslut & label & vbcrlf
				sReslut=Replace(sReslut, "+", "-")
			End If
			regEx.Pattern = patrn
		Next
	End if
	set Matches=Nothing
End function

Function Word_to_Txt(Path) 'Save doc file as text file
	Dim oWord, oDic
	Set oWord = CreateObject("Word.Application")
	Set oDoc = oWord.Documents.Open(sPath)
	oDoc.SaveAs sPath&".txt", wdFormatText
	oWord.Quit
	Set oWord = Nothing
	Set oDoc = Nothing
End Function

Function OpenExcel(Path)
	Dim oExcel
	Set oExcel = CreateObject("excel.application")
	oExcel.workbooks.open(Path)
	oExcel.visible = true
	oExcel.worksheets(1).activate
	Set oExcel =Nothing
End Function

Dim fso, oRefile, sReslut, sNewPath, oTxt
Const ForReading = 1, ForWriting = 2, wdFormatText = 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set oRefile = fso.OpenTextFile("D:\Batch\" & sFilename & ".csv", ForWriting, True)
sPath=SelectFile( )
Call Word_to_Txt(sPath)
sNewPath=sPath & ".txt"
Call ReadTextFile
oRefile.WriteLine sReslut
oRefile.close
Set oTxt=fso.getfile(sNewPath)
oTxt.Delete
Set oTxt = Nothing
Set fso =Nothing
Set oRefile=Nothing
WScript.Echo "The labels have been transferred to D:\Batch\.  Script is Finished!"
Call OpenExcel("D:\Batch\" & sfilename & ".csv")
'Any question, pls. contact Justin.Thanks.