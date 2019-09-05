set Wshshell =WScript.CreateObject("WScript.Shell")
With Wshshell
	n = inputbox("Enter the number of times of iteration", "Enter a number",40)
	recovery = inputbox("Enter the number of input", "Enter a number", 87.5 & "{%}")
	Wscript.sleep 5000
	For i = 1 to n
		.Sendkeys recovery 
		.Sendkeys "{Down}"
	Next
End With
