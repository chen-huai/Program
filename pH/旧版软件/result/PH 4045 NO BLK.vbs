set Wshshell =WScript.CreateObject("WScript.Shell")
With Wshshell
	pH_1 = inputbox("Enter the number of input pH 1", "Enter a number")
	pH_2 = inputbox("Enter the number of input pH 2", "Enter a number")
	pH = inputbox("Enter the number of input pH-", "Enter a number")
	LabTemp =  24.0
	SolTemp = 25.5
	Wscript.sleep 1000
	.Sendkeys LabTemp 
	.Sendkeys "{RIGHT}"
	.Sendkeys SolTemp 
	.Sendkeys "{UP 3}"
	Wscript.sleep 500
	.Sendkeys "{LEFT 10}"
	.Sendkeys "{DOWN}"
	.Sendkeys pH_1
	.Sendkeys "{DOWN}"
	.Sendkeys pH_2
	.Sendkeys "{DOWN 3}"
	.Sendkeys pH

End With
