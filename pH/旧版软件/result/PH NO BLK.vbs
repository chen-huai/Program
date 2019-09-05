set Wshshell =WScript.CreateObject("WScript.Shell")
With Wshshell
	pH_1 = inputbox("Enter the number of input pH 1", "Enter a number")
	pH_2 = inputbox("Enter the number of input pH 2", "Enter a number")
	LabTemp = 23.2
	SolTemp = 23.9
	Wscript.sleep 1000
	.Sendkeys LabTemp 
	.Sendkeys "{RIGHT}"
	.Sendkeys SolTemp 
	.Sendkeys "{UP 2}"
	Wscript.sleep 500
	.Sendkeys "{LEFT 9}"
	.Sendkeys pH_1
	.Sendkeys "{DOWN}"
	.Sendkeys pH_2
	.Sendkeys "{DOWN 3}"
	.Sendkeys 0

End With
