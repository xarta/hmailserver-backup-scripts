' just to encapsulate some objects
' doing this way so I can pass the object around
' ... and can extend it in the future hopefully with no side-effects
Class XartaComputer
	Private objShell
	Private objEnv
	Private strComputer

	Public Property Get HostName()
		HostName = strComputer
	End Property

	Private Sub Class_Initialize()
		Set objShell = CreateObject("Wscript.Shell")
		Set objEnv = objShell.Environment("Process")
		strComputer = objEnv("COMPUTERNAME")
	End Sub
End Class