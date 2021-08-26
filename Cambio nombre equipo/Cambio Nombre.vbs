'Script para modificar el nombre del equipo Windows
'Autor: Jose Conde 

'Variables con interación del usuario
Nombre = INPUTBOX("Introduce un nombre nuevo para tu equipo")
Usuario = INPUTBOX("Nombre del usuario")
Contra = INPUTBOX("Contraseña del usuario")


'Creamos in instancia WMI para consultar información
Set objWMIService = GetObject("Winmgmts:root\cimv2")


'Código para modificar nombre equipo
For Each objComputer in _
    objWMIService.InstancesOf("Win32_ComputerSystem")

        nombre_nuevo = objComputer.rename(Nombre,Usuario,Contra)
        If nombre_nuevo <> 0 Then
           WScript.Echo "Error al modificar el nombre. Error = " & Err.Number

        Else
	   respuesta = INPUTBOX("Quieres reiniciar ahora para aplicar los cambios? (s/n)") 
	    
	   'Aplicar cambios de nombre
           If (respuesta = "s") Then
		Set objShell = CreateObject("WScript.Shell")
		strCommand = "cmd /c shutdown -r -t 0"
		objShell.Run strCommand,vbhide
	   Else 
	       WScript.Echo "[!] Okay, no olvides reiniciarlo luego!"
           End If

	End If

Next