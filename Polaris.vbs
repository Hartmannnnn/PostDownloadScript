' Verione 1.0 da usare 

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
'
' WshShell.Run "net use \\" & IpControllo & "\d SYS2000 /USER:Administrator > NUL 2>&1"
' fso.CopyFile pathPlc & ".cod" , "\\" & IpControllo & pathPLCsystem, True
' WshShell.Run "net use " & IpControllo & "\d /DELETE > NUL 2>&1"
'
' MsgBox("File Copiati")
 set WshShell = WScript.CreateObject("WScript.Shell")
 
 x=MsgBox("Copiare il file PLC sul controllo?",4)
 
 ' Rispondo Si
 IF x = 6 Then
	 'Avvio timer per capire quanto ci metto a scaricare progetto
	 StartTime = Timer()
     ' Prelevo IP controllo da progetto PLC
     IpControllo = WshShell.ExpandEnvironmentStrings("%PRJCONN%")
     IpControllo = Mid(IpControllo, InStr(5, IpControllo, ":") + 1, 20)

     ' Cerco fine stringa  IP (può finire con "_" o con "/")
     posfine = InStr(1, IpControllo, "_")
     if posfine <= 0 then
         posfine = InStr(1, IpControllo, "/")
     end if
     
     if posfine > 0 then
         IpControllo = left(IpControllo,PosFine - 1)
     end if

     ' Compongo nome e percorso file PLC
     nomePlc = WshShell.ExpandEnvironmentStrings("%PRJTITLE%")
     pathPlc = WshShell.ExpandEnvironmentStrings("%PRJPATH%")
     ' Verifico che stringa termini con "\"
     If Right(pathPlc,1) <> "\" Then
         pathPlc = pathPlc & "\"
     End If
    
     pathPlcName = pathPlc & nomePlc

	 'MsgBox("net use \\" & IpControllo & "\d SYS2000 /USER:Administrator")
     ' Mappo l'unità di rete e attendo fine mappatura
     WshShell.RUN "net use \\" & IpControllo & "\d SYS2000 /USER:Administrator",1, TRUE
     
     ' Copio file .PLC e .COD
     Set fso = CreateObject("Scripting.FileSystemObject")

	' Setto il percorso dove andrà il plc nel controllo
	 pathPLCsystem = "\d\Polaris\plc\"

'     MsgBox(pathPlc & ".plc")
'     MsgBox("\\" & IpControllo & pathPLCsystem)
	 If fso.FileExists(pathPlcName & ".plcprj") Then
		fso.CopyFile pathPlcName & ".plcprj" , "\\" & IpControllo & pathPLCsystem, True
	 End If

	 ' Verifico se sul controllo esiste la cartella /Build (LogicLab 5.12 e successivi) altrimenti la creo
	 ' Se sul controllo ci sono versioni precedenti, troveranno il PLC modificato e ricompileranno ignorando la cartella/Build.
	 If NOT fso.FolderExists("\\" & IpControllo & pathPLCsystem & "Build") Then
		fso.CreateFolder("\\" & IpControllo & pathPLCsystem & "Build")
	 End if

	 fso.CopyFile pathPlc & "Build\" & nomePlc & ".cod" , "\\" & IpControllo & pathPLCsystem & "Build\", True
	 fso.CopyFile pathPlc & "Build\" & "*.xml" , "\\" & IpControllo & pathPLCsystem & "Build\", True

	 ' Per LogicLab 5.16 e successivi con progetto a cartelle
	If fso.FolderExists(pathPlc & "src") Then
		'Se sul controllo esiste la cartella /src la cancello per eliminare eventuali file eliminati in locale
		 If fso.FolderExists("\\" & IpControllo & pathPLCsystem & "src") Then
			 fso.DeleteFolder "\\" & IpControllo & pathPLCsystem & "src", true
		 End if

			 ' Copio la cartella locale \src sul controllo
		 fso.CopyFolder pathPlc & "src", "\\" & IpControllo & pathPLCsystem, True
		 
		'Se sul controllo esiste la cartella /Library la cancello per eliminare eventuali file eliminati in locale
		 If fso.FolderExists("\\" & IpControllo & pathPLCsystem & "Library") Then
			 fso.DeleteFolder "\\" & IpControllo & pathPLCsystem & "Library", true
		 End if

		 'Se ho la cartella /Library nel progetto locale la copio
		 If fso.FolderExists(pathPlc & "Library") Then
			 ' Copio la cartella locale \Library sul controllo
			 fso.CopyFolder pathPlc & "Library", "\\" & IpControllo & pathPLCsystem, True
		 End if

		'Se sul controllo esiste la cartella /config la cancello
		 If fso.FolderExists("\\" & IpControllo & pathPLCsystem & "config") Then
			 fso.DeleteFolder "\\" & IpControllo & pathPLCsystem & "config", true
		 End if	 

		 'Se ho la cartella /config nel progetto locale la copio
		 If fso.FolderExists(pathPlc & "config") Then
			 ' Copio la cartella locale \config sul controllo
			 fso.CopyFolder pathPlc & "config", "\\" & IpControllo & pathPLCsystem , True
		 End if

	End if

	 fso.CopyFile pathPlc & "*.pll" , "\\" & IpControllo & pathPLCsystem, True
	 fso.CopyFile pathPlc & "*.exp" , "\\" & IpControllo & pathPLCsystem, True
     
	 ' Copio i file plclib solo se esistono altirmenti ho un errore
	 fileExists = False
	 Set folder = fso.GetFolder(pathPlc)
	 For Each file In folder.Files
	 	If LCase(fso.GetExtensionName(file.Name)) = "plclib" Then
	 		fileExists = True
	 		Exit For
	 	End If
	 Next
	 if fileExists Then
	 	fso.CopyFile pathPlc & "*.plclib" , "\\" & IpControllo & pathPLCsystem, True
	 End if

     ' Smappo l'unità di rete
     WshShell.RUN "net use \\" & IpControllo & "\d /DELETE"
     'objNetwork.RemoveNetworkDrive "Q:"

     EndTime = Timer()
     MsgBox("File Copiati, Tempo impiegato: " & FormatNumber(EndTime - StartTime, 3) & " s")
 End IF

