Set objfso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject( "WScript.Shell" )



Dim JenkinHome,HomePath, smRunPath ,JenkinStartFolder, JobName

JobName = "PackageDeployment"   		'*********************Enter The Jenkins Job Name here****************				

smRunPath="C:\Program Files (x86)\HPE\Service Manager 9.50\Server\RUN"    '*****************Enter The SM Run diretory Path here****************


HomePath = oShell.ExpandEnvironmentStrings( "%JENKINS_HOME%" )
JenkinHome = Replace(HomePath,"\","\\")
JenkinStartFolder = HomePath & "\workspace\" & JobName




     objStartFolder = JenkinStartFolder 
     Set batFile = objfSO.CreateTextFile(HomePath & "\workspace\" & JobName & "\unload.bat")
	batFile.Write("cd " & smRunPath & vbCrLf)
     Set objFolder = objfso.GetFolder(objStartFolder)
     For Each objFile In objFolder.Files
          If objfso.GetExtensionName(objFile) = "unl" Then
               
		batFile.Write("sm file.load " & chr(34) & JenkinHome &"\\workspace\\" & JobName & "\\"& objFile.Name & chr(34) & " NULL NULL winnt -bg" & vbCrLf)
          End If
     Next


batFile.Close


