
Param (
    [string] $From ="", # Path que tenia el acceso directo
    [string] $To ="",   # Path que tenia el acceso directo
    [string] $ProgramName ="", # Nombre del programa
    [string] $Destination =""  # path donde se guarda el acceso directo
)
$UserPath = Get-ChildItem -Path C:\Users -Directory | select Name; 

Foreach ($i in $UserPath){ 

        $files = Get-ChildItem -Path "C:\Users\$($i.Name)\$($Destination)\" -Include $ProgramName -File -Recurse
        Write-Output "Files: $($files)"
        $obj = New-Object -ComObject WScript.Shell 

    
        ForEach($file in $files){ 
	    $link = $obj.CreateShortcut($file) 
        Write-Output "File: $($file)"      
	    Write-Host $link.TargetPath -ForegroundColor green
	    $link.TargetPath = $link.TargetPath.Replace($From,$To)
	    $link.Save() 
        } 
 }