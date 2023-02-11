$path = @()

#Echo enter file path
Write-Output "Enter the file paths to pin to Quick Access, "done" to end list."

while ($input = Read-Host) {
    if ($input -eq "done") 
    { break }

    $path += $input
}

foreach ($filePath in $path) {
    if (Test-Path $filePath) {
        #iterate command for array
        $shell = New-Object -ComObject Shell.Application
        
        $folder = $shell.Namespace((Split-Path $filePath))
        
        #NameSpace(11) - reference to Msft Quick Access
        $shell.NameSpace(11).Self.InvokeVerb("Pin to Quick access")
    } else {
        #error if does not exist
        Write-Output "The file path '$filePath' does not exist."
    }
}
