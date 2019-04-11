write-host "Please choose a directory to store the script log"
$awesomesauce = "
__________     ___.         .__ __        _____                   __                _________ .__                        
\______   \__ _\_ |_________|__|  | __   /     \ _____    _______/  |_  ___________ \_   ___ \|  | _____    ______ ______
 |       _/  |  \ __ \_  __ \  |  |/ /  /  \ /  \\__  \  /  ___/\   __\/ __ \_  __ \/    \  \/|  | \__  \  /  ___//  ___/
 |    |   \  |  / \_\ \  | \/  |    <  /    Y    \/ __ \_\___ \  |  | \  ___/|  | \/\     \___|  |__/ __ \_\___ \ \___ \ 
 |____|_  /____/|___  /__|  |__|__|_ \ \____|__  (____  /____  > |__|  \___  >__|    \______  /____(____  /____  >____  >
        \/          \/              \/         \/     \/     \/            \/               \/          \/     \/     \/ 
"
function ChooseFolder([string]$Message, [string]$InitialDirectory)
{
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, 0, $InitialDirectory)
    $selectedDirectory = $folder.Self.Path 
    return $selectedDirectory
}
$logfolder = ChooseFolder -Message "Please select a log file directory" -InitialDirectory 'MyComputer' 
$logfile = $logfolder + '\' + (Get-Date -Format o |ForEach-Object {$_ -Replace ':', '.'}) + "Guardians.txt"
write-host "Script result log can be found at $logfile" -ForegroundColor Green

if ( !(Get-InstalledModule -Name Rubrik -ErrorAction SilentlyContinue) )
    {
        write-host ("Rubrik Module not installed. Please verify installation and retry.") -BackgroundColor Red
        write-host "Terminating Script" -BackgroundColor Red
        add-content $logfile ("Rubrik Module not found. Please verify installation and retry.")
        add-content $logfile "You can install the module by running : Install-Module -Name Rubrik"
        return
    }
write-host "Getting Credentials from user prompt" -ForegroundColor Green
add-content $logfile "Getting Credentials from user prompt"
$Credentials = Get-Credential
$RubrikClusterIP = read-host "Please enter a Rubrik Cluster IP or FQDN"
try
{
    Connect-Rubrik -Server $RubrikClusterIP -Credential $Credentials -ErrorAction Stop |out-null
    add-content $logfile ('Connected to Rubrik Cluster at ' + $RubrikClusterIP)
    add-content $logfile '----------------------------------------------------------------------------------------------------'
}
catch
{
    write-host "Failed to connect to Rubrik Cluster" -BackgroundColor Red
    write-host $RubrikClusterIP
    write-host $Error[0]
    write-host "Terminating Script" -BackgroundColor Red
    add-content $logfile "Failed to connect to Rubrik Cluster"
    add-content $logfile $RubrikClusterIP
    add-content $logfile $Error[0]
    add-content $logfile "Terminating Script"
    return
}
#Get the report data from the report created earlier
$reportname = "Rubrik Protection Summary"
$reportcheck = Get-RubrikReport -name $reportname
#check the report exists
if ($reportcheck.total -eq 0) {
	write-host "Failed to find the Report data" -BackgroundColor Red
	add-content $logfile ("Failed to find the Report data")
	return
}
#export the report
$CSV = Get-RubrikReport -name $reportname | Export-RubrikReport
Invoke-WebRequest $CSV -OutFile "c:\temp\protectionreporting.csv"
write-host $awesomesauce
write-host "Your report data has been refreshed please load the powerBI report for some #awesomesauce"