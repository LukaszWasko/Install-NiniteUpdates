#requires -Version 2.0

<#
.SYNOPSIS

 The script takes care of the NinitePro application with the scheduler help.


.DESCRIPTION

  The script can update the local copy of the application ninitepro.exe.
  The script will select the fastest cache location (if many will be delivered) or will use Ninite site if caches are unavailable.
  The script writes its own logs and installed applications.
  With support for multiple parameters, the script works best with Scheduler.
  Scipt will be independent of the network infrastructure when you first create chedler task and copying ninitepro.exe application to local disk (for example with GPO).
  Task schedule is best to configure this by this website: https://brashear.me/blog/2015/01/23/deploy-application-updates-with-ninite-with-a-cache-on-your-lan/


.PARAMETER NiniteSourcePath

  Optional network path Ninite application in the FQDN UNC format, which will be used for optional updates the local copy of the application


.PARAMETER Apps

  The list of applications to install / update between the two quotes. Divide applications by spacebar. If the application has its own quotes, before each of them put a backslash.


.PARAMETER CachePaths

  Optional ninite cache network locations in UNC FQDN format. The script selects the fastest. Locations divided by spacebar. All locations should be placed between quotation marks.


.PARAMETER UpdateOnly

  This parameter is optional. Enter the parameter without value if you just want to update existing applications.


.PARAMETER RootDir

  Optional local work path, which will be placed ninitepro.exe and logs.


.PARAMETER DeletedFiles

  Enter the number of days after which logs will be removed.


.NOTES
  Version:       1.0
  Author:        Lukasz Wasko
  Creation Date: 2016.12.23


.EXAMPLE

  .\Install-NiniteUpdates.ps1 -apps "Firefox Flash \"Flash (IE)\" Chrome Opera Greenshot Pidgin" -NiniteSourcePath "\\dcsrv01.contoso.com\netlogon\Ninite\NinitePro.exe" -CachePaths "\\fssrv01.contoso.com\Ninite\NiniteDownloads\files \\fssrv02.contoso.com\Ninite\NiniteDownloads\files"


.EXAMPLE
  .\Install-NiniteUpdates.ps1 -apps "\"Flash (IE)\""
#>

[CmdletBinding()]
Param (
    [Parameter(HelpMessage='Enter the ninitepro UNC path in FQDN format, which be used to updating local ninitepro file. `nExample: -NiniteSourcePath \\dcsrv01.contoso.com\netlogon\Ninite\NinitePro.exe')]
    [ValidateNotNullOrEmpty()]
    [string] $NiniteSourcePath = '\\dcsrv01.contoso.com\netlogon\Ninite\NinitePro.exe',
    
    [Parameter(Mandatory = $True,
               HelpMessage='The list of applications to install / update between the two quotes. Divide applications by spacebar. If the application has its own quotes, before each of them put a backslash. `nExample: -Apps "firefox chrome \"Flash (IE)\" aimp"')]
    [ValidateNotNullOrEmpty()]
    [string] $Apps,
    
    [Parameter(HelpMessage='Optional ninite cache network locations in UNC FQDN format. The script selects the fastest. Locations divided by spacebar. All locations should be placed between quotation marks. `nExample: -CachePaths "\\fssrv01.contoso.com\Ninite\NiniteDownloads\files \\fssrv02.contoso.com\Ninite\NiniteDownloads\files"')]
    [string] $CachePaths,
    
    [Parameter(HelpMessage='This parameter is optional. Enter the parameter without value if you just want to update existing applications. `nExample: -updateonly')]
    [switch] $UpdateOnly = $false,
    
    [Parameter(HelpMessage='Optionally enter the local working path where the application ninitepro and log files will be placed. `nExample: -rootdir "c:\Tools\Ninite"')]
    [string] $RootDir = 'c:\Tools\Ninite',
    [Parameter(HelpMessage='Optionally enter a number of days, after which the log files will be deleted.')]
    [decimal] $LogsAge
)
# Powershell path handy tip for scheduler task:
# C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe


#region functions
Function ErrorsToFile
{
    if($error.count -gt 0)
    {
        $Error | out-file -FilePath "$LogPathLocal\$($env:computername)_$((Get-Date).ToString("yyyyMMdd"))_errors.txt" -Append
        $error.clear() #Clears current errors
    }
}

Function Out-Log
{
    Param (
        [Parameter(Mandatory=$True)]
        [string]$Message,
        [switch]$OutToFile = $false
    )

    Write-Verbose -Message $Message -Verbose:$MyVerbose
    if($OutToFile)
    {
        "[$((get-date).ToString("G",$(New-Object globalization.cultureinfo("pl-PL"))))]: $Message" | out-file -FilePath "$LogPathLocal\scriptLogs.txt" -Append -Encoding utf8
    }
}

Function Get-NiniteProcess
{
    Param (
        [Parameter(Mandatory=$True)]
        [string]$FileName,
        [Parameter(Mandatory=$True)]
        [string]$Arguments
    )
    
    $Process = New-Object System.Diagnostics.ProcessStartInfo
    $Process.FileName = $FileName
    $Process.RedirectStandardOutput = $true
    $Process.UseShellExecute = $false
    $Process.Arguments = $Arguments
    $Diag = New-Object System.Diagnostics.Process
    $Diag.StartInfo = $Process
    $Diag.Start() | Out-Null
    $Diag.WaitForExit()
    $Diag.StandardOutput.ReadToEnd() #result capture
}

Function Delete-OldFiles
{
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [decimal] $FilesAge,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("Log")]
        [string] $FolderPath = 'c:\Tools\Ninite\log\'
    )

    $DeleteBefore = (Get-Date).AddDays(-$FilesAge)
    $FilesToDelete = Get-ChildItem -Path $FolderPath -Filter *.txt | Where-Object { $_.CreationTime -lt $DeleteBefore -and (!($_.PSIsContainer))}
    if($FilesToDelete)
    {
        $FilesToDelete | Remove-Item -Force
    }
    @($FilesToDelete).count
}

#endregion

#region declaration of variables
$AppPathLocal  = "$RootDir\App\NinitePro.exe"
$LogPathLocal  = "$RootDir\Log"
$MyVerbose     = $true
$Cache         = $null
if($CachePaths)
{
    [string[]]$CachePaths = $CachePaths.split(' ') #converting the captured parameter from commandline to multi string
}
$DoNotRunBefore = 4 #Do not perform script's execute if you have previously performed the script less than $DoNotRunBefore hours ago. Handy option.
$StartTime = [datetime]::Now
#endregion

#region Creating structure and copy / update NinitePro.exe application
Out-Log "================== Start ==================" -OutToFile
Out-Log 'Received variables:' -OutToFile
if($NiniteSourcePath)
{
    Out-Log " NiniteSourcePath: $NiniteSourcePath" -OutToFile
}
Out-Log " Apps            : $Apps" -OutToFile
if($LogsAge)
{
    Out-Log " LogsAge         : $LogsAge" -OutToFile
}
if($CachePaths)
{
    Out-Log " CachePaths      : $CachePaths" -OutToFile
}
Out-Log " UpdateOnly      : $(if($UpdateOnly){'true'}else{'false'})" -OutToFile
Out-Log " RootDir         : $RootDir" -OutToFile

if(!(test-path $RootDir\Log))
{
    Out-Log "Building local folder structure.." -OutToFile
    New-Item $RootDir\App -ItemType directory 2> $null | Out-Null #2> $null It provides no message for error
    New-Item $RootDir\Log -ItemType directory 2> $null | Out-Null
}

if($NiniteSourcePath)
{
    if(!(Test-path -Path $AppPathLocal)) #if the are no local ninitepro file and remote path was entered, do....
    {
        if(test-path $NiniteSourcePath) #check if remote file exists
        {
            Out-Log "Coping ninitepro.exe file from remote location.." -OutToFile
            copy-item -Path $NiniteSourcePath -Destination $AppPathLocal
        }
    }
    elseif(test-path $NiniteSourcePath) #if ninitepro exists on remote server and exists in local machnie, do..
    {
        $FSO = New-Object -ComObject Scripting.FileSystemObject
        if($FSO.GetFileVersion($NiniteSourcePath) -gt $FSO.GetFileVersion($AppPathLocal))
        {
            Out-Log "Local ninitepro.exe file: Updating from remote location.." -OutToFile
            copy-item -Path $NiniteSourcePath -Destination $AppPathLocal -Force    #..update local file. source: unc
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($FSO)
        }    
        else
        {
            Out-Log "Local ninitepro.exe file: exists and does not require an update." 
        }
    }
}

if(!(test-path $AppPathLocal)) #re-check whether a exists local ninitepro file
{
    Out-Log 'ERROR! Local ninitepro.exe file missing!' -OutToFile
    ErrorsToFile #saving errors to file
    BREAK #breaking the scipt
}
if(!($CachePaths))
{
    if(Get-Command Invoke-WebRequest 2>$null) #PS 3.0
    {
        $webStatusCode = (Invoke-WebRequest 'https://ninite.com/' -UseBasicParsing -TimeoutSec 10 -Method Get).statusCode        
    }
    else #PS 2.0
    {
        $webStatusCode = [int](([System.Net.WebRequest]::Create('https://ninite.com/')).GetResponse().StatusCode)
    }
    if($webStatusCode -ne 200)
    {
        Out-Log "ERROR! Cache path was not entered and ninite.com reported unexpected status: $webStatusCode." -OutToFile
        ErrorsToFile #saving errors to file
        BREAK #breaking the scipt
    }
    else
    {
        Out-Log "ninite.com returned expected status: $webStatusCode" -OutToFile
    }
}

ErrorsToFile #saving errors to file
#endregion

#region Testing cache paths and selecting the fastest
if($CachePaths)
{
    Foreach ($server in $CachePaths){
        $hostName = ($server.Split('`\') | where {$_})[0]
        if(Test-Connection $hostName -Count 1 -Quiet)
        {
            Out-Log "Testing connection: $hostName .."
            $testCache = [pscustomobject]@{
                Path = $server
                AveragePing = (Test-Connection $hostName -count 4 -BufferSize 64 | Measure-Object ResponseTime -average).average
            }
            Out-Log  " Average ping: $($testCache.AveragePing)"
            if(!($cache)) #if there is nothing in the cache, then enter the current test
            {
                $Cache = $testCache
            }
            else #But if it is, then ..
            {
                if($testCache.AveragePing -lt $cache.AveragePing) #if it is faster then overwrite
                {
                    $Cache = $testCache
                }
            }
        }
        else
        {
            Out-Log "Cache folder unavailable: $server" -OutToFile
        }
    }
}
if($cache)
{
    Out-Log "The $(if(@($CachePaths).count -gt 1){'fastest '})cache path: $(($Cache.path.Split('`\') | where {$_})[0])" -OutToFile
}
else
{
    
    if(@($CachePaths).count -gt 1)
    {
        Out-Log 'All cache path are unavailable. Apps will be downloaded from ninite.com.' -OutToFile
    }
    else
    {
        Out-Log 'Cache path are unavailable or did not entered. Apps will be downloaded from ninite.com.' -OutToFile
    }
}
#endregion

#region submitting arguments
$Arguments = '/disableshortcuts /disableautoupdate /allusers /silent .'
if($UpdateOnly)
{
    $Arguments += ' /updateonly'
}
$Arguments += " /select $([string]$Apps)" #selected applications
if($cache) #if cache location is available (submitting arguments)
{
    $Arguments += " /cachepath $($cache.path)"
}
<# optionally..
else
{
    $Arguments += "/nocache"
}
#>
#endregion

#region log retention
if($LogsAge -gt 0)
{
    Out-Log "Launching logs retention.."
    $DeletedFiles = Delete-OldFiles -FilesAge $LogsAge -FolderPath $LogPathLocal
    Out-Log "Deleted $DeletedFiles log files from $LogPathLocal folder" -OutToFile
}
#endregion

#region ninitepro execution and log saving
Out-Log "Executed command : $AppPathLocal $Arguments" -OutToFile
$Result = Get-NiniteProcess -FileName $AppPathLocal -Arguments $Arguments
Out-Log "Ninite log file  : $("$($env:computername)_$((Get-Date).ToString("yyyyMMddHHmmss"))_$(if($UpdateOnly){'update'}else{'install'}).txt")" -OutToFile
Out-Log "Ninite status    : $( if($Result) {(($Result -split '\n')[0]).Substring(0,((($Result -split '\n')[0]).Length -1))} else {'N/A'} )" -OutToFile
$Result | out-file -FilePath "$LogPathLocal\$($env:computername)_$((Get-Date).ToString("yyyyMMddHHmmss"))_$(if($UpdateOnly){'update'}else{'install'}).txt" -Append

ErrorsToFile #saving errors to file
$EndTime = [datetime]::Now
Out-Log "=========== END [Elapsed time: $(($endTime - $StartTime).Hours)h. $(($endTime - $StartTime).Minutes)m. $(($endTime - $StartTime).Seconds)s.] ===========" -OutToFile
Out-Log " " -OutToFile
#endregion

#region HANDY: obtaining the list of applications from ninite.com
<#
forEach($datum in ((Invoke-WebRequest 'https://ninite.com/applist/pro.html').ParsedHtml.getElementsByTagName("table") | Select-Object -First 1 ).rows){
    if($datum.tagName -eq "tr"){
        $thisRow = @()
        $cells = $datum.children
        forEach($cell in $cells){
            if($cell.tagName -imatch "t[dh]"){
                $thisRow += $cell.innerText
            }
        }
        $table += $thisRow -join ","
    }
}

$acceptable = $table | Foreach{
    $_.split(',')[-1]
}
#>
#endregion
