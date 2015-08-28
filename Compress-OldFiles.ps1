[CmdletBinding()]
param (
    [parameter(ParameterSetName='NoEmail',Mandatory=$true)][parameter(ParameterSetName='Email',Mandatory=$true)][string]$Folder,
    [parameter(ParameterSetName='Email', Mandatory=$true)][string]$To,
    [parameter(ParameterSetName='Email', Mandatory=$true)][string]$From,
    [parameter(ParameterSetName='Email', Mandatory=$true)][string]$SMTPServer = $PSEmailServer
)


Function Format-FileSize() {
    Param ([int]$size)
    If     ($size -gt 1TB) {[string]::Format('{0:0.00} TB', $size / 1TB)}
    ElseIf ($size -gt 1GB) {[string]::Format('{0:0.00} GB', $size / 1GB)}
    ElseIf ($size -gt 1MB) {[string]::Format('{0:0.00} MB', $size / 1MB)}
    ElseIf ($size -gt 1KB) {[string]::Format('{0:0.00} kB', $size / 1KB)}
    ElseIf ($size -gt 0)   {[string]::Format('{0:0.00} B', $size)}
    Else                   {''}
}

 

Write-Host "Processing $Folder"

Write-Host 'Searching for items...'
$Items = Get-ChildItem -Path $Folder -Exclude *.7z,*.rar,*.zip

# TODO:  Add filtering here based on last modified time! For directories, search for mod time of newest file.
$Items = $Items
Write-Host "Found $($items.Count) items to compress"

$List = @();
$Errors = $false;

$ErrorLog = [System.IO.Path]::GetTempFileName()

foreach ($Item in $Items) {
    if($Item.PSIsContainer) {
        $OriginalSize = (Get-ChildItem $Item -recurse | Measure-Object -Property Length -Sum).Sum
    } else {
        $OriginalSize = $Item.Length
    }


	Write-Host "Compressing $Item"
	& $PSScriptRoot\7za.exe a -mmt -mx9 -sdel -t7z "$Item.7z" "$Item" 2>>$ErrorLog  | Out-Null
	if($LASTEXITCODE -eq 0) {
        $CompressedSize = (Get-ChildItem "$Item.7z").Length
        $Savings = $OriginalSize - $CompressedSize
	} else {
		Write-Warning "Failed to compress $Item"
        $Errors = $True
        $CompressedSize = -1
        $Savings = 0
	}

    $List += New-Object PSObject -Property @{'Item' = $item; 'OriginalSize' = $OriginalSize; 'CompressedSize' = $CompressedSize; 'Savings' = $Savings; 'ExitCode' = $LASTEXITCODE}
}


Write-Output $list

if($Errors) {
    $ErrorMessages = Get-Content $ErrorLog
    Write-Output $ErrorMessages
}$x


If($To -and $Items) {
    Write-Host "Sending report to $To"
    $Subject = "Compression Report for $Folder on $($ENV:ComputerName)"
    $Head = '<style>table {border-collapse: collapse; width: 100%} table, th, td {border:1px solid black;}</style>'
    $PreContent = "<h2>Compression Report for $Folder on $($ENV:ComputerName)</h2>"
    $PostContent = ''
    If($Errors) { 
        $Subject = "Errors: $Subject"
        $PreContent = "$PreContent <p>Some files failed to compress</p>"
        $PostContent = "<h3>Error Messages</h3> <pre>$($ErrorMessages|Out-String)</pre>"
   }

   $message = $List | ConvertTo-Html -As Table -PreContent $PreContent -PostContent $PostContent -Head $Head -Property Item,
        @{Label='Original Size';Expression={Format-FileSize -Size $_.OriginalSize}},
        @{Label='Compressed Size';Expression={Format-FileSize -Size $_.CompressedSize}},
        @{Label='Spaced Saved';Expression={Format-FileSize -Size $_.Savings}},
        ExitCode | Out-String
    Send-MailMessage -To $To -Subject $Subject -Body $message -SmtpServer $SMTPServer -From $From -BodyAsHtml 
}

Remove-Item $ErrorLog