[CmdletBinding()]
param (
    [parameter(Mandatory=$true)][string]$folder
)

Write-Output "Processing $folder"

Write-Output 'Searching for items...'
$items = Get-ChildItem -Path $folder -Exclude *.7z,*.rar,*.zip

# TODO:  Add filtering here based on last modified time!
# Calculate original file/folder sizes

Write-Output "Found $($items.Count) items to compress"

foreach ($item in $items) {
	Write-Output "Compressing $item"
	& $PSScriptRoot\7za.exe a -mmt -mx9 -sdel -t7z "$item.7z" "$item"
	if($LASTEXITCODE -eq 0) {
		Write-Output "Compressed $item and deleted original"
	} else {
		Write-Warning "Failed to compress $item"
	}
}


