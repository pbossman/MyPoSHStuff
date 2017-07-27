### Credit: RamblingCookieMonster
### Module: PSExcel
### https://github.com/RamblingCookieMonster/PSExcel


## define Manifest Output File
Write-Verbose "Mounting Sharepoint (SP:\)" -Verbose
New-PSDrive -Name SP -PSProvider FileSystem -Root "\\SharepointServer\DavWWWRoot\SitePath" -Description "SharePoint"
$ManifestPath = "SP:\Shared Documents\Manifest.xlsx"
$DFSManifestPath = "\\DOMAIN\HIGHLY_AVAIL_DFS_Namespase\Path\Manifest.xlsx"

## import Required modules\snapins
Import-Module PSExcel
if (-not (Get-Command Save-Excel -ErrorAction SilentlyContinue)) {
    Import-Module 'C:\Program Files\Windows Powershell\Modules\PSExcel\PSExcel.psm1' -verbose
}
if (-not (Get-Command get-XAFarm -ErrorAction SilentlyContinue)) {
    Add-PSSnapin *citrix* -ver
}

$currDate = Get-Date -Format "MM-dd-yyyy"

## build a collection of the individual applications
$appData = Get-XAApplication | Get-XAApplicationReport | 
    % {
        Write-Verbose "Getting details for $($_.BrowserName)" -Verbose
        $WorkGroupServers = @(If ($_.WorkerGroupNames) { Get-XAWorkerGroupServer -WorkerGroupName $_.WorkerGroupNames -ErrorAction SilentlyContinue | Sort ServerName | select -ExpandProperty ServerName } )
        $allServers = @($_.Servernames) + $WorkGroupServers
        [PSCustomObject] @{
            'FolderPath' = $_.folderPath;
            'Application' = $_.BrowserName;
            'ProvisionGroups' = @($_.Accounts | Where AccountdisplayName -ne "DOMAIN\ExcludedAccountName") -join ', ' -replace 'DOMAIN\\','';
            'UserAccounts' = $_.Accounts -join ', '
            'AllServers' = $allServers -join ', ' ;
            'ServerAssignment' = $_.ServerNames -join ', '
            'WorkGroup' = $_.WorkerGroupNames -join ', ';
            'WGServers' = $WorkGroupServers -join ', ';
            'Enabled' = $_.Enabled;
            'AppSettings' = $_
         }
    } |
    Sort FolderPath,Application

## Filter only the "provisioned" apps
## select the fields and order for the list
## Write date to Excel Document (worksheet: CTXData)
Write-Verbose "Writing data to $ManifestPath" -Verbose
$appData | 
    Select FolderPath, Application, ProvisionGroups, AllServers, Workgroup, ServerNames, Enabled |
    sort folderpath,Application | 
    Export-XLSX -Path $ManifestPath -WorksheetName CTXData -AutoFit -ClearSheet

## Create a Powershell object (variable) from the the written Excel document
##   so that rows and columns can be re-formatted
##   and saved back into that document
$excel = new-excel -Path $ManifestPath 

## set format of individual rows/colums
Write-Verbose "Updating format" -Verbose
$excel | Get-Worksheet -Name CTXData | Format-Cell -Size 10 -StartRow 2 -StartColumn 1 -Autofit -BorderStyle Thin
$excel | Get-Worksheet -Name CTXData | Format-Cell -Header -Bold:$true -Size 10 -BackgroundColor Yellow -Color Black
$excel | Get-Worksheet -Name CTXData | Format-Cell -StartColumn 2 -EndColumn 2 -StartRow 2 -Bold:$true -Size 10 -Autofit 
$excel | Get-Worksheet -Name CTXData | Format-Cell -StartColumn 3 -AutofitMaxWidth 45 -Autofit

### Special format for Disabled apps
Write-Verbose "Updating Enabled Data" -Verbose
## clear the formatting for the "Enabled" column (column 7)
##   so that with changes, it will refresh the formatting
##   Formatting:  Set False Values to font color RED
$excel | Get-Worksheet -Name CTXData | Format-Cell -StartColumn 7 -StartRow 2 -BackgroundColor White -Bold:$false -Autofit
$DisabledApps = Search-CellValue -Excel $excel -FilterScript { $_ -like 'FALSE' }
$DisabledApps | Where Column -EQ 7 | % { $excel | Get-Worksheet -name $_.WorkSheetName | 
            Format-Cell -StartRow $_.Row -EndRow $_.row -StartColumn $_.Column -EndColumn $_.Column -Color Red -Verbose}

## write data back to Excel file
##   Passthru the saved document back into the working variable
Write-Verbose "Saving file to $ManifestPath" -Verbose
$Excel = $Excel | Save-Excel -Passthru

## set view / Freeze Panes (top row frozen)
Write-Verbose "Updating Freeze Panes" -Verbose
$excel | Get-Worksheet -Name CTXData | Set-FreezePane -Row 2 

## write data back to Excel file
##   Passthru the saved document back into the working variable
Write-Verbose "Saving file to $ManifestPath" -Verbose
$Excel = $Excel | Save-Excel -Passthru

## write Excel file to backup DFS Path 
Write-Verbose "Saving file to $DFSManifestPath" -Verbose
$Excel | Save-Excel -Path $DFSManifestPath -Passthru

## Cleanup Close Excel Object
$excel | Close-Excel 
## Cleanup PSDrive
Write-Verbose "Removing SP:\" -Verbose
Remove-PSDrive -Name SP

