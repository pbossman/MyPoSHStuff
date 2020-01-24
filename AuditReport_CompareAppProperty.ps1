## this script is part of a sweet of audit scripts. and will compare Citrix Application properties
##  then send an email on the differences

### Function: Compare-ObjectProperties
### Credit: Jamie Nelson MSFT
### https://blogs.technet.microsoft.com/janesays/2017/04/25/compare-all-properties-of-two-objects-in-windows-powershell/

### Cmdlet: Invoke-Parallel
### Credit: RamblingCookiemoster
### https://github.com/RamblingCookieMonster/Invoke-Parallel



# ALL 
$previosConfig = Import-Clixml -Path (Get-Item c:\support\a*.xml | select -First 1).FullName
$newConfig = Import-Clixml -Path (Get-Item c:\support\a*.xml | select -Last 1).FullName
$previosConfig | select -First 150 -Skip 100 | Export-Clixml C:\Support\sample01.xml
$newConfig | select -Skip 100 -First 150 | Export-Clixml C:\Support\sample09.xml

##Sample
$previosConfig = Import-Clixml -Path (Get-Item c:\support\S*.xml | select -First 1).FullName
$newConfig = Import-Clixml -Path (Get-Item c:\support\s*.xml | select -Last 1).FullName

$propertyName = "BrowserName"
$AppSettings = Compare-Object -ReferenceObject $previosConfig.appsettings -DifferenceObject $newConfig.appsettings -Property $propertyName -IncludeEqual -PassThru | 
                    Where SideIndicator -eq '==' | Select -Property * -ExcludeProperty SideIndicator
$AppSettings | Invoke-Parallel {
    Function Compare-ObjectProperties {
        Param(
            [PSObject]$ReferenceObject,
            [PSObject]$DifferenceObject 
        )
        $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
        $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
        $objprops = $objprops | Sort | Select -Unique
        $diffs = @()
        foreach ($objprop in $objprops) {
            $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
            if ($diff) {            
                $diffprops = @{
                    PropertyName=$objprop
                    RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                    DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
                }
                $diffs += New-Object PSObject -Property $diffprops
            }        
        }
        if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
    }

    Write-Verbose -Message "$($_.BrowserName)" 
    $currApp = $_
    $ref = $newConfig.appsettings | Where Browsername -eq $_.$propertyName
    $DiffProps = Compare-ObjectProperties -ReferenceObject $_ -DifferenceObject $ref
    if ($DiffProps) { 
        $currApp | Add-Member -MemberType NoteProperty -Name DifferentProperties -Value $DiffProps
        $currapp | select BrowserName, DifferentProperties
    }
} -ImportVariables | Tee-Object -Variable AppChanges


If ($AppChanges) {
    $MailMessage =""
    #region Pre-Work - setup details to send status email
    $smtpServer = "mail.domain.com"
    $MailFrom = "CitrixScript@domain.com"
    #$MailTo = @("nobody@domain.com")
    $MailTo = @("user.name@domain.com")
    $MailSubject = "XenApp Changes"
    #endRegion

    ## build email message body
    $MailMessage = @"
    Citrix Applications Variance:<br>
    <br>
"@

    $AppChanges | %{ 
        $MailMessage =  $MailMessage + "$($_.BrowserName)"

        $msgText = ($_.DifferentProperties| Out-String)

        $MailMessage =  $MailMessage + $msgText

    } 

    $MailMessage =  $MailMessage + @"
    </B><br><br>
    Sent by: $($env:USERNAME)
    <br>
      on:$($env:COMPUTERNAME)
    <br>
      Script: $($MyInvocation.MyCommand.Definition)
"@

    Send-MailMessage -SmtpServer $smtpServer -From $MailFrom -To $MailTo -Subject $MailSubject -BodyAsHtml $MailMessage
}
#endregion
