$previosConfig = Import-Clixml -Path (Get-Item c:\support\a*.xml | select -First 1).FullName

$newConfig = Import-Clixml -Path (Get-Item c:\support\a*.xml | select -Last 1).FullName

$DiffBrowserName = Compare-Object -ReferenceObject $previosConfig.appsettings -DifferenceObject $newConfig.appsettings -Property browsername -PassThru
$newApps = $DiffBrowserName | Where SideIndicator -eq '=>'
$delApps = $DiffBrowserName | ? SideIndicator -eq '<='
#$delApps | Select BrowserName | ft
#$newApps | select BrowserName | ft

If ($newApps) {
    $MailMessage =""
    #region Pre-Work - setup details to send status email
    $smtpServer = "mail.domain.com"
    $MailFrom = "CitrixScript@domain.com"
    #$MailTo = @("nobody@domain.com")
    $MailTo = @("user.name@domain.com")
    $MailSubject = "New XenApp Apps"
    #endRegion


    $MailMessage = @"
    Citrix Applications Added:<br>
    <br>
"@

    $newApps | % {
        $MailMessage =  $MailMessage + "<br>$($_.BrowserName)"
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
}#endregion