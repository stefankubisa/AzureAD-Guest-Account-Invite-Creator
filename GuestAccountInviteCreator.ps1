cls

$IsItDone_Pos = 1
$GuestUserName_Pos = 2
$GuestUserEmail_Pos = 3

$Path = $args[0]
$CurrentWorksheet = $args[1]

Write-Host "Creating Excel App Object" 
$excel = new-object -comobject Excel.Application 
$excel.visible = $true 
$excel.DisplayAlerts = $false 
$excel.WindowState = "xlMaximized"
Write-Host "Opening Workbook"
Write-Host "____________________"

try {
    $workbook = $excel.workbooks.open($path)
}
catch {
    Write-Host $_.Exception.Message
    Write-Host "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}
try {
    $Worksheet = $workbook.Worksheets.item($CurrentWorksheet)
}
catch {
    Write-Output $_.Exception.Message
    Write-Output "Closing Workbook"
    $workbook.Close()
    Write-Output "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}

$verticalCount = (($Worksheet.UsedRange.Rows).count - 1 )
Write-Host -ForegroundColor DarkGreen "User Count: $verticalCount"
Write-Host "If this makes sense press any key, if not CTRL + C"
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Write-Host "
Which authentication method would you like to use:
--------------------------------------------------
1. Autoauthentication via UPN
2. Autoauthentication via prompt
3. Authentication via static data in the script
4. Type UPN manually"

$number = Read-Host "Time to choose Dr. Freeman"
switch ($number){
1 {
    $UPN = whoami /upn
    Connect-AzureAD -AccountId $UPN
}
2 {
    Connect-AzureAD
}
3 {
    Connect-AzureAD -AccountId "stefan.kubisa@sellerx.com"
}
4 {
    $UPN = Read-Host
    Connect-ExchangeOnline -UserPrincipalName $UPN
}
}

$keepGoing = $true
while ($keepGoing) {
    for ($i = 1; $i -lt $verticalCount + 1; $i++) { 
        if (($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "OK" -or ($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "SKIP") {
            continue
        }
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "In Progress"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 44

        $GuestUserName = $Worksheet.Cells.Item($i + 1, $GuestUserName_Pos).Text 
        $GuestUserEmail = $Worksheet.Cells.Item($i + 1, $GuestUserEmail_Pos).Text 
        
        Write-Host -ForegroundColor DarkGreen "_____________________________" 
        Write-Host -ForegroundColor DarkGreen "$GuestUserName" 
        Write-Host -ForegroundColor DarkGreen "$GuestUserEmail" 
        Write-Host -ForegroundColor DarkGreen "-----------------------------" 

        New-AzureADMSInvitation -InvitedUserDisplayName $GuestUserName -InvitedUserEmailAddress $GuestUserEmail -InviteRedirectURL https://www.microsoft.com/en-ww/microsoft-teams/log-in -SendInvitationMessage $true

        Start-Sleep -Seconds 1
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "OK"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 43

    }
    $keepGoing = $false
}

$workbook.Save()
Write-Host "All done, closing workbook"
Start-Sleep -Seconds 5
$excel.Quit()