
$DomainName = $env:USERDNSDOMAIN  
$GPOsInDomain = Get-GPO -All -Domain $DomainName  #Get-GPO -Name ""  - Use this command to find Specific GPO to target.

#$GPOsInDomain = Get-GPO -Name "R&D - PR - Computer and User - DC01_Citrix Windows Folder Redirection"
# Target specific OUs
<#
$GPOsInDomain @ (
     
)
#>

$ScriptPath = "C:\Users\ABaquilo\OneDrive - City National Bank\Documents\GPO Scripts"


if (!(Get-Module -Name ImportExcel) -and !(Get-Module -Name GroupPolicy)) {
    Install-Module -Name ImportExcel -Force
    Install-Module -Name GroupPolicy -Force
}


$GpoLinks = foreach ($g in $GPOsInDomain){              
        [xml]$Gpo = Get-GPOReport -ReportType Xml -Guid $g.Id
        foreach ($i in $Gpo.GPO.LinksTo) {
                [PSCustomObject]@{
                    "GPO Name" = $Gpo.GPO.Name  -join ';'
                    "Path" = $i.SOMPath  -join ';'
                    "Link Enabled" = $i.Enabled  -join ';'
                    "OU Name" = $i.SOMName  -join ';'
                    "Created Time" = $Gpo.GPO.CreatedTime  -join ';'
                    "Modified Time" = $Gpo.GPO.ModifiedTime  -join ';'
                    
                }
            }
        }



Write-host "Completed running the GPO report script at: $(get-date -Format G)."
$GpoLinks | Sort-Object Name | Export-Csv -Path $ScriptPath\GPOReport.csv -Append
#$GpoLinks | Sort-Object Name | Export-Excel -KillExcel -Path $ScriptPath\GPOReport.xlsx -Append -AutoSize

