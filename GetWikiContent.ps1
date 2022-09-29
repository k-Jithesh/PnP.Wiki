$SiteURL = "https://aucklandtransport.sharepoint.com/sites/KBBK"
$CSVFilePath = "C:\Jithesh\Customer\AT\SPO\Wiki\Files"


# Install-Module -Name "PnP.PowerShell"
# Register-PnPManagementShellAccess

$ErrorActionPreference = "Stop"

$SiteConn = Connect-PnPOnline -Url $SiteURL -UseWebLogin # -Credentials $Cred -ReturnConnection

$RootWeb = Get-PnPWeb -Connection $SiteConn 

$List = Get-PnPList -Web $RootWeb | Where {$_.Title -eq "Site Pages"} 

$ListItem = Get-PnPListItem -List $List

#$counter = 1
$ListsInventory = New-Object System.Collections.Generic.List[Object]

$ListItem |  ForEach-Object {
         $ListsInv = New-Object -TypeName PSObject -Property @{
                    FileLeafRef = $_.FieldValues.FileLeafRef
                    WikiField = $_.FieldValues.WikiField 
                    Title = $_.FieldValues.Title
                    Categ = $_.FieldValues.New_x0020_Nav
                    KeyWords = $_.FieldValues.Keywords.Label -join ';' 
                    Id = $_.Id
            }
            
            # $ListsInv  | Out-File "$CSVFilePath\$counter.txt" 
            Add-Content -Path "$CSVFilePath\$($ListsInv.Id).html" -Value $ListsInv.WikiField -Encoding UTF8 
            # $counter++
            $ListsInventory.Add($ListsInv)          
          }

          $ListsInv

          $ListsInventory | Select-Object Id, Title,Categ,FileLeafRef,KeyWords  | Export-Csv "C:\Jithesh\Customer\AT\SPO\Wiki\list.csv" -NoTypeInformation


