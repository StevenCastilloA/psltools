<#Requires 

Reports.ReadAll
User.Read.All
#>

function connectToGraph{
param (
        [string]$idTenant,
        [string]$client_id,
        [string]$client_secret
    )
    
    
    
    $url = 'https://login.microsoftonline.com/'+$idTenant+'/oauth2/token'
    
    $body = @{
     'client_id'=$client_id
     'client_secret'=$client_secret
     'resource'='https://graph.microsoft.com'
     'grant_type'='client_credentials'
    }

    $header = @{
        'Content-Type'='application/x-www-form-urlencoded'
    }

    $response = Invoke-RestMethod -Uri $url -Method 'POST' -Body $body -Headers $header 

    return $response.access_token

}


function getMailboxes{
param (
        [string]$bearer
    )
    
    $url = 'https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period=''D7'')?$format=application/json'    
    
    $mailboxes = @()

        $header = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer "+$bearer
        }

        while( $url.Length -gt 5 ){
  
            try{

                $response = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $header -ContentType 'application/json'
                #write-host $response.value| ConvertFrom-Json
                #echo $response

                $mailboxes += $response.value 
                $url = $response.'@odata.nextLink'
        
            }catch {
    
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()                
                Write-Host 'Error getMailboxes'
                Write-Host $ErrResp.error.message
                $url = ''
       
            }
        }
        
        return $mailboxes

    }
function getOneDrives{
param (
        [string]$bearer
    )
    
    $url = 'https://graph.microsoft.com/beta/reports/getOneDriveUsageAccountDetail(period=''D7'')?$format=application/json'    
    
    $onedrives = @()

        $header = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer "+$bearer
        }

        while( $url.Length -gt 5 ){
  
            try{

                $response = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $header -ContentType 'application/json'
                #write-host $response.value| ConvertFrom-Json
                #echo $response

                $onedrives += $response.value 
                $url = $response.'@odata.nextLink'
        
            }catch {
    
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()
                Write-Host 'Error getOnedrives'
                Write-Host $ErrResp.error.message
                $url = ''
       
            }
        }
        
        return $onedrives

    }

function getSitesStats{
param (
        [string]$bearer
    )
    
    $url = 'https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period=''D7'')'    
    
    $sites = @()

        $header = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer "+$bearer
        }

        while( $url.Length -gt 5 ){
  
            try{
            
                $response = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $header -ContentType 'application/json'

                $csvData = $response | ConvertFrom-Csv
                $jsonData = $csvData 

                $sites += $jsonData
                $url = $response.'@odata.nextLink'
        
            }catch {
    
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()
                Write-Host 'Error getSiteStats'
                Write-Host $ErrResp.error.message
                $url = ''
       
            }
        }
        
        return $sites

    }

function getSites{
param (
        [string]$bearer
    )
    
    $url = 'https://graph.microsoft.com/v1.0/sites'    
    
    $sites = @()

        $header = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer "+$bearer
        }

        while( $url.Length -gt 5 ){
  
            try{
            
                $response = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $header -ContentType 'application/json'

                $sites += $response.value 
                $url = $response.'@odata.nextLink'
        
            }catch {
    
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()                
                Write-Host 'Error getSites'
                Write-Host $ErrResp.error.message
                $url = ''
       
            }
        }
        
        return $sites

    }

function getUsers{
param (
        [string]$bearer
    )
    
    $url = 'https://graph.microsoft.com/v1.0/users?$top=999'
    
    $users = @()

        $header = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer "+$bearer
        }

        while( $url.Length -gt 5 ){
  
            try{

                $response = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $header -ContentType 'application/json'
                #echo $response

                $users += $response.value
                $url = $response.'@odata.nextLink'
        
            }catch {
    
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
                $streamReader.Close()
                Write-Host 'Error getUsers'
                Write-Host $ErrResp.error.message
                $url = ''
       
            }
        }
        return $users

    }



$idTenant = ''
$idApp = ''
$secretApp = ''

$b = connectToGraph -idTenant $idTenant -client_id $idApp -client_secret $secretApp

$users = getUsers -bearer $b
$mailboxes = getMailboxes -bearer $b
$onedrives = getOneDrives -bearer $b
$sites = getSites -bearer $b
$sitesStats = getSitesStats -bearer $b

$nameFile = 'StatsOrigen'

$users | Export-Excel -Path "D:\$nameFile.xlsx" -WorksheetName "Users" 
$mailboxes | Export-Excel -Path "D:\$nameFile.xlsx" -WorksheetName "Mailbox" 
$onedrives | Export-Excel -Path "D:\$nameFile.xlsx" -WorksheetName "Onedrive" 
$sites | Export-Excel -Path "D:\$nameFile.xlsx" -WorksheetName "Sites" 
$sitesStats | Export-Excel -Path "D:\$nameFile.xlsx" -WorksheetName "SitesStats" 
