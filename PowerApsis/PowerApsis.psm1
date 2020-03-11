[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
    
Function Invoke-ApsisAPI {
Param(
    [Parameter(Mandatory)]
    [string]$Function,
    [pscustomobject]$Body,
    [switch]$ForceBodyToArray,
    [ValidateSet('Get','Post','Put','Patch','Delete')]
    [string]$Method = 'Post',
    [Switch]$Queued,
    [string]$Key = ''
)
    $header = @{
        Authorization = "Basic $([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$key`:")))"
        "Content-type" = "Application/json"
        Accept = "json"
        charset= "utf-8"
    }

    #Convert body to utf8
    If($Body){
        If($ForceBodyToArray){[array]$Body = $($Body)}
        $BodyString = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $Body -Depth 100))
    }

    Switch($Queued){
        $true{
            [uri]$ApsisAPI = "https://se.api.anpdm.com/$Function"
            If($Body){
                $ret = Invoke-RestMethod -Uri $ApsisAPI -Headers $header -Body $BodyString -Method $Method
            }
            Else{
                $ret = Invoke-RestMethod -Uri $ApsisAPI -Headers $header -Method $Method
            }

            Do{
                $qRet = Invoke-RestMethod -Uri ($ret.Result.PollUrl.Replace('http://','https://')) -Method Get
                Switch -Regex ($qRet.State) {
                    "[01]"{
                        $qRetWorking = $true
                        Start-Sleep 1
                    }
                    "[2]"{
                        $qRetWorking = $False
                    }
                    Default{
                        $qRetWorking = $False
                        Throw $qRet
                    }
                }
            }
            While($qRetWorking)

            If($qRet.State -eq '2'){
                Return Invoke-RestMethod $qRet.DataUrl
            }
            Else{
                Throw $qRet
            }
        }
        $false{
            [uri]$ApsisAPI = "https://se.api.anpdm.com/$Function"
            If($Body){
                $ret = Invoke-RestMethod -Uri $ApsisAPI -Headers $header -Body $BodyString -Method $Method
            }
            Else{
                $ret = Invoke-RestMethod -Uri $ApsisAPI -Headers $header -Method $Method
            }
        }
    }
    If($ret.Code -eq '1'){
        Return $ret
    }
    Else{
        Throw $ret
    }
}

Function Get-ApsisSubscribers{
Param(
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    Return Invoke-ApsisAPI -Key $Key -Function '/v1/subscribers/all' -Body ([pscustomobject]@{AllDemographics = $true}) -Queued
}

Function Set-ApsisSubscriber{
Param(
    [Parameter(Mandatory=$true)]
    [string]$email,
    [Parameter(Mandatory=$true)]
    [string]$name,
    [Parameter(Mandatory=$true)]
    [string]$siteName,
    [Parameter(Mandatory=$true)]
    [string]$MailingListId,
    [Parameter(Mandatory=$true)]
    [string]$Key,
    [Parameter(Mandatory=$false)]
    [string]$ClientName = ''
)
    
    $Body = [pscustomobject]@{
        Email = $email
        Name = $name
        
        DemDataFields = @(
            [pscustomobject]@{
                Key = "Fornamn"
                Value= $name
            }
            [pscustomobject]@{
                Key = "Site"
                Value = $siteName
            }
            [pscustomobject]@{
                Key = "KlientID"
                Value = $ClientName
            }
        )
    }
    $res = Invoke-ApsisAPI -Key $Key -Function "/v1/subscribers/mailinglist/$MailingListId/create?updateIfExists=true" -Body $Body
    Return $res.Message
}

Function Remove-ApsisSubscriber{
Param(
    [Parameter(Mandatory=$true)]
    [array]$Ids,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    Return Invoke-ApsisAPI -Key $Key -Function '/subscribers/v2/id' -Method Delete -Body $Ids -ForceBodyToArray -Queued
}

Function Get-ApsisMailinglists{
Param(
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    Return (Invoke-ApsisAPI -Key $Key '/mailinglists/v2/all').Result
}

Function New-ApsisMailingLists{
Param(
    [Parameter(Mandatory=$true)]
    [string]$Name,
    [Parameter(Mandatory=$true)]
    [string]$FromEmail,
    [Parameter(Mandatory=$true)]
    [string]$FromName,
    [Parameter(Mandatory=$false)]
    [string]$FolderID = 0,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
        
    $Body = [pscustomobject]@{
        Name = $Name
        FromEmail = $FromEmail
        FromName = $FromName
        FolderID = $FolderID
    }
        
    Invoke-ApsisAPI -Key $Key -Function '/v1/mailinglists/' -Body $Body
}

Function Remove-ApsisMailingLists{
Param(
    [Parameter(Mandatory=$true)]
    [array]$Ids,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    Return Invoke-ApsisAPI -Key $Key -Function '/v1/mailinglists/' -Method Delete -Body $Ids -ForceBodyToArray -Queued
}

Function Get-ApsisEvents{
Param(
    [Parameter(Mandatory=$true)]
    [string]$Key
)

    #Get all events with sessions
    $Body = [pscustomobject]@{
        ExcludeDisabled = $false
    }
    Return (Invoke-ApsisAPI -Key $Key -Function '/event/v2/sessions' -Body $Body).Result
}

Function Get-ApsisEventAttendees{
Param(
    [Parameter(Mandatory=$true)]
    [string]$EventId,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    $Body = [pscustomobject]@{
        EventId = $EventId
    }
    Return (Invoke-ApsisAPI -Key $Key -Function "/event/v2/attendees" -Body $Body).Result
}

Function Get-ApsisEventOptions{
Param(
    [Parameter(Mandatory=$true)]
    [string]$EventId,
    [Parameter(Mandatory=$true)]
    [string]$Key
)    
    Return (Invoke-ApsisAPI -Function "/event/v2/$EventId/optionsdatacategories" -Key $key -Method Get).Result
}
    
Function Add-ApsisEventAttendee{
Param(
    [Parameter(Mandatory=$true)]
    [string]$EventId,
    [Parameter(Mandatory=$true)]
    [string]$SessionId,
    [Parameter(Mandatory=$false)]
    [ValidateSet('Registered','WaitingList','CheckedIn','Cancel')]
    [string]$Status = 'Registered',
    [Parameter(Mandatory=$true)]
    [string]$Email,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    $Body = [pscustomobject]@{
        Attendee = [pscustomobject]@{
            ControlValues = @(
                [pscustomobject]@{
                    Name = "email"
                    Value = $Email
                }
            )
        }
        Guests = @()
        Status = "Registered"
        NumberOfAnonymousGuests = 0
        DebugMode = $true
    }
    $res = (Invoke-ApsisAPI -Key $Key -Function "/event/v2/$EventId/session/$SessionId/attendee" -Body $Body).Result

    If($res.Succeeded -eq $true){
        Return $res
    }
    Else{
        Throw $res
    }
}

Function Register-ApsisEventAttendee{
Param(
    [Parameter(Mandatory=$true)]
    [string]$EventId,
    [Parameter(Mandatory=$true)]
    [string]$SessionId,
    [Parameter(Mandatory=$true)]
    [string]$Email,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    $Body = [pscustomobject]@{
        Attendee = [pscustomobject]@{
            ControlValues = @(
                [pscustomobject]@{
                    Name = "email"
                    Value = $Email
                }
            )
        }
        Guests = @()
        NumberOfAnonymousGuests = 0
        DebugMode = $true
    }
    $res = (Invoke-ApsisAPI -Key $Key -Function "/event/v2/$EventId/session/$SessionId/register" -Body $Body -ErrorAction continue -WarningAction SilentlyContinue).Result

    If($res.Succeeded -eq $true -or $res.DebugInfo -like "*ParticipantAlreadyRegistered"){
        Return 0
    }
    Else{
        Throw $res | ConvertTo-Json -Depth 100
    }
}

Function Set-ApsisEventAttendeeStatus{
Param(
    [Parameter(Mandatory=$true)]
    [string]$EventId,
    [Parameter(Mandatory=$true)]
    [string]$SessionId,
    [Parameter(Mandatory=$true)]
    [string]$Email,
    [Parameter(Mandatory=$true)]
    [string]$Key
)
    $Body = [pscustomobject]@{
        Attendee = [pscustomobject]@{
            ControlValues = @(
                [pscustomobject]@{
                    Name = "email"
                    Value = $Email
                }
            )
        }
        Guests = @()
        NumberOfAnonymousGuests = 0
        DebugMode = $true
    }
    $res = (Invoke-ApsisAPI -Key $Key -Function "/event/v2/$EventId/attendee/$AttendeeId/status" -Body $Body -ErrorAction continue -WarningAction SilentlyContinue).Result

    If($res.Succeeded -eq $true -or $res.DebugInfo -like "*ParticipantAlreadyRegistered"){
        Return 0
    }
    Else{
        Throw $res | ConvertTo-Json -Depth 100
    }
}