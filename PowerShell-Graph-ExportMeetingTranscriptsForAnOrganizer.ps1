# PowerShell-Graph-ExportMeetingTranscriptsForAnOrganizer.ps1

<#
.SYNOPSIS
    Search Teams online meeting transcripts in a time range for a specific organizer
    using Microsoft Graph application OAuth, then save transcript content to files.

.NOTES
    REQUIRED APP PERMISSION (Application):
      - OnlineMeetingTranscript.Read.All # Required for reading meeting transcripts. Do admin Grant.

    IMPORTANT:
      - Microsoft documents that application access to online meeting transcripts
        requires an Application Access Policy granted to a user.
      - The getAllTranscripts function is for SCHEDULED online meetings organized by
        the specified organizer.
      - getAllTranscripts currently does NOT support channel meetings.
      - This script uses raw REST calls only (no Graph SDK).
      - By default, meeting transcripts are only stored in Teams for 120 days unless the retention policy is configured.
        See: https://learn.microsoft.com/en-us/microsoftteams/manage-teams-recording-expiration-policy

    EXAMPLE APP ACCESS POLICY COMMANDS (run separately as admin in Teams PowerShell):
      Do this 1 time only:  Install-Module -Name MicrosoftTeams
      Connect-MicrosoftTeams

      New-CsApplicationAccessPolicy -Identity "GraphTranscriptPolicy" `
        -AppIds "YOUR-APP-ID" `
        -Description "Allow app to access Teams meeting transcripts"
      # Example: New-CsApplicationAccessPolicy -Identity "MySomethingSomethingPolicy" --AppIds "xxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Description "Allow app to access Teams meeting transcripts"
      # Note: MySomethingSomethingPolicy is the name of the policy you created - you make up the name here.

      Grant-CsApplicationAccessPolicy -PolicyName "GraphTranscriptPolicy" `
        -Identity "organizer@contoso.com"
      # Example: Grant-CsApplicationAccessPolicy -PolicyName "MySomethingSomethingPolicy" -Identity "organizer@contoso.com"
      # Note that the -PolicyName here matches the policy name created earlier.

    Note: 
        GraphTranscriptPolicy is the name your callling this policy.
        Call this to check if the policy exists: 
            Get-CsApplicationAccessPolicy -Identity "GraphTranscriptPolicy"
        Call this to check if the policy is assingned to the user:
            Get-CsUserPolicyAssignment -Identity "userId"
        If you want broader coverage, assign the policy to the organizer(s) you will query.
    
    How to proxy to Fiddler:
      netsh winhttp set proxy localhost:8888  # To proxy traffic to Fiddler's default address and port.
      netsh winhttp reset proxy               # To reset proxying - be sure to run this to turn off proxying through Fiddler.

    Origionaly generated with CoPilot then modified.
#>

# =========================
# CONFIGURATION
# =========================

# Tenant / App registration
$TenantId     = "YOUR_TENANT_ID"
$ClientId     = "YOUR_APP_CLIENT_ID"
$ClientSecret = "YOUR_APP_CLIENT_SECRET"

# Organizer whose scheduled Teams meetings you want to search
# This user must be covered by the Teams application access policy.
$OrganizerUserId = "ORGANIZER_USER_OBJECT_ID"    
    # For delegate flow - user id/upn /security group/"-global"
    # For app flow - use user ID and not upn. Can also use a security group or "-global"
# Date range (UTC / ISO 8601 recommended)
# Example:
#   2026-06-01T00:00:00Z
#   2026-06-15T23:59:59Z

# UTC time range
$StartDateTimeUtc  = "2026-06-01T00:00:00Z"
$EndDateTimeUtc    = "2026-06-30T23:59:59Z"

# Output
$OutputRoot        = "C:\Temp\TeamsTranscriptExport"
$IncludeMeetingDetails = $true   # requires OnlineMeetings.Read.All

# ---------------------------
# HELPERS
# ---------------------------

$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")]
        [string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    Write-Host "[$timestamp] [$Level] $Message"
}

function Ensure-Folder {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-SafeFileName {
    param(
        [string]$Name,
        [int]$MaxLength = 120
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "UnnamedMeeting"
    }

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalidChars) {
        $Name = $Name.Replace($c, '_')
    }

    $Name = $Name.Trim()
    if ($Name.Length -gt $MaxLength) {
        $Name = $Name.Substring(0, $MaxLength)
    }

    return $Name
}

function Get-GraphAppToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    Write-Log "Requesting app-only access token from Microsoft identity platform..."
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
    return $tokenResponse.access_token
}

function Invoke-GraphJsonGet {
    param(
        [string]$Uri,
        [string]$AccessToken
    )

    $guidString = [guid]::NewGuid().ToString()
    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "application/json"
        "client-request-id" = $guidString
    }

 
    Write-Log "GET JSON: $Uri" "DEBUG"
$resp = Invoke-WebRequest -Method Get -Uri $Uri -Headers $headers -UseBasicParsing
    return ($resp.Content | ConvertFrom-Json)
}

function Invoke-GraphTextGet {
    param(
        [string]$Uri,
        [string]$AccessToken
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "text/vtt"
    }

    Write-Log "GET TEXT: $Uri" "DEBUG"
    $resp = Invoke-WebRequest -Method Get -Uri $Uri -Headers $headers -UseBasicParsing
    return $resp.Content
}

function Get-AllGraphPages {
    param(
        [string]$InitialUri,
        [string]$AccessToken
    )

    $allItems = New-Object System.Collections.Generic.List[object]
    $next = $InitialUri

    while ($null -ne $next -and $next -ne "") {
        $page = Invoke-GraphJsonGet -Uri $next -AccessToken $AccessToken

        if ($page.value) {
            foreach ($item in $page.value) {
                [void]$allItems.Add($item)
            }
        }

        $next = $page.'@odata.nextLink'
        if ($next) {
            Write-Log "Following @odata.nextLink..." "DEBUG"
        }
    }

    return $allItems
}

# ---------------------------
# MAIN
# ---------------------------

Ensure-Folder -Path $OutputRoot
$TranscriptFolder = Join-Path $OutputRoot "Transcripts"
Ensure-Folder -Path $TranscriptFolder

$token = Get-GraphAppToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

# Use the function form documented by Graph for time-range filtering
$getAllTranscriptsUri = @(
    "https://graph.microsoft.com/v1.0/users/$OrganizerUserId/onlineMeetings/",
    "getAllTranscripts(",
    "meetingOrganizerUserId='$OrganizerUserId',",
    "startDateTime=$StartDateTimeUtc,",
    "endDateTime=$EndDateTimeUtc",
    ")?`$top=100"
) -join ""

Write-Log "Searching for transcript-bearing online meetings in the requested UTC time range..."
$transcripts = Get-AllGraphPages -InitialUri $getAllTranscriptsUri -AccessToken $token

if (-not $transcripts -or $transcripts.Count -eq 0) {
    Write-Log "No transcripts were found for organizer $OrganizerUserId in the requested time range." "WARN"
    return
}

Write-Log ("Found {0} transcript object(s)." -f $transcripts.Count)

# Optional: cache meeting details by meetingId so repeated transcript objects for the same meeting
# do not cause repeated GET /onlineMeetings/{meetingId} calls.
$meetingCache = @{}

$indexRows = New-Object System.Collections.Generic.List[object]

foreach ($t in $transcripts) {
    try {
        $meetingIdEncoded    = [System.Uri]::EscapeDataString([string]$t.meetingId)
        $transcriptIdEncoded = [System.Uri]::EscapeDataString([string]$t.id)

        $meetingSubject   = $null
        $meetingStart     = $null
        $meetingEnd       = $null
        $joinWebUrl       = $null

        if ($IncludeMeetingDetails) {
            if (-not $meetingCache.ContainsKey([string]$t.meetingId)) {
                $meetingUri = "https://graph.microsoft.com/v1.0/users/$OrganizerUserId/onlineMeetings/$meetingIdEncoded"
                try {
                    $meetingObj = Invoke-GraphJsonGet -Uri $meetingUri -AccessToken $token
                    $meetingCache[[string]$t.meetingId] = $meetingObj
                }
                catch {
                    Write-Log "Could not resolve meeting details for meetingId '$($t.meetingId)'. Continuing without subject/start/end. Error: $($_.Exception.Message)" "WARN"
                    $meetingCache[[string]$t.meetingId] = $null
                }
            }

            $meetingObj = $meetingCache[[string]$t.meetingId]
            if ($null -ne $meetingObj) {
                $meetingSubject = $meetingObj.subject
                $meetingStart   = $meetingObj.startDateTime
                $meetingEnd     = $meetingObj.endDateTime
                $joinWebUrl     = $meetingObj.joinWebUrl
            }
        }

        $safeSubject = Get-SafeFileName -Name $meetingSubject
        $createdPart = if ($t.createdDateTime) {
            ([DateTime]$t.createdDateTime).ToString("yyyyMMdd_HHmmss")
        }
        else {
            "UnknownCreatedTime"
        }

        $transcriptFileName = "{0}__{1}__{2}.vtt" -f $createdPart, $safeSubject, (([string]$t.id).Substring(0, [Math]::Min(16, ([string]$t.id).Length)))
        $transcriptFilePath = Join-Path $TranscriptFolder $transcriptFileName

        # Download the transcript content as VTT
        $contentUri = "https://graph.microsoft.com/v1.0/users/$OrganizerUserId/onlineMeetings/$meetingIdEncoded/transcripts/$transcriptIdEncoded/content?`$format=text/vtt"
        $vtt = Invoke-GraphTextGet -Uri $contentUri -AccessToken $token

        # Save as UTF-8
        [System.IO.File]::WriteAllText($transcriptFilePath, $vtt, [System.Text.UTF8Encoding]::new($false))

        Write-Log "Saved transcript: $transcriptFilePath"

        $row = [PSCustomObject]@{
            OrganizerUserId        = $OrganizerUserId
            MeetingId              = [string]$t.meetingId
            TranscriptId           = [string]$t.id
            TranscriptCreatedUtc   = [string]$t.createdDateTime
            TranscriptEndUtc       = [string]$t.endDateTime
            MeetingSubject         = $meetingSubject
            MeetingStartUtc        = $meetingStart
            MeetingEndUtc          = $meetingEnd
            JoinWebUrl             = $joinWebUrl
            TranscriptContentUrl   = [string]$t.transcriptContentUrl
            SavedTranscriptPath    = $transcriptFilePath
        }

        [void]$indexRows.Add($row)
    }
    catch {
        Write-Log "Failed processing transcript id '$($t.id)'. Error: $($_.Exception.Message)" "ERROR"
    }
}

$csvPath = Join-Path $OutputRoot "TranscriptIndex.csv"
$indexRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Log "Done."
Write-Log "Transcript files folder: $TranscriptFolder"
Write-Log "Index CSV: $csvPath"
