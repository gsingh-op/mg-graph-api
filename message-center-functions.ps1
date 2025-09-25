Import-Module Microsoft.Graph.Devices.ServiceAnnouncement

# Function to format service announcement messages with consistent structure
function Format-ServiceAnnouncementMessage {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )
    
    process {
        $InputObject | Select-Object -Property @{Name='StartDateTime'; Expression={if ($_.startDateTime) { [DateTime]::Parse($_.startDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                                               @{Name='EndDateTime'; Expression={if ($_.endDateTime) { [DateTime]::Parse($_.endDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                                               @{Name='LastModifiedDateTime'; Expression={if ($_.lastModifiedDateTime) { [DateTime]::Parse($_.lastModifiedDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                                               title,
                                               id,
                                               category, 
                                               severity,
                                               isMajorChange,
                                               @{Name='Tags'; Expression={($_.tags -join ", ")}},
                                               actionRequiredByDateTime,
                                               @{Name='Services'; Expression={($_.services -join ", ")}},
                                               @{Name='RoadmapIds'; Expression={($_.details | Where-Object {$_.Name -eq "RoadmapIds"}).Value}},
                                               @{Name='Summary'; Expression={($_.details | Where-Object {$_.Name -eq "Summary"}).Value}},
                                               @{Name='Platforms'; Expression={($_.details | Where-Object {$_.Name -eq "Platforms"}).Value}},
                                               @{Name='Source'; Expression={"Message Center"}}
    }
}

# Function to get Copilot service announcements with major changes
function Get-CopilotMajorChanges {
    param(
        [int]$Top = 3
    )
    
    Get-MgServiceAnnouncementMessage -Filter "services/any(s: contains(s, 'Copilot')) and isMajorChange eq true" -Top $Top |
        Format-ServiceAnnouncementMessage |
        Sort-Object LastModifiedDateTime -Descending
}

# Function to filter by specific RoadmapId
function Get-ServiceAnnouncementByRoadmapId {
    param(
        [string]$RoadmapId
    )
    
    $filteredResults = Get-MgServiceAnnouncementMessage | 
        Where-Object { 
            ($_.details | Where-Object {$_.Name -eq "RoadmapIds"}).Value -like "*$RoadmapId*" 
        } |
        Format-ServiceAnnouncementMessage |
        Sort-Object LastModifiedDateTime -Descending

    if ($filteredResults) {
        return $filteredResults
    } else {
        Write-Output "Not found for RoadmapId: $RoadmapId"
        return $null
    }
}

# Usage Examples:

# 1. Get Top 3 Copilot major changes
Write-Host "=== Top 3 Copilot Major Changes ===" -ForegroundColor Green
$results = Get-CopilotMajorChanges -Top 3
$results | ConvertTo-Json

Write-Host "`n=== Filter by RoadmapId ===" -ForegroundColor Green
# 2. Filter by specific RoadmapId
$targetRoadmapId = "478611"  # Change this to the roadmap ID you want to filter by
$roadmapResults = Get-ServiceAnnouncementByRoadmapId -RoadmapId $targetRoadmapId
if ($roadmapResults) {
    $roadmapResults | ConvertTo-Json
}
