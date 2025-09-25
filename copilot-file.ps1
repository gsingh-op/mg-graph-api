Import-Module Microsoft.Graph.Devices.ServiceAnnouncement

# Instructions:
# Comment out the code you don't need to run, and run the code you need to run.

# filters = isMasjorChange = true
# filters = services/any(s: contains(s, 'Copilot'))

# 1. Get the Top 3 service announcement message for the Copilot service

Get-MgServiceAnnouncementMessage -Filter "services/any(s: contains(s, 'Copilot'))" -All | 
    Select-Object -Property @{Name='LastModifiedDateTime'; Expression={if ($_.lastModifiedDateTime) { [DateTime]::Parse($_.lastModifiedDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                            @{Name='StartDateTime'; Expression={if ($_.startDateTime) { [DateTime]::Parse($_.startDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                            @{Name='EndDateTime'; Expression={if ($_.endDateTime) { [DateTime]::Parse($_.endDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
                            
                           title,
                           id,
                           category, 
                           severity,
                           isMajorChange,
                        #    tags,
                           @{Name='Tags'; Expression={($_.tags -join ", ")}},
                           actionRequiredByDateTime,
                        #    services,
                           @{Name='Services'; Expression={($_.services -join ", ")}},
                        #    details,
                           @{Name='RoadmapIds'; Expression={($_.details | Where-Object {$_.Name -eq "RoadmapIds"}).Value}},
                           @{Name='Summary'; Expression={($_.details | Where-Object {$_.Name -eq "Summary"}).Value}},
                           @{Name='Platforms'; Expression={($_.details | Where-Object {$_.Name -eq "Platforms"}).Value}} |
                        #    @{Name='link'; Expression={("https://mc.merill.net/message/{0}" -f $_.id)}}
                        #    body                
    Sort-Object lastModifiedDateTime -Descending | 
    ConvertTo-Json |
    Out-File -FilePath "copilot-announcements-all.json"

# filter out where lastmodifiedtime > Jun 2025
# Get-Content "copilot-announcements-all.json" | 
#     ConvertFrom-Json |
#     Where-Object { [DateTime]::Parse($_.LastModifiedDateTime) -gt [DateTime]::Parse("June 1, 2025") } |
#     ConvertTo-Json |
#     Out-File -FilePath "copilot-announcements-since-june.json"


# 1b. Filter by specific RoadmapId
# $targetRoadmapId = "478611"  # Change this to the roadmap ID you want to filter by
# $filteredResults = Get-MgServiceAnnouncementMessage | 
#     Where-Object { 
#         ($_.details | Where-Object {$_.Name -eq "RoadmapIds"}).Value -like "*$targetRoadmapId*" 
#     } |
#     Select-Object -Property @{Name='StartDateTime'; Expression={if ($_.startDateTime) { [DateTime]::Parse($_.startDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
#                            @{Name='EndDateTime'; Expression={if ($_.endDateTime) { [DateTime]::Parse($_.endDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
#                            @{Name='LastModifiedDateTime'; Expression={if ($_.lastModifiedDateTime) { [DateTime]::Parse($_.lastModifiedDateTime).ToString("MMMM dd, yyyy") } else { $null }}},
#                            title,
#                            id,
#                            category, 
#                            severity,
#                            isMajorChange,
#                            @{Name='Tags'; Expression={($_.tags -join ", ")}},
#                            actionRequiredByDateTime,
#                            @{Name='Services'; Expression={($_.services -join ", ")}},
#                            @{Name='RoadmapIds'; Expression={($_.details | Where-Object {$_.Name -eq "RoadmapIds"}).Value}},
#                            @{Name='Summary'; Expression={($_.details | Where-Object {$_.Name -eq "Summary"}).Value}},
#                            @{Name='Platforms'; Expression={($_.details | Where-Object {$_.Name -eq "Platforms"}).Value}},
#                            @{Name='Source'; Expression={"Message Center"}} |
#     Sort-Object lastModifiedDateTime -Descending

# if ($filteredResults) {
#     $filteredResults | ConvertTo-Json
# } else {
#     Write-Output "Not found for RoadmapId: $targetRoadmapId"
# }

# 2. Get 10 service announcement messages for the Copilot service, convert to JSON and save to file
# Get-MgServiceAnnouncementMessage -Filter "services/any(s: contains(s, 'Copilot'))" | 
# Select-Object -Property title,
#                        id,
#                        category, 
#                        severity,
#                        isMajorChange,
#                        tags,
#                        services,
#                        actionRequiredByDateTime
#                     #    details,
#                     #    body 
#                        | 
# ConvertTo-Json |
# Out-File -FilePath "copilot-announcements.json"

# 3. filter by id, convert to JSON
# Get-MgServiceAnnouncementMessage -Filter "id eq 'MC1152323'" |
# Select-Object -Property title,
#                        id,
#                        category, 
#                        severity,
#                        isMajorChange,
#                        tags,
#                        services,
#                        actionRequiredByDateTime,
#                        details
#                        | ConvertTo-Json

# 4. Filter messages by RoadmapId

# $roadmapId = "501451"

# Get-MgServiceAnnouncementMessage -Filter "details/any(d: d/value eq '501451')" |
# Select-Object -Property title,
#                        id,
#                        category, 
#                        severity,
#                        isMajorChange,
#                        tags,
#                        services,
#                        actionRequiredByDateTime,
#                        details |
# ConvertTo-Json


# Get all service announcement messages
# $messages = Get-MgServiceAnnouncementMessage

# # Filter messages where Details contains RoadmapIds with the target value
# $filteredMessages = $messages | Where-Object {
#     $_.Details -match $null -or $_.Details | Where-Object {
#         $_.Name -eq "RoadmapIds" -and $_.Value -eq $roadmapId
#     }
# }

# # Display the filtered messages
# $filteredMessages
