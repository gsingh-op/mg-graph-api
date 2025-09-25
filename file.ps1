# Get the messages
$messages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages"


# # Select and display only the first message with specific fields
$messages.value | Where-Object {$_.id -eq "MC1152323"} | Select-Object -Property `
    title,
    id,
    category,
    severity,
    isMajorChange
    # @{Name="services"; Expression={($_.services -join ", ")}},
    # @{Name="RoadmapIds"; Expression={($_.details | Where-Object {$_.name -eq "RoadmapIds"}).value}},
    # @{Name="Summary"; Expression={($_.details | Where-Object {$_.name -eq "Summary"}).value}},
    # @{Name="Platforms"; Expression={($_.details | Where-Object {$_.name -eq "Platforms"}).value}} | ConvertTo-Json
    # @{Name="ExternalLink"; Expression={($_.details | Where-Object {$_.name -eq "ExternalLink"}).value}} | ConvertTo-Json
    # @{Name="body"; Expression={$_.body.content}}



# Get all distinct severity from all messages
# $messages.value | Select-Object -Property severity -Unique | Sort-Object severity







# Get all distinct categories from all messages
# $messages.value | Select-Object -Property severity -Unique | Sort-Object category





# Extract specific details: RoadmapIds, Summary, and Platforms in JSON format
# $messages.value | Select-Object -First 1 -Property `
#     @{Name="RoadmapIds"; Expression={($_.details | Where-Object {$_.name -eq "RoadmapIds"}).value}},
#     @{Name="Summary"; Expression={($_.details | Where-Object {$_.name -eq "Summary"}).value}},
#     @{Name="Platforms"; Expression={($_.details | Where-Object {$_.name -eq "Platforms"}).value}} | ConvertTo-Json