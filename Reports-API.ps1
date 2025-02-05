##############
## Report API
function RunQueryandEnumerateResults {
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get)
    [array]$ResultsValue = $Results.value
    if ($results."@odata.nextLink" -ne $null) {
        $NextPageUri = $results."@odata.nextLink"
        While ($NextPageUri -ne $null) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }
    return $ResultsValue
}

##List All Tenant Users
$apiuri = "https://graph.microsoft.com/v1.0/users?`$select=id,userprincipalname,mail,displayname,givenname,surname,licenseAssignmentStates,proxyaddresses,usagelocation,usertype,accountenabled"
$users = RunQueryandEnumerateResults

##List all Tenant Groups
$apiuri = "https://graph.microsoft.com/v1.0/groups"
$Groups = RunQueryandEnumerateResults


##Get Teams details
$TeamGroups = $Groups | ? { ($_.grouptypes -Contains "unified") -and ($_.resourceProvisioningOptions -contains "Team") }


foreach ($teamgroup in $TeamGroups) {
    $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/allchannels"
    $Teamchannels = RunQueryandEnumerateResults
    $standardchannels = ($teamchannels | ? { $_.membershipType -eq "standard" })
    $privatechannels = ($teamchannels | ? { $_.membershipType -eq "private" })
    $outgoingsharedchannels = ($teamchannels | ? { ($_.membershipType -eq "shared") -and (($_."@odata.id") -like "*$($teamgroup.id)*") })
    $incomingsharedchannels = ($teamchannels | ? { ($_.membershipType -eq "shared") -and ($_."@odata.id" -notlike "*$($teamgroup.id)*") })
    $teamgroup | Add-Member -MemberType NoteProperty -Name "StandardChannels" -Value $standardchannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannels" -Value $privatechannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannels" -Value $outgoingsharedchannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "IncomingSharedChannels" -Value $incomingsharedchannels.id.count -Force
    $privatechannelSize = 0
    foreach ($Privatechannel in $privatechannels) {
        $PrivateChannelObject = $null
        $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/channels/$($Privatechannel.id)/FilesFolder"
        Try {
            $PrivateChannelObject = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get)
            $Privatechannelsize += $PrivateChannelObject.size
        } Catch {
            $Privatechannelsize += 0
        }
    }
    $sharedchannelSize = 0
    foreach ($sharedchannel in $outgoingsharedchannels) {
        $sharedChannelObject = $null
        $apiuri = "https://graph.microsoft.com/beta/teams/$($teamgroup.id)/channels/$($Sharedchannel.id)/FilesFolder"
        Try {
            $SharedChannelObject = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get)
            $Sharedchannelsize += $SharedChannelObject.size
        } Catch {
            $Sharedchannelsize += 0
        }
    }
    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannelsSize" -Value $privatechannelSize -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannelsSize" -Value $sharedchannelSize -Force
    $TeamDetails = $null
    $apiuri = "https://graph.microsoft.com/v1.0/groups/$($teamgroup.id)/drive/"
    $TeamDetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get)
    $teamgroup | Add-Member -MemberType NoteProperty -Name "DataSize" -Value $TeamDetails.quota.used -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "URL" -Value $TeamDetails.webUrl.replace("/Shared%20Documents", "") -Force
}


##Get All License SKUs
$apiuri = "https://graph.microsoft.com/v1.0/subscribedskus"
$SKUs = RunQueryandEnumerateResults


##Get Org Details
$apiuri = "https://graph.microsoft.com/v1.0/organization"
$OrgDetails = RunQueryandEnumerateResults

##List All Azure AD Service Principals
$apiURI = "https://graph.microsoft.com/beta/servicePrincipals"
[array]$AADApps = RunQueryandEnumerateResults

foreach ($user in $users) {
    $user | Add-Member -MemberType NoteProperty -Name "License SKUs" -Value ($user.licenseAssignmentStates.skuid -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Group License Assignments" -Value ($user.licenseAssignmentStates.assignedByGroup -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Disabled Plan IDs" -Value ($user.licenseAssignmentStates.disabledplans -join ";") -Force
}

##Translate License SKUs and groups
foreach ($user in $users) {
    foreach ($Group in $Groups) {$user.'Group License Assignments' = $user.'Group License Assignments'.replace($group.id, $group.displayName)}
    foreach ($SKU in $SKUs) {$user.'License SKUs' = $user.'License SKUs'.replace($SKU.skuid, $SKU.skuPartNumber)}
    foreach ($SKUplan in $SKUs.servicePlans) {$user.'Disabled Plan IDs' = $user.'Disabled Plan IDs'.replace($SKUplan.servicePlanId, $SKUplan.servicePlanName)}
}

##Get Conditional Access Policies
$apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
[array]$ConditionalAccessPolicies = RunQueryandEnumerateResults

##Get Directory Roles
$apiURI = "https://graph.microsoft.com/beta/directoryRoleTemplates"
[array]$DirectoryRoleTemplates = RunQueryandEnumerateResults

##Get Trusted Locations
$apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
[array]$NamedLocations = RunQueryandEnumerateResults


##Tidy GUIDs to names
$ConditionalAccessPoliciesJSON = $ConditionalAccessPolicies | ConvertTo-Json -Depth 5
if ($ConditionalAccessPoliciesJSON -ne $null) {
    ##TidyUsers
    foreach ($User in $Users) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($user.id, ("$($user.displayname) - $($user.userPrincipalName)"))
    }
    ##Tidy Groups
    foreach ($Group in $Groups) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($group.id, ("$($group.displayname) - $($group.id)"))
    }
    ##Tidy Roles
    foreach ($DirectoryRoleTemplate in $DirectoryRoleTemplates) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($DirectoryRoleTemplate.Id, $DirectoryRoleTemplate.displayname)
    }
    ##Tidy Apps
    foreach ($AADApp in $AADApps) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.appid, $AADApp.displayname)
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.id, $AADApp.displayname)
    }
    ##Tidy Locations
    foreach ($NamedLocation in $NamedLocations) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($NamedLocation.id, $NamedLocation.displayname)
    }
    $ConditionalAccessPolicies = $ConditionalAccessPoliciesJSON | ConvertFrom-Json
    $CAOutput = @()
    $CAHeadings = @(
        "displayName",
        "createdDateTime",
        "modifiedDateTime",
        "state",
        "Conditions.users.includeusers",
        "Conditions.users.excludeusers",
        "Conditions.users.includegroups",
        "Conditions.users.excludegroups",
        "Conditions.users.includeroles",
        "Conditions.users.excluderoles",
        "Conditions.clientApplications.includeServicePrincipals",
        "Conditions.clientApplications.excludeServicePrincipals",
        "Conditions.applications.includeApplications",
        "Conditions.applications.excludeApplications",
        "Conditions.applications.includeUserActions",
        "Conditions.applications.includeAuthenticationContextClassReferences",
        "Conditions.userRiskLevels",
        "Conditions.signInRiskLevels",
        "Conditions.platforms.includePlatforms",
        "Conditions.platforms.excludePlatforms",
        "Conditions.locations.includLocations",
        "Conditions.locations.excludeLocations"
        "Conditions.clientAppTypes",
        "Conditions.devices.deviceFilter.mode",
        "Conditions.devices.deviceFilter.rule",
        "GrantControls.operator",
        "grantcontrols.builtInControls",
        "grantcontrols.customAuthenticationFactors",
        "grantcontrols.termsOfUse",
        "SessionControls.disableResilienceDefaults",
        "SessionControls.applicationEnforcedRestrictions",
        "SessionControls.persistentBrowser",
        "SessionControls.cloudAppSecurity",
        "SessionControls.signInFrequency"
    )

    Foreach ($Heading in $CAHeadings) {
        $Row = $null
        $Row = New-Object psobject -Property @{
            PolicyName = $Heading
        }
    
        foreach ($CAPolicy in $ConditionalAccessPolicies) {
            $Nestingcheck = ($Heading.split('.').count)

            if ($Nestingcheck -eq 1) {
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value $CAPolicy.$Heading -Force
            }
            elseif ($Nestingcheck -eq 2) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 3) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 4) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()).($SplitHeading[3].ToString()) -join ';' )-Force       
            }
        }

        $CAOutput += $Row

    }

}

#######################
##Get OneDrive Report##
$apiUri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')"
$OneDrive = ((Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get) | ConvertFrom-Csv)

#########################
##Get SharePoint Report##
$apiUri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
$SharePoint = ((Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get) | ConvertFrom-Csv)
$SharePoint | Add-Member -MemberType NoteProperty -Name "TeamID" -Value "" -force
foreach ($Site in $Sharepoint) {

    $TeamID = $null
    $Site.TeamID = ($TeamGroups | ? { $_.url -contains $site.'site url' }).id


}

######################
##Get Mailbox Report##
$apiUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D30')"
$MailboxStatsReport = ((Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get) | ConvertFrom-Csv)

##############################
##Get M365 Apps usage report##
$apiUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='D30')"
$M365AppsUsage = ((Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get) | ConvertFrom-Csv)

##Optional Items - Graph

##Process Group Membership
If ($IncludeGroupMembership) {
    $GroupMembersObject = @()
    foreach ($group in $groups) {
        $apiuri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members"
        $Members = RunQueryandEnumerateResults
        foreach ($member in $members) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.displayname
                MemberUserPrincipalName = $member.userprincipalname
                MemberType              = "Member"
                MemberObjectType        = $member.'@odata.type'.replace('#microsoft.graph.', '')

            }
            $GroupMembersObject += $memberEntry
        }

        $apiuri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners"
        $Owners = RunQueryandEnumerateResults

        foreach ($member in $Owners) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.displayname
                MemberUserPrincipalName = $member.userprincipalname
                MemberType              = "Owner"
                MemberObjectType        = $member.'@odata.type'.replace('#microsoft.graph.', '')

            }

            $GroupMembersObject += $memberEntry

        }
    }


    
}

##Tidy up Proxyaddresses
foreach ($user in $users) {
    $user.proxyaddresses = $user.proxyaddresses -join ';'
}
##Tidy up Proxyaddresses
foreach ($group in $groups) {
    $group.proxyaddresses = $group.proxyaddresses -join ';'
}



###################EXCHANGE ONLINE############################
Try {
    Connect-ExchangeOnline -Certificate $Certificate -AppID $clientid -Organization ($orgdetails.verifieddomains | ? { $_.isinitial -eq "true" }).name -ShowBanner:$false
}
catch {
    write-host "Error connecting to Exchange Online...Exiting..." -ForegroundColor red
    Pause
    Exit
}

##Get Shared and Resource Mailboxes
[array]$RoomMailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
[array]$EquipmentMailboxes = Get-EXOMailbox -RecipientTypeDetails EquipmentMailbox -ResultSize unlimited
[array]$SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited


##Get Resource Mailbox Sizes
foreach ($room in $RoomMailboxes) {
    $ProgressStatus = "Getting room mailbox statistics $i of $($RoomMailboxes.count)..."
    $i++
    UpdateProgress

    $RoomStats = $null
    $RoomStats = get-EXOmailboxstatistics $room.primarysmtpaddress
    $room | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $RoomStats.TotalItemSize -Force
    $room | Add-Member -MemberType NoteProperty -Name ItemCount -Value $RoomStats.ItemCount -Force

    ##Clean email addresses value
    $room.EmailAddresses = $room.EmailAddresses -join ';'
}


foreach ($equipment in $EquipmentMailboxes) {
    $ProgressStatus = "Getting Equipment mailbox statistics $i of $($EquipmentMailboxes.count)..."
    $i++
    UpdateProgress

    $EquipmentStats = $null
    $EquipmentStats = get-EXOmailboxstatistics $equipment.primarysmtpaddress
    $equipment | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $EquipmentStats.TotalItemSize -Force
    $equipment | Add-Member -MemberType NoteProperty -Name ItemCount -Value $EquipmentStats.ItemCount -Force

    ##Clean email addresses value
    $equipment.EmailAddresses = $equipment.EmailAddresses -join ';'
}


##Get Shared Mailbox Sizes
foreach ($SharedMailbox in $SharedMailboxes) {
    $ProgressStatus = "Getting shared mailbox statistics $i of $($SharedMailboxes.count)..."
    $i++
    UpdateProgress

    $SharedStats = $null
    $SharedStats = get-EXOmailboxstatistics $SharedMailbox.primarysmtpaddress
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $SharedStats.TotalItemSize -Force
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name ItemCount -Value $SharedStats.ItemCount -Force
    
    ##Clean email addresses value
    $SharedMailbox.EmailAddresses = $SharedMailbox.EmailAddresses -join ';'
}

$ProgressStatus = "Getting user mailbox statistics..."
UpdateProgress


##Collect Mailbox statistics
$MailboxStats = @()
foreach ($user in ($users | ? { ($_.mail -ne $null ) -and ($_.userType -eq "Member") })) {
    $stats = $null
    $stats = $MailboxStatsReport | ? { $_.'User Principal Name' -eq $user.userprincipalname }
    $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $user.id -Force
    $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $user.mail -Force
    $MailboxStats += $stats
    

}

##Collect Archive Statistics
$ArchiveStats = @()
[array]$ArchiveMailboxes = get-EXOmailbox -Archive -ResultSize unlimited
foreach ($archive in $ArchiveMailboxes) {
    $ProgressStatus = "Getting archive mailbox statistics $i of $($ArchiveMailboxes.count)..."
    $i++
    UpdateProgress
    $stats = $null
    $stats = get-EXOmailboxstatistics $archive.PrimarySmtpAddress -Archive #-erroraction SilentlyContinue
    $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $archive.ExternalDirectoryObjectId -Force
    $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $archive.primarysmtpaddress -Force
    $ArchiveStats += $stats
    
}


##Collect Mail Contacts
##Collect transport rules

$MailContacts = Get-MailContact -ResultSize unlimited | select displayname, alias, externalemailaddress, emailaddresses, HiddenFromAddressListsEnabled
foreach ($mailcontact in $MailContacts) {
    $mailcontact.emailaddresses = $mailcontact.emailaddresses -join ';'
}


##Collect transport rules

$Rules = $null
[array]$Rules = Get-TransportRule -ResultSize unlimited | select name, state, mode, priority, description, comments
$RulesOutput = @()
##Output rules to variable
foreach ($Rule in $Rules) {

    $RulesOutput += $Rule

}


#######Optional Items - EXO#######

##Process Mailbox Permissions
If ($IncludeMailboxPermissions) {
    $PermissionOutput = @()
    ##Get all mailboxes
    $MailboxList = Get-EXOMailbox -ResultSize unlimited
    $PermissionProgress = 1
    foreach ($mailbox in $MailboxList) {
        [array]$Permissions = Get-EXOMailboxPermission -UserPrincipalName $mailbox.UserPrincipalName | ? { $_.User -ne "NT AUTHORITY\SELF" }
        foreach ($permission in $Permissions) {
            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.user

            }
            $PermissionOutput += $PermissionObject
        }
        [array]$RecipientPermissions = Get-EXORecipientPermission $mailbox.UserPrincipalName |  ? { $_.Trustee -ne "NT AUTHORITY\SELF" }
        foreach ($permission in $RecipientPermissions) {
            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.trustee

            }
            $PermissionOutput += $PermissionObject
        }
    }
}

#######Report Export#######


##Collect Mailflow Connectors

$InboundConnectors = Get-InboundConnector | select enabled, name, connectortype, connectorsource, SenderIPAddresses, SenderDomains, RequireTLS, RestrictDomainsToIPAddresses, RestrictDomainsToCertificate, CloudServicesMailEnabled, TreatMessagesAsInternal, TlsSenderCertificateName, EFTestMode, Comment 
foreach ($inboundconnector in $InboundConnectors) {
    $inboundconnector.senderipaddresses = $inboundconnector.senderipaddresses -join ';'
    $inboundconnector.senderdomains = $inboundconnector.senderdomains -join ';'
}
$OutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors:$true | select enabled, name, connectortype, connectorsource, TLSSettings, RecipientDomains, UseMXRecord, SmartHosts, Comment
foreach ($OutboundConnector in $OutboundConnectors) {
    $OutboundConnector.RecipientDomains = $OutboundConnector.RecipientDomains -join ';'
    $OutboundConnector.SmartHosts = $OutboundConnector.SmartHosts -join ';'
}
$ProgressStatus = "Getting MX records..."
UpdateProgress


##MX Record Check
$MXRecordsObject = @()
foreach ($domain in $orgdetails.verifieddomains) {
    Try {
        [array]$MXRecords = Resolve-DnsName -Name $domain.name -Type mx -ErrorAction SilentlyContinue
    }
    catch {
        write-host "Error obtaining MX Record for $($domain.name)"
    }
    foreach ($MXRecord in $MXRecords) {
        $MXRecordsObject += $MXRecord
    }
}

$ProgressStatus = "Updating references..."
UpdateProgress


##Update users tab with Values
$users | Add-Member -MemberType NoteProperty -Name MailboxSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name MailboxItemCount -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveFileCount -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name Mailboxtype -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveItemCount -Value "" -Force

foreach ($user in ($users | ? { $_.usertype -ne "Guest" })) {
    ##Set Mailbox Type
    if ($roommailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Room"
    }    
    elseif ($EquipmentMailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Equipment"
    }
    elseif ($sharedmailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Shared"
    }
    else {
        $user.Mailboxtype = "User"
    }

    ##Set Mailbox Size and count
    If ($MailboxStats | ? { $_.objectID -eq $user.id }) {
        $user.MailboxSizeGB = (((($MailboxStats | ? { $_.objectID -eq $user.id }).'Storage Used (Byte)' / 1024) / 1024) / 1024) 
        $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
        $user.MailboxItemCount = ($MailboxStats | ? { $_.objectID -eq $user.id }).'item count'
    }

    ##Set Shared Mailbox size and count
    If ($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set Equipment Mailbox size and count
    If ($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }


    ##Set Room Mailbox size and count
    If ($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set archive size and count
    If ($ArchiveStats | ? { $_.objectID -eq $user.id }) {
        $user.ArchiveSizeGB = (((($ArchiveStats | ? { $_.objectID -eq $user.id }).totalitemsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
        $user.ArchiveSizeGB = [math]::Round($user.ArchiveSizeGB, 2)
        $user.ArchiveItemCount = ($ArchiveStats | ? { $_.objectID -eq $user.id }).ItemCount
    }

    ##Set OneDrive Size and count
    if ($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }) {
        if (($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'Storage Used (Byte)') {
            If ($user.OneDriveSizeGB ) {
                $user.OneDriveSizeGB = (((($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'Storage Used (Byte)' / 1024) / 1024) / 1024)
                $user.OneDriveSizeGB = [math]::Round($user.OneDriveSizeGB, 2)
                $user.OneDriveFileCount = ($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'file count'
            }
        }
    }
}





