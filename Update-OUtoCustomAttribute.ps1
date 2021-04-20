<#
.NOTES
	Name: Update-OUtoCustomAttribute.ps1
	Author: Daniel Sheehan
	Requires: PowerShell v3 or higher, the Active Directory module, and the
	account running this script needs to have permissions to modify the
	custom attribute that will house the OU/Container information.
	Version History:
	1.0 - 6/26/2018 - Initial release.
    1.1 - 10/10/2018 - Added exclusion for O365/Modern Group write-back objects.
	############################################################################
	The sample scripts are not supported under any Microsoft standard support
	program or service. The sample scripts are provided AS IS without warranty
	of any kind. Microsoft further disclaims all implied warranties including,
	without limitation, any implied warranties of merchantability or of fitness
	for a particular purpose. The entire risk arising out of the use or
	performance of the sample scripts and documentation remains with you. In no
	event shall Microsoft, its authors, or anyone else involved in the creation,
	production, or delivery of the scripts be liable for any damages whatsoever
	(including, without limitation, damages for loss of business profits,
	business interruption, loss of business information, or other pecuniary
	loss) arising out of the use of or inability to use the sample scripts or
	documentation, even if Microsoft has been advised of the possibility of such
	damages.
	############################################################################
.SYNOPSIS
	For Exchange hybrid deployments, this script populates the OU/Container path
	for all mailboxes and mail-enabled groups in the specified domains.
.DESCRIPTION
	In order to facilitate management of synchronized mailboxes and mail-enabled
	groups in Exchange Online (EXO), based upon their OU/Container path in the
	on-premises AD environment, that OU/Container information has to be
	synchronized somehow into Azure AD and EXO so that custom write scopes can
	be defined. This script populates that information, and keeps it up to date
	if run regularly, into the designated Custom Attribute of all mailboxes and
	mail-enabled groups in the targeted domains.
	By default all domains in an AD forest are targeted, but one or more domains
	can be targeted using the supplied Domains parameter.
.PARAMETER Domains
	Optional: Specify one or more AD Domain names, comma separated with short or
	FQDN names, to process by this script. If no domains are specified, then all
	domains in the environment will be processed.
.PARAMETER WhatIf
	Optional: Don't make any changes to AD objects, just simulate what
	would happen if the changes were made.
.PARAMETER OutCSVFile
	Optional: Output all the modified mailboxes and
	their details to the specified CSV file. The details recorded for each are
	DisplayName, Updated, UPN, AccountStatus, WhenChanged, WhenCreated, and
	Domain.
.EXAMPLE
	Update-OUtoCsutomAttribute.ps1 -OutCSVFile .\UpdatedOUMappings.CSV
	All mailboxes and mail-enabled groups across all domains will have their
	OU/Container compared to the value in the specified custom attribute, with
	changes being made as necessary, and the results will be exported to the CSV
	file in the local directory.
.EXAMPLE
	Update-OUtoCsutomAttribute.ps1 -WhatIf -Verbose
	All mailboxes and mail-enabled groups across all domains will have their
	OU/Container value compared against what is int he specified custom
	attribute, with the Verbose information being written to the screen, but
	they will not be modified.
.EXAMPLE
	Update-OUtoCustomAttribute.ps1 -Domain contoso.com
	All mailboxes and mail-enabled groups in the cotoso.com domain will have
	their OU/Container compared to the value in the specified custom attribute,
	with changes being made as necessary.
.LINK
	https://gallery.technet.microsoft.com/Tracking-and-Controlling-f2b14d0f
#>

#Requires -Version 3.0

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $False)]
	[String[]]$Domains,
	[Parameter(Mandatory = $False)]
	[Switch]$WhatIf,
	[Parameter(Mandatory = $False)]
	[String]$OutCSVFile
)

# --- Begin Defined Variables ---
# Define the custom attribute to store the object OU path, using the attribute name in AD. I.E. Custom Attribute 1 = extensionAttribute1
$CustomAttribute = "extensionAttribute7"
# --- End Defined Variables ---

# Start tracking the time this script takes to run.
$StopWatch = New-Object System.Diagnostics.Stopwatch
$StopWatch.Start()

# Check to see if the ActiveDirectory module is not already loaded as it.
If (-not(Get-Module "ActiveDirectory")) {
	# Its not, so try to load it.
	Try {
		Write-Host -ForegroundColor Cyan "Importing the Active Directory PowerShell module."
		Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$False
	# See if an error was caught loading it.
	} Catch {
		# There was an error so report it, stop the stopwatch, and exit the script.
		Write-Host -ForegroundColor Red "There was an error importing the Active Directory module, which is required for the script to function."
		$StopWatch.Stop()
		EXIT
	}
}

# If one or more domains was specified in the parameter, than extract their FQDN and store them in the AllDomains variable.
If ($Domains) {
	$AllDomains = ForEach ($DomainEntry in $Domains) {
		(Get-ADDomain $DomainEntry).DNSRoot
	}
# Otherwise gather all domains in the forest.
} Else {
	$AllDomains = (Get-ADForest).Domains
}

# Establish some tracking variables.
[System.Collections.ArrayList]$AllObjects = @()
$AllObjectsCount = 0
$AllNewCount = 0
$AllUpdateCount = 0

# If the -Verbose parameter was used, then save the default foreground text color and then change it to Cyan and enable account status checking.
If ($PSBoundParameters["Verbose"]) {
	$VerboseForeground = $Host.PrivateData.VerboseForegroundColor
	$Host.PrivateData.VerboseForegroundColor = "Cyan"
}

# Loop through each domain in the collection.
Write-Host ""
ForEach ($Domain in $AllDomains) {
	Write-Host "Processing objects in the domain `"$Domain`"." -ForegroundColor Green
	# Find a local DC for the domain, even if it is in the next site.
	[String]$DomainDC = (Get-ADDomainController -DomainName $Domain -Discover -NextClosestSite).HostName

    # Gather all user and groups in the domain that are Exchange enabled (ProxyAddresses is populated), except for HealthMailboxes and Modern/O365 Groups.
    $Objects = Get-ADObject -LDAPFilter "(&(ProxyAddresses=*)(|(&(ObjectClass=User)(ObjectCategory=person)(!(name=HealthMailbox*)))(&(ObjectClass=Group)(!(msExchRecipientTypeDetails=8796093022208)))))" -Properties canonicalName,DisplayName,$CustomAttribute,cn,whenChanged,whenCreated,mail -Server $DomainDC
	# Set some tracking information per loop.
	$ObjectsCount = ($Objects | Measure-Object).Count
	$AllObjectsCount += $ObjectsCount
	$NewCount = 0
	$UpdateCount = 0

	# Process each object in the collection.
	ForEach ($Object in $Objects) {
		# Validate the object has a Canonical Name. This is supposed to be dynamically generated by AD, but once any a while it is not and the script can't process the object if it is missing.
		If ($Object.canonicalName) {
			# Extract the CN attribute and escape any forward slashes so the format matches what's in the CanonicalName attribute.
			$CN = $Object.CN.ToString().Replace("/","\/")
			# Extract the Parent OU/Container by taking the CanonincalName attribute and stripping out the object's CN.
			$ParentObject = $Object.canonicalName.ToString().Replace($CN,"")

			<# Only uncomment this section if you have OU/Container paths that are longer than 448 characters (HIGHLY UNLIKELY).
			# Check to see if the string is longer than 448 characters, and if so trim it as Azure AD only supports 448 characters for custom attributes.
			If ($ParentObject.Length -gt 448) {
				$ParentObject = $ParentObject.Substring(0,448)
			}
			#>

			# Check to see if there is nothing set in the custom attribute field, and if not note it will be an Add operation.
			If (-not($Object.$CustomAttribute)) {
				$UpdateAction = "Add"
				$ExistingEntry = "N/A"
				Write-Verbose "$($Object.canonicalName) | New | WhenCreated: $($Object.WhenCreated)"
			# Otherwise check to see if what is in the custom attribute field is already set to the ParentObject value, and if not note it will be an Update operation.
			} ElseIf ($Object.$CustomAttribute -notlike $ParentObject) {
				$UpdateAction = "Update"
				$ExistingEntry = "$($Object.$CustomAttribute)"
				Write-Verbose "$($Object.canonicalName) | $ParentObject | WhenChanged: $($Object.WhenChanged)"
			# Otherwise there is no action to take as the account already has the correct value in the custom attribute field.
			} Else {
				$UpdateAction = $Null
			}

			# If there is an action to take, then proceed forward.
			If ($UpdateAction) {

				# Check to see if the WhatIf switch was used, and if so mark that as the status as WhatIf.
				If ($WhatIf) {
					$UpdateStatus = "WhatIf"
				# Otherwise try to perform the actual modification to the custom attribute regardless of it its an Add or Update operation.
				} Else {
	   				$UpdateStatus = "Successful"
					Try {
						Set-ADObject $Object -Replace @{$CustomAttribute=$ParentObject} -ErrorAction Stop -Server $DomainDC
					} Catch {
						$UpdateStatus = "Failed"
						Write-Host -ForegroundColor Red "There was an error modifying the custom attribute for the object `"$($Object.DisplayName)`" with the error code:"
						Write-Host -ForegroundColor Red "$($_.Exception)"
					}
				}

				# Check to see if the modification didn't fail (meaning it was a success or WhatIf was used) and update the appropriate numbers.
				If (-not($UpdateStatus -eq "Failed")) {
					If ($UpdateAction -eq "Update") {
						$UpdateCount++
					} Else {
						$NewCount++
					}
				}
			}

		# Otherwise the whole operation was skipped due to a lack of a canonicaName and record that as appropriate.
		} Else {
			$UpdateStatus = "Skipped"
			$UpdateAction = "Missing CanonicalName"
			$ParentObject = "N/A"
			$ExistingEntry = "N/A"
			Write-Warning "The object `"$($Object.DistinguishedName)`" did not have an AD generated CanonicalName. Please review this account and address this issue."
		}

		# If there was an UpdateAction recorded AND the objects are supposed to be output to a CSV file, then record them to the AllObjects array.
		If ($UpdateAction -and $OutCSVFile) {
			[Void]$AllObjects.Add([PSCustomObject]@{
				DisplayName = $Object.DisplayName
				Action = $UpdateAction
				Updated = $UpdateStatus
				NewEntry = $ParentObject
				ExistingEntry = $ExistingEntry
				EmailAddress = $Object.mail
				ObjectClass = $Object.ObjectClass
				WhenChanged = $Object.whenChanged
				WhenCreated = $Object.whenCreated
				Domain = $Domain
			})
		}
	}

	# Report the total objects that were processed for the domain.
	If ($WhatIf) {
		Write-Host "There would have been $NewCount unprocessed objects updated, and $UpdateCount previously processed objects updated, out of $ObjectsCount objects in the `"$Domain`" domain."
	} Else {
		Write-Host "There were $NewCount unprocessed objects updated, and $UpdateCount previously processed objects updated, out of $ObjectsCount objects in the `"$Domain`" domain."
	}
	# Add the domain counts into the overall counts for reporting at the end.
	$AllNewCount += $NewCount
	$AllUpdateCount += $UpdateCount
	Write-Host ""
}

# Report the total mailboxes that were processed for all domains.
Write-Host ""
If ($WhatIf) {
	Write-Host "There would have been $AllNewCount unprocessed objects updated, and $AllUpdateCount previously processed objects updated, out of $AllObjectsCount objects across all $($Domain.Count) domains."
} Else {
	Write-Host "There were $AllNewCount unprocessed objects updated, and $AllUpdateCount previously processed objects updated, out of $AllObjectsCount objects across all $($Domain.Count) domains."
}

# If the OutCSVFile switch was used, then export the information to the target file.
If ($OutCSVFile) {
	$AllObjects | Export-CSV -NoTypeInformation $OutCSVFile
	Write-Host ""
	Write-Host -ForegroundColor Green "The collected data was exported to the `"$OutCSVFile`" CSV file."
}

# If the VerboseForeground variable was set at the top of the script, then change the color back to the default.
If ($VerboseForeground) {
	$Host.PrivateData.VerboseForegroundColor = $VerboseForeground
}

# Calculate the amount of time the script took to run and write the information to the screen.
$StopWatch.Stop()
$ElapsedTime = $StopWatch.Elapsed
$TotalHours = ($ElapsedTime.Days * 24) + $ElapsedTime.Hours
Write-Host ""
Write-Host "The script took $TotalHours hour(s), $($ElapsedTime.Minutes) minute(s), and $($ElapsedTime.Seconds) second(s) to run."