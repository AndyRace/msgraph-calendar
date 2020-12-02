# NOTE: The following does NOT use batching (see: https://docs.microsoft.com/en-us/graph/json-batching)
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $UserId,    # e.g. andy.race@myazureadtenant.com

    # The following filter is applied to identify the events to be updated
    # See: https://docs.microsoft.com/en-us/graph/query-parameters#filter-parameter
    [Parameter(Mandatory=$false)]
    [string]
    $EventFilter = "Start/dateTime lt '$((Get-Date).Date.ToString("s"))'"
)

$VerbosePreference = "SilentlyContinue"
#$ErrorActionPreference = "Stop"
$ErrorActionPreference = "Break"
Set-StrictMode -Version Latest
Set-PSDebug -Strict

# See:
#   https://docs.microsoft.com/en-us/graph/use-the-api
#   https://docs.microsoft.com/en-us/graph/query-parameters
#   https://docs.microsoft.com/en-us/graph/powershell/installation
#   https://docs.microsoft.com/en-us/graph/powershell/get-started
#   https://docs.microsoft.com/en-us/graph/powershell/navigating
#   https://docs.microsoft.com/en-us/graph/api/resources/calendar?view=graph-rest-1.0
#   https://github.com/microsoftgraph/msgraph-sdk-powershell

# Install-Module Microsoft.Graph

# ('Microsoft.Graph') | ForEach-Object {
#     $module = $_
#     If (Get-Module -Name $module) {
#         return
#     }
#     elseif (Get-Module -ListAvailable -Name $module) {
#         Import-Module -name $module -Scope Local -Force
#     } else {
#         Install-module -name $module -AllowClobber -Force -Scope CurrentUser -SkipPublisherCheck
#         Import-Module -name $module -Scope Local -Force
#     }

#     If (!$(Get-Module -Name $module)) {
#         Write-Error "Could not load dependant module: $module"
#         throw
#     }
# }

# See: https://docs.microsoft.com/en-us/graph/api/resources/calendar?view=graph-rest-1.0
Connect-MgGraph -Scopes "User.Read", "Calendars.ReadWrite" | Out-Null #"User.Read.All"

Select-MgProfile -Name "beta"

# Get-MgUser_List: Insufficient privileges to complete the operation.
# Get-MgUser

# List available commands
# Get-Command -Module Microsoft.Graph* *calendar*

$cal = Get-MgUserCalendar -UserId $UserId

$skip = 0
$pageSize = 100
$count = 0

# Set historic events to be private
$Sensitivity = 'private'

Write-Host "$(Get-Date): Started"
try {
    do {
        $events = Get-MgUserCalendarEvent -CalendarId $cal.Id -UserId $UserId -Filter $EventFilter -OrderBy 'Start/dateTime desc' -Skip $skip -PageSize $pageSize

        $events | ForEach-Object {
            $count++
            if ($_.Sensitivity -ne $Sensitivity) {
                Write-Progress -Activity 'Updating calendar' -Status "${count}: $($_.Start.dateTime): $($_.Subject)"

                # As an example we update the 'sensitivity' of the event here
                Update-MgUserEvent -EventId $_.Id -UserId $UserId -Sensitivity $Sensitivity
            }
        } 
        $skip += $events.Count
        Write-Host "$(Get-Date): Updated $count"
    } while ($events.Count -gt 0)
} finally {
    Write-Host "$(Get-Date): Done ($count)"
}