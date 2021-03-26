# .SYNOPSIS
#    Good Times!
#
# .DESCRIPTION
#    Dieses Skript zeigt die Uptime-Zeiten der vergangenen Tage an, und berechnet
#    daraus die in der Zeitmanagement zu buchenden Zeiten, sowie Gleitzeit-Differenzen.
#
# .NOTES
#
#    Copyright 2015 Thomas Rosenau
#    Enhanced by Achim Winkler
#
#    Licensed under the Apache License, Version 2.0 (the "License");
#    you may not use this file except in compliance with the License.
#    You may obtain a copy of the License at
#
#        http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS,
#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#    See the License for the specific language governing permissions and
#    limitations under the License.
#
# .PARAMETER  install
#    Installiert einen Scheduler in der Windows Aufgabenplanung, um eine Warnung anzuzeigen, wenn die maximal erlaubte Anzahl der Arbeitsstunden erreicht wurde.
# .PARAMETER  uninstall
#    Deinstalliert einen zuvor installierten Scheduler wieder.
# .PARAMETER  check
#    Prüft ob die maximal zulässige Anzahl der Arbeitsstunden bereits erreicht wurde und gibt gegebenenfalls eine Warnung aus (wird normalerweise nur intern vom Scheduler aufgerufen).
# .PARAMETER  historyLength
#    Anzahl der angezeigten Tage in der Vergangenheit.
#    Standardwert: 60
#    Alias: -l
# .PARAMETER  workingHours
#    Anzahl der zu arbeitenden Stunden pro Tag.
#    Standardwert: 8
#    Alias: -h
# .PARAMETER  lunchBreak
#    Länge der Mittagspause in Stunden pro Tag.
#    Standardwert: 0.75
#    Alias: -b
# .PARAMETER  precision
#    Rundungspräzision in %, d.h. 1 = Rundung auf volle Stunde, 4 = Rundung auf 60/4=15 Minuten, …, 100 = keine Rundung
#    Standardwert: 60
#    Alias: -p
# .PARAMETER  dateFormat
#    Datumsformat gemäß https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx?cs-lang=vb#content
#    Standardwert: ddd dd/MM/yyyy
#    Alias: -d
# .PARAMETER  joinIntervals
#    Ignoriert die Pausen zwischen den Intervallen und rechnet nur das erste Intervall und das letzte Intervall zusammen. (0 = ausgeschaltet, 1 = eingeschaltet)
#    Standardwert: 1
#    Alias: -j
# .PARAMETER  maxWorkingHours
#    Anzahl der maximal erlaubten Stunden pro Tag die gearbeitet werden darf.
#    Standardwert: 10
#    Alias: -m
# .PARAMETER  showLogoff
#    Zeigt Logoff/Login Events an.
#    Standardwert: 1
#    Alias: -i
#
# .INPUTS
#    Keine
# .OUTPUTS
#    Keine
#
# .EXAMPLE
#    .\goodTimes.ps1
#    (Aufruf mit Standardwerten)
# .EXAMPLE
#    .\goodTimes.ps1 install
#    (Installiert einen Task in der Windows Aufgabenverwaltung und prüft alle 5min, ob die maximal zulässige Anzahl der täglichen Arbeitsstunden bereits erreicht wurde und gibt gegebenenfalls eine Warnung aus)
# .EXAMPLE
#    .\goodTimes.ps1 -historyLength 30 -workingHours 8 -lunchBreak 1 -precision 4 -joinIntervals 1 -maxWorkingHours 10
#    (Aufruf mit explizit gesetzten Standardwerten)
#    (30 Tage anzeigen, Arbeitszeit 8 Stunden täglich, 1 Stunde Mittagspause, Rundung auf 15 (=60/4) Minuten, Intervalle zusammen rechnen, 10h maximale Arbeitszeit)
# .EXAMPLE
#    .\goodTimes.ps1 -l 30 -h 8 -b 1 -p 4
#    (Aufruf mit explizit gesetzten Standardwerten, Kurzschreibweise)
# .EXAMPLE
#    .\goodTimes.ps1 -l 14 -h 7 -b .5 -p 6
#    (14 Tage anzeigen, Arbeitszeit 7 Stunden täglich, 30 Minuten Mittagspause, Rundung auf 10 (=60/6) Minuten)

param (
    [string]
    [Parameter(mandatory = $false)]
    [ValidateSet('install','uninstall','check')]
        $mode,
    [int]
    [validateRange(1, [int]::MaxValue)]
    [alias('l')]
        $historyLength = 60,
    [byte]
    [validateRange(0, 24)]
    [alias('h')]
        $workinghours = 8,
    [decimal]
    [validateRange(0, 24)]
    [alias('b')]
        $lunchbreak = 0.75,
    [byte]
    [validateRange(1, 100)]
    [alias('p')]
        $precision = 60,
    [string]
    [ValidateScript({$_ -cnotmatch '[HhmsfFt]'})]
    [alias('d')]
        $dateFormat = 'ddd dd/MM/yyyy', # "/" ist Platzhalter für lokalisierten Trenner
    [byte]
    [validateRange(0, 1)]
    [alias('j')]
        $joinIntervals = 1,
    [byte]
    [validateRange(0, 255)]
    [alias('m')]
        $maxWorkingHours = 10,
    [byte]
    [validateRange(0, 1)]
    [alias('i')]
        $showLogoff = 1
)

# helper functions to calculate the required attributes

# total uptime
function getUptimeAttr($entry) {
    $result = New-TimeSpan

    if ($joinIntervals -eq 0) {
        foreach ($interval in $entry) {
            $result = $result.add($interval[1] - $interval[0])
        }
    } else {
        $result = $entry[-1][-1] - $entry[0][0]
    }

    $result
}

# uptime intervals
function getIntervalAttr($entry) {
    $result = @()

    if ($joinIntervals -eq 0) {
        foreach ($interval in $entry) {
            $result += '{0:HH:mm}-{1:HH:mm}' -f $interval[0], $interval[1]
        }
    } else {
        $result = '{0:HH:mm}-{1:HH:mm}' -f $entry[0][0], $entry[-1][-1]
    }

    $result -join ', '
}

# booking hours
function getBookingHoursAttr($interval) {
    $netTime = $interval.totalHours - $lunchbreak
    [math]::Round($netTime * $precision) / $precision
}

# flex time delta
function getFlexTimeAttr($bookedHours) {
    $delta = $bookedHours - $workinghours
    $result = $delta.toString('+0.00;-0.00; 0.00')
    if ($delta -eq 0) {
        write $result, $null
    } elseif ($delta -gt 0) {
        write $result, 'darkgreen'
    } else {
        write $result, 'darkred'
    }
}
# end helper functions

# generate a hashmap of the abovementioned attributes
function getLogAttrs($entry) {
    $result = @{
        uptime = getUptimeAttr $entry
        intervals = getIntervalAttr $entry
    }
    $result.bookingHours = getBookingHoursAttr $result.uptime
    $result.flexTime = getFlexTimeAttr $result.bookingHours
    $result
}

# convenience function to write to screen with or without color
function print($string, $color) {
    if ($color) {
        write-host -f $color -n $string
    } else {
        write-host -n $string
    }
}

# convenience function to write to screen with or without color
function println($string, $color) {
    print ($string + "`r`n") $color
}

# If running in the console, wait for input before continuing.
function wait() {
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host 'Press any key to continue...'
        $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyUp') > $null
    }
}

# helper to determine whether a given EventLogRecord is a boot or wakeup event
function isStartEvent($event) {
    return ($event.ID -eq 12 -and $event.ProviderName -eq 'Microsoft-Windows-Kernel-General') -or
            ($event.ID -eq 1 -and $event.ProviderName -eq 'Microsoft-Windows-Power-Troubleshooter') -or
            ($showLogoff -eq 1 -and $event.ID -eq 811 -and $event.ProviderName -eq 'Microsoft-Windows-Winlogon' -and $event.Message -Match "<Sens>" -and $event.Message -Match "\(5\)")
}

# helper to determine whether a given EventLogRecord is a shutdown or suspend event
function isStopEvent($event) {
    return ($event.ID -eq 13 -and $event.ProviderName -eq 'Microsoft-Windows-Kernel-General') -or
            ($event.ID -eq 42 -and $event.ProviderName -eq 'Microsoft-Windows-Kernel-Power') -or
            ($showLogoff -eq 1 -and $event.ID -eq 811 -and $event.ProviderName -eq 'Microsoft-Windows-Winlogon' -and $event.Message -Match "<Sens>" -and $event.Message -Match "\(4\)")
}




if ($mode -eq 'install') {
    "Installation started..."

    $script = 'schedule.vbs'
    $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    $action = New-ScheduledTaskAction -Execute $script -WorkingDirectory $scriptPath -Argument "check -l 1 -h $workinghours -b $lunchbreak -p $precision -m $maxWorkingHours"
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddSeconds(10)) -RepetitionInterval (New-TimeSpan -Minutes 5)
    $settings = New-ScheduledTaskSettingsSet -Hidden -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -StartWhenAvailable
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings

    if (Get-ScheduledTask 'Check-Worktime' -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask 'Check-Worktime'
    }
    Register-ScheduledTask 'Check-Worktime' -InputObject $task | Out-Null

    "Ready."

    Exit $LASTEXITCODE
}
if ($mode -eq 'uninstall') {
    "Deinstallation started ..."

    if (Get-ScheduledTask 'Check-Worktime' -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask 'Check-Worktime'
    }

    "Ready."

    Exit $LASTEXITCODE
}

# create an array of filterHashTables that filter boot and shutdown events from the desired period
$startTime = (get-date).addDays(-$historyLength)
$filters = (
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-General'
        ID = 12, 13
    },
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-Power'
        ID = 42
    },
    @{
        StartTime = $startTime
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-Power-Troubleshooter'
        ID = 1
    },
    @{
        StartTime = $startTime
        LogName = 'Microsoft-Windows-Winlogon/Operational'
        ProviderName = 'Microsoft-Windows-Winlogon'
        ID = 811
    }
)

# get system log entries for boot/shutdown
# sort (reverse chronological order) and convert to ArrayList
if ($showLogoff -eq 1) {
    [Collections.ArrayList]$events = Get-WinEvent -FilterHashtable $filters | select ID, TimeCreated, ProviderName, Message | sort TimeCreated
} else {
    [Collections.ArrayList]$events = Get-WinEvent -FilterHashtable $filters | select ID, TimeCreated, ProviderName | sort TimeCreated
}

# create an empty list, which will hold one entry per day
$log = New-Object Collections.ArrayList

# fill the $log list by searching for start/stop pairs
:outer while ($events.count -ge 2) {
    if ($log) {
        # find the latest stop event
        do {
            if ($events.count -lt 2) {
                # if there is only one stop event left, there can't be any more start event (e.g. when system log was cleared)
                break outer
            }
            $end = $events[$events.count - 1]
            $events.remove($end)
        } while (-not (isStopEvent $end)) # consecutive start events. This may happen when the system crashes (power failure, etc.)
    } else {
        # add a fake shutdown event for this very moment
        $end = @{TimeCreated = get-date}
    }

    # find the corresponding start event
    do {
        if ($events.count -lt 1) {
            # no more events left
            break outer
        }
        $start = $events[$events.count - 1]
        $events.remove($start)
    } while (-not (isStartEvent $start)) # not sure if there can indeed be consecutive stop events, but let's better be safe than sorry

    # check if the current start/stop pair has occurred on the same day as the previous one
    $last = $log[0]
    $interval = ,($start.TimeCreated, $end.TimeCreated)
    if ($last -and $start.TimeCreated.Date.equals($last[0][0].Date)) {
        # combine uptimes
        $log[0] = $interval + $last
    } else {
        # create new day
        $log.insert(0, $interval)
    }

}

if ($mode -eq 'check') {
    $entry = $log[-1]
    $attrs = getLogAttrs($entry)

    if ($attrs.bookingHours -ge ($maxWorkingHours)) {
        $Shell = new-object -comobject wscript.shell -ErrorAction Stop
        $Shell.popup("Maximum worktime reached!!!`nYou must leave now!!!", 0, 'Maximum Worktime', 48 + 4096) | Out-Null
    }
    elseif ($attrs.bookingHours -ge ($maxWorkingHours - 0.25)) {  # 0.25 = 15 minutes
        $Shell = new-object -comobject wscript.shell -ErrorAction Stop
        $Shell.popup("Maximum worktime reached in a few minutes.`nYou should leave now!", 0, 'Maximum Worktime Warning', 48 + 4096) | Out-Null
    }

    Exit $LASTEXITCODE
}

# colors
$oldFgColor= $host.UI.RawUI.ForegroundColor
$host.UI.RawUI.ForegroundColor = 'gray'
$oldBgColor = $host.UI.RawUI.BackgroundColor
$host.UI.RawUI.BackgroundColor = 'black'

# write the output
$screenWidth = $host.UI.RawUI.BufferSize.width

Write-Host ("{0,-$($screenWidth - 1)}" -f '    Date      Workt. Flexitime  Uptime (incl. breaks)')
Write-Host ("{0,-$($screenWidth - 1)}" -f '------------- ------ ---------  ---------------------')

foreach ($entry in $log) {
    $firstStart = $entry[0][0]
    $dayOfWeek = ([int]$firstStart.dayOfWeek + 6) % 7
    $dayFormatted = $firstStart.Date.toString($dateFormat)
    $attrs = getLogAttrs($entry)
    
    if ($dayOfWeek -lt $lastDayOfWeek) {
        println ("{0,-$($screenWidth - 1)}" -f '-------------')
    }
    $lastDayOfWeek = $dayOfWeek
    
    if ($dayOfWeek -ge 5) {
        Write-Host $dayFormatted -n -backgroundColor darkred -foregroundcolor gray
    } else {
        print $dayFormatted
    }
    
    if ($attrs.bookingHours -ge ($maxWorkingHours + $lunchbreak)) {
        print ('  {0,5}  ' -f
                $attrs.bookingHours.toString('#0.00', [Globalization.CultureInfo]::getCultureInfo('de-DE'))
              ) red
    } else {
        print ('  {0,5}  ' -f
                $attrs.bookingHours.toString('#0.00', [Globalization.CultureInfo]::getCultureInfo('de-DE'))
              ) cyan
    }
    
    print ('  {0,6}  ' -f $attrs.flexTime[0]) $attrs.flexTime[1]
    
    print ("{0,3:#0}:{1:00} | {2,-$($screenWidth - 42)}" -f
        $attrs.uptime.hours,
        [Math]::Round($attrs.uptime.minutes + $attrs.uptime.seconds / 60),
        $attrs.intervals) darkGray
    Write-Host
}

# restore previous colors
$host.UI.RawUI.BackgroundColor = $oldBgColor
$host.UI.RawUI.ForegroundColor = $oldFgColor

wait
