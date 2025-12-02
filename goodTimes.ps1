# .SYNOPSIS
#    Good Times!
#
# .DESCRIPTION
#    This script displays the uptime times of the past days, and calculates from them the
#    times to be booked in the time management, as well as flextime differences.
#    The script does not store any data on your computer. It only uses the data that Windows automatically provides.
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
#    Installs a scheduler in Windows Task Scheduler to display a warning when the maximum allowed working hours has been reached.
# .PARAMETER  install_widget
#    Installs a scheduler in Windows Task Scheduler to display a widget at logon (installation needs admin rights).
# .PARAMETER  uninstall
#    Uninstalls the scheduler to display a warning when the maximum allowed number of working hours has been reached.
# .PARAMETER  uninstall_widget
#    Uninstalls the scheduler to display a widget at logon (deinstallation needs admin rights).
# .PARAMETER  check
#    Checks if the maximum allowed number of working hours has already been reached and displays a warning if necessary (normally only called internally by the scheduler).
# .PARAMETER  widget
#    Launches a widget to display all relevant information about your working time.
# .PARAMETER  historyLength
#    Number of days to show in uptime history.
#    Default: 60
#    Alias: -l
# .PARAMETER  workingHours
#    Working hours per day, used for overtime calculation. This will be added to your daily work time.
#    Default: 8
#    Alias: -h
# .PARAMETER  breakfastBreak
#    Length of breakfast break in hours per day. This will be added to your daily work time.
#    Default: 0.25
#    Alias: -b1
# .PARAMETER  lunchBreak
#    Length of lunch break in hours per day. This will be added to your daily work time.
#    Default: 0.50
#    Alias: -b2
# .PARAMETER  precision
#    Rounding precision in %, i.e. 1 = rounding to full hour, 4 = rounding to 60/4=15 minutes, ..., 100 = no rounding
#    Default: 60
#    Alias: -p
# .PARAMETER  dateFormat
#    Date format according to https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx?cs-lang=vb#content
#    Default: ddd dd/MM/yyyy
#    Alias: -d
# .PARAMETER  joinIntervals
#    Ignores the breaks between the intervals and combines only the start of the first interval and the end of the last interval (0 = switched off, 1 = switched on).
#    Default: 1
#    Alias: -j
# .PARAMETER  maxWorkingHours
#    Number of maximum allowed hours per day that can be worked.
#    Default: 10
#    Alias: -m
# .PARAMETER  showLogoff
#    Show logoff/login events and lockscreen on/off events (0 = switched off, 1 = switched on).
#    Default: 1
#    Alias: -i
#
# .INPUTS
#    None
# .OUTPUTS
#    None
#
# .EXAMPLE
#    .\goodTimes.ps1
#    (run with default values)
# .EXAMPLE
#    .\goodTimes.ps1 install
#    (Installs a task in the Windows Task Manager and checks every 5min if the maximum allowed number of daily working hours has already been reached and displays a warning if necessary.)
# .EXAMPLE
#    .\goodTimes.ps1 -historyLength 60 -workingHours 8 -breakfastBreak 0.25 -lunchBreak 0.5 -precision 60 -joinIntervals 1 -maxWorkingHours 10 -showLogoff 1
#    (Call with explicitly set default values)
#    (Show 60 days, working time 8 hours daily, 15 minutes breakfast break, 30 minutes lunch break, rounding to 1 minute, join intervals, 10h maximum working hours, show logon/logoff and lockscreen events)
# .EXAMPLE
#    .\goodTimes.ps1 -l 60 -h 8 -b1 0.25 -B2 0.5 -p 60 -m 10 -i 1
#    (Call with explicitly set default values, short form)
# .EXAMPLE
#    .\goodTimes.ps1 -l 14 -h 7 -b1 0 -b2 0.5 -p 6
#    (show 14 days, working time 7 hours daily, breakfast break disabled, 30 minutes lunch break, rounding to 10 (=60/6) minutes)

param (
    [string]
    [Parameter(mandatory = $false)]
    [ValidateSet('install', 'uninstall', 'install_widget', 'uninstall_widget', 'check', 'widget')]
        $mode,
    [int]
    [validateRange(1, [int]::MaxValue)]
    [alias('l')]
        $historyLength = 60,
    [decimal]
    [validateRange(0, 12)]
    [alias('h')]
        $workinghours = 8,
    [decimal]
    [validateRange(0, 12)]
    [alias('b1')]
        $breakfastBreak = 0.25,
    [decimal]
    [validateRange(0, 12)]
    [alias('b2')]
        $lunchBreak = 0.50,
    [byte]
    [validateRange(1, 100)]
    [alias('p')]
        $precision = 60,
    [string]
    [ValidateScript({$_ -cnotmatch '[HhmsfFt]'})]
    [alias('d')]
        $dateFormat = 'ddd dd/MM/yyyy',
    [byte]
    [validateRange(0, 1)]
    [alias('j')]
        $joinIntervals = 1,
    [decimal]
    [validateRange(0, 12)]
    [alias('m')]
        $maxWorkingHours = 10,
    [byte]
    [validateRange(0, 1)]
    [alias('i')]
        $showLogoff = 1
)

# Load all required assemblies once at the start
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing -ErrorAction SilentlyContinue

# global configuration variables (can be overridden by goodTimes.json next to the script)
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Definition
$configFile  = Join-Path $scriptDir 'goodTimes.json'

$defaultConfig = @{
    cultureInfo     = 'de-DE'
    breakDeduction1 = 3.0
    breakDeduction2 = 6.0
    breakThreshold  = 3.0
    topMost         = $true
    topPosition     = -1
    leftPosition    = -1
}

function Remove-HashCommentLines {
    param([string]$json)
    # split into lines, drop lines that start with optional whitespace then #
    $lines = $json -split "`r?`n"
    ($lines | Where-Object { $_ -notmatch '^\s*#' }) -join "`n"
}

if (Test-Path $configFile) {
    try {
        $raw = Get-Content -Raw -Path $configFile -ErrorAction Stop
        $clean = Remove-HashCommentLines $raw
        $cfg = $clean | ConvertFrom-Json -ErrorAction Stop
    } catch {
        $cfg = $null
    }
} else {
    $cfg = $null
}

function Get-ConfigValue {
    param($Name, $Default)
    if ($null -ne $cfg -and $cfg.PSObject.Properties.Name -contains $Name) {
        return $cfg.$Name
    }
    return $Default
}

$cultureInfo     = Get-ConfigValue 'cultureInfo'              $defaultConfig.cultureInfo
$breakDeduction1 = [double](Get-ConfigValue 'breakDeduction1' $defaultConfig.breakDeduction1)
$breakDeduction2 = [double](Get-ConfigValue 'breakDeduction2' $defaultConfig.breakDeduction2)
$breakThreshold  = [double](Get-ConfigValue 'breakThreshold'  $defaultConfig.breakThreshold)
$topMost         = [bool]  (Get-ConfigValue 'topMost'         $defaultConfig.topMost)
$topPosition     = [int]   (Get-ConfigValue 'topPosition'     $defaultConfig.topPosition)
$leftPosition    = [int]   (Get-ConfigValue 'leftPosition'    $defaultConfig.leftPosition)


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
    $netTime = $interval.totalHours
    if ($interval.totalHours -ge $breakDeduction1) {
        $netTime -= $breakfastBreak
    }
    if ($interval.totalHours -ge $breakDeduction2) {
        $netTime -= $lunchBreak
    }
    [math]::Round($netTime * $precision) / $precision
}

# flex time delta
function getFlexTimeAttr($bookedHours) {
    $delta = $bookedHours - $workinghours
    $result = $delta.toString('+0.00;-0.00; 0.00', [Globalization.CultureInfo]::getCultureInfo($cultureInfo))
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

function Show-Widget {

    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [double] $startTime,
        [Parameter(Mandatory=$true, Position=1)]
        [double] $normalWorkHours,
        [Parameter(Mandatory=$true, Position=2)]
        [double] $maxWorkHours,
        [Parameter(Mandatory=$true, Position=3)]
        [double] $break1,
        [Parameter(Mandatory=$true, Position=4)]
        [double] $break2,
        [Parameter(Mandatory=$true, Position=5)]
        [double] $breakDeduction1,
        [Parameter(Mandatory=$true, Position=6)]
        [double] $breakDeduction2,
        [Parameter(Mandatory=$true, Position=7)]
        [double] $unplannedBreaks
    )

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="goodTimes" Height="110" Width="110" WindowStartupLocation="CenterScreen" Top="0" Left="0" WindowStyle="None" ResizeMode="NoResize" ShowInTaskbar="False" AllowsTransparency="True" Background="Transparent" Opacity="1" Topmost="True">
    <Canvas x:Name="Canvas" ToolTipService.InitialShowDelay="1000" ToolTipService.ShowDuration="10000" ToolTipService.Placement="Bottom" ToolTipService.HasDropShadow="False" ToolTipService.IsEnabled="True">
        <Path x:Name="Path_NormalWorktime" Stroke="Black" StrokeThickness="1" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Opacity="1" SnapsToDevicePixels="True">
            <Path.Data>
                <PathGeometry>
                    <PathGeometry.Figures>
                        <PathFigureCollection>
                            <PathFigure x:Name="PathFigure_NormalWorktime" IsClosed= "True" StartPoint="0,0">
                                <PathFigure.Segments>
                                    <PathSegmentCollection>
                                        <ArcSegment x:Name="ArcSegment_NormalWorktime" Size="0,0" IsLargeArc="True" IsStroked="False" SweepDirection="Clockwise" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentA_NormalWorktime" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentB_NormalWorktime" Point="0,0" />
                                    </PathSegmentCollection>
                                </PathFigure.Segments>
                            </PathFigure>
                        </PathFigureCollection>
                    </PathGeometry.Figures>
                </PathGeometry>
            </Path.Data>
            <Path.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#FFFF80" Offset="0.0" />
                    <GradientStop Color="#FFFF00" Offset="1.0" />
                </RadialGradientBrush>
            </Path.Fill>
        </Path>
        <Path x:Name="Path_MaxWorktime" Stroke="Black" StrokeThickness="1" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Opacity="1" SnapsToDevicePixels="True">
            <Path.Data>
                <PathGeometry>
                    <PathGeometry.Figures>
                        <PathFigureCollection>
                            <PathFigure x:Name="PathFigure_MaxWorktime" IsClosed= "True" StartPoint="0,0">
                                <PathFigure.Segments>
                                    <PathSegmentCollection>
                                        <ArcSegment x:Name="ArcSegment_MaxWorktime" Size="0,0" IsLargeArc="True" IsStroked="False" SweepDirection="Clockwise" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentA_MaxWorktime" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentB_MaxWorktime" Point="0,0" IsStroked="False" />
                                    </PathSegmentCollection>
                                </PathFigure.Segments>
                            </PathFigure>
                        </PathFigureCollection>
                    </PathGeometry.Figures>
                </PathGeometry>
            </Path.Data>
            <Path.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#FF8080" Offset="0.0" />
                    <GradientStop Color="#FF4040" Offset="1.0" />
                </RadialGradientBrush>
            </Path.Fill>
        </Path>
        <Path x:Name="Path_Break1" Stroke="Black" StrokeThickness="1" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Opacity="1" SnapsToDevicePixels="True">
            <Path.Data>
                <PathGeometry>
                    <PathGeometry.Figures>
                        <PathFigureCollection>
                            <PathFigure x:Name="PathFigure_Break1" IsClosed= "True" StartPoint="0,0">
                                <PathFigure.Segments>
                                    <PathSegmentCollection>
                                        <ArcSegment x:Name="ArcSegment_Break1" Size="0,0" IsLargeArc="True" IsStroked="False" SweepDirection="Clockwise" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentA_Break1" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentB_Break1" Point="0,0" />
                                    </PathSegmentCollection>
                                </PathFigure.Segments>
                            </PathFigure>
                        </PathFigureCollection>
                    </PathGeometry.Figures>
                </PathGeometry>
            </Path.Data>
            <Path.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#80FF80" Offset="0.0" />
                    <GradientStop Color="#40FF40" Offset="1.0" />
                </RadialGradientBrush>
            </Path.Fill>
        </Path>
        <Path x:Name="Path_Break2" Stroke="Black" StrokeThickness="1" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Opacity="1" SnapsToDevicePixels="True">
            <Path.Data>
                <PathGeometry>
                    <PathGeometry.Figures>
                        <PathFigureCollection>
                            <PathFigure x:Name="PathFigure_Break2" IsClosed= "True" StartPoint="0,0">
                                <PathFigure.Segments>
                                    <PathSegmentCollection>
                                        <ArcSegment x:Name="ArcSegment_Break2" Size="0,0" IsLargeArc="True" IsStroked="False" SweepDirection="Clockwise" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentA_Break2" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentB_Break2" Point="0,0" />
                                    </PathSegmentCollection>
                                </PathFigure.Segments>
                            </PathFigure>
                        </PathFigureCollection>
                    </PathGeometry.Figures>
                </PathGeometry>
            </Path.Data>
            <Path.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#80FF80" Offset="0.0" />
                    <GradientStop Color="#40FF40" Offset="1.0" />
                </RadialGradientBrush>
            </Path.Fill>
        </Path>
        <Path x:Name="Path_ActualWorktime" Stroke="Black" StrokeThickness="1" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Opacity="0.8" SnapsToDevicePixels="True">
            <Path.Data>
                <PathGeometry>
                    <PathGeometry.Figures>
                        <PathFigureCollection>
                            <PathFigure x:Name="PathFigure_ActualWorktime" IsClosed= "True" StartPoint="0,0">
                                <PathFigure.Segments>
                                    <PathSegmentCollection>
                                        <ArcSegment x:Name="ArcSegment_ActualWorktime" Size="0,0" IsLargeArc="True" SweepDirection="Clockwise" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentA_ActualWorktime" Point="0,0" />
                                        <LineSegment x:Name="LineSegmentB_ActualWorktime" Point="0,0" IsStroked="False" />
                                    </PathSegmentCollection>
                                </PathFigure.Segments>
                            </PathFigure>
                        </PathFigureCollection>
                    </PathGeometry.Figures>
                </PathGeometry>
            </Path.Data>
            <Path.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#0080FF" Offset="0.0" />
                    <GradientStop Color="#0040FF" Offset="1.0" />
                </RadialGradientBrush>
            </Path.Fill>
        </Path>
        <Ellipse x:Name="Path_ActualWorktime12h" Stroke="Black" StrokeThickness="1" SnapsToDevicePixels="True" Canvas.Left="8" Canvas.Top="8" Height="84" Width="84" Opacity="0.2">
            <Ellipse.Fill>
                <RadialGradientBrush GradientOrigin="0.5,0.5" Center="0.5,0.5" RadiusX="1.0" RadiusY="1.0">
                    <GradientStop Color="#0080FF" Offset="0.0" />
                    <GradientStop Color="#0040FF" Offset="1.0" />
                </RadialGradientBrush>
            </Ellipse.Fill>
        </Ellipse>
        <Ellipse Stroke="Black" StrokeThickness="1" SnapsToDevicePixels="True" Height="101" Width="101" />
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="4" />
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="30" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="60" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="4">
            <Line.RenderTransform>
                <RotateTransform Angle="90" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="120" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="150" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="4">
            <Line.RenderTransform>
                <RotateTransform Angle="180" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="210" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="240" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="4">
            <Line.RenderTransform>
                <RotateTransform Angle="270" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="300" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>
        <Line Stroke="Black" StrokeThickness="2" SnapsToDevicePixels="True" X1="50" Y1="0" X2="50" Y2="3">
            <Line.RenderTransform>
                <RotateTransform Angle="330" CenterX="50" CenterY="50" />
            </Line.RenderTransform>
        </Line>

        <Line Stroke="Gray" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" SnapsToDevicePixels="True" Opacity="0.9" X1="105" Y1="3" X2="99" Y2="9" />
        <Line Stroke="Gray" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" SnapsToDevicePixels="True" Opacity="0.9" X1="99" Y1="3" X2="105" Y2="9" />
        <Rectangle x:Name="Close_Widget" Stroke="Gray" StrokeThickness="1" Fill="Gray" SnapsToDevicePixels="True" Opacity="0.01" Width="8" Height="8" Canvas.Left="98" Canvas.Top="2" />

        <Line x:Name="Minimize_Icon" Stroke="Gray" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" SnapsToDevicePixels="True" Opacity="0.9" X1="90" Y1="9" X2="96" Y2="9" />
        <Rectangle x:Name="Maximize_Icon" Stroke="Gray" StrokeThickness="2" SnapsToDevicePixels="True" Opacity="0.9" Width="8" Height="8" Canvas.Left="89" Canvas.Top="2" />

        <Rectangle x:Name="MinMax_Widget" Stroke="Gray" StrokeThickness="1" Fill="Gray" SnapsToDevicePixels="True" Opacity="0.01" Width="8" Height="8" Canvas.Left="89" Canvas.Top="2" />

        <Canvas.Effect>
            <DropShadowEffect Color="Black" ShadowDepth="3" BlurRadius="6" Opacity="0.4" />
        </Canvas.Effect>

        <Canvas.ToolTip>
            <ToolTip Background="Silver">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Foreground="#0080FF" Width="100">Start time:</TextBlock>
                        <TextBlock x:Name="Tooltip_StartTime" Foreground="#0080FF" Width="35" TextAlignment="Right">00:00</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Foreground="#FFFF00" Width="100">Normal worktime:</TextBlock>
                        <TextBlock x:Name="Tooltip_NormalWorktime" Foreground="#FFFF00" Width="35" TextAlignment="Right">00:00</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Foreground="#FF4040" Width="100">Max worktime:</TextBlock>
                        <TextBlock x:Name="Tooltip_MaxWorktime" Foreground="#FF4040" Width="35" TextAlignment="Right">00:00</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Foreground="#0040FF" Width="100">Time to work:</TextBlock>
                        <TextBlock x:Name="Tooltip_TimeToWork" Foreground="#0040FF" Width="35" TextAlignment="Right">00:00</TextBlock>
                    </StackPanel>
                    <StackPanel x:Name="Tooltip_OptionalUnplannedBreaks" Orientation="Horizontal">
                        <TextBlock Foreground="White" Width="100">Unplanned breaks:</TextBlock>
                        <TextBlock x:Name="Tooltip_UnplannedBreaks" Foreground="White" Width="35" TextAlignment="Right">00:00</TextBlock>
                    </StackPanel>
                </StackPanel>
            </ToolTip>
        </Canvas.ToolTip>
    </Canvas>
</Window>
"@

    $ComputeCartesianCoordinate =
    {
        Param
        (
            [Parameter(Mandatory=$true, Position=0)]
            [double] $angle,
            [Parameter(Mandatory=$true, Position=1)]
            [double] $radius
        )

        $angleRad = ([Math]::pi / 180.0) * ($angle - 90.0)

        # x,y
        return ($radius * [Math]::cos($angleRad)), ($radius * [Math]::sin($angleRad))
    }

    $GetPieCoordinates =
    {
        Param
        (
            [Parameter(Mandatory=$true, Position=0)]
            [int] $centreX,
            [Parameter(Mandatory=$true, Position=1)]
            [int] $centreY,
            [Parameter(Mandatory=$true, Position=2)]
            [double] $rotationAngle,
            [Parameter(Mandatory=$true, Position=3)]
            [double] $wedgeAngle,
            [Parameter(Mandatory=$true, Position=4)]
            [double] $radius,
            [Parameter(Mandatory=$true, Position=5)]
            [double] $innerRadius
        )

        # 0.5 is added to increase the sharpness of the stroke
        # https://developer.mozilla.org/en-US/docs/Web/API/Canvas_API/Tutorial/Applying_styles_and_colors
        $innerArcStartPointX, $innerArcStartPointY = &$ComputeCartesianCoordinate $rotationAngle $innerRadius
        $innerArcStartPointX += $centreX + 0.5
        $innerArcStartPointY += $centreY + 0.5

        $innerArcEndPointX, $innerArcEndPointY = &$ComputeCartesianCoordinate ($rotationAngle + $wedgeAngle) $innerRadius
        $innerArcEndPointX += $centreX + 0.5
        $innerArcEndPointY += $centreY + 0.5

        $outerArcStartPointX, $outerArcStartPointY = &$ComputeCartesianCoordinate $rotationAngle $radius
        $outerArcStartPointX += $centreX + 0.5
        $outerArcStartPointY += $centreY + 0.5

        $outerArcEndPointX, $outerArcEndPointY = &$ComputeCartesianCoordinate ($rotationAngle + $wedgeAngle) $radius
        $outerArcEndPointX += $centreX + 0.5
        $outerArcEndPointY += $centreY + 0.5

        return $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY
    }

    $UpdateWidget =
    {
        # hour = 360 / 12 = 30
        # minute = 360 / 12 / 60 = 0.5

        #normal worktime
        $start = $startTime
        $rotationAngle = $start * 30
        $end = $normalWorkHours + $break1 + $break2
        $wedgeAngle = $end * 30

        $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY = &$GetPieCoordinates 50 50 $rotationAngle $wedgeAngle 50 0
        $PathFigure_NormalWorktime.StartPoint = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
        $ArcSegment_NormalWorktime.IsLargeArc = ($wedgeAngle -gt 180)
        $ArcSegment_NormalWorktime.Size = [System.Windows.Size]::new(50,50)
        $ArcSegment_NormalWorktime.Point = [System.Windows.Point]::new($outerArcEndPointX,$outerArcEndPointY)
        $LineSegmentA_NormalWorktime.Point = [System.Windows.Point]::new($innerArcEndPointX,$innerArcEndPointY)
        $LineSegmentB_NormalWorktime.Point = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)

        #max worktime
        $start = $startTime + $end
        $rotationAngle = $start * 30
        $end = $maxWorkHours - $normalWorkHours
        $wedgeAngle = $end * 30

        $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY = &$GetPieCoordinates 50 50 $rotationAngle $wedgeAngle 50 0
        $PathFigure_MaxWorktime.StartPoint = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
        $ArcSegment_MaxWorktime.IsLargeArc = ($wedgeAngle -gt 180)
        $ArcSegment_MaxWorktime.Size = [System.Windows.Size]::new(50,50)
        $ArcSegment_MaxWorktime.Point = [System.Windows.Point]::new($outerArcEndPointX,$outerArcEndPointY)
        $LineSegmentA_MaxWorktime.Point = [System.Windows.Point]::new($innerArcEndPointX,$innerArcEndPointY)
        $LineSegmentB_MaxWorktime.Point = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)

        #break1
        if ($break1 -gt 0) {
            $start = $startTime + $breakDeduction1
            $rotationAngle = $start * 30
            $end = $break1
            $wedgeAngle = $end * 30

            $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY = &$GetPieCoordinates 50 50 $rotationAngle $wedgeAngle 50 0
            $PathFigure_Break1.StartPoint = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
            $ArcSegment_Break1.IsLargeArc = ($wedgeAngle -gt 180)
            $ArcSegment_Break1.Size = [System.Windows.Size]::new(50,50)
            $ArcSegment_Break1.Point = [System.Windows.Point]::new($outerArcEndPointX,$outerArcEndPointY)
            $LineSegmentA_Break1.Point = [System.Windows.Point]::new($innerArcEndPointX,$innerArcEndPointY)
            $LineSegmentB_Break1.Point = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
        }
        else {
            $Path_Break1.Visibility = [System.Windows.Visibility]::Hidden
        }

        #break2
        if ($break2 -gt 0) {
            $start = $startTime + $breakDeduction2
            $rotationAngle = $start * 30
            $end = $break2
            $wedgeAngle = $end * 30

            $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY = &$GetPieCoordinates 50 50 $rotationAngle $wedgeAngle 50 0
            $PathFigure_Break2.StartPoint = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
            $ArcSegment_Break2.IsLargeArc = ($wedgeAngle -gt 180)
            $ArcSegment_Break2.Size = [System.Windows.Size]::new(50,50)
            $ArcSegment_Break2.Point = [System.Windows.Point]::new($outerArcEndPointX,$outerArcEndPointY)
            $LineSegmentA_Break2.Point = [System.Windows.Point]::new($innerArcEndPointX,$innerArcEndPointY)
            $LineSegmentB_Break2.Point = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
        }
        else {
            $Path_Break2.Visibility = [System.Windows.Visibility]::Hidden
        }

        #actual worktime
        $start = $startTime
        $rotationAngle = $start * 30
        $end = (Get-Date).TimeOfDay.TotalHours - $startTime - $unplannedBreaks
        if ($end -ge 12.0) {
            $end = $end % 12.0
        }
        else {
            $Path_ActualWorktime12h.Visibility = [System.Windows.Visibility]::Hidden
        }
        $wedgeAngle = $end * 30

        $innerArcStartPointX, $innerArcStartPointY, $innerArcEndPointX, $innerArcEndPointY, $outerArcStartPointX, $outerArcStartPointY, $outerArcEndPointX, $outerArcEndPointY = &$GetPieCoordinates 50 50 $rotationAngle $wedgeAngle 42 0
        $PathFigure_ActualWorktime.StartPoint = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)
        $ArcSegment_ActualWorktime.IsLargeArc = ($wedgeAngle -gt 180)
        $ArcSegment_ActualWorktime.Size = [System.Windows.Size]::new(42,42)
        $ArcSegment_ActualWorktime.Point = [System.Windows.Point]::new($outerArcEndPointX,$outerArcEndPointY)
        $LineSegmentA_ActualWorktime.Point = [System.Windows.Point]::new($innerArcEndPointX,$innerArcEndPointY)
        $LineSegmentB_ActualWorktime.Point = [System.Windows.Point]::new($outerArcStartPointX,$outerArcStartPointY)

        #tooltip
        $Tooltip_StartTime.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes ($startTime * 60))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        $Tooltip_NormalWorktime.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes (($startTime * 60) + ($normalWorkHours * 60) + ($break1 * 60) + ($break2 * 60)))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        $Tooltip_MaxWorktime.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes (($startTime * 60) + ($maxWorkHours * 60) + ($break1 * 60) + ($break2 * 60)))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        $worktime = (Get-Date).TimeOfDay.TotalHours - $startTime - $unplannedBreaks
        $worktimeAdj = $worktime
        if ($worktime -ge $breakDeduction1) {
            $worktimeAdj -= $break1
        }
        if ($worktime -ge $breakDeduction2) {
            $worktimeAdj -= $break2
        }
        if ($worktimeAdj -le $normalWorkHours) {
            $Tooltip_TimeToWork.Text = "-" + ((New-Object DateTime) + (New-TimeSpan -Minutes (($normalWorkHours * 60) - ($worktimeAdj * 60)))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        }
        else {
            $Tooltip_TimeToWork.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes (($normalWorkHours * 60) - ($worktimeAdj * 60))).Negate()).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        }
        if ($joinIntervals -eq 0) {
            if ($null -eq $unplannedBreaks -or [double]::IsNaN([double]$unplannedBreaks)) {
                $Tooltip_UnplannedBreaks.Text = '00:00'
            } else {
                $Tooltip_UnplannedBreaks.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes ($unplannedBreaks * 60))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
            }
        }

        if ($Widget.Topmost -eq $true) {
            $Minimize_Icon.Visibility = [System.Windows.Visibility]::Visible
            $Maximize_Icon.Visibility = [System.Windows.Visibility]::Hidden
            # force topmost again if lost
            $Widget.Topmost = $false
            $Widget.Topmost = $true
        }
        else {
            $Minimize_Icon.Visibility = [System.Windows.Visibility]::Hidden
            $Maximize_Icon.Visibility = [System.Windows.Visibility]::Visible
        }

        $Widget.UpdateLayout()
    }

    $SyncWidget =
    {
        Param([Parameter(Mandatory=$true)][ref]$refUnplannedBreaks)

        $log = &$script:updateWorktimes -eventPeriod 1
        $entry = $log[-1]
        $attrs = getLogAttrs($entry)

        $unplannedBreaks = (Get-Date).TimeOfDay.TotalHours - $entry[0][0].TimeOfDay.TotalHours - $attrs.uptime.TotalHours

        $refUnplannedBreaks.Value = $unplannedBreaks
    }

    #Read the form
    $Reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Widget = [Windows.Markup.XamlReader]::Load($reader)

    #AutoFind all controls
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {
        New-Variable -Name $_.Name -Value $Widget.FindName($_.Name) -Force
    }

    $Widget.Add_MouseLeftButtonDown({
        $Widget.DragMove()
    })

    $Close_Widget.Add_MouseLeftButtonDown({
        $Widget.Close()
    })

    $MinMax_Widget.Add_MouseLeftButtonDown({
        if ($Widget.Topmost -eq $true) {
            $Widget.Topmost = $false

            $Minimize_Icon.Visibility = [System.Windows.Visibility]::Hidden
            $Maximize_Icon.Visibility = [System.Windows.Visibility]::Visible

            $Widget.UpdateLayout()
        }
        else {
            $Widget.Topmost = $true

            $Minimize_Icon.Visibility = [System.Windows.Visibility]::Visible
            $Maximize_Icon.Visibility = [System.Windows.Visibility]::Hidden

            $Widget.UpdateLayout()
        }
    })

    #<Ellipse x:Name="SuspendResume_Worktime" Width="8" Height="8" Fill="Green" Stroke="Black" StrokeThickness="0" SnapsToDevicePixels="True" Opacity="0.9" Canvas.Left="2" Canvas.Top="2" />
    #$SuspendResume_Worktime.Add_MouseLeftButtonDown({
    #    if ($SuspendResume_Worktime.Fill -eq [System.Windows.Media.Brushes]::Green) {
    #        $SuspendResume_Worktime.Fill = [System.Windows.Media.Brushes]::Red
    #    }
    #    else {
    #        $SuspendResume_Worktime.Fill = [System.Windows.Media.Brushes]::Green
    #    }
    #})

    $Widget.Add_MouseDoubleClick({
    #    $_.Button -eq [System.Windows.Forms.MouseButtons]::Left
    #    $Widget.Close()
        $ScriptName = $MyInvocation.ScriptName
        Start-Process PowerShell.exe -ArgumentList "-noexit -EP Bypass", "-command $ScriptName -l 60 -h $workinghours -b1 $breakfastBreak -b2 $lunchBreak -p $precision -j $joinIntervals -m $maxWorkingHours -i $showLogoff"
    })

    $updateTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $updateTimer.Interval = New-TimeSpan -Minutes 1
    $updateTimer.Add_Tick($UpdateWidget)
    $updateTimer.Start()

    $syncTimer
    if ($joinIntervals -eq 0) {
        $syncTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $syncTimer.Interval = New-TimeSpan -Minutes 15
        $syncTimer.Add_Tick({Invoke-Command -ScriptBlock $SyncWidget -ArgumentList ([ref]$unplannedBreaks)})
        $syncTimer.Start()

        if ($null -eq $unplannedBreaks -or [double]::IsNaN([double]$unplannedBreaks)) {
            $Tooltip_UnplannedBreaks.Text = '00:00'
        } else {
            $Tooltip_UnplannedBreaks.Text = ((New-Object DateTime) + (New-TimeSpan -Minutes ($unplannedBreaks * 60))).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($script:cultureInfo))
        }
    }
    else {
        $Tooltip_OptionalUnplannedBreaks.Visibility = [System.Windows.Visibility]::Collapsed
    }
    $Widget.Topmost = $topMost
    if ($topPosition -ne -1 -and $leftPosition -ne -1) {
        # 0 = Manual
        $Widget.WindowStartupLocation = 0
        $Widget.Top = $topPosition
        $Widget.Left = $leftPosition
    }

    &$UpdateWidget

    $Widget.ShowDialog() | Out-Null

    $updateTimer.Stop()
    if ($joinIntervals -eq 0) {
        $syncTimer.Stop()
    }
}

function New-WPFMessageBox {

    # For examples for use, see my blog:
    # https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/

    # CHANGES
    # 2017-09-11 - Added some required assemblies in the dynamic parameters to avoid errors when run from the PS console host.

    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 10,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)


        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segoe UI"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)


        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute)
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        # TopMost
        $TopMost = 'TopMost'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)
        $PSBoundParameters.TopMost = "True"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TopMost, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TopMost, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1" Topmost="$($PSBoundParameters.TopMost)">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
            New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }

        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content)
        }
        Catch
        {
            $_
        }
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }


    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }


    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}

$updateWorktimes = {
    Param(
        [Parameter(Mandatory=$false)]
        [int]$eventPeriod = $historyLength
    )
    
    # create an array of filterHashTables that filter boot and shutdown events from the desired period
    $startTime = (get-date).addDays(-$eventPeriod)

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

    # try to get data from archived events also, which should be available up to 10 days (administrator rights needed?!?)
    if (Test-Path "C:\Windows.old\WINDOWS\System32\winevt\Logs\System.evtx" -PathType leaf) {
        if ([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544') {
            $filters += @{ StartTime = $startTime
                           Path = 'C:\Windows.old\WINDOWS\System32\winevt\Logs\System.evtx'
                           ProviderName = 'Microsoft-Windows-Kernel-General'
                           ID = 12, 13
                        }
            $filters += @{ StartTime = $startTime
                           Path = 'C:\Windows.old\WINDOWS\System32\winevt\Logs\System.evtx'
                           ProviderName = 'Microsoft-Windows-Kernel-Power'
                           ID = 42
                        }
            $filters += @{ StartTime = $startTime
                           Path = 'C:\Windows.old\WINDOWS\System32\winevt\Logs\System.evtx'
                           ProviderName = 'Microsoft-Windows-Power-Troubleshooter'
                           ID = 1
                        }
        }
    }
    if (Test-Path "C:\Windows.old\WINDOWS\System32\winevt\Logs\Microsoft-Windows-Winlogon%4Operational.evtx" -PathType leaf) {
        if ([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544') {
            $filters += @{ StartTime = $startTime
                           Path = "C:\Windows.old\WINDOWS\System32\winevt\Logs\Microsoft-Windows-Winlogon%4Operational.evtx"
                           ProviderName = 'Microsoft-Windows-Winlogon'
                           ID = 811
                        }
        }
    }

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
            $diff = $last[0][0] - $end.TimeCreated
            if ($diff.TotalMinutes -ge $breakThreshold) {
              $log[0] = $interval + $last
            } else {
              $log[0][0][0] = $start.TimeCreated
            }
        } else {
            # create new day
            $log.insert(0, $interval)
        }
    }

    # , prevents array unrolling!
    return ,$log
}


if ($mode -eq 'install') {
    "Installation started..."

    #$vbscript = 'schedule.vbs'
    #$vbscriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    #$action = New-ScheduledTaskAction -Execute $vbscript -WorkingDirectory $vbscriptPath -Argument "check -l 1 -h $workinghours -b1 $breakfastBreak -b2 $lunchBreak -p $precision -m $maxWorkingHours -j $joinIntervals -i $showLogoff"
    $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    $action = New-ScheduledTaskAction -Execute conhost.exe -WorkingDirectory $scriptPath -Argument "--headless powershell.exe -File goodTimes.ps1 check -l 1 -h $workinghours -b1 $breakfastBreak -b2 $lunchBreak -p $precision -m $maxWorkingHours -j $joinIntervals -i $showLogoff"
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
elseif ($mode -eq 'uninstall') {
    "Deinstallation started ..."

    if (Get-ScheduledTask 'Check-Worktime' -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask 'Check-Worktime'
    }

    "Ready."

    Exit $LASTEXITCODE
}
elseif ($mode -eq 'install_widget') {
    if ([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544') {
        "Widget installation started..."

        #$vbscript = 'schedule.vbs'
        #$vbscriptPath = split-path -parent $MyInvocation.MyCommand.Definition
        #$action = New-ScheduledTaskAction -Execute $vbscript -WorkingDirectory $vbscriptPath -Argument "widget -l 1 -h $workinghours -b1 $breakfastBreak -b2 $lunchBreak -p $precision -m $maxWorkingHours -j $joinIntervals -i $showLogoff"
        $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
        $action = New-ScheduledTaskAction -Execute conhost.exe -WorkingDirectory $scriptPath -Argument "--headless powershell.exe -File goodTimes.ps1 widget -l 1 -h $workinghours -b1 $breakfastBreak -b2 $lunchBreak -p $precision -m $maxWorkingHours -j $joinIntervals -i $showLogoff"
        $trigger = New-ScheduledTaskTrigger -AtLogon
        $settings = New-ScheduledTaskSettingsSet -Hidden -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -StartWhenAvailable
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings

        if (Get-ScheduledTask 'Start-WorktimeWidget' -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask 'Start-WorktimeWidget'
        }
        Register-ScheduledTask 'Start-WorktimeWidget' -InputObject $task | Out-Null

        "Ready."

        Exit $LASTEXITCODE
    } else {
        Write-Host "Administrator rights needed to install the automatic widget startup at windows logon! (scheduler task with trigger AtLogon)"
    }
}
elseif ($mode -eq 'uninstall_widget') {
    if ([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544') {
        "Widget deinstallation started ..."

        if (Get-ScheduledTask 'Start-WorktimeWidget' -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask 'Start-WorktimeWidget'
        }

        "Ready."

        Exit $LASTEXITCODE
    } else {
        Write-Host "Administrator rights needed to uninstall the automatic widget startup at windows logon!"
    }
}

if ($mode -in ('check', 'widget')) {
    $log = &$updateWorktimes -eventPeriod 1
} else {
    $log = &$updateWorktimes
}

$entry = $log[-1]
$attrs = getLogAttrs($entry)

# show dialog only within the first 5 minutes after startup
if (($attrs.uptime.TotalMinutes -lt 5.0) -and (Test-Path "C:\Windows.old\WINDOWS\System32\winevt\Logs\System.evtx" -PathType leaf)) {
    Write-Host "ATTENTION: Windows upgrade detected and data may be lost in a few days!"

    if (-not ([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544')) {
        Write-Host "Administrator rights needed to get data from Windows.old directory (available up to 10 days after upgrading windows)!"
    }

    $ErrorMsgParams = @{
        Title = 'ATTENTION'
        TitleFontSize = 20
        TitleBackground = 'Red'
        TitleTextForeground = 'WhiteSmoke'
        TitleFontWeight = 'UltraBold'
        Sound = 'Windows Exclamation'
        Timeout = 300
    }
    Try {
        New-WPFMessageBox @ErrorMsgParams -Content "Windows upgrade detected and data may be lost in a few days!&#10;&#10;Administrator rights needed to get data from Windows.old directory (available up to 10 days after upgrading windows) $attrs.uptime.TotalMinutes!"
    }
    Catch {
        $Shell = new-object -comobject wscript.shell -ErrorAction Stop
        $Shell.popup("Windows upgrade detected and data may be lost in a few days!`n`nAdministrator rights needed to get data from Windows.old directory (available up to 10 days after upgrading windows)!", 0, 'ATTENTION', 48 + 4096) | Out-Null
    }
}

if ($mode -eq 'check') {
    $minutes = [Math]::Round($maxWorkingHours * 60) - [Math]::Round($attrs.uptime.TotalMinutes)

    if ($attrs.uptime.TotalHours -ge $breakDeduction1) {
        $minutes += [Math]::Round($breakfastBreak * 60)
    }
    if ($attrs.uptime.TotalHours -ge $breakDeduction2) {
        $minutes += [Math]::Round($lunchBreak * 60)
    }

    if ($attrs.bookingHours -ge $workinghours -and $attrs.bookingHours -lt ($workinghours + 0.08333)) {   # 0.08333 = 5 minutes (interval for check)
        $time = ((Get-Date) + (New-TimeSpan -Minutes $minutes)).ToString("HH:mm", [Globalization.CultureInfo]::getCultureInfo($cultureInfo))

        $InfoParams = @{
            Title = 'INFORMATION'
            TitleFontSize = 20
            TitleBackground = 'LightSkyBlue'
            TitleTextForeground = 'Black'
            Sound = 'Windows Exclamation'
        }
        Try {
            New-WPFMessageBox @InfoParams -Content "Normal worktime reached!&#10;&#10;Max worktime reached at $time.&#10;&#10;You can leave now!"
        }
        Catch {
            $Shell = new-object -comobject wscript.shell -ErrorAction Stop
            $Shell.popup("Normal worktime reached!`n`nMax worktime reached at $time.`n`nYou can leave now!", 0, 'Normal Worktime', 48 + 4096) | Out-Null
        }
    }
    elseif ($attrs.bookingHours -ge $maxWorkingHours) {
        $ErrorMsgParams = @{
            Title = 'ATTENTION'
            TitleFontSize = 20
            TitleBackground = 'Red'
            TitleTextForeground = 'WhiteSmoke'
            TitleFontWeight = 'UltraBold'
            Sound = 'Windows Exclamation'
            Timeout = 300
        }
        $absMinutes = [Math]::Abs($minutes)
        Try {
            New-WPFMessageBox @ErrorMsgParams -Content "Maximum worktime reached since $absMinutes minutes!!!&#10;&#10;You must leave now!!!"
        }
        Catch {
            $Shell = new-object -comobject wscript.shell -ErrorAction Stop
            $Shell.popup("Maximum worktime reached since $absMinutes minutes!!!`n`nYou must leave now!!!", 0, 'Maximum Worktime', 48 + 4096) | Out-Null
        }
    }
    elseif ($attrs.bookingHours -ge ($maxWorkingHours - 0.25)) {  # 0.25 = 15 minutes
        $WarningParams = @{
            Title = 'WARNING'
            TitleFontSize = 20
            TitleBackground = 'Orange'
            TitleTextForeground = 'Black'
            Sound = 'Windows Exclamation'
            Timeout = 300
        }
        Try {
            New-WPFMessageBox @WarningParams -Content "Maximum worktime reached in $minutes minutes.&#10;&#10;You should leave now!"
        }
        Catch {
            $Shell = new-object -comobject wscript.shell -ErrorAction Stop
            $Shell.popup("Maximum worktime reached in $minutes minutes.`n`nYou should leave now!", 0, 'Maximum Worktime Warning', 48 + 4096) | Out-Null
        }
    }

    Exit $LASTEXITCODE
}
elseif ($mode -eq 'widget') {
    $unplannedBreaks = 0

    #Add-Type -Name Window -Namespace Console -MemberDefinition '
    #[DllImport("Kernel32.dll")]
    #public static extern IntPtr GetConsoleWindow();
    #[DllImport("User32.dll")]
    #public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    #'

    #$hWindow = [Console.Window]::GetConsoleWindow()
    #[Console.Window]::ShowWindow($hWindow, 0) | Out-Null

    if ($joinIntervals -eq 0) {
        $unplannedBreaks = (Get-Date).TimeOfDay.TotalHours - $entry[0][0].TimeOfDay.TotalHours - $attrs.uptime.TotalHours
    }

    Show-Widget $entry[0][0].TimeOfDay.TotalHours $workinghours $maxWorkingHours $breakfastBreak $lunchBreak $breakDeduction1 $breakDeduction2 $unplannedBreaks

    #[Console.Window]::ShowWindow($hWindow, 4) | Out-Null

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
    $dayFormatted = $firstStart.Date.toString($dateFormat, [Globalization.CultureInfo]::getCultureInfo($cultureInfo))
    $attrs = getLogAttrs($entry)

    if ($dayOfWeek -lt $lastDayOfWeek) {
        println ("{0,-$($screenWidth - 1)}" -f '------------- ------ ---------  ---------------------')
    }
    $lastDayOfWeek = $dayOfWeek

    if ($dayOfWeek -ge 5) {
        Write-Host $dayFormatted -n -backgroundColor darkred -foregroundcolor gray
    } else {
        print $dayFormatted
    }

    if ($attrs.bookingHours -ge $maxWorkingHours) {
        print ('  {0,5}  ' -f
                $attrs.bookingHours.toString('#0.00', [Globalization.CultureInfo]::getCultureInfo($cultureInfo))
              ) red
    } else {
        print ('  {0,5}  ' -f
                $attrs.bookingHours.toString('#0.00', [Globalization.CultureInfo]::getCultureInfo($cultureInfo))
              ) cyan
    }

    print ('  {0,6}  ' -f $attrs.flexTime[0]) $attrs.flexTime[1]

    print ("{0,3:#0}:{1:00} | {2,-$($screenWidth - 42)}" -f
        $attrs.uptime.hours,
        [Math]::Round($attrs.uptime.minutes + ($attrs.uptime.seconds / 60)),
        $attrs.intervals) darkGray
    Write-Host
}

# restore previous colors
$host.UI.RawUI.BackgroundColor = $oldBgColor
$host.UI.RawUI.ForegroundColor = $oldFgColor

wait
