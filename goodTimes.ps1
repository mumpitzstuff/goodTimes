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
        $dateFormat = 'ddd dd/MM/yyyy',
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

# global configuration variables (e.g. en-GB or en-US would be possible)
$cultureInfo = 'de-DE'

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
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
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

    Begin {
        Add-Type -AssemblyName PresentationFramework
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
    $minutes = [Math]::Round($maxWorkingHours * 60) + [Math]::Round($lunchbreak * 60) - [Math]::Round(($attrs.uptime.hours * 60) + $attrs.uptime.minutes + ($attrs.uptime.seconds / 60))
    
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
