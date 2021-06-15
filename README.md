# Good Times!

This is a little PowerShell script that helps you track your working hours by analyzing your machine's uptime.
It reads the required information from the system log and outputs it as a table.

## Usage

* Read and follow the instructions on [PowerShell’s execution policy][1]

then

* Start PowerShell
* Type `.\goodTimes.ps1` 

or

* Run `powershell.exe -file goodTimes.ps1`


### Command-Line Arguments

* `install`
  Installs a scheduler to display a warning when the maximum allowed number of working hours has been reached (default parameters are used if no further parameters are specified).
* `uninstall`
  Installs a scheduler to display a warning when the maximum allowed number of working hours has been reached.
* `check`
  Checks if the maximum allowed number of working hours has already been reached and gives a warning if necessary (normally only called internally by the scheduler).
* `-historyLength` (Alias `-l`)
  Number of days to show in uptime history. Defaults to `60`.
* `-workingHours` (Alias `-h`)
  Working hours per day, used for overtime calculation. Defaults to `8`.
* `-breakfastBreak` (Alias `-b1`)
  Length of breakfast break in hours. This will be subtracted from your work time. Defaults to `0.25`.
* `-lunchBreak` (Alias `-b2`)
  Length of lunch break in hours. This will be subtracted from your work time. Defaults to `0.50`.
* `-precision` (Alias `-p`)
  Rounding precision in percent, where 1 = round to the hour, 2 = round to 30 minutes, etc. Defaults to `60`.
* `-dateFormat` (Alias `-d`)
  Date format as defined in [the .NET reference][2]. Defaults to `ddd dd/MM/yyyy`.
* `-joinIntervals` (Alias `-j`)
  Ignores the breaks between intervals and joins all intervals together. Defaults to `1`.
* `-maxWorkingHours` (Alias `-m`)
  Maximum working hours per day. Defaults to `10`.
* `-showLogoff` (Alias `-i`)
  Show Logon/Logoff events (joinIntervals must also be set to 0 to activate this feature). Defaults to `1`.

### Example Pictures
![normal worktime reached message](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/normal_worktime_reached.png?raw=true)
![max worktime reached warning](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/max_worktime_reached.png?raw=true)
![max worktime reached error](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/max_worktime_reached1.png?raw=true)
![widget with unplanned breaks](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/widget_with_unplanned_breaks.png?raw=true)
![widget without unplanned breaks](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/widget_without_unplanned_breaks.png?raw=true)


## i18n

Sorry, this version is in English only, so if you want some internationalization, you can easily edit the script. 
 
 
[1]: http://stackoverflow.com/questions/10635/why-are-my-powershell-scripts-not-running
[2]: https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx?cs-lang=vb#content
