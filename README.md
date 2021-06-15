# Good Times!

This is a little PowerShell script that helps you track your working hours by analyzing your machine's uptime.
It reads the required information from the system log and outputs it as a table.\
The script does not store any data on your computer. It only uses the data that Windows automatically provides.

## Usage

Use the batch files to get a first impression!

* Read and follow the instructions on [PowerShell’s execution policy][1]

then

* Start PowerShell
* Type `.\goodTimes.ps1 <any parameters>` 

or

* Run `powershell.exe -file goodTimes.ps1 <any parameters>`


### Command-Line Arguments

* `install`
  Installs a scheduler in Windows Task Scheduler to display a warning when the maximum allowed working hours has been reached (default parameters are used if no further parameters are specified).
* `install_widget`
  Installs a scheduler in Windows Task Scheduler to display a widget at logon (installation needs admin rights).
* `uninstall`
  Uninstalls the scheduler to display a warning when the maximum allowed number of working hours has been reached.
* `uninstall_widget`
  Uninstalls the scheduler to display a widget at logon (deinstallation needs admin rights).
* `check`
  Checks if the maximum allowed number of working hours has already been reached and displays a warning if necessary (normally only called internally by the scheduler).
* `widget`
  Launches a widget to display all relevant information about your working time.
* `-historyLength` (Alias `-l`)
  Number of days to show in uptime history. Defaults to `60`.
* `-workingHours` (Alias `-h`)
  Working hours per day, used for overtime calculation. Defaults to `8`.
* `-breakfastBreak` (Alias `-b1`)
  Length of breakfast break in hours per day. This will be added to your daily work time. Defaults to `0.25`.
* `-lunchBreak` (Alias `-b2`)
  Length of lunch break in hours per day. This will be added to your daily work time. Defaults to `0.50`.
* `-precision` (Alias `-p`)
  Rounding precision in percent, where 1 = round to the hour, 2 = round to 30 minutes, etc. Defaults to `60`.
* `-dateFormat` (Alias `-d`)
  Date format according to [the .NET reference][2]. Defaults to `ddd dd/MM/yyyy`.
* `-joinIntervals` (Alias `-j`)
  Ignores the breaks between the intervals and combines only the start of the first interval and the end of the last interval (0 = switched off, 1 = switched on). Defaults to `1`.
* `-maxWorkingHours` (Alias `-m`)
  Number of maximum allowed hours per day that can be worked. Defaults to `10`.
* `-showLogoff` (Alias `-i`)
  Show logoff/login events and lockscreen on/off events (0 = switched off, 1 = switched on). Defaults to `1`.

### Script Settings
* `$cultureInfo` (Default `de-DE`)
  Can be set e.g. to en-GB or en-US.
* `$breakDeduction1` (Default `3.0`)
  After how many hours should the breakfastBreak be deducted.
* `$breakDeduction2` (Default `6.0`)
  After how many hours should the lunchBreak be deducted.

### Example Pictures
![worktimes](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/worktimes.png?raw=true)\
![normal worktime reached message](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/normal_worktime_reached.png?raw=true)
![max worktime reached warning](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/max_worktime_reached.png?raw=true)
![max worktime reached error](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/max_worktime_reached1.png?raw=true)\
![widget with unplanned breaks](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/widget_with_unplanned_breaks.png?raw=true)
![widget without unplanned breaks](https://github.com/mumpitzstuff/goodTimes/blob/master/docu/widget_without_unplanned_breaks.png?raw=true)


## i18n

Sorry, this version is in English only, so if you want some internationalization, you can easily edit the script. 
 
 
[1]: http://stackoverflow.com/questions/10635/why-are-my-powershell-scripts-not-running
[2]: https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx?cs-lang=vb#content
