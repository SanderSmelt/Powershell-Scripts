Function Get-DateFromWeekNumber {
	<#
	.SYNOPSIS
	Get the date of the specified day, week and year

	.DESCRIPTION
	Get the date of the specified day, week and year
	
	.NOTES
	Filename: Get-DateFromWeekNumber.psm1 
	Version: 1.0 
	Author: Sander Smelt 
	Creation Date: 21-06-2019
	
	.LINK
	https://github.com/SanderSmelt
	
	.PARAMETER Week
	specifie number of the week between 1 and 52

	.PARAMETER Day
	Specifie the day. the current day is the default.
	
	.PARAMETER Year
	Specifie the year. the current year is the default.

	.INPUTS
	None. You cannot pipe objects to get-datefromweeknumber.

	.OUTPUTS
	System.DateTime object of the requested day

	.EXAMPLE
	PS> get-datefromweeknumber -week 1
	Friday, January 4, 2019 12:43:42 PM
	
	.EXAMPLE
	PS> get-datefromweeknumber -week 1 -day "monday"
	Monday, December 31, 2018 12:43:42 PM
	
	.EXAMPLE
	PS> get-datefromweeknumber -week 10 -day "wednesday" -year 2020
	Wednesday, March 4, 2020 12:43:42 PM
	#>
	param(
		[Parameter(Mandatory = $true)]
		[ValidateRange(1,52)]
		[int]$week,
		[int]$Year = ((Get-Date).year),
		[string]$Day = ((Get-Date).DayOfWeek)
	)
	
	$daynumber = (New-Object system.globalization.cultureinfo((Get-Culture).name)).datetimeformat.daynames.ToLower().IndexOf($day.tolower())
	if ($daynumber -lt 0){$daynumber = (New-Object system.globalization.cultureinfo("en-EN")).datetimeformat.daynames.ToLower().IndexOf($day.ToLower())}
	$timestamp = Get-Date -Day 1 -Month 1 -year $year
	$timestamp = ($timestamp).AddDays(-(($timestamp).DayOfWeek.value__))
	([datetime]"$timestamp").AddDays((($week - 1)*7) + $daynumber)
}
Export-ModuleMember -Function Get-DateFromWeekNumber