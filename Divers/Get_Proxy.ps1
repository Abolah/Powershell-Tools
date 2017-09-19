clear
import-module activedirectory
$result = Get-ADComputer -Filter 'Name -like ""'
$result.GetValue(2) | Export-Csv D:\Powershell\export2.csv