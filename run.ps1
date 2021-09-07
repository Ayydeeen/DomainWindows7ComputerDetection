#Getting list of computers in AD that have any Windows 7 edition OS installed
$computers1 = get-adcomputer -filter {OperatingSystem -like 'Windows 7 *'} -properties Name -SearchBase 'ou=<YOUR OU HERE>,DC=<YOUR DOMAIN HERE>,DC=<YOUR DOMAIN HERE>'

#Function to grab logged in user if the computer is validates as being online
function Get-LoggedOnUser {
   [CmdletBinding()]
   param (
     [Parameter()]
     [ValidateScript({ Test-Connection -ComputerName $_ -Quiet -Count 3 })]
     [ValidateNotNullOrEmpty()]
     [string[]]$ComputerName = $env:COMPUTERNAME
   )
   foreach ($comp in $ComputerName) {
     $output = @{ 'ComputerName' = $comp }
     $output.UserName = (Get-WmiObject -Class win32_computersystem -ComputerName $comp).UserName
     [PSCustomObject]$output
   }
 }

#Gather data for each computer in $computers1
foreach ($computer in $computers1) { 
  $pcinfo = Get-ADComputer $computer.Name -Properties lastlogontimestamp | ` 
       Select-Object @{Name="Computer";Expression={$_.Name}}, ` 
       @{Name="Lastlogon";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}

  #Execute Get-LoggedOnUser function for current computer
  $username = Get-LoggedOnUser -ComputerName $computer.Name | Select-Object username

  #If the computer was offline and the $username variable is blank, see if there is an SID identifier from the lastlogon property and notate
  $SecIdentifier = New-Object System.Security.Principal.SecurityIdentifier($lastuserlogoninfo.SID)
  if (!$username) { $username = $SecIdentifier.Translate([System.Security.Principal.NTAccount]) }

  #Creating object property table with gathered info
  $properties = @{'Computer'=$pcinfo.Computer;
          'LastLogon'=$pcinfo.Lastlogon;
          'User'=$username.value
          } #end $properties

  #Append info to win7.csv file in current directory
  $out = write-output (New-Object -Typename PSObject -Property $properties)
  add-content -Path win7.txt -Value $out 
};
