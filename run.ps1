#Variable Declaration

$RawPC=''
$ComputerList=''

-------------------------
#Functions

#Function to grab logged in user if the computer is validates as being online
function Get-LoggedOnUser($name) {
   if (Test-Connection -ComputerName $name -Quiet -Count 1) {
      $output = @{ 'ComputerName' = $Name }
      $output.UserName = (Get-WmiObject -Class win32_computersystem -ComputerName $name).UserName
      [PSCustomObject]$output
   }
   else {
      write-output($name + " is offline.")
   }
 }

------------------------
#Main

#Getting list of computers in AD that have any Windows 7 edition OS installed
$RawPC = get-adcomputer -filter {OperatingSystem -like 'Windows 7 *'} -properties Name -SearchBase 'ou=MHCOMPUTERS,DC=MHCLINICAL,DC=local'

#Add headers to .csv file
add-content -Path win7.csv -Value "Computer Name , Last Logon , Current User (if found)"

#For every Windows 7 computer detected in the get-adcomputer command
foreach ($comp in $RawPC) {
   #Check to see if computer is Enabled on the company domain
   if ($comp.Enabled) {
        #Gather data for each computer that's enabled
        $pcinfo = Get-ADComputer $comp.Name -Properties lastlogontimestamp | Select-Object @{Name="Computer";Expression={$_.Name}}, @{Name="Lastlogon";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}      
        
        #Execute Get-LoggedOnUser function for current computer
        $username = Get-LoggedOnUser($comp.Name) | Select-Object username

        if ($username) { $username = "n/a" }
        
        $out = write-output ($pcinfo.Computer + "," + $pcinfo.Lastlogon + "," + $username.value)
        write-output($out)
        add-content -Path win7.csv -Value $out 
    }
}

