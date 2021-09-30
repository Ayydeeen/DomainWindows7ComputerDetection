$RawPC=''
$ComputerList=''



#Getting list of computers in AD that have any Windows 7 edition OS installed
$RawPC = get-adcomputer -filter {OperatingSystem -like 'Windows 7 *'} -properties Name -SearchBase 'ou=CHANGEME,DC=CHANGEME,DC=CHANGEME'

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




foreach ($comp in $RawPC) {
    if ($comp.Enabled) {
        
        #Gather data for each computer that's enabled
        $pcinfo = Get-ADComputer $comp.Name -Properties lastlogontimestamp | Select-Object @{Name="Computer";Expression={$_.Name}}, @{Name="Lastlogon";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}      
        
        #Execute Get-LoggedOnUser function for current computer
        $username = Get-LoggedOnUser -ComputerName $comp.Name | Select-Object username

        #If the computer was offline and the $username variable is blank, see if there is an SID identifier from the lastlogon property and notate
        $lastuserlogoninfo = Get-WmiObject -Class Win32_UserProfile -ComputerName $comp.name | Select-Object -First 1
        if ($lastuserlogoninfo.SID) {$SecIdentifier = New-Object System.Security.Principal.SecurityIdentifier($lastuserlogoninfo.SID)}
        if (!$username) { $username = $SecIdentifier.Translate([System.Security.Principal.NTAccount]) }


        $out = write-output ($pcinfo.Computer + "," + $pcinfo.Lastlogon + "," + $username.value)
        add-content -Path win7.csv -Value $out 
    }
}

