#Common Powershell Functions
#Written by Luis Orta
#Contains a few tools here and there to help with time consuming tasks.
function Get-OpenFile{
<#
.Synopsis
   Checks file servers for open files by file name. 
.DESCRIPTION
   Outputs object based data to the pipeline. Object Fields are Locks, OpenMode, File, Hostname, and Accessedby. Computername can accept multiple inputs.
.EXAMPLE
   Get-OpenFiles -ComputerName fileserver.contoso.com -FileName "Word.doc"
.EXAMPLE
   Get-OpenFiles -ComputerName Fileserver1, Fileserver2, Fileserver3 -FileName "reports"
.EXAMPLE
   Get-ADComputer fileserver1 | Select-Object -Property DNSHostName | Get-OpenFiles -FileName "docx"
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/01/21/get-openfiles/')]
    param (
        # valid fileserver name here. Can accept multiple values
        [Parameter (Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position = 0)]
        [Alias('Hostname','DNSHostName')]
        [string[]]$Computername,

        # Filename or part of filename. Single value only.
        [Parameter(Mandatory=$false,
                   Position = 1)]
        [string]$FileName
    )
    Process{
    foreach ($Computer in $Computername){
                try{
                    openfiles.exe /query /s $Computer /fo csv /V | Out-File -Force $env:TEMP\openfiles.csv -ErrorAction Stop
                    $Files = Import-CSV $env:TEMP\openfiles.csv
                        foreach ($File in $Files){
                            $properties = @{'Hostname'=$Computer
                                            'AccessedBy'=$File."Accessed By"
                                            'Locks'=$File."#Locks"
                                            'OpenMode'=$File."Open Mode"
                                            'File'=$File."Open File (Path\executable)"}
                            $obj = New-Object -TypeName PSObject -Property $properties -ErrorAction Stop
                            
                            Write-Output $obj | Where-Object {$obj.File -match $FileName}
                        }
                    Remove-Item $env:TEMP\openfiles.csv
                }
                catch{
                    $properties = @{'Hostname'=$Computer
                                    'AccessedBy'=$Null
                                    'Locks'=$Null
                                    'OpenMode'=$Null
                                    'File'=$Null}
                    Write-Warning "Error getting open files."
                }
       }
    }
    }
function Get-AddRemoveProgram{
<#
.Synopsis
   Looks for installed programs on a computer by a program name. Only part of the name is required to perform match.
   A quick warning, this cmdlet is slow.
.DESCRIPTION
   Outputs object based data to the pipeline.Computername can accept multiple inputs.
.EXAMPLE
   Get-AddRemovePrograms -ComputerName server.contoso.com -ProgramName "*Microsoft*"
.EXAMPLE
   Get-AddRemovePrograms -ComputerName server, server2, server3 -ProgramName "*Microsoft*"
.EXAMPLE
   Get-ADComputer server1 | Select-Object -Property DNSHostname | Get-AddRemovePrograms -ProgramName *Microsoft*
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/01/26/get-addremoveprograms/')]
    Param(
        # valid server name here. Can accept multiple values
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage="Enter a Valid Computer Name")]
        [Alias('Hostname','DNSHostName')]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$False,
                   HelpMessage="Enter a part of the program name Example:Office")]
        $ProgramName = "*"
        )
        PROCESS {
                foreach ($computer in $ComputerName){
            try{
            $programs = Invoke-Command -ComputerName $computer{
             Param($ProgramName)
             $32bit = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
             $64bit = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*
             return $32bit + $64bit} -ArgumentList $ProgramName -ErrorAction Stop
             #$programs = $32bit + $64bit

    foreach ($program in $programs){
                    $properties = @{ComputerName = $computer
                                    ProgramName = $program.DisplayName
                                    Publisher = $program.Publisher
                                    Version = $program.DisplayVersion}
                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj | Where-Object {$obj -like $ProgramName}
                    }

            }

        catch{
                    $properties = @{ComputerName = $computer
                                    ProgramName = $null
                                    Publisher = $null
                                    Version = $null}
                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
        

        }

   }
}
}
function Get-NetLocalGroup{
<#
.Synopsis
   Checks local groups on local or remote computers and displays the users in the group.
.DESCRIPTION
   Command check which users are in the specified group on a local or remote computer. 
   Parameters are not mandatory in this command. ComputerName by default is the localhost
   Group by default is Administrators. Accepts Pipeline Input.
.EXAMPLE
   Get-NetLocalGroup
Description     
------------
Returns local administrators of the current computer
.EXAMPLE
    Get-NetLocalGroup -ComputerName Server1 -Group Administrators
Description
-----------
Returns local administrators of the remote computer.
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/02/19/get-netlocalgroup/')]
    Param(
        # ComputerName
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias('ComputerNames','Hostname','DNSHostName')] 
        [string[]]$Computername="LocalHost",

        # Local Computer Group
        [Parameter(Mandatory=$false,
                   Position=1)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Group = "Administrators"
    )

    Begin{
    }
    Process{
        foreach ($Computer in $Computername){
                #Connect to remote computer and begin mining data
                $ADGroup = Invoke-Command -ComputerName $Computer -ScriptBlock {
                net localgroup $args[0]
                Write-Output $ADGroup
                } -ArgumentList $Group
                    for ($i=6; $i -lt $ADGroup.length-3; $i++){
                        #Return as hash table and turn into PSObjects.
                        $properties = @{ComputerName = $Computer
                                                UserName = $ADGroup[$i]}
                        $obj = New-Object -TypeName PSObject -Property $properties
                        Write-Output $obj
                    }
                
 
                }
                                                        
        }
}
function Set-PrinterLocation{
<#
.Synopsis
   Sets Location information on remote printers or local hosts
.DESCRIPTION
   Sets location property on remote print queues or local print queues. This command takes 3 parameters. 
   Only the location is a mandatory parameter. Server will default to localhost and printer will default to all printers.
.EXAMPLE
   The Following sets the location on remote print queus matching the sharename Print
   Set-PrinterLocation -Server Printserver1 -Location "Redmond, WA" -Printer *Print*
.EXAMPLE
   The following sets the location on all local print queues to Redmond, WA
   Set-PrinterLocation -Server Printserver1 -Location "Redmond, WA"
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/02/19/get-printerlocation-set-printerlocation/')]
    Param
    (
        # A Valid Print Server Name
        [Parameter(Mandatory=$True,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Server = "LocalHost",

        # Please Enter the location you want to set
        [Parameter(Mandatory=$true,
                   Position=1)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Location,

        # Filter for which Printers. Use wildcards if necessary
        [Parameter(Position=2)]
        [Alias("ShareName")]
        [string[]]$Printer = "*"
    )

    Begin{
    $PrintQueues = Get-WmiObject -Class Win32_Printer -ComputerName $Server | Where-Object {$_.ShareName -like $Printer}
    }
    Process{
            foreach ($PrintQueue in $PrintQueues){
                $PrintQueue.Location = $Location
                $PrintQueue.Put();
            }
    }
}
function Get-PrinterLocation{
<#
.Synopsis
   Gets Location information on remote printers or local hosts
.DESCRIPTION
   Gets location property on remote print queues or local print queues 
   Server will default to localhost and printer will default to all printers.
.EXAMPLE
   The Following gets the location on remote print queus matching the sharename Print
   Get-PrinterLocation -Server Printserver1 -Printers *Print*
.EXAMPLE
   The following gets the location on all local print queues
   Get-PrinterLocation
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/02/19/get-printerlocation-set-printerlocation/')]
    Param
    (
        # Server
        [Parameter(Mandatory=$False,  
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Server="LocalHost",

        # Printer ShareName
        [Parameter(Position=1)]
        [string]$Printers = "*"
    )

    Begin{
    $PrintQueues = Get-WmiObject -Class Win32_Printer -ComputerName $Server | Where-Object {$_.ShareName -like $Printers}
    }
    Process{
            foreach ($PrintQueue in $PrintQueues){
                            $properties = @{'Server'=$Server
                                            'ShareName'=$PrintQueue."ShareName"
                                            'Location'=$PrintQueue."Location"}
                            $obj = New-Object -TypeName PSObject -Property $properties
                            Write-Output $obj
            }
    }
}
function Get-ADSubnet{
<#
.Synopsis
   Finds Active directory sites that match the requested subnet.
.DESCRIPTION
    Use this tool to find a corresponding active directory site for an IP address. Wildcards can be placed at any octet in this command. 
.EXAMPLE
   The Following gets the active directory site for the subnet 10.0.0.*. 
   Get-ADSubnet -IPAddress 10.0.0.*
.EXAMPLE
   The following gets the active directory sites for multiple subnets
   Get-ADSubnet -IPAddress 10.0.0.*,192.168.9.*
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/03/11/get-adsubnet/')]
    Param
    (
        # Enter an IP Address Space. Ex: 10.0.0.*
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage="Enter a Valid Computer Name",
                   Position=0)]
        #Need to consider what other cmdlets provide IP addresses in the pipeline.
        #[Alias('Name')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]

        [string[]]$IPAddress
    )
    Begin{
    #This grabs the configuration database in the active directory schema in a text format we can feed into the get-adobject command.
    $Configuration = (Get-ADDomain | Select-Object SubordinateReferences).SubOrdinateReferences | Select-String -Pattern "Configuration"
    }

    Process{
        #First loop for every IP address entered as a parameter.
        foreach ($IP in $IPAddress){
         $Sites = (get-adobject -filter 'ObjectClass -eq "site"' -SearchBase $Configuration -Properties siteObjectBL) | where {$_.siteObjectBL -like ("*" + $IP)}#).siteObjectBL
           #Next we loop through the all of the possible return sites. This allows us to separate them in pipeline output for single objects.
            foreach ($Site in $Sites){
            #One more loop to go through all of the subnets that return in each site object. They are nested arrays, so this part separates each IP address to make clean pipeline output.
                            foreach ($SubnetCN in $Site.siteObjectBL){
                            #Cleanup the string and return only the IP and subnet
                            $Subnet = $SubnetCN.split("="",")[1]
                            #turn it into a a hash table and return as objects.
                            $properties = @{'Site'=$Site.Name
                                           'Subnet'=$Subnet
                                            }
                            $obj = New-Object -TypeName PSObject -Property $properties
                            Write-Output $obj
                            }
            }
        }
    }
}

function Get-LoggedOnUser{
<#
.Synopsis
   Retreives currently logged in domain users on remote computers
.DESCRIPTION
   Retreives the list of currently logged in computers in the WMI object Win32_loggedonuser and outputs domain users logged in that are not the current user running the script.
.EXAMPLE
   Get-LoggedOn -Computername computer1
.EXAMPLE
   get-adcomputer -filter {name -like "*computer*"} | select -expandproperty name | get-loggedon
  #>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/06/02/get-loggedonuser/')]
    [OutputType([String])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("name","cn","computer")] 
        [string[]]$computername
    )

    Begin
    {
    #Regex Expression to match the domain names and the user name
    $regex = '.+Domain="(.+)",Name="(.+)"$'
    #environment variable for the current user to filter later
    $currentuser = $env:username
    }
    Process
    {
        try
        {
        foreach ($computer in $computername)
            {
                #Enumerate the logged in users
                $users = Get-WmiObject -ComputerName $computer -Class Win32_LoggedOnUser -ErrorAction Stop | Select Antecedent -unique 
                #Check each output
                foreach($user in $users)
                    {
                        #Match against the regex
                        $user.antecedent -match $regex > $null
                        #If the matches variable contains the computeraname (local accounts) or the user running the command, do not output
                        if ($matches[1] -ne $computer -and $matches[2] -notlike $currentuser)
                        {
                            $properties = @{'ComputerName'=$computer
                                            'User'=$matches[2]}
                            $obj = New-Object -TypeName PSObject -Property $properties
                            Write-Output $obj
                        }
                        #clear the matches variable for the next user
                        $matches.clear()
                    }
            }
        }
        catch
        {
        Write-Error "Unable to connect to $computer to retreive user names"
        }
    }
}

function Get-ADPrinter 
{
<#
.SYNOPSIS
    Finds printers that have been published in active directory.
.DESCRIPTION
    Finds printers published in AD and returns information about the printer. Things like the server they are on, portname, UNC path, Driver, and Location.
.EXAMPLE
    Get-ADPrinter -Printer TestPrinter01
.EXAMPLE
    Get-ADPrinter -Printer TestPrinter01,TestPrinter02,TestPrinter03
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/06/11/get-adprinter')]
    [OutputType([String])]
    Param (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("CN")] 
        [string[]]$Printer
    )
    
    process 
    {
        foreach ($Print in $Printer)
        {
            try 
            {
                    $Print = "*$Print"
                    $DN = Get-ADObject -Filter {ObjectClass -eq "printQueue" -and Name -like $Print} -Properties printerName,serverName,portName,uNCName,driverName,location
                    foreach ($D in $DN)
                    {
                        $properties = @{'Printer'=$D.printerName
                                        'Server'=$D.serverName
                                        'PortName'=$D.portName
                                        'UNC'=$D.uNCName
                                        'Driver'=$D.driverName
                                        'Location'=$D.location}
                        $obj = New-Object -TypeName PSObject -Property $properties
                        Write-Output $obj
                    }
            }
            
            catch 
            {
                    Write-Warning "No Valid Printer found"
                    $properties = @{'Printer'=$Print
                                    'Server'=$Null
                                    'PortName'=$Null
                                    'UNC'=$Null
                                    'Driver'=$Null
                                    'Location'=$Null}
                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
            }
        }    
    }
    
}
function Get-ADFolderACL {
<#
.SYNOPSIS
    Gets Active Directory Groups and Users of a file directory
.EXAMPLE
    This example gets the the top level groups and users ACLs.
    Get-ADFolderACL -Path \\Test-Server\Folder Location
.EXAMPLE
    This example will get all users and recurse through the groups to return the users in those groups.
    Get-ADFolderACL -Path \\Test-Server\Folder -Recurse
#>
    [CmdletBinding()]
    Param (
        # Enter a valid local or UNC path
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string[]]$Path,
        
        # Return all users of the groups
        [Parameter(Mandatory=$false,
                    Position=0)]
        [switch]$Recurse
    )
    
    process {
        foreach ($Pat in $Path)
        {
            Write-Verbose "Obtaining ACLS"
            $acls = Get-ACL -Path $Pat | ForEach-Object {$_.Access}
            if ($Recurse)
            {
            $Users = foreach ($acl in $acls)
                {   
                    $Filter = $acl.identityreference.tostring().split("\",[System.StringSplitOptions]::RemoveEmptyEntries)[1]
                    if ($Filter -ne $null)
                        {
                        Write-Verbose "Getting $Filter"
                        $User = Get-ADGroupMember -Identity $Filter -Recursive
                        $User = $User | Select-Object -Property Name,distinguishedName,ObjectClass
                        Write-Output $User
                        }
                }
                $Users = $Users | Select-Object -Property Name,distinguishedName,ObjectClass -Unique
                Write-Output $Users
            } 
            else
            {
                    foreach ($acl in $acls)
                    {
                        $Filter = $acl.identityreference.tostring().split("\",[System.StringSplitOptions]::RemoveEmptyEntries)[1]
                        if ($Filter -ne $null)
                            {
                            Write-Verbose "Getting $Filter"
                            $Users = Get-ADObject -Filter {SamAccountName -eq $Filter}
                            $Users = $Users | Select-Object -Property Name,distinguishedName,ObjectClass -Unique
                            Write-Output $Users
                            }
                    }
            }     
        }
    }
}
function Get-ADFolderACL {
<#
.SYNOPSIS
    Gets Active Directory Groups and Users of a file directory
.EXAMPLE
    This example gets the the top level groups and users ACLs.
    Get-ADFolderACL -Path \\Test-Server\Folder Location
.EXAMPLE
    This example will get all users and recurse through the groups to return the users in those groups.
    Get-ADFolderACL -Path \\Test-Server\Folder -Recurse
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/07/25/get-adfolderacl/')]
    Param (
        # Enter a valid local or UNC path
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string[]]$Path,
        
        # Return all users of the groups
        [Parameter(Mandatory=$false,
                    Position=0)]
        [switch]$Recurse
    )
    
    process {
        foreach ($Pat in $Path)
        {
            Write-Verbose "Obtaining ACLS"
            $acls = Get-ACL -Path $Pat | ForEach-Object {$_.Access}
            if ($Recurse)
            {
            $Users = foreach ($acl in $acls)
                {   
                    $Filter = $acl.identityreference.tostring().split("\",[System.StringSplitOptions]::RemoveEmptyEntries)[1]
                    if ($Filter -ne $null)
                        {
                        Write-Verbose "Getting $Filter"
                        $User = Get-ADGroupMember -Identity $Filter -Recursive
                        $User = $User | Select-Object -Property Name,distinguishedName,ObjectClass
                        Write-Output $User
                        }
                }
                $Users = $Users | Select-Object -Property Name,distinguishedName,ObjectClass -Unique
                Write-Output $Users
            } 
            else
            {
                    foreach ($acl in $acls)
                    {
                        $Filter = $acl.identityreference.tostring().split("\",[System.StringSplitOptions]::RemoveEmptyEntries)[1]
                        if ($Filter -ne $null)
                            {
                            Write-Verbose "Getting $Filter"
                            $Users = Get-ADObject -Filter {SamAccountName -eq $Filter}
                            $Users = $Users | Select-Object -Property Name,distinguishedName,ObjectClass -Unique
                            Write-Output $Users
                            }
                    }
            }     
        }
    }
}
function Get-GlobalPrinter {
<#
.SYNOPSIS
    Gets Globally Installed printers on local or remote computers.
.EXAMPLE
    This example gets all printers on local computer.
    Get-GlobalPrinter

Printer      UNC                            Server            Computername  
-------      ---                            ------            ------------  
TestPrinter1 \\Serv1.test1.com\TestPrinter1 \\Serv1.test1.com TestPC01
TestPrinter2 \\Serv2.test1.com\TestPrinter2 \\Serv2.test1.com TestPC01
.EXAMPLE
    This example will get the printer TestPrinter2 from the remote computer Test-PC02
    Get-GlobalPrinter -Computername Test-PC01 -Printer TestPrinter2
Printer      UNC                            Server            Computername  
-------      ---                            ------            ------------  
TestPrinter2 \\Serv2.test1.com\TestPrinter2 \\Serv2.test1.com TestPC01
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/07/26/global-printer-bundle/')]
    Param (
        # Provide a valid computername
        [Parameter(Mandatory=$false,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("cn")] 
        [string[]]$Computername = 'localhost',
		# Provide a valid printer name
        [Parameter(Mandatory=$false,
                   Position=1)]
        [string[]]$Printer = '*')
    
    process 
    {
        foreach($Computer in $Computername)
        {
            Write-Verbose "Invoking Command to get printers on $Computer"
            $Printers = Invoke-Command -ComputerName $Computer -ScriptBlock{
                $Printers = Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Print\Connections"
                foreach ($Printer in $Printers)
                {
                    $properties = @{'Printer'=$Printer.GetValue("Printer").Split("\")[-1]
									'UNC'=$Printer.GetValue("Printer")
									'Server'=$Printer.GetValue("Server")
                                    'Computername'=$env:COMPUTERNAME}
                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj | Where-Object Printer -like $args[0]
                }
            } -ArgumentList $Printer
            foreach ($Print in $Printers)
				{
					Write-Output $Print | Select-Object Printer,UNC,Server,ComputerName 
				} 
        }
    }
}
function Add-GlobalPrinter 
{
<#
.SYNOPSIS
    Adds global printers on local or remote computers.
.EXAMPLE
    This example Adds a global printer on the local computer.
    Add-GlobalPrinter -UNC \\Serv1.test1.com\TestPrinter1
.EXAMPLE
    This example Adds multiple global printers on the local computer.
    Add-GlobalPrinter -UNC \\Serv1.test1.com\TestPrinter1,\\Serv2.test1.com\TestPrinter2
.EXAMPLE
    This example Adds multiple global printers on a remote computer.
    Add-GlobalPrinter -Computername TestPC01 -UNC \\Serv1.test1.com\TestPrinter1,\\Serv2.test1.com\TestPrinter2
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/07/26/global-printer-bundle/')]
    Param (
        # Enter a valid computer name
        [Parameter(Mandatory=$false,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("cn")] 
        [string[]]$ComputerName = 'localhost',
        # Enter a valid UNC path to a printer
        [Parameter(Mandatory=$false,
                   Position=1,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
       [string[]]$UNC)      
    
    process 
    {
        foreach($Computer in $Computername)
        {
            Write-Verbose "Invoking Command to Add printers on $Computer"
            Invoke-Command -ComputerName $Computer -ScriptBlock{
                foreach($arg in $args)
                {
                    rundll32 printui.dll PrintUIEntry /q /ga /n $arg
                }
                    
            } -ArgumentList $UNC
        }
    }
}
function Remove-GlobalPrinter 
{
<#
.SYNOPSIS
    Removes global printers on local or remote computers.
.EXAMPLE
    This example removes a global printer on the local computer.
    Remove-GlobalPrinter -UNC \\Serv1.test1.com\TestPrinter1
.EXAMPLE
    This example removes multiple global printers on the local computer.
    Remove-GlobalPrinter -UNC \\Serv1.test1.com\TestPrinter1,\\Serv2.test1.com\TestPrinter2
.EXAMPLE
    This example removes multiple global printers on a remote computer.
    Remove-GlobalPrinter -Computername TestPC01 -UNC \\Serv1.test1.com\TestPrinter1,\\Serv2.test1.com\TestPrinter2
#>
    [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/07/26/global-printer-bundle/')]
    Param (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("cn")] 
        [string[]]$Computername = 'localhost',
        # Param1 help description
        [Parameter(Mandatory=$false,
                   Position=0,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
       [string[]]$UNC)      
    
    process 
    {
        foreach($Computer in $Computername)
        {
            Write-Verbose "Invoking Command to remove printers on $Computer"
            Invoke-Command -ComputerName $Computer -ScriptBlock{
                foreach($arg in $args)
                {
                    rundll32 printui.dll PrintUIEntry /q /gd /n $arg
                }
                    
            } -ArgumentList $UNC
        }
    }
}
