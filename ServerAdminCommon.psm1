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
       Get-OpenFiles -ComputerName Fileserver1, Fileserver2, Fileserver3 -FileName "*reports*"
    .EXAMPLE
       Get-ADComputer fileserver1 | Select-Object -Property DNSHostName | Get-OpenFiles -FileName "*.docx"
    #>
        [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/01/21/get-openfiles/')]
        param (
            # valid fileserver name here. Can accept multiple values
            [Parameter (Mandatory=$true,
                        ValueFromPipeline=$true,
                        ValueFromPipelineByPropertyName=$true,
                        Position = 0)]
            [Alias('Hostname','DNSHostName')]
            [string[]]$ComputerName,
    
            # Filename or part of filename. Single value only.
            [Parameter(Mandatory=$false,
                       Position = 1)]
            [string]$FileName
        )
        Process{
        foreach ($Computer in $Computername){
                    try{
                        $Files = $Files = openfiles.exe /query /s $ComputerName /fo csv /V | ConvertFrom-Csv -ErrorAction Stop
                            foreach ($File in $Files){
                                $File | Where-Object {$PSItem.'Open File (Path\executable)' -like $FileName}
                            }
                    }
                    catch{
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
            [Parameter(Mandatory=$False,
                       ValueFromPipeline=$True,
                       ValueFromPipelineByPropertyName=$True,
                       HelpMessage="Enter a Valid Computer Name")]
            [Alias('Hostname','DNSHostName')]
            [string[]]$ComputerName = "localhost",
    
            [Parameter(Mandatory=$False,
                       HelpMessage="Enter a part of the program name Example:Office")]
            $ProgramName = "*"
            )
            PROCESS {
                    foreach ($computer in $ComputerName){
                try{
                $programs = Invoke-Command -ComputerName $computer{
                 $32bit = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
                 $64bit = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*
                 return $32bit + $64bit} -ErrorAction Stop
    
        foreach ($program in $programs){
                    $program = Write-Output $program | Where-Object -Property Displayname -like $ProgramName
                        if ($program.DisplayName -ne $Null)
                            {
                            $properties = @{ComputerName = $computer
                                            ProgramName = $program.DisplayName
                                            Publisher = $program.Publisher
                                            Version = $program.DisplayVersion
                                            UninstallString = $program.UninstallString}
                            $obj = New-Object -TypeName PSObject -Property $properties
                            Write-Output $obj
                            }
        }    
                }
    
            catch{ Write-Warning "$Computer was not reachable."
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
             $Sites = (get-adobject -filter 'ObjectClass -eq "site"' -SearchBase $Configuration -Properties siteObjectBL) | Where-Object {$_.siteObjectBL -like ("*" + $IP)}#).siteObjectBL
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
    
    function Get-LoggedOnUser {
    <#
    .Synopsis
       Retreives currently logged in domain users on remote computers
    .DESCRIPTION
       Retreives the list of currently logged in computers in the WMI object Win32_loggedonuser and outputs domain users logged in that are not the current user running the script.
    .EXAMPLE
       Get-LoggedOnUser -Computername computer1
    .EXAMPLE
       get-adcomputer -filter {name -like "*computer*"} | select -expandproperty name | get-loggedonUser
    .EXAMPLE
       Get-LoggedonUser -Computername computer1 -includelocal
      #>
        [CmdletBinding(HelpUri = 'https://luisrorta.com/2017/06/02/get-loggedonuser/')]
        [OutputType([String])]
        Param
        (
            # Param1 help description
            [Parameter(Mandatory=$false, 
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true, 
                       ValueFromRemainingArguments=$false, 
                       Position=0)]
            [ValidateNotNull()]
            [ValidateNotNullOrEmpty()]
            [Alias("name","cn","computer")] 
            [string[]]$Computername = "localhost",
    
            [Parameter(Mandatory=$false,
                       Position=1)]
            [Switch]$IncludeLocal
        )
    
        Process
        {
            try
            {
            foreach ($computer in $computername)
                {
                    #Enumerate the logged in users
                    $users = Get-CimInstance -ComputerName $computer -ClassName Win32_LoggedOnUser | Select-Object Antecedent -Unique 
                    #Check each output and filter out the current user and local service accounts.
                    if ($IncludeLocal) 
                    {
                        foreach($user in $users)
                        {
                            if ($user.Antecedent.name -ne $env:username) 
                            {
                                $obj = New-Object -TypeName PSCustomObject -Property @{'ComputerName' = $user.Antecedent.PSComputerName
                                    'Name' = $user.Antecedent.Name
                                    'Domain' = $user.Antecedent.Domain}
                                Write-Output $obj
                            }
                        }    
                    }
                    else {
                        foreach($user in $users)
                        {
                            if (($user.Antecedent.Domain -ne $user.Antecedent.PSComputerName) -and ($user.Antecedent.Name -ne $env:username) ) 
                            {
                                $obj = New-Object -TypeName PSCustomObject -Property @{'ComputerName' = $user.Antecedent.PSComputerName
                                    'Name' = $user.Antecedent.Name
                                    'Domain' = $user.Antecedent.Domain}
                                Write-Output $obj
                            }
                        }
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
                        $DN = Get-ADObject -Filter {ObjectClass -eq "printQueue" -and PrinterName -like $Print} -Properties printerName,serverName,portName,uNCName,driverName,location
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
    
    function Test-LocalCredential {
    <#
    .SYNOPSIS
        Tests local user account passwords against remote computers to verify credentials
    .EXAMPLE
        $Cred = Get-Credential
        Test-LocalCredential -Credential $Cred -ComputerName TestServ1
    .EXAMPLE
        Test-LocalCredential -Computername TestServ1, TestServ2, TestServ3
    #>
        [CmdletBinding(HelpUri = 'https://luisrorta.com/')]
        Param (
            # Enter valid credentials using get-credential
            [Parameter(Mandatory=$true,
                       Position=0)]
            [ValidateNotNull()]
            [ValidateNotNullOrEmpty()]
            [System.Management.Automation.PSCredential]$Credential = (Get-Credential),
            
            # Enter a computer name.
            [Parameter(Mandatory=$false,
                       Position=1)]
            [Alias("name","cn","computer","PSComputerName")] 
            [String[]]$ComputerName = $env:COMPUTERNAME
        )
        process 
        {
            foreach ($Computer in $ComputerName) 
            {
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $Credential = $Using:Credential
                    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                    $DirectoryObject = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('machine','localhost')
                    $Check = $DirectoryObject.ValidateCredentials($Credential.GetNetworkCredential().Username,$Credential.GetNetworkCredential().Password)
                    $Obj = New-Object -TypeName PSCustomObject -Property @{
                        UserName = $Credential.GetNetworkCredential().Username
                        CredentialCheck = $Check
                    }
                    Write-Output $Obj
                }
    
            }
        }
    }
    
    function Get-Netstat {
    <#
    .SYNOPSIS
        A powershell version of the command line utility netstat.
    .DESCRIPTION
        Uses invoke-command to run netstat remotely and perform string handling to create powershell objects.
    .EXAMPLE
        Get-Netstat -Computername TestServ1, TestServ2
    #>
        [CmdletBinding(HelpUri = 'https://luisrorta.com/')]
        Param (
            # Enter a valid computer Name
            [Parameter(Mandatory=$False,
                        Position=0)]
            [Alias("p1")] 
            [string[]]$Computername = "LocalHost",
            [Parameter(Mandatory=$false,
                        Position=1,
                        ParameterSetName='Listening')]
            [ValidateSet("LISTENING", "ESTABLISHED", "TIME_WAIT", "*")]
            $State = "*",
            [Parameter(Mandatory=$False,
                        Position=2)]
            [ValidateSet("InterNetwork", "InterNetworkV6", "*")]
            $AddressFamily = "*"
        )
        
        process 
        {
            foreach ($Computer in $Computername) 
            {
                Write-Verbose -Message "Connecting to $ComputerName to run Netstat. Will Return State $State and using AddressFamily $AddressFamily"
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $Netstats = NETSTAT.EXE -ANO
                    for ($i = 4; $i -lt $Netstats.Count; $i++) 
                    {
                        $split = $Netstats[$i].split("",[System.StringSplitOptions]::RemoveEmptyEntries)
                        if ($split[0] -eq "TCP")
                        {
                            $obj = new-object -typename pscustomobject -Property @{Proto = $split[0]
                                                                               LocalAddress = [IPAddress]($split[1].Substring(0,$split[1].lastindexof(":")))
                                                                               LocalPort = [int]($split[1].split(":")[-1])
                                                                               RemoteAddress = [IPAddress]($split[2].Substring(0,$split[2].lastindexof(":")))
                                                                               RemotePort = [int]($split[2].split(":")[-1])
                                                                               State = $split[3]
                                                                               ProcessName = (Get-Process -Id $split[4]).Name
                                                                               ProcessID = [int]($split[4])}
                        }
                        if ($split[0] -eq "UDP"){
                            $obj = new-object -typename pscustomobject -Property @{Proto = $split[0]
                                                                               LocalAddress = [IPAddress]($split[1].Substring(0,$split[1].lastindexof(":")))
                                                                               LocalPort = [int]($split[1].split(":")[-1])
                                                                               RemoteAddress = [IPAddress]$IP = "0.0.0.0"
                                                                               RemotePort = 0
                                                                               State = "LISTENING"
                                                                               ProcessName = (Get-Process -Id $Split[3]).Name
                                                                               ProcessID = [int]($split[4])}
                        }
                        #$Filter = {$_.ProcessName -ne "System" -and $_.ProcessName -ne "svchost" -and $_.ProcessName -ne "RouterNT" -and $_.ProcessName -ne "wininit" -and $_.ProcessName -ne "lsass"}   
                        Write-Output $obj | Where-Object -FilterScript {$_.State -like $Using:State -and $_.LocalAddress.AddressFamily -like $Using:AddressFamily }
                    }  
                }
            }
        }
    }
    function ConvertTo-ACLMapping {
        <#
    .SYNOPSIS
        Short description
    .DESCRIPTION
        Long description
    .EXAMPLE
        Example of how to use this cmdlet
    .EXAMPLE
        Another example of how to use this cmdlet
    .INPUTS
        Inputs to this cmdlet (if any)
    .OUTPUTS
        Output from this cmdlet (if any)
    .NOTES
        General notes
    .COMPONENT
        The component this cmdlet belongs to
    .ROLE
        The role this cmdlet belongs to
    .FUNCTIONALITY
        The functionality that best describes this cmdlet
    #>
        [CmdletBinding(DefaultParameterSetName='Parameter Set 1',
                       SupportsShouldProcess=$true,
                       PositionalBinding=$false,
                       HelpUri = 'http://www.microsoft.com/',
                       ConfirmImpact='Medium')]
        [Alias()]
        [OutputType([String])]
        Param (
            # Param1 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [IPAddress[]]$LocalAddress,
            # Param2 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [Int[]]$LocalPort,
            #Param4 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [IPAddress[]]$RemoteAddress,
            # Param5 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [Int[]]$RemotePort,
            # Param5 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [String[]]$Proto,
            # Param6 help description
            [Parameter(Mandatory=$true,
                       ValueFromPipeline=$true,
                       ValueFromPipelineByPropertyName=$true)]
            [Alias("ProcessName")]
            [String[]]$Name,
            # Param7 help description
            [Parameter(Mandatory=$false)]
            [String]$Direction = "Inbound"
        )
        
        begin {
    
        }
        
        process {
            for ($i = 0; $i -lt $LocalPort.Count; $i++) {
    
                Write-Output "netsh advfirewall firewall add rule name=`"$($Name[$i])`" dir=in action=allow protocol=$($Proto[$i]) localport=$($LocalPort[$i])"
            }
        }
        
        end {
        }
    }