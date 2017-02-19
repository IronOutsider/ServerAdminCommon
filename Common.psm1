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
    [CmdletBinding()]
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
    [CmdletBinding()]
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
    [CmdletBinding()]
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
                $Computer = $env:COMPUTERNAME
                #Make ADSI Connection to local computer and begin enumerating Objects
                $ADSIComputer = [ADSI]("WinNT://$Computer,computer")
                $group = $ADSIComputer.psbase.children.find('Administrators',  'Group')
                $ADGroup = $group.psbase.invoke("members")  | ForEach {
                                                                      $_.GetType().InvokeMember("Name",  'GetProperty',  $null,  $_, $null)
                                                                      }
                Write-Output $ADGroup
                }
                    foreach($Group in $ADGroup) {
                            #Return as hash table and turn into PSObjects.
                            $properties = @{ComputerName = $Computer
                                            UserName = $Group}
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
    [CmdletBinding()]
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
    [CmdletBinding()]
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