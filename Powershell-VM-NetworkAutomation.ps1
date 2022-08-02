# Import active directory module for running AD cmdlets
Import-Module activedirectory

$MainMenu = {
Write-Host " ***************************"
Write-Host " *           Menu          *" 
Write-Host " ***************************" 
Write-Host 
Write-Host " 1.) Import OU's from CVS file" 
Write-Host " 2.) Populate OU's from CVS file"
Write-Host " 3.) Apply a GPO to an OU" 
Write-Host " 4.) VM Menu" 
Write-Host " 5.) Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}  

$VMMenu = {
Write-Host " ***************************"
Write-Host " *      Hyper-V Menu       *" 
Write-Host " ***************************" 
Write-Host 
Write-Host " 1.) Export virtual machines" 
Write-Host " 2.) Delete virtual machines"
Write-Host " 3.) Start virtual machines" 
Write-Host " 4.) Stop virtual machines" 
Write-Host " 5.) Create a virtual machine"
Write-Host " 6.) Create a VM switch"
Write-Host " 7.) Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}
$TypeMenu = {
Write-Host " ***************************"
Write-Host " *       Switch Type       *" 
Write-Host " ***************************" 
Write-Host 
Write-Host " 1.) Internal" 
Write-Host " 2.) Private"
Write-Host " 3.) External" 
Write-Host " 4.) Quit"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}
function OU {
try {
        $CSVPATH = Read-Host "List CSV file location"
        #$CSVPATH = "C:\Users\Administrator\Desktop\Scripts\Project\project.csv"
        if (Test-Path $CSVPATH -PathType leaf) {
        Write-Host
        Write-Host "File does exist"
        Write-Host
        }
        else {
        Write-Host
        Write-Host "File does not exist"
        Write-Host
        }
        $csvUsers = Import-Csv -Path $CSVPATH
        $ousToCreate = $csvUsers.Department | Select-Object -Unique
        $ousToCreate.foreach({
            if (Get-AdOrganizationalUnit -Filter "Name -eq '$_'") {
                Write-Host "The OU with the name [$_] already exists."
            } 
            else {
                New-AdOrganizationalUnit -Name $_
            }
            })
            Get-ADOrganizationalUnit -Filter 'Name -like "*"' | Format-Table Name, DistinguishedName -A
}
catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }
}
function User {
try {
        $CSVPATH = Read-Host "List CSV file location"
        #$CSVPATH = "C:\Users\Administrator\Desktop\Scripts\Project\project.csv"
        if (Test-Path $CSVPATH -PathType leaf) {
        Write-Host
        Write-Host "File does exist" -ForegroundColor Green
        Write-Host
        }
        else {
        Write-Host
        Write-Host "File does not exist"
        Write-Host
        }
        $csvUsers = Import-Csv -Path $CSVPATH
        $ousToCreate = $csvUsers.Department | Select-Object -Unique
        foreach ($csvUser in $csvUsers) {
        $proposedUsername = '{0}{1}' -f $csvUser.'First', $csvUser.'Last'.Substring(0, 1)
        $DisplayName = '{0}{1}' -f $csvUser.'First', $csvUser.'Last'
        $PlainPassword = "P@ssw0rd"
        $SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
        ## Check to see if the proposed username exists
        if (Get-AdUser -Filter "Name -eq '$proposedUsername'") {
        Write-Host "The AD user [$proposedUsername] already exists."
        } 
        else {
        $newUserParams = @{
            
            Name        = $proposedUsername
            SamAccountName = $proposedUsername
            Path        = "OU=$($csvUser.Department),DC=prgm,DC=local"
            Enabled     = $true
            GivenName   = $csvUser.'First'
            Surname     = $csvUser.'Last'
            OfficePhone = $csvUser.'Telephone'
            Department = $csvUser.'Department'
            AccountPassword = (ConvertTo-SecureString “P@ssw0rd” –AsPlainText –Force)
            DisplayName = $DisplayName
            ChangePasswordAtLogon =$true
                           }
            New-AdUser @newUserParams
            }
        }
       }
catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }

}
function GPO {
    $Domain = Get-ADDomain -Current LoggedOnUser
    $Domain = $Domain | Format-Table -Property DNSRoot -HideTableHeaders
    $Domain = Out-String -InputObject $Domain
    $Domain = $Domain.Trim()
    $GPO = Get-GPO -All -Domain $Domain | Out-GridView -Title "Select GPO" -PassThr
    $GPO = $GPO | Format-Table -Property DisplayName -HideTableHeaders
    $GPO = Out-String -InputObject $GPO
    $GPO = $GPO.Trim()
    $OU = Get-ADOrganizationalUnit -Filter 'Name -like "*"' | Out-GridView -Title "Select OU to apply GPO to" -PassThru
    $OU = $OU | Format-Table -Property DistinguishedName -HideTableHeaders
    $OU = Out-String -InputObject $OU
    $OU = $OU.Trim()
    New-GPLink -name $GPO -Target $OU -Enforced Yes
    }
function MakeVM {
try{
    $FolderNameVM = "C:\VM"
    $FolderNameVHD = "C:\VHD"
    if (Test-Path $FolderNameVM) {
    Write-Host 
    }
    else {
  
    #PowerShell Create directory if not exists
    New-Item $FolderName -ItemType Directory
    Write-Host "VM Folder Created successfully"
    }
    if (Test-Path $FolderNameVHD) {
    Write-Host 
    }
    else {
    #Create directory if not exists
    New-Item $FolderName -ItemType Directory
    Write-Host "VHD Folder Created successfully"
    }     
    $NAME = Read-Host "What is the VM's name?"
    $MEM = Read-Host "What is the RAM (in GB)?"
    $MEM64 = [int64]$MEM.Replace('GB','') * 1GB
    $VHD = Read-Host "What is the hard drive size (in GB)"
    $VHD64 = [int64]$VHD.Replace('GB','') * 1GB
    $GEN = Read-Host "What is the generation? (1 or 2)"
    $VHDPATH = Get-ChildItem -Path C:\ -Name | Out-GridView -Title "Select location for VHD" -PassThru
    $VHDPATH = "C:\" + $VHDPATH + "\" + "$NAME" + ".vhdx" 
    $VMPATH = Get-ChildItem -Path C:\ -Name | Out-GridView -Title "Select location for VM" -PassThru
    $PATHTEMP = "C:\" + $VMPATH
    New-Item -Path $PATHTEMP -Name $NAME -ItemType Directory | Out-Null
    $VMPATH = "C:\" + $VMPATH + "\" + $NAME
    $SWITCH = Get-VMSwitch -SwitchType External | Out-GridView -Title "Select network switch" -PassThru
    $SWITCH = $SWITCH | Format-Table -Property Name -HideTableHeaders
    $SWITCH = Out-String -InputObject $SWITCH
    $SWITCH = $SWITCH.Trim()
    $ISO = Get-ChildItem -Path C:\ISO -Name | Out-GridView -Title "Select the ISO file" -PassThru
    $ISO = "C:\ISO\"+$ISO
    New-VM -Name $NAME -MemoryStartupBytes $MEM64 -Path $VMPATH -Switch $SWITCH -Generation $GEN 
    New-VHD -Path $VHDPATH -SizeBytes $VHD64 -Dynamic 
    Add-VMHardDiskDrive -VMName $NAME -path $VHDPATH
    Set-VMDvdDrive -VMName $NAME -ControllerNumber 1 -Path $ISO
    }
    catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }
}
function ExportVM {
try {
     $FilePath = Get-ChildItem -Path C:\ -Name | Out-GridView -Title "Select Path" -PassThru
     Write-Host " Select virtual machines to Export."
     Get-VM | Out-GridView -Title "Select virtual machines to export" -PassThru | Export-VM -Path $FilePath
     }
catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }
}
function DeleteVM {
try {
     Write-Host
     Write-Host " Select virtual machines to to deleted."
     Get-VM | Out-GridView -Title "Select virtual machines to be deleted" -PassThru | Remove-VM -Confirm
    }
catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }
}
function VMStart {
try {
       Write-Host
       Get-VM | Out-GridView -Title "Select virtual machines to be started"  -PassThru | Start-VM
     }
catch {      
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }      
}
function VMStop { 
try {
       Write-Host
       Get-VM | Out-GridView -Title "Select virtual machines to be stopped"  -PassThru | Stop-VM
    }
catch {
     Write-Host
     Write-Host $Error[0] -ForegroundColor Red
     Write-Host
     }       
}

function VMSwitch ($Switch)
{ 
      Write-Host
       $Name = Read-Host "What is the switches name?"
        Do { 
        Invoke-Command $TypeMenu
        $Select = Read-Host
        Switch ($Select)
        {
        1 {
           New-VMSwitch -name $Name -SwitchType Internal
           $vmswitch = Get-VMSwitch
           write-output $vmswitch
           exit
           }
        2 {
           New-VMSwitch -name $Name -SwitchType Private
           $vmswitch = Get-VMSwitch
           Write-output $vmswitch
           exit
           }
        3 {
           Do
           {
           $c = Read-Host "Allow management OS to share this network adapter? (Yes/No)"
           }
           Until (($c -eq "yes") -or ($c -eq "y") -or ($c -eq "no") -or ($c -eq "n")) 
           $OS = $c
           if (($OS -eq "yes") -or ($OS -eq "y")) { $OS = $true}
           if (($OS -eq "no") -or ($OS -eq "n")) { $OS = $false}
           $Interface = Get-NetAdapter | select Name, InterfaceDescription | Out-GridView -Title "Select interface to be used as a switch" -PassThru
           $Interface = $Interface | Format-Table -Property Name -HideTableHeaders
           $Interface = Out-String -InputObject $Interface
           $Interface = $Interface.Trim() 
           New-VMSwitch -name $Name  -NetAdapterName $Interface -AllowManagementOS $OS
           $vmswitch = Get-VMSwitch
           write-output $vmswitch
           exit
           }

    }
    } 
    Until ($Select -eq 4)
}

    


Do { 
Invoke-Command $MainMenu
$Select = Read-Host
Switch ($Select)
    {
    1 {
       OU
       }
    2 {
       User
       }
    3 {
       GPO
       }
    4 {
       Do { 
            Invoke-Command $VMMenu
            $Select = Read-Host
            Switch ($Select)
                {
                1 {
                   ExportVM 
                  }
                2 {
                   DeleteVM
                  }
                3 {
                   VMStart
                  }
                4 {
                   VMStop
                  }
                5 {
                   MakeVM
                   }
                6 {
                   VMSwitch
                   Get-VMSwitch
                   }
                }
            }
            While ($Select -ne 7)
         }
       }
     }
While ($Select -ne 5)
