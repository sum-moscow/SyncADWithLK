<#
    .DESCRIPTION
       Script for update (create) AD user

    .AUTHOR    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
    .COMPANY   State University of Management aka GUU

    .LINK
       http://...
#>

function New-ADPatch {
    param (
        [parameter(Mandatory = $true)]
        [string] $FirstName,

        [parameter(Mandatory = $true)]
        [string] $LastName,

        [parameter(Mandatory = $false)]
        [string] $MiddleName,

        [parameter(Mandatory = $true)]
        [string] $Domain,
                
        [parameter(Mandatory = $false)]
        [array] $OUs,
        
        [parameter(Mandatory = $false)]
        [string] $Title,

        [parameter(Mandatory = $false)]
        [string] $Department,
                
        [parameter(Mandatory = $false)]
        [string] $Company,
                
        [parameter(Mandatory = $false)]
        [string] $Avatar,

        [parameter(Mandatory = $true)]
        [string] $EmployeeNumber,

        [parameter(Mandatory = $true)]
        [string] $EmployeeID,
                
        [parameter(Mandatory = $false)]
        [string] $SAM,

        [parameter(Mandatory = $true)]
        [boolean] $Enabled,

        [parameter(Mandatory = $false)]
        [string] $Country, 

        [parameter(Mandatory = $false)]
        [string] $PasswordPath,

        [parameter(Mandatory = $false)]
        [string] $UPNPrefix,

        [parameter(Mandatory = $false)]
        [string] $UPNPostfix
    )

    $Name = $LastName + " " + $FirstName
    $Initials = $FirstName[0] + "."
    if ($MiddleName){
        $Name += " " + $MiddleName
        #non-breaking space
        $Initials += ' ' + $MiddleName[0] + "."
    }

    [System.Collections.ArrayList]$SetToUser = @{
        Name = $Name
        DisplayName = $Name
        GivenName = $FirstName
        Surname = $LastName
        Enabled = $Enabled
        Initials = $Initials
        EmployeeID = $EmployeeID
    }
     
    if ($Enabled){
        if (!$Department){
            $Department = $OUs[$OUs.Count-1].name
        }
        if($Department.Length -ge 64){
            $Department = $Department.Substring(0,63) + "…";
        }
        [void]$SetToUser.Add(@{Key="Department";Value=$Department})

        $Path = Get-Path -OUs $OUs

        if ($SAM){
            $DistinguishedName = "CN=" + $SAM + "," + $Path
        } else {
            $DistinguishedName = "CN=" + $Name + "," + $Path
        }
        
        [void]$SetToUser.Add(@{Key="DistinguishedName";Value=$DistinguishedName})
        [void]$SetToUser.Add(@{Key="Title";Value=$Title})
        [void]$SetToUser.Add(@{Key="Company";Value=$Company})
        [void]$SetToUser.Add(@{Key="PasswordNeverExpires";Value=$true})
        [void]$SetToUser.Add(@{Key="Country";Value=$Country})

    } else {
        #no add new disabled user
        $ADUser = Get-ADUser -Filter { EmployeeNumber -eq $EmployeeNumber } 
        if (!$ADUser){
            return 
        }
    } 

    $ADUser = Get-ADUser -Filter { EmployeeNumber -eq $EmployeeNumber } -Properties *
    if (!$ADUser){
        $Login = Get-FreeADLogin -FirstName $FirstName -LastName $LastName -MiddleName $MiddleName -Domain $Domain -UPNPrefix $UPNPrefix -UPNPostfix $UPNPostfix
        if ($SAM){
            $Login.SAM = $SAM
        }
        Log("Create user $Id with UPN: $($Login.UPN), SAM: $($Login.SAM)")
        Patch("New-ADUser -Name '$Name' -UserPrincipalName $($Login.UPN) -EmailAddress $($Login.UPN) -SamAccountName $($Login.SAM) -EmployeeNumber '$Id' -ChangePasswordAtLogon `$false -Path '$Path'")
        # only for "Enabled"
        $Password = New-RandomPassword
        Patch("Set-ADAccountPassword -Identity $($Login.SAM)  -NewPassword (ConvertTo-SecureString -AsPlainText '$Password' -Force)")
        if ($PasswordPath){
            echo "$($Login.UPN),$Password" >> $PasswordPath 
        }
        $ADUser = @{Name=$SetToUser.Name;SamAccountName=$Login.SAM;DistinguishedName=$DistinguishedName}
    }

    $SetToUser.GetEnumerator() | % { 
        #echo "$($_.Key) -Value $($_.Value)"
        Check-ADUserAttr -ADUser $ADUser -Property $_.Key -Value $_.Value
    }
    
   
}

function Get-Path{
    param (
        [parameter(Mandatory = $true)]
        $OUs
    )

    $Path = (Get-ADDomain).DistinguishedName;
    $NeedCreate = $false
    foreach($OU in $OUs){
        #maximum OU Length = 64
        if($OU.name.Length -ge 64 ){
            $OU.name = $OU.name.Substring(0,63) + "…";
        }

        $LastPath = $Path
        $Path = "OU=" + $OU.name.Replace("""","\""").Replace(",","\,") + "," + $Path

        try {
            [void](Get-ADOrganizationalUnit -Identity $Path) 
        } catch {
            Patch("New-ADOrganizationalUnit -Name '$($OU.name)' -Path '$LastPath'")
        }
    }

    return $Path
}

function Check-ADUserAttr{
    param (
        [parameter(Mandatory = $true)]
        $ADUser,

        [parameter(Mandatory = $true)]
        [string] $Property,

        [parameter(Mandatory = $true)]
        $Value
    )  

    if($ADUser.$Property -ne $Value){


        switch ($Property){
            "DistinguishedName" {
                $Value = $Value -replace 'CN=[^,]*,'
                $Value = "$Value"
                Patch("Move-ADObject -Identity '$($ADUser.DistinguishedName)' -TargetPath '$Value'")
            }
            "Name" {
                
            }
            "Enabled" {
                if ($Value) {
                    Patch("Enable-ADAccount -Identity $($ADUser.SamAccountName)")
                } else {
                    Patch("Disable-ADAccount -Identity $($ADUser.SamAccountName)")  
                }
            }
            default { 
                switch($Value.GetType().Name){
                    "String" {$Value = "'" + $Value + "'"}
                    "Boolean" {$Value = '$' + $Value}
                }
                       
                Patch("Set-ADUser $($ADUser.SamAccountName) -$Property $Value")
            }
        }  
    }
}