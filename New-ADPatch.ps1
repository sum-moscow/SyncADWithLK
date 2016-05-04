<#
    .DESCRIPTION
       Script for update (create) AD user

    .AUTHOR    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
    .COMPANY   State University of Management aka GUU

    .LINK
       http://...
#>

[System.Collections.ArrayList]$global:newLogins = @()
function New-ADPatch {
    param (
        [parameter(Mandatory = $true)]
        [string] $FirstName,

        [parameter(Mandatory = $true)]
        [string] $LastName,

        [parameter(Mandatory = $false)]
        [string] $MiddleName = $Null,

        [parameter(Mandatory = $false)]
        [string] $Name = $Null,

        [parameter(Mandatory = $false)]
        [string] $NamePrefix = $Null,

        [parameter(Mandatory = $false)]
        [string] $NamePostfix = $Null,

        [parameter(Mandatory = $true)]
        [string] $Domain = $Null,
                
        [parameter(Mandatory = $false)]
        [array] $OUs = $Null,
        
        [parameter(Mandatory = $false)]
        [string] $Title = $Null,

        [parameter(Mandatory = $false)]
        [string] $Department = $Null,
                
        [parameter(Mandatory = $false)]
        [string] $Company = $Null,
                
        [parameter(Mandatory = $false)]
        [string] $Avatar = $Null,

        [parameter(Mandatory = $true)]
        [string] $EmployeeNumber = $Null,

        [parameter(Mandatory = $true)]
        [string] $EmployeeID,
                
        [parameter(Mandatory = $false)]
        [string] $SAM = $Null,

        [parameter(Mandatory = $true)]
        [boolean] $Enabled = $Null,

        [parameter(Mandatory = $false)]
        [string] $Country = $Null, 

        [parameter(Mandatory = $false)]
        [string] $PasswordPath = $Null,

        [parameter(Mandatory = $false)]
        [string] $UPNPrefix = $Null,

        [parameter(Mandatory = $false)]
        [string] $UPNPostfix = $Null
    )

    if (-not $Enabled){
        #no add new disabled user
        $ADUser = Get-ADUser -Filter { EmployeeNumber -eq $EmployeeNumber }
        if (!$ADUser){
            return 
        } 
    } 

    if (!$Name){
        $Name = $LastName + " " + $FirstName

        if ($MiddleName){
            $Name += " " + $MiddleName
        }
        
        $Name = $NamePrefix + $Name + $NamePostfix
    }

    if($Name.Length -ge 64){
        $Name = $Name.Substring(0,63) + "…";
    }

    if($Title.Length -ge 64){
        $Title = $Title.Substring(0,63) + "…";
    }
    
    if (!$Department){
        $Department = $OUs[$OUs.Count-1].name
    }

    if($Department.Length -ge 64){
        $Department = $Department.Substring(0,63) + "…";
    }

    $Initials = $FirstName[0] + "."
    if ($MiddleName){
        #non-breaking space
        $Initials += ' ' + $MiddleName[0] + "."
    }

    [System.Collections.ArrayList]$SetToUser = @{
        DisplayName = $Name
        GivenName = $FirstName
        Surname = $LastName
        Enabled = $Enabled
        Initials = $Initials
        EmployeeID = $EmployeeID
    }
     

    [void]$SetToUser.Add(@{Key="Department";Value=$Department})

    $Path = Get-Path -OUs $OUs

    $DistinguishedName = "CN=" + $Name + "," + $Path
        
    [void]$SetToUser.Add(@{Key="DistinguishedName";Value=$DistinguishedName})
    [void]$SetToUser.Add(@{Key="Title";Value=$Title})
    [void]$SetToUser.Add(@{Key="Company";Value=$Company})
    [void]$SetToUser.Add(@{Key="PasswordNeverExpires";Value=$true})
    [void]$SetToUser.Add(@{Key="Country";Value=$Country})

    $ADUser = Get-ADUser -Filter { EmployeeNumber -eq $EmployeeNumber } -Properties *
    if (!$ADUser){
        $Login = Get-FreeADLogin -FirstName $FirstName -LastName $LastName -MiddleName $MiddleName -Domain $Domain -UPNPrefix $UPNPrefix -UPNPostfix $UPNPostfix -IgnoreLogins $global:newLogins
        if ($SAM){
            $Login.SAM = $SAM
        }
        Log("Create user $Id with UPN: $($Login.UPN), SAM: $($Login.SAM)")

        $global:newLogins.Add($Login)

        Patch("New-ADUser -Name '$Name' -UserPrincipalName $($Login.UPN) -EmailAddress $($Login.UPN) -SamAccountName $($Login.SAM) -EmployeeNumber '$Id' -ChangePasswordAtLogon `$false -Path '$Path'")
        # only for "Enabled"
        $Password = New-RandomPassword
        [void]$SetToUser.Add(@{Key="Password";Value=$Password})

        Patch("Set-ADAccountPassword -Identity $($Login.SAM)  -NewPassword (ConvertTo-SecureString -AsPlainText '$Password' -Force)")
        $ADUser = @{Name=$SetToUser.Name; `
                    SamAccountName=$Login.SAM; ` 
                    DistinguishedName=$DistinguishedName; `
                    UserPrincipalName=$Login.UPN
                   }
        if ($PasswordPath){
            $BackUpDataInCsv = New-Object psobject

            $SetToUser.GetEnumerator() | % { 
                Add-Member -InputObject $BackUpDataInCsv -MemberType noteproperty `
                        -Name $_.Key -Value $_.Value
            }

            If ((Get-Content $PasswordPath) -eq $Null) {
                $BackUpDataInCsv | Export-Csv -NoTypeInformation -Encoding UTF8 $C.ad.passwordPath 
            } else {
                $BackUpDataInCsv | Export-Csv -NoTypeInformation -Encoding UTF8 -Append $C.ad.passwordPath 
            }
    
        }
    }
    
    [void]$SetToUser.Add(@{Key="proxyAddresses";Value=$ADUser.UserPrincipalName})
    [void]$SetToUser.Add(@{Key="EmailAddress";Value=$ADUser.UserPrincipalName})
 
      
    $SetToUser.GetEnumerator() | % { 
        #log("$($_.Key) -Value $($_.Value)")
        #this is bad (need fix)
        if ($_.Key -eq "Password"){
            continue
        }

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
        $OU.name = $OU.name.Replace("""","").Replace("'","")
        #maximum OU Length = 64
        if($OU.name.Length -ge 64 ){
            $OU.name = $OU.name.Substring(0,63) + "…";
        }

        $LastPath = $Path
        $Path = "OU=" + $OU.name.Replace(",","\,") + "," + $Path

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

    # custom check value
    switch ($Property){
        "proxyAddresses" {
            if (-not ($ADUser.$Property -contains "SMTP:$Value")){
                Patch("Set-ADUser $($ADUser.SamAccountName) -Add @{ProxyAddresses='SMTP:$Value'}")  
            }    
            return
        } 
    }

    if ($ADUser.$Property -ne $Value){
        switch ($Property){
            "DistinguishedName" {
                # check change CN
                $Name = ($Value -split ',OU=')[0] -replace "^CN="
                $NameCurrent = ($ADUser.$Property -split ',OU=')[0] -replace "^CN="
                if ($Name -ne $NameCurrent){
                    Patch("Rename-ADObject -Identity '$($ADUser.ObjectGUID)' -NewName '$Name'")
                }

                # check change OU
                if ($ADUser.doNotMoveADObjectByScript){
                    return
                }

                $Path = $Value -replace 'CN=[^,]*,' -replace "'","''"
                $PathCurrent = $ADUser.$Property -replace 'CN=[^,]*,' -replace "'","''"
                if ($Path -ne $PathCurrent){
                    Patch("Move-ADObject -Identity '$($ADUser.ObjectGUID)' -TargetPath '$Path'")
                } 
            }
            "Enabled" {
                if ($Value) {
                    Patch("Enable-ADAccount -Identity $($ADUser.SamAccountName)")
                } else {
                    if (-Not ($ADUser.doNotDisableByScript)){
                        Patch("Disable-ADAccount -Identity $($ADUser.SamAccountName)")
                    }
                }
            }
            default { 
                switch($Value.GetType().Name){
                    "String" {
                        $Value = $Value -replace "'","''"
                        $Value = "'" + $Value + "'"
                     }
                    "Boolean" {$Value = '$' + $Value}
                }
                       
                Patch("Set-ADUser $($ADUser.SamAccountName) -$Property $Value # Previous: $($ADUser.$Property)")
            }
        }  
    }
}