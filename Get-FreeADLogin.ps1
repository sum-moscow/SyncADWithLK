﻿<#
    .DESCRIPTION
       Script get free AD username from first, last and mibble name

    .AUTHOR    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
    .COMPANY   State University of Management aka GUU

    .LINK
       http://...
#>

# Check only UPN
function Get-FreeADLogin{
     Param (        [parameter(Mandatory = $true)]        [String]$FirstName,        [parameter(Mandatory = $true)]        [String]$LastName,        [parameter(Mandatory = $false)]        [String]$MiddleName = $Null,

        [parameter(Mandatory = $true)]
        [String]$Domain,

        [parameter(Mandatory = $false)]
        [Switch]$NoCkeckSAM = $Null,

        [parameter(Mandatory = $false)]
        [string] $UPNPrefix = $Null,

        [parameter(Mandatory = $false)]
        [string] $UPNPostfix = $Null,

        [parameter(Mandatory = $false)]
        $IgnoreLogins = $Null
    )

    $trFirst = (Get-Translit $FirstName.Trim()).toLower()[0]
    $trMiddle = (Get-Translit $MiddleName.Trim()).toLower()[0]
    $trLast = Get-Translit $LastName.Trim().toLower()

    $SAMs = (
        "$($trFirst)$($trMiddle)_$($trLast)",
        "$($trLast)_$($trFirst)$($trMiddle)",
        "$($trFirst)$($trMiddle).$($trLast)",
        "$($trLast).$($trFirst)$($trMiddle)",
        "$($trFirst)_$($trLast)",
        "$($trLast)_$($trFirst)"
    )


    foreach ($SAM in $SAMs){
        $SAM = ($SAM.Replace(" ","_") -replace "^\W+" ) -replace "^_"

        $UPN = "$UPNPrefix$SAM$UPNPostfix@$Domain"
        
        # SAM is not used by users, but need for system
        if($SAM.Length -gt 20){
            $SAM = $SAM.Substring(0, 20)
        }

        if($NoCkeckSAM){
            # always return 0 records
            $SAM = $false
        }

        

        if ( (Get-ADUser -Filter {UserPrincipalName -eq $UPN -Or 
                                  SamAccountName -eq $SAM -or
                                  ProxyAddresses -eq "SMTP:$UPN" -or
                                  ProxyAddresses -eq "smtp:$UPN" }
           ).UserPrincipalName.count -eq 0 ) {

            if ($IgnoreLogins){
                if (($IgnoreLogins.SAM.IndexOf($SAM) -ge 0) -Or
                    ($IgnoreLogins.UPN.IndexOf($UPN) -ge 0)){
                    continue
                }    
            } 

            Return @{"UPN"=$UPN;"SAM" =$SAM}
        }
    }
    Log-Critical("Can't create UPN")
    Return $false
}