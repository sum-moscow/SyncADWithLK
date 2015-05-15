<#
    .DESCRIPTION
       This script set licenses to Office 365

    .AUTHOR    David Brabant
                 from http://stackoverflow.com/a/10500327/2854758 

    .MODIFY    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
    .COMPANY   State University of Management aka GUU
   
    .LINK
       https://github.com/sum-moscow/AAD-sync
#>


Function Load-Module {
    param (
        [parameter(Mandatory = $true)]
        [string] $name
    )

    if (Get-Module -Name $name){ 
        return $true
    } else {
        $retVal = Get-Module -ListAvailable | where { $_.Name -eq $name }
        if ($retVal){
            try {
                Import-Module $name -ErrorAction SilentlyContinue
                return $true
            } catch { }
        } 
    }

    return $false
}