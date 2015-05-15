<#
    .DESCRIPTION
       This script remove all empty OU from SearchRoot

    .AUTHOR    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
    .COMPANY   State University of Management aka GUU

    .LINK
       http://...
#>

$Searcher = New-Object System.DirectoryServices.DirectorySearcher -Property @{
    Filter = '(objectclass=organizationalunit)'
    PageSize = 500
    SearchRoot = "LDAP://OU=Staff,DC=ad,DC=guu,DC=ru" 
}

$EmptyOUs = ($Searcher.FindAll() | Where-Object {(([adsi]$_.Path).psbase.children | Measure-Object).Count -eq 0}).Properties.distinguishedname

foreach($EmptyOU in $EmptyOUs){
    echo $EmptyOU
    Set-ADOrganizationalUnit -Identity $EmptyOU -ProtectedFromAccidentalDeletion $false
    Remove-ADOrganizationalUnit -Identity $EmptyOU -Confirm:$false
}

