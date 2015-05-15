<#
.DESCRIPTION
   Upload user logins to LK (json)
.AUTHOR   Zubarev Alexander aka Strike (av_zubarev@guu.ru)
   
.LINK
   http://...
#>


# Import modules and config ($C)
$LocalDir = $MyInvocation.MyCommand.Definition | split-path -parent
. $LocalDir\Import-AllRun.ps1

# Setup dashing
Log-Set -Service "CopyADLoginsToLK" -Url $C.dashing.url -Token $C.dashing.token 
Log-Begin

<#############################>Log("Downloading users")<##########################################>
$time = "1";
$Resource = "$($C.server.path)/loginIsNull?token=$($C.server.token)&time=$time&public_key=$($C.server.public_key)&since=1"
try {
    $RawText = (curl "$Resource").Content
} catch {
    Log-Critical("Can't download users: ") -End
    exit    
}

<#############################>Log("Parsing downloaded data")<####################################>
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")        
$Jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
$Jsonserial.MaxJsonLength = 100000000

try {
    $Json = $Jsonserial.DeserializeObject($RawText)
} catch {
    Log-Critical("Can't deserialize downloaded data") -End
    exit    
}

if ($Json.status){
    $Users = $Json.data
} else {
    Log-Critical("Can't get users (status is $($Json.status)") -End
    exit  
}

<#############################>Log("Fetching $($Users.Count) users")<####################################>
foreach ($User in $Users){
    Log("$($Users.IndexOf($User)+1) of $($Users.count)")

    $UserTypeOU = New-Object 'System.Collections.Generic.Dictionary[String,String]'
    if ($User.is_staff){
        $Id = $User."1c_id"
    } elseif ($User.is_student){
        $Id = $User.personal_file
    } else {
        Log-Warning("User type not found: id=$($User.id), is_staff=$($User.is_staff), is_student=$($User.is_student)")
        continue
    }

    if (!$Id){
        Log-Warning("!User ID is NULL: id=$($User.id)")
        continue
    }

    $U = Get-ADUser -Filter 'EmployeeNumber -eq $Id'

    if ($U.count -ne 1){
        Log-Warning("Problem with user: id=$($User.id)")
        continue   
    }

    $UPN = $U.UserPrincipalName

    $Body = "login=$UPN"       
    $Resource = "$($C.server.path)/$($User.id)?token=$($C.server.token)&public_key=$($C.server.public_key)&time=1"
    $RawText = Invoke-RestMethod -Method Put -Uri "$Resource" -Body $Body

}

Log-Stop
