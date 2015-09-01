<#
.DESCRIPTION
   Sync AD users with LK (json)
.AUTHOR    Zubarev Alexander aka Strike (av_zubarev@guu.ru)
   
.LINK
   http://...
#>

# Import modules and config ($C)
$LocalDir = $MyInvocation.MyCommand.Definition | split-path -parent
. $LocalDir\Import-AllRun.ps1

# Setup dashing
Log-Set -Service "LKAD" -Url $C.dashing.url -Token $C.dashing.token 
Log-Begin

# Helper function
[System.Collections.ArrayList]$global:FullPatch = @()
function Patch{
   if (!$global:FullPatch.Contains($args[0])){
       [void]$global:FullPatch.Add($args[0])
   }
}

<#############################>Log("Downloading users")<##########################################>
foreach($Domain in $C.server.domains.ChildNodes){
    if ($Domain.env -eq "PROD") {
        break;
    }
    $Domain = $null
}

if (!$Domain){
    Log-Critical("Can't find domain in Config.xml")
    exit
}

$time = "1";
$FullPath = "$($Domain.protocol)://$($Domain.fqdn)/$($C.server.path)"
$Resource = "$FullPath/changedSince?token=$($C.server.token)&time=$time&public_key=$($C.server.public_key)&since=1"
try {
    $RawText = (curl "$Resource").Content
} catch {
    Log-Critical("Can't download users: $($_.Exception.Message)") -End
    exit    
}

<#############################>Log("Parsing downloaded data")<####################################>
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")        
$Jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
$Jsonserial.MaxJsonLength = 100000000

try {
    $Json = $Jsonserial.DeserializeObject($RawText)
} catch {
    Log-Critical("Can't deserialize downloaded data: $($_.Exception.Message)") -End
    exit    
}

if ($Json.status){
    $Users = $Json.data
} else {
    Log-Critical("Can't get users (status is $($Json.status)") -End
    exit  
}

<#############################>Log("Fetching $($Users.Count) users")<#############################>
foreach ($User in $Users){
    Log("$($Users.IndexOf($User)+1) of $($Users.count)")

    # binding students or staff (need fix)
    $Id = $Null
    $Domain = $Null
    $SAM = $Null
    $Role = $Null
    $SAM = $Null
    $Department = $Null # = the deepest OU$
    [System.Collections.ArrayList]$OUs = $Null

    $UserTypeOU = New-Object 'System.Collections.Generic.Dictionary[String,String]'
    if ($User.is_staff){
        $Domain = "guu.ru"
        $UserTypeOU.Add("name",$C.ad.ou.staff)
        $Id = $User."1c_id"
        
        #Find user main role
        $PluralistOUs = $Null
        $PluralistRole = $Null

        $DisabledOUs = $Null
        $DisabledRole = $Null
        
        foreach ($Position in $User.staff){
            if ($Position.active){
                if ($Position.main_role) {
                    $OUs = $Position.OU
                    $Role = $Position.role
                } else {
                    $PluralistOUs = $Position.OU
                    $PluralistRole = $Position.role
                }
            } else {
                   $DisabledOUs = $Position.OU
                   $DisabledRole = $Position.role
            }
        }

        if (!$OUs){
            $OUs = $PluralistOUs
            $Role = $PluralistRole
        } 

        if (!$OUs){
            $OUs = $DisabledOUs
            $Role = $DisabledRole
        } 

        if (!$OUs){
            Log-Warning("User without OU: id=$($User.id)")
            continue            
        } 
    } elseif ($User.is_student){
        $Domain = "edu.guu.ru"
        $Role = "Студент"
        $Department = $User.institute
        $OUs = @()
        $SAM = "student_$($User.personal_file)"
        $UserTypeOU.Add("name",$C.ad.ou.students)
        $Id = $User.personal_file
    } else {
        Log-Warning("User type not found: id=$($User.id), is_staff=$($User.is_staff), is_student=$($User.is_student)")
        continue
    }

    if (!$Id){
        Log-Warning("User ID is NULL: id=$($User.id)")
        continue
    }

    
    $OUs.Reverse()
    [void]$OUs.Add($UserTypeOU)
    $OUs.Reverse()

    $Company = "ФГБОУ ВПО «Государственный университет управления»"
    try { 
        New-ADPatch -FirstName  $User.first_name -LastName $User.last_name -MiddleName $User.middle_name `
             -OUs $OUs -Title $Role -Company $Company -EmployeeNumber $Id -EmployeeID $user.id `
             -Avatar $User.avatar -Enabled $user.active -SAM $SAM -Department $Department `
             -Domain $Domain -Country RU
    } catch {
        Log-Warning("Can't sync user with ID: id=$($User.id): $($_.Exception.Message)")
    }
}


<#############################>Log("Saving patch")<###############################################>
$Date = Get-Date -Format $C.patch.filedataformat
$Filename = "$($C.patch.folder)\$Date.adpatch"
echo $global:FullPatch > $Filename
Log-Stop
