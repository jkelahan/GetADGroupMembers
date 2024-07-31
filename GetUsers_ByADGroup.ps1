Param 
( #Report type should match the [section] from the ini file.
   [Parameter(ValueFromPipelineByPropertyName)]
   [string] $GroupName
)
$ReportData = @"
Department|Individual|Account Name
"@
$Now = Get-Date -Format "MM.dd.yy_HHmm"; #ex: 08.10.23_1253 = 8/10/23 @ 12:53 

<###################################################################
#
# Written by Josh Kelahan for Denise Doner 10/18/23
#
# Gets members of specific AD groups. Can choose group by menu, or specify on command line.
# Creates "pretty" report
#
# Requires ImportExcel module. Run the following as admin to install:
# Install-Module ImportExcel -AllowClobber -Force -Scope AllUsers
#
####################################################################>
#                      Functions
####################################################################

function ShowMenu()
{
    Write-Host "            Choose Group Name              "
    Write-Host "###########################################"
    Write-Host "1: WUIT_EntApps_Autosys_Console_Op_Prod"
    Write-Host "2: WUIT_EntApps_Autosys_Scheduler_Prod"
    Write-Host "3: WUIT_EntApps_Autosys_Commander"
    Write-Host "4: WUIT_EntApps_Autosys_Console_Op_Test"
    Write-Host "5: WUIT_EntApps_Autosys_Scheduler_Test"
    Write-Host "6: WCC Console Op"
    Write-Host "7: WCC Scheduler"
    Write-Host "8: WCC Console Commander"
    Write-Host "0: To Exit"

    $Choice = "";
    $Selection = Read-Host "Choose a number to continue"
    switch($Selection)
    {
        0 {exit}
        1 {$Choice = "WUIT_EntApps_Autosys_Console_Op_Prod"}
        2 {$Choice = "WUIT_EntApps_Autosys_Scheduler_Prod"}
        3 {$Choice = "WUIT_EntApps_Autosys_Commander"}
        4 {$Choice = "WUIT_EntApps_Autosys_Console_Op_Test"}
        5 {$Choice = "WUIT_EntApps_Autosys_Scheduler_Test"}
        6 {$Choice = "WCC Console Op"}
        7 {$Choice = "WCC Scheduler"}
        8 {$Choice = "WCC Console Commander"}
        Default {
            Write-Host "Invalid Selction."
            exit
        }
    }
    
    return $Choice
}

function Prettyify($MembersList, $GroupName)
{
    $Filename = $GroupName+"_$Now";
    $ReportName = "$FileName.xlsx";
    $BufferFile = "Buff_$FileName.txt"

    $ReportData | Out-File $BufferFile -Append; #Column headers = Department | Name | Account

    foreach($Account in $Members)
    {
        #If I dont break it out, it prints full object notation. 
        $Dept = $Account.Department;
        $Name = $Account.Name;
        $User = $Account.SamAccountName
        
        Write-Host "Found Account: $Dept|$Name|$User";
        $Row = "$Dept|$Name|$User"
        $Row | Out-File $BufferFile -Append
    }

    #Cleanup
    $Buffer = Get-Content $BufferFile
    $BufferData = ConvertFrom-csv $Buffer -Delimiter '|'
    $BufferData | Export-Excel -path $ReportName -FreezeTopRow -BoldTopRow
    Remove-item $BufferFile;
    
}


########################################################################

if($GroupName)
{#If a group name was submitted on the command line
    $Filename = $GroupName+"_$Now.txt";
    try 
    {
        $Members = Get-ADGroupMember -Identity $GroupName -Recursive | Get-ADUser -Properties SamAccountName, Name;# | Out-File $Filename;
        Prettyify $Members $GroupName
    }
    catch 
    {
        Write-Host "Invalid group name. Verify spelling and try again. Please wrap the group name in quotes."
    }
}
else
{
    $GroupName = ShowMenu
    $Members = Get-ADGroupMember -Identity $GroupName -Recursive | Get-ADUser -Properties SamAccountName, Name, Department;
    Prettyify $Members $GroupName;
}

