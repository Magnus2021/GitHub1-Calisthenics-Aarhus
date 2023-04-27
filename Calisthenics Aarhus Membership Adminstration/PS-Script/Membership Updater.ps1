#Data Import
#################################################################################################################################################
#$relativePath = "Membership Official.xlsx" # Define the path to the file relative to the script directory
#$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path # Use the $PSScriptRoot variable to get the full path to the script directory
#$Excelpath = Join-Path -Path $scriptDir -ChildPath $relativePath
$Excelpath = "C:\Users\Magnus\OneDrive - Aarhus universitet\Skrivebord\Calisthenics Aarhus Membership Adminstration\Data\Memberlist For PS-script.xlsx"
$Table = Import-XLSX -path $Excelpath #Reads the content of the membership table

#Email service down below
#################################################################################################################################################
$From = "Kontakt@calisthenicsaarhus.dk" #Email to send from
$Subject = "Membership Update" #Subject of Email
$SMTPServer = "send.one.com" #Email Service
$SMTPPort = "587" #Port for email service
$SMTPUsername = "kontakt@calisthenicsaarhus.dk" #Login Options for the email service
$SMTPPassword = "Caliaarhus2021" #Login Options for the email service

$year = (get-date).Year #Gets the current year
$month = (Get-Date).Month #Gets the current month
$day = 1 #Gets the current day
$MembershipStart = Get-Date -Day $day -Month $month -Year $year #Puts the above variable into one variable. This describes the start date of the membership
$MembershipFinish = $MembershipStart.AddMonths(3) #Describes the end date of the membership 
$MembershipDelete = $MembershipFinish.AddDays(14) #Descrobes the deletion date for their data
$MembershipStartString = $MembershipStart.ToString("dd/MM/yyyy") #converts into short readable string
$MembershipFinishString = $MembershipFinish.ToString("dd/MM/yyyy") #converts into short readable string
$MembershipDeleteString = $MembershipDelete.ToString("dd/MM/yyyy") #converts into short readable string
$threeMonths = New-TimeSpan -Days 90  #used as logic for the if statement in the forloop below
$threeMonthsAndFifteenDays = New-TimeSpan -Days 105 #used as logic for the if statement in the forloop below

foreach ($Member in $table) #Loop for the table 
{       
$date = Get-Date $member.Date #Gets current membership signup date
$age = (Get-Date) - $date #calculates age of membership

write-host "age is $Age"

if ($age -gt $threeMonths -and $age -lt $threeMonthsAndFifteenDays) #Checks if membership is between 3 months and 3 month and 15 days. if true, sends email to member.
    {
    $To = $Member.Email #Members mail
    $Body ="Dear $($Member.Navn),

Your membership has expired, and you are kindly asked to transfer 450 DKK to the following account:

Registration number : 2877
Account number      : 4399493582
    
This will extend your membership from $($MembershipStartString) to $($MembershipFinishString).

Make sure you include your full name in the message when completing the transfer, so we can register your payment.
    
If you do not wish to continue your membership, please let us know, or refrain from paying, and we will delete your data on $($MembershipDeleteString).
    
Best regards,
The Board of Calisthenics Aarhus." #The email context
        Write-Host "$(Get-Date -Format 'HH:mm:ss'): Sending E-mail to $($member.Email)" #Informs script user
        #MUST STAY COMMENTED OUT IF TESTING
        #Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Credential (New-Object System.Management.Automation.PSCredential($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force))) -UseSsl -WarningAction Ignore
    }

    else #If membership is not older than 3 month is goes into this else statement, which is effectively emptive.
    {
        Write-Host "$(Get-Date -Format 'HH:mm:ss'): Membership for $($member.Navn) is still active. Did not send E-mail" #Informs script user   
    }
}
#################################################################################################################################################
