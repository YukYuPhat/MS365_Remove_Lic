#John Schuster
#john.schuster1978@gmail.com
#MIT License
#no warranty or liability implied or given


#Install-Module -Name AzureAD   <---- must have this installed on your computer with admin creds
#must run this as user that has admin rights to Microsoft 365 Online

#loads the whole office list
$path = "$($env:USERPROFILE)\Downloads\Office365.csv"
$csv = Import-Csv -path $path

#this is the do not touch list.

$donotuse = @(

'receptionists@dmmta.com'

'ScheduleInformation@dmmta.com'

'dartsharepoint@ridedart.com'

'SVC_USR_DW_Hall_Ph@ridedart.com'

'mailroom@ridedart.com'

'training4@dmmta.com'

'SpiceWorks_SVC_Acct@dmmta.com'

'FRDispatch@ridedart.com'

'Sharepoint_DONOTReply@ridedart.com'

'SWBGTickets_SVC_Acct@dmmta.com'

'bcyclemailbox@dmmta.com'

'TM@dartoffice365.onmicrosoft.com'

'svc_acc_crmadmin@ridedart.com'

'Rideshare_SVC_Acct@dmmta.com'

'SWTickets_SVC_Acct@dmmta.com'

'Maint_iPad1@ridedart.com'

'PTDispatch@ridedart.com'

'itmailbox@dmmta.com'

'woodmanalarmsinbox@ridedart.com'

'callrecording@dmmta.com'

'Barracuda_Archive@dmmta.com'

'rshare@dmmta.com'

'Dial-A-Ride@dmmta.com'

'ptweekend@ridedart.com'

'SVC_USR_ComAreaPhone@dmmta.com'

'arcticwolf@dartoffice365.onmicrosoft.com'

'SVC_USR_DWEntPhone@ridedart.com'

'emailarchive@dmmta.com'

'SpanishVM@dmmta.com'

'Partnerships@dmmta.com'
)

#array to see who we removed licenses
$removedLicenses =@()

#Path to the text doc
$RemoveLicPath = "C:\Temp\RemoveLic.txt"

#Connect to Microsoft 365 in Azure
Connect-AzureAD


#loops through each line in the CSV file
foreach($line in $csv)
{ 
        # This checks to see if last activity date was not year 2021, The user has an Exchange Licence and is not in the DONOTUSE array
        If (($line.'Exchange Last Activity Date' -notmatch "2021") -and ($line.'Has Exchange License' -eq "True") -and ($donotuse -notcontains $line.'User Principal Name'))
        { 
            #the main group that removes the licenses.
            $userUPN=$line.'User Principal Name'
            $planName="ENTERPRISEPACK_GOV"
            $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $License.RemoveLicenses = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
            Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $license

            #adding UPN to the array
            $removedLicenses += $line.'User Principal Name'
        } 
   
} 

$removedLicenses | Out-File -Append $RemoveLicPath

read-host "press enter to quit."
