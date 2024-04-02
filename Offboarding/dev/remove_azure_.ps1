Connect-AzureAD
Connect-ExchangeOnline

$userid = (Get-AzureADuser -objectid "test.user@testdomain.test").objectid

$Groups = Get-AzureADUserMembership -ObjectId $userID 
foreach($Group in $Groups){ 
    try { 
        Remove-AzureADGroupMember -ObjectId $Group.ObjectID -MemberId $userID -erroraction Stop 
    }
    catch {
        write-host "$($Group.displayname) membership cannot be removed via Azure cmdlets."
        Remove-DistributionGroupMember -identity $group.mail -member $userid -BypassSecurityGroupManagerCheck # -Confirm:$false
    }
}