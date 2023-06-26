$header='Email'
echo $header
$CUD_old = Import-CSV C:\scripts\student\old.csv | Group-Object -AsHashTable -AsString -Property 'Email'
$CUD_new = Import-CSV C:\scripts\student\new.csv | Group-Object -AsHashTable -AsString -Property 'Email'

$OnlyInold = @()
$OnlyInNew = @()

ForEach ($Device in $CUD_old.Values) {
    if (!$CUD_new[$Device.Email]) {
        $OnlyInold += $Device
    }
}

ForEach ($Device in $CUD_new.Values) {
    if (!$CUD_old[$Device.Email]) {
        $OnlyInNew += $Device
    }
}

$OnlyInOld | Export-CSV -NoTypeInformation C:\scripts\student\Dropouts.csv
$OnlyInNew | Export-CSV -NoTypeInformation C:\scripts\student\Newstudents.csv