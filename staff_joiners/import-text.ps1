$arrayFromFile = @(Get-Content "C:\scripts\staff_joiners\joiner_data.txt")
$splitArray =$arrayFromFile.Split(":").Trim()
#$splitArray[1]
#$splitArray
$username = $splitArray[3]
$manager = $splitArray[5]
$title = $splitArray[11]
$site = $splitArray[15]
$location = $splitArray[17]
$department = $splitArray[45]

echo "name is $username"
echo "manager is $manager"
echo "job title is $title"
echo "Site is $site"
echo "office is $location"
echo "department is $department"
