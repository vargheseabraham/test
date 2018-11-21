$fullPath = (Get-ChildItem -Path "C:\DataCollection\" -Force -Recurse).FullName
$fullName = $fullPath | ?{$_ -like '*ApplicationUsage*'}

$allApps = @()
foreach($fName in $fullName)
{
    $csvData = Import-Csv $fName
    $allApps += ($csvData).CurrentApplication
}