<#
Author: AlanDSaster
Github: https://github.com/AlanDSaster/
Date:   2021-03-12
Summary:
  This script resolves the following issue when trying to print from outlook:
    Cannot print unless an item is selected

  See the following link for more information:
    https://services.dartmouth.edu/TDClient/1806/Portal/KB/ArticleDet?ID=96554
#>
echo "Script Starting: Fix-Outlook-Cannot_print_unless_an_item_is_selected.ps1"
$rootdirectory = "C:\Users\"
$user = $env:UserName
$subfolder = "\AppData\Roaming\Microsoft\Outlook\"
$filename = "outlprnt"
$old = ".old"
$originalfilename = "$rootdirectory$user$subfolder$filename"
echo $originalfilename
if (Test-Path -Path "$originalfilename") {
  $iteration = 0
  DO {
    $iteration++
  } While (Test-Path -Path "$originalfilename$old$iteration")
  $newfilename = "$originalfilename$old$iteration"
  Rename-Item -Path "$originalfilename" -NewName "$newfilename"
  echo "$filename found. Renamed to $newfilename"
} else {
  echo "$filename does not exist"
}
echo "Script Finished: Fix-Outlook-Cannot_print_unless_an_item_is_selected.ps1"
