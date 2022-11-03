     Function Search-OutlookCalendar
     {
      Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null
      $olFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
      $outlook = new-object -comobject outlook.application
      $namespace = $outlook.GetNameSpace(“MAPI”)
      $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
      $folder.items | Select-Object -Property Subject, Start, Duration, Location, Body
      $myNewfolder.items
     }
# $Past and $Future define a time frame for the search process
$Past=-30
$Future=0
$Pattern="Réunion"
# replace $_.Subject with $_.Body to change the search scope
Search-OutlookCalendar `
                    |where-object { $_.start -gt (Get-date).AddDays($Past) -AND $_.start -lt (Get-Date).AddDays($Future)} `
                    |Where-Object {select-string -pattern $Pattern -InputObject $_.Subject} `
                    |Select-Object -Property Subject, Start, Duration, Location
