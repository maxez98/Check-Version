forEach ($appId in 'Access', 'Excel', 'Lync', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Visio', 'Word') {

   $regKey =  "hklm:\Software\Classes\$($appId).application\curVer"

   if (-not (test-path $regKey)) {
     $version = 'n/a'
   }
   else {
      $defaultValue = (get-item $regKey).getValue('')
   
      $version =  $defaultValue -replace '.*\.(\d+)', '$1'
   }
   write-output('  {0,-10} {1,3}' -f $appId, $version)
}