## Try all user module paths and return first one
foreach ($modulePath in $env:PSModulePath.Split(';')) {
    if ($modulePath -inotlike '*\Program Files\*' -and $modulePath -inotlike '*\Windows\*') {
        $userPSModulePath = $modulePath;
        break;
    }
}

