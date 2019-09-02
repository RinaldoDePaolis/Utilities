# Recover lost service account passwords from IIS

[see full article here](https://nov.ms/lost)

`cmd.exe /c $env:windir\system32\inetsrv\appcmd.exe list apppool "<NAME OF APP POOL>" /text:ProcessModel.Password`
