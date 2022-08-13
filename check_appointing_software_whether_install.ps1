Function Get-Software  {
 
  [OutputType('System.Software.Inventory')]
 
  [Cmdletbinding()] 
 
  Param( 
 
  [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
 
  [String[]]$Computername=$env:COMPUTERNAME
 
  )         
 
  Begin {
 
  }
 
  Process  {     

  $PythonInstalledState = $false
  $PyCharmInstalledState = $false

  ForEach  ($Computer in  $Computername){ 
 
  If  (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {
 
      $Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
 
      ForEach($Path in $Paths) { 
 
      Write-Verbose  "Checking Path: $Path"
 
      #  Create an instance of the Registry Object and open the HKLM base key 
 
  Try  { 
 
     $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 
 
  } Catch  { 
 
     Write-Error $_ 
 
     Continue 
 
  } 
 
  #  Drill down into the Uninstall key using the OpenSubKey Method 
 
  Try  {
 
      $regkey=$reg.OpenSubKey($Path)  
 
      # Retrieve an array of string that contain all the subkey names 
 
      $subkeys=$regkey.GetSubKeyNames()      
 
      # Open each Subkey and use GetValue Method to return the required  values for each 
 
      ForEach ($key in $subkeys){   
 
      Write-Verbose "Key: $Key"
 
      $thisKey=$Path+"\\"+$key 
 
  Try {  
 
    $thisSubKey=$reg.OpenSubKey($thisKey)   
 
    # Prevent Objects with empty DisplayName 
 
    $DisplayName =  $thisSubKey.getValue("DisplayName")
 
    If ($DisplayName  -AND $DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {
 
        $Date = $thisSubKey.GetValue('InstallDate')
 
        If ($Date) {
 
            Try {
 
                $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
 
            } Catch{
 
                Write-Warning "$($Computer): $_ <$($Date)>"
 
                $Date = $Null
 
  }
 
  } 
 
  # Create New Object with empty Properties 
 
  $Publisher =  Try {
 
    $thisSubKey.GetValue('Publisher').Trim()
 
  } 
 
  Catch {
   
    $thisSubKey.GetValue('Publisher')
 
  }
 
  $Version = Try {
 
     #Some weirdness with trailing [char]0 on some strings
 
     $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
 
  } 
 
  Catch {
 
     $thisSubKey.GetValue('DisplayVersion')
 
  }
 
  $UninstallString =  Try {
 
     $thisSubKey.GetValue('UninstallString').Trim()
 
  } 
 
  Catch {
 
     $thisSubKey.GetValue('UninstallString')
 
  }
 
  $InstallLocation =  Try {
 
     $thisSubKey.GetValue('InstallLocation').Trim()
 
  } 
 
  Catch {
 
    $thisSubKey.GetValue('InstallLocation')
 
  }
 
  $InstallSource =  Try {
 
     $thisSubKey.GetValue('InstallSource').Trim()
 
  } 
 
  Catch {
 
      $thisSubKey.GetValue('InstallSource')
 
  }
 
  $HelpLink = Try {
 
     $thisSubKey.GetValue('HelpLink').Trim()
 
  } 
 
  Catch {
 
     $thisSubKey.GetValue('HelpLink')
 
  }
 
  $Object = [pscustomobject]@{
 
  Computername = $Computer
 
  DisplayName = $DisplayName
 
  Version  = $Version
 
  InstallDate = $Date
 
  Publisher = $Publisher
 
  UninstallString = $UninstallString
 
  InstallLocation = $InstallLocation
 
  InstallSource  = $InstallSource
 
  HelpLink = $thisSubKey.GetValue('HelpLink')
 
  EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
 
  }

 $Object.pstypenames.insert(0,'System.Software.Inventory')
 
 If($DisplayName -match "Python 3.*") {

    $PythonInstalledState = $True
    Write-Output $Object

 }

 If($DisplayName -match "PyCharm*") {

    $PyCharmInstalledState = $True
    Write-Output $Object

 }
 
  } Catch {
 
    Write-Warning "$Key : $_"
 
  }   
 
  }Catch  {}   
 
  } Catch  {}   
 
  $reg.Close() 
 
  } Catch  {}                  
 
  } 
 
  } 
 If($PythonInstalledState -eq $True) {
 
    $PythonState = Read-Host "检测到此电脑已经安装Python，是否跳过此步骤（Y/N）? "
    While($PythonState -ne "n" -and $PythonState -ne "N" -and $PythonState -ne "y" -and $PythonState -ne "Y") {
    
        $PythonState = Read-Host "输入有误，请重新输入（Y/N）： "
        if($PythonState -eq "n" -or $PythonState -eq "N" -or $PythonState -eq "y" -or $PythonState -eq "Y"){
            break
        }


    }

    If($PythonState -eq "n" -or $PythonState -eq "N"){
 
        # 执行安装Python的函数
        Write-Host "正在安装Python"

 }
    Elseif($PythonState -eq "y" -or $PythonState -eq "Y"){
    Write-Host "不在安装Python"
    }
 }
 Else {
 
    # 直接执行安装python函数
    Write-Host "正在安装python"
 }

 }
 If($PyCharmInstalledState -eq $True) {
 
    $PyCharmState = Read-Host "检测到此电脑已经安装PyCharm，是否跳过此步骤（Y/N）? "
    While($PyCharmState -ne "n" -and $PyCharmState -ne "N" -and $PyCharmState -ne "y" -and $PyCharmState -ne "Y"){
    
        $PyCharmState = Read-Host "输入有误，请重新输入（Y/N）： "
        if($PyCharmState -eq "n" -or $PyCharmState -eq "N" -or $PyCharmState -eq "y" -or $PyCharmState -eq "Y"){
        break
        }
    }

    If($PyCharmState -eq "n" -or $PyCharmState -eq "N"){
 
        # 执行安装PyCharm的函数
        Write-Host "正在安装pycharm"

 }
    Elseif($PyCharmState -eq "y" -or $PyCharmState -eq "Y"){
    Write-Host "不在安装pycharm"
    }
 }
  Else{
 
        # 直接执行安装Pycharm函数
        Write-Host "正在安装pycharm"
 }
 }
  } 
Get-Software