$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$Payload = Join-Path $Root "payload\DExcelAssist.xlam"

$ExtraBas = Join-Path $Root "tools\DExcelAssistExtra.bas"

$EventsCls = Join-Path $Root "tools\DExcelAssistAppEvents.cls"

$AddInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"

$XLStartDir = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"

$AddinPath = Join-Path $AddInsDir "DExcelAssist.xlam"

$XLStartPath = Join-Path $XLStartDir "DExcelAssist.xlam"

$AppDir = Join-Path $env:APPDATA "DExcelAssist"

$LogDir = Join-Path $AppDir "logs"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Info($m){ Write-Host "[DExcelAssist] $m" }

function Warn($m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }

function Kill-Excel {

  Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill(); $_.WaitForExit(5000) } catch {} }

  Info "Excelプロセスを終了しました。"

}

function Remove-DExcelAssistOpenRegs {

  foreach($ver in '16.0','15.0','14.0'){

    $key="HKCU:\Software\Microsoft\Office\$ver\Excel\Options"

    if(Test-Path $key){

      $props=Get-ItemProperty $key

      foreach($p in $props.PSObject.Properties){

        if($p.Name -like 'OPEN*' -and [string]$p.Value -match 'DExcelAssist\.xlam'){

          Remove-ItemProperty -Path $key -Name $p.Name -ErrorAction SilentlyContinue

        }

      }

    }

  }

}

function Add-OpenReg([string]$path){

  foreach($ver in '16.0','15.0','14.0'){

    $key="HKCU:\Software\Microsoft\Office\$ver\Excel\Options"

    if(!(Test-Path $key)){ New-Item -Path $key -Force | Out-Null }

    $props=Get-ItemProperty $key -ErrorAction SilentlyContinue

    $used=@{}

    if($props){ foreach($p in $props.PSObject.Properties){ if($p.Name -like 'OPEN*'){ $used[$p.Name]=$true } } }

    $i=0; do { $name= if($i -eq 0){'OPEN'}else{"OPEN$i"}; $i++ } while($used.ContainsKey($name))

    New-ItemProperty -Path $key -Name $name -Value ('/R "' + $path + '"') -PropertyType String -Force | Out-Null

  }

}

function Xlam-HasSelectedUi([string]$path){

  Add-Type -AssemblyName System.IO.Compression.FileSystem

  $zip=[System.IO.Compression.ZipFile]::OpenRead($path)

  try{

    $entry=$zip.GetEntry('customUI/customUI.xml')

    if($null -eq $entry){ return $false }

    $sr=New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)

    try{ $xml=$sr.ReadToEnd() } finally { $sr.Close() }

    return ($xml -match 'tab id="DExcelAssistTab"' -and $xml -match 'label="DExcelAssist"' -and $xml -notmatch 'RelaxToolsTab' -and $xml -notmatch 'RelaxShapesTab' -and $xml -notmatch 'RelaxAppsTab' -and $xml -match 'SelectedFavoriteGroup' -and $xml -match 'hotkey' -and $xml -match 'searchFusen' -and $xml -match 'DExcelAssistExtraGroup' -and $xml -match 'dxaHolidaySheet' -and $xml -match 'dxaAutoFitRows' -and $xml -match 'dxaCreateSheetIndex' -and $xml -match 'BacklogTab' -and $xml -match 'dxaBacklogGroupByParent' -and $xml -notmatch 'dxaBacklogFormatGantt' -and $xml -notmatch 'dxaBacklogCreateGanttSummary' -and $xml -notmatch 'dxaBacklogCreateDelayList' -and $xml -notmatch 'dxaBacklogCreateMeetingView' -and $xml -notmatch 'dxaBacklogCreateAssigneeLoad' -and $xml -match 'dxaExportVbaWithFolderPicker' -and $xml -match 'dxaCreateFolderTreeWithFolderPicker' -and $xml -match 'dxaCreateFileList' -and $xml -match 'dxaCreateChangeHistory' -and $xml -match 'dxaCheckNotationVariants' -and $xml -match 'dxaDiagnoseHeavyWorkbook' -and $xml -match 'dxaImportTimecardNormalWork' -and $xml -match 'dxaImportTimecardShiftWork')

  } finally { $zip.Dispose() }

}



function Enable-ExcelVbomAndTrustedLocation {

  foreach($ver in '16.0','15.0','14.0'){

    $sec="HKCU:\Software\Microsoft\Office\$ver\Excel\Security"

    if(!(Test-Path $sec)){ New-Item -Path $sec -Force | Out-Null }

    New-ItemProperty -Path $sec -Name 'AccessVBOM' -Value 1 -PropertyType DWord -Force | Out-Null

    $tl="HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\DExcelAssist"

    if(!(Test-Path $tl)){ New-Item -Path $tl -Force | Out-Null }

    New-ItemProperty -Path $tl -Name 'Path' -Value ($AddInsDir + '\') -PropertyType String -Force | Out-Null

    New-ItemProperty -Path $tl -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force | Out-Null

    New-ItemProperty -Path $tl -Name 'Description' -Value 'DExcelAssist AddIn Location' -PropertyType String -Force | Out-Null

  }

}



function Save-DExcelAssistSettings {

  $reg = "HKCU:\Software\DExcelAssist"

  if(!(Test-Path $reg)){ New-Item -Path $reg -Force | Out-Null }

  New-ItemProperty -Path $reg -Name 'InstallRoot' -Value $Root -PropertyType String -Force | Out-Null

  New-ItemProperty -Path $reg -Name 'LocalVersion' -Value (Get-LocalVersionText) -PropertyType String -Force | Out-Null

  New-ItemProperty -Path $reg -Name 'AutoUpdateEnabled' -Value 0 -PropertyType DWord -Force | Out-Null

}



function Import-ExtraVbaModule([string]$xlamPath){

  if(!(Test-Path $ExtraBas)){ throw "tools\DExcelAssistExtra.bas が見つかりません。" }

  if(!(Test-Path $EventsCls)){ throw "tools\DExcelAssistAppEvents.cls が見つかりません。" }

  Enable-ExcelVbomAndTrustedLocation

  $xl = $null

  $wb = $null

  try{

    $xl = New-Object -ComObject Excel.Application

    $xl.DisplayAlerts = $false

    $xl.Visible = $false

    $wb = $xl.Workbooks.Open($xlamPath)

    $vbproj = $wb.VBProject

    for($i=$vbproj.VBComponents.Count; $i -ge 1; $i--){

      $comp=$vbproj.VBComponents.Item($i)

      if($comp.Name -eq 'DExcelAssistExtra' -or $comp.Name -eq 'DExcelAssistAppEvents' -or $comp.Name -like 'DExcelAssistAppEvents*'){

        $vbproj.VBComponents.Remove($comp)

      }

    }

    [void]$vbproj.VBComponents.Import($ExtraBas)

    [void]$vbproj.VBComponents.Import($EventsCls)

    $wb.Save()

    $wb.Close($true)

    $xl.Quit()

    Info "追加VBAモジュールをXLAMへ取り込みました。"

  } catch {

    try{ if($wb -ne $null){ $wb.Close($false) } }catch{}

    try{ if($xl -ne $null){ $xl.Quit() } }catch{}

    throw "追加機能用VBAモジュールの取り込みに失敗しました。Excelの『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』を有効にして再実行してください。詳細: $($_.Exception.Message)"

  } finally {

    try{ if($wb -ne $null){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } }catch{}

    try{ if($xl -ne $null){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } }catch{}

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

  }

}





function Remove-ExistingAddinRegistrations {

  Info "DExcelAssist同名アドインの既存登録のみ確認します。他のアドインは削除しません。"



  # 1) ExcelのAddIns一覧に同名/旧名が存在する場合は、先にInstalled=Falseで解除します。

  #    同名アドインが存在しない場合は何もせず、後続で通常追加のみ行います。

  $xl = $null

  try{

    $xl = New-Object -ComObject Excel.Application

    $xl.DisplayAlerts = $false

    $xl.Visible = $false

    $removed = 0

    foreach($a in $xl.AddIns){

      try{

        $name = [string]$a.Name

        $full = [string]$a.FullName

        if($name -match '^DExcelAssist\.xlam$' -or $full -match '\\DExcelAssist\.xlam$'){

          if($a.Installed){ $a.Installed = $false }

          $removed++

        }

      }catch{}

    }

    if($removed -gt 0){ Info "既存のDExcelAssistアドイン登録を解除しました。件数=$removed" }

    else{ Info "既存のDExcelAssistアドイン登録はありません。追加のみ実行します。" }

    $xl.Quit()

  } catch {

    Warn "既存アドイン解除のExcel COM処理はスキップしました: $($_.Exception.Message)"

    try{ if($xl -ne $null){ $xl.Quit() } }catch{}

  } finally {

    try{ if($xl -ne $null){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } }catch{}

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

  }



  # 2) OPEN起動登録、Add-in Manager上の同名パス登録を削除します。

  Remove-DExcelAssistOpenRegs

  foreach($ver in '16.0','15.0','14.0'){

    $mgr="HKCU:\Software\Microsoft\Office\$ver\Excel\Add-in Manager"

    if(Test-Path $mgr){

      $props=Get-ItemProperty $mgr -ErrorAction SilentlyContinue

      if($props){

        foreach($pr in $props.PSObject.Properties){

          if($pr.Name -match '^(PS|Path|Parent|Child|Drive|Provider)' ){ continue }

          if(([string]$pr.Name -match 'DExcelAssist') -or ([string]$pr.Value -match 'DExcelAssist\.xlam')){

            Remove-ItemProperty -Path $mgr -Name $pr.Name -ErrorAction SilentlyContinue

          }

        }

      }

    }

  }



  # 3) 実体ファイルがある場合だけ削除します。なければ何もしません。

  foreach($p in @($AddinPath,$XLStartPath)){

    if(Test-Path $p){

      Remove-Item $p -Force -ErrorAction SilentlyContinue

      Info "既存ファイルを削除しました: $p"

    }

  }

}





function Remove-LegacyAutoUpdateTask {

  # v92以前で登録された自動アップデート用タスクが残っている場合だけ削除します。

  # 新規の自動アップデート登録は行いません。

  $taskName = "DExcelAssist Auto Update"

  try{

    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if($null -ne $existing){

      Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

      Info "旧自動アップデート用タスクを削除しました。"

    }

  } catch {

    try{ schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null }catch{}

  }

}



function Install-SelectedRelaxTools {

  if(!(Test-Path $Payload)){ throw "payload\DExcelAssist.xlam が見つかりません。" }

  Info "インストール/修復の前処理として、Excelプロセスを強制終了します。"

  Kill-Excel

  Start-Sleep -Milliseconds 500

  Remove-LegacyAutoUpdateTask

  New-Item -ItemType Directory -Force -Path $AddInsDir,$XLStartDir | Out-Null

  Remove-ExistingAddinRegistrations

  Copy-Item $Payload $AddinPath -Force

  Import-ExtraVbaModule $AddinPath

  Copy-Item $AddinPath $XLStartPath -Force

  if(!(Xlam-HasSelectedUi $AddinPath)){ throw "customUIのDExcelAssist 1タブ統合に失敗しています。" }

  Add-OpenReg $AddinPath

  foreach($ver in '16.0','15.0','14.0'){

    $sec="HKCU:\Software\Microsoft\Office\$ver\Excel\Options"

    if(!(Test-Path $sec)){ New-Item -Path $sec -Force | Out-Null }

    New-ItemProperty -Path $sec -Name 'ShowDevTools' -Value 1 -PropertyType DWord -Force | Out-Null

  }

  # COM AddIns registration, if Excel can be started invisibly.

  try{

    $xl = New-Object -ComObject Excel.Application

    $xl.DisplayAlerts = $false

    $xl.Visible = $false

    $found = $false

    foreach($a in $xl.AddIns){ try{ if([string]$a.FullName -ieq $AddinPath){ $a.Installed = $true; $found=$true } }catch{} }

    if(-not $found){ $a=$xl.AddIns.Add($AddinPath, $true); $a.Installed=$true }

    $xl.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

  } catch { Warn "Excel COM登録はスキップしました: $($_.Exception.Message)" }

  Save-DExcelAssistSettings

  Info "インストール/修復が完了しました。Excelを起動して DExcelAssist タブを確認してください。"

}

function Diagnose {

  Write-Host "==== DExcelAssist v112 統合リボン 診断 ===="

  Write-Host "Root: $Root"

  Write-Host "Payload: $Payload exists=$(Test-Path $Payload)"

  Write-Host "ExtraBas: $ExtraBas exists=$(Test-Path $ExtraBas)"
  Write-Host "EventsCls: $EventsCls exists=$(Test-Path $EventsCls)"

  Write-Host "Addin: $AddinPath exists=$(Test-Path $AddinPath) size=$((Get-Item $AddinPath -ErrorAction SilentlyContinue).Length) selectedUI=$(if(Test-Path $AddinPath){Xlam-HasSelectedUi $AddinPath}else{$false})"

  Write-Host "XLSTART: $XLStartPath exists=$(Test-Path $XLStartPath) size=$((Get-Item $XLStartPath -ErrorAction SilentlyContinue).Length) selectedUI=$(if(Test-Path $XLStartPath){Xlam-HasSelectedUi $XLStartPath}else{$false})"

  foreach($ver in '16.0','15.0','14.0'){

    $key="HKCU:\Software\Microsoft\Office\$ver\Excel\Options"

    if(Test-Path $key){

      $props=Get-ItemProperty $key

      foreach($p in $props.PSObject.Properties){ if($p.Name -like 'OPEN*' -and [string]$p.Value -match 'DExcelAssist\.xlam'){ Write-Host "$ver $($p.Name)=$($p.Value)" } }

    }

  }

  try{

    $xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')

    Write-Host "Excel: 起動中"

    foreach($a in $xl.AddIns){ try{ if(([string]$a.FullName) -match 'DExcelAssist\.xlam'){ Write-Host "AddIn: $($a.Name) Installed=$($a.Installed) FullName=$($a.FullName)" } }catch{} }

  }catch{ Write-Host "Excel: 未起動または取得不可" }

}



function Get-LocalVersionText {

  $vfile = Join-Path $Root "VERSION.txt"

  if(Test-Path $vfile){ return ((Get-Content $vfile -Raw).Trim()) }

  return "v0.0.0"

}

function Convert-ToVersionObject([string]$text){

  $t = ([string]$text).Trim()

  $t = $t -replace '^v',''

  if($t -match '^[0-9]+$'){ $t = "$t.0.0" }

  elseif($t -match '^[0-9]+\.[0-9]+$'){ $t = "$t.0" }

  try { return [version]$t } catch { return [version]'0.0.0' }

}

function Compare-VersionText([string]$a,[string]$b){

  $va = Convert-ToVersionObject $a

  $vb = Convert-ToVersionObject $b

  return $va.CompareTo($vb)

}

function Read-ReleaseVersion {

  $current = Get-LocalVersionText

  if([string]::IsNullOrWhiteSpace($current)){ $current = "v0.0.0" }

  while($true){

    $inputVersion = Read-Host "リリースするバージョンを入力してください（例: v1.0.0 / 1.0.0、未入力なら現在値 $current）"

    if([string]::IsNullOrWhiteSpace($inputVersion)){ $inputVersion = $current }

    $inputVersion = ([string]$inputVersion).Trim()

    if($inputVersion -notmatch '^v'){ $inputVersion = 'v' + $inputVersion }

    if($inputVersion -match '^v[0-9]+(\.[0-9]+){0,3}([._-][0-9A-Za-z]+)?$'){ return $inputVersion }

    Write-Host "バージョン形式が不正です。例: v1.0.0 / 1.0.0 / v1.0.0-beta" -ForegroundColor Yellow

  }

}



function Set-VersionFile([string]$dir,[string]$version){

  $vf = Join-Path $dir "VERSION.txt"

  Set-Content -Path $vf -Value $version -Encoding UTF8

}



function Create-ReleaseFiles {

  $version = Read-ReleaseVersion

  if([string]::IsNullOrWhiteSpace($version)){ $version = "v0.0.0" }

  $safeVersion = $version -replace '[^0-9A-Za-z._-]','_'

  $releaseDir = Join-Path $Root "_release"

  $uploadDir = Join-Path $releaseDir "main_branch_upload"

  $stageParent = Join-Path $releaseDir "_stage"

  $stageRoot = Join-Path $stageParent "DExcelAssistSafeInstaller_$safeVersion"

  if(Test-Path $releaseDir){ Remove-Item $releaseDir -Recurse -Force -ErrorAction SilentlyContinue }

  New-Item -ItemType Directory -Force -Path $releaseDir,$uploadDir,$stageRoot | Out-Null



  $items = @('DExcelAssist.bat','README.md','VERSION.txt','payload','tools','licenses')

  foreach($it in $items){

    $src = Join-Path $Root $it

    if(Test-Path $src){

      Copy-Item $src (Join-Path $stageRoot $it) -Recurse -Force

      Copy-Item $src (Join-Path $uploadDir $it) -Recurse -Force

    }

  }



  # バッチ実行者が指定したバージョンを、配布物とmainブランチアップロード用ファイルへ反映します。

  Set-VersionFile $Root $version

  Set-VersionFile $stageRoot $version

  Set-VersionFile $uploadDir $version



  $guide = @"

# DExcelAssist release files



Version: $version



## 配布用ファイル



このフォルダには、DExcelAssist の配布用ZIPと、GitHub main ブランチへ配置できる `main_branch_upload` フォルダを作成します。



## main_branch_upload に含まれるもの



- DExcelAssist.bat

- VERSION.txt

- README.md

- payload/DExcelAssist.xlam

- tools/DExcelAssist.ps1

- tools/DExcelAssistExtra.bas

- tools/DExcelAssistAppEvents.cls

- licenses/RelaxTools_LICENSE_NOTE.txt



## 注意



自動アップデート機能は含めていません。

インストール・アンインストール時に操作する対象は DExcelAssist.xlam のみです。RelaxToolsなど他のExcelアドインは削除しません。

"@

  Set-Content -Path (Join-Path $releaseDir "README_RELEASE.md") -Value $guide -Encoding UTF8



  $zipPath = Join-Path $releaseDir ("DExcelAssistSafeInstaller_$safeVersion.zip")

  if(Test-Path $zipPath){ Remove-Item $zipPath -Force }

  Compress-Archive -Path $stageRoot -DestinationPath $zipPath -Force

  Remove-Item $stageParent -Recurse -Force -ErrorAction SilentlyContinue

  Info "リリース用ファイルを作成しました: $releaseDir"

  Write-Host "ZIP: $zipPath"

  Write-Host "mainブランチアップロード用: $uploadDir"

}



function Uninstall {

  Remove-LegacyAutoUpdateTask

  Remove-DExcelAssistOpenRegs

  foreach($p in @($AddinPath,$XLStartPath)){ if(Test-Path $p){ Remove-Item $p -Force -ErrorAction SilentlyContinue } }

  Info "アンインストールしました。"

}



while($true){

  Write-Host ""

  Write-Host "DExcelAssist v112 統合リボン"

  Write-Host "1: インストール/修復（Excel強制終了後、DExcelAssistのみ置換）"

  Write-Host "2: 診断"

  Write-Host "3: アンインストール"

  Write-Host "4: Excel残留プロセスを強制終了"

  Write-Host "5: リリース用ファイル作成（バージョン入力＋mainブランチアップロード用＋ZIP）"

  Write-Host "0: 終了"

  $n=Read-Host "番号を入力してください"

  switch($n){

    '1' { Install-SelectedRelaxTools }

    '2' { Diagnose }

    '3' { Uninstall }

    '4' { Kill-Excel }

    '5' { Create-ReleaseFiles }

    '0' { return }

    default { Write-Host "不正な番号です。" }

  }

}

