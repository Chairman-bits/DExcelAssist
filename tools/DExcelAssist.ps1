param()
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Payload = Join-Path $Root "payload\DExcelAssist.xlam"
$ExtraBas = Join-Path $Root "tools\DExcelAssistExtra.bas"
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
    return ($xml -match 'tab id="DExcelAssistTab"' -and $xml -match 'label="DExcelAssist"' -and $xml -notmatch 'RelaxToolsTab' -and $xml -notmatch 'RelaxShapesTab' -and $xml -notmatch 'RelaxAppsTab' -and $xml -match 'SelectedFavoriteGroup' -and $xml -match 'hotkey' -and $xml -match 'searchFusen' -and $xml -match 'execSourceExport' -and $xml -match 'DExcelAssistExtraGroup' -and $xml -match 'dxaHolidaySheet' -and $xml -match 'dxaAutoFitRows')
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
function Import-ExtraVbaModule([string]$xlamPath){
  if(!(Test-Path $ExtraBas)){ throw "tools\DExcelAssistExtra.bas が見つかりません。" }
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
      if($comp.Name -eq 'DExcelAssistExtra'){
        $vbproj.VBComponents.Remove($comp)
      }
    }
    [void]$vbproj.VBComponents.Import($ExtraBas)
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

function Install-SelectedRelaxTools {
  if(!(Test-Path $Payload)){ throw "payload\DExcelAssist.xlam が見つかりません。" }
  Info "インストール/修復の前処理として、Excelプロセスを強制終了します。"
  Kill-Excel
  Start-Sleep -Milliseconds 500
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
  Info "インストール/修復が完了しました。Excelを起動して DExcelAssist タブを確認してください。"
}
function Diagnose {
  Write-Host "==== DExcelAssist v89 統合リボン 診断 ===="
  Write-Host "Root: $Root"
  Write-Host "Payload: $Payload exists=$(Test-Path $Payload)"
  Write-Host "ExtraBas: $ExtraBas exists=$(Test-Path $ExtraBas)"
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
function Uninstall {
  Remove-DExcelAssistOpenRegs
  foreach($p in @($AddinPath,$XLStartPath)){ if(Test-Path $p){ Remove-Item $p -Force -ErrorAction SilentlyContinue } }
  Info "アンインストールしました。"
}
while($true){
  Write-Host ""
  Write-Host "DExcelAssist v89 統合リボン"
  Write-Host "1: インストール/修復（Excel強制終了後にDExcelAssist 1タブ統合版を登録）"
  Write-Host "2: 診断"
  Write-Host "3: アンインストール"
  Write-Host "4: Excel残留プロセスを強制終了"
  Write-Host "0: 終了"
  $n=Read-Host "番号を入力してください"
  switch($n){
    '1' { Install-SelectedRelaxTools }
    '2' { Diagnose }
    '3' { Uninstall }
    '4' { Kill-Excel }
    '0' { return }
    default { Write-Host "不正な番号です。" }
  }
}
