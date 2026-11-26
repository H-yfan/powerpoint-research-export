param()
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$source=Join-Path $root 'build/ResearchExport.ppam'
if(-not(Test-Path -LiteralPath $source)){throw '缺少 build/ResearchExport.ppam，请先运行 Build-Addin.ps1。'}
$target=Join-Path $env:APPDATA 'ResearchExport'
New-Item -ItemType Directory -Path $target -Force | Out-Null
$app=New-Object -ComObject PowerPoint.Application
try {
 for($i=$app.AddIns.Count;$i -ge 1;$i--){
  $old=$app.AddIns.Item($i)
  if([IO.Path]::GetFileName($old.FullName) -ieq 'ResearchExport.ppam'){$old.Loaded=0;$old.Registered=0;$app.AddIns.Remove($i)}
 }
 $installed=Join-Path $target 'ResearchExport.ppam'
 Copy-Item -LiteralPath $source -Destination $installed -Force
 Copy-Item -LiteralPath (Join-Path $root 'Finalize-Png.ps1') -Destination $target -Force
 $addin=$app.AddIns.Add($installed)
 $addin.Registered=-1
 $addin.AutoLoad=-1
 $addin.Loaded=-1
 if($addin.Loaded -ne -1){throw 'PowerPoint 未确认加载成功，请检查宏信任设置。'}
 Write-Host "已安装并加载：$installed"
} finally {if($app){[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}}


