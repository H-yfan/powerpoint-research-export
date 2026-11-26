param([switch]$Install)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root=$PSScriptRoot
$build=Join-Path $root 'build'
New-Item -ItemType Directory -Force $build | Out-Null
$macro=Join-Path $build 'ResearchExport.ppam'
if(Test-Path -LiteralPath $macro){throw 'build/ResearchExport.ppam 已存在，请备份并删除后重试。'}
$p=$null;$deck=$null
try {
 $p=New-Object -ComObject PowerPoint.Application
 if($null -eq $p -or $null -eq $p.Presentations){throw '无法连接 PowerPoint。请关闭其他 PowerPoint 窗口后重试，或改用 PowerShell 7 运行本脚本。'}
 $deck=$p.Presentations.Add(0)
 try { $components=$deck.VBProject.VBComponents; if($null -eq $components){throw 'VBA 工程访问被阻止'} } catch {throw '无法访问 VBA 工程。请按说明临时开启“信任对 VBA 工程对象模型的访问”，再运行。脚本不会自动改动信任设置。'}
 $bas=Join-Path $root 'ResearchExport.bas'
 # Import via Unicode CodeModule text: BAS is GBK for manual import on Chinese Windows.
 $text=[Text.Encoding]::GetEncoding(936).GetString([IO.File]::ReadAllBytes($bas))
 $text=$text -replace '(?m)^Attribute VB_Name[^\r\n]*\r?\n',''
 $module=$components.Add(1);$module.Name='ResearchExport';$module.CodeModule.AddFromString($text)
 $blank=$deck.Slides.Add(1,12)
 $deck.SaveAs($macro,30)
 $deck.Close();$deck=$null
} finally {
 if($deck){$deck.Saved=-1;$deck.Close()}
 if($p){[Runtime.InteropServices.Marshal]::ReleaseComObject($p)|Out-Null}
}
# Add Ribbon part to the OOXML package. Preserve all existing VBA/package content.
$zip=[IO.Compression.ZipFile]::Open($macro,[IO.Compression.ZipArchiveMode]::Update)
try {
 $entry=$zip.CreateEntry('customUI/customUI14.xml')
 $stream=$entry.Open();$data=[IO.File]::ReadAllBytes((Join-Path $root 'customUI14.xml'));$stream.Write($data,0,$data.Length);$stream.Dispose()
 $rel=$zip.GetEntry('_rels/.rels');$reader=New-Object IO.StreamReader($rel.Open());[xml]$xml=$reader.ReadToEnd();$reader.Dispose();$rel.Delete()
 $node=$xml.CreateElement('Relationship',$xml.DocumentElement.NamespaceURI)
 $node.SetAttribute('Id','ResearchExportRibbon');$node.SetAttribute('Type','http://schemas.microsoft.com/office/2007/relationships/ui/extensibility');$node.SetAttribute('Target','customUI/customUI14.xml');[void]$xml.DocumentElement.AppendChild($node)
 $rel=$zip.CreateEntry('_rels/.rels');$writer=New-Object IO.StreamWriter($rel.Open(),(New-Object Text.UTF8Encoding($false)));$xml.Save($writer);$writer.Dispose()
 # XML already normally has a default type, add one if absent.
 $ct=$zip.GetEntry('[Content_Types].xml');$reader=New-Object IO.StreamReader($ct.Open());[xml]$types=$reader.ReadToEnd();$reader.Dispose()
 if(-not @($types.DocumentElement.ChildNodes | Where-Object {$_.Extension -eq 'xml'}).Count){
  $node=$types.CreateElement('Default',$types.DocumentElement.NamespaceURI);$node.SetAttribute('Extension','xml');$node.SetAttribute('ContentType','application/xml');[void]$types.DocumentElement.AppendChild($node)
  $ct.Delete();$ct=$zip.CreateEntry('[Content_Types].xml');$writer=New-Object IO.StreamWriter($ct.Open());$types.Save($writer);$writer.Dispose()
 }
} finally {$zip.Dispose()}
$target=Join-Path $env:APPDATA 'ResearchExport'
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item -LiteralPath (Join-Path $root 'Finalize-Png.ps1') -Destination $target -Force
if($Install){
 $installed=Join-Path $target 'ResearchExport.ppam'
 Copy-Item -LiteralPath $macro -Destination $installed -Force
 $app=New-Object -ComObject PowerPoint.Application
 try {$addin=$app.AddIns.Add($installed);$addin.Registered=-1;$addin.AutoLoad=-1;$addin.Loaded=-1} finally {[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
 Write-Host "已安装：$installed。请重启 PowerPoint，检查“科研高清导出”选项卡。"
} else {Write-Host "已生成：$macro；PNG 组件已复制到 $target。按说明手动加载，或再次构建时使用 -Install。"}
Write-Host '请关闭“信任对 VBA 工程对象模型的访问”。此加载项未签名；请按组织策略信任。'

