param(
 [Parameter(Mandatory=$true)][string]$InputPath,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [Parameter(Mandatory=$true)][ValidateRange(36,2400)][int]$Dpi,
 [Parameter(Mandatory=$true)][int]$Width,
 [Parameter(Mandatory=$true)][int]$Height,
 [int]$CanvasWidth=0,
 [int]$CanvasHeight=0,
 [ValidateRange(0,32767)][int]$CropX=0,
 [ValidateRange(0,32767)][int]$CropY=0,
 [switch]$RequireAlpha
)
$ErrorActionPreference='Stop'
$stage=$null
try {
Add-Type -TypeDefinition @"
using System;
using System.IO;
public static class PngDensity {
 static uint U(byte[] b,int p) { return ((uint)b[p]<<24)|((uint)b[p+1]<<16)|((uint)b[p+2]<<8)|b[p+3]; }
 static void Put(Stream s,uint x) {s.WriteByte((byte)(x>>24));s.WriteByte((byte)(x>>16));s.WriteByte((byte)(x>>8));s.WriteByte((byte)x);}
 static uint Crc(byte[] b,int p,int n) {uint c=0xffffffff;for(int i=p;i<p+n;i++){c^=b[i];for(int j=0;j<8;j++) c=(c&1)!=0 ? 0xedb88320^(c>>1):c>>1;}return c^0xffffffff;}
 public static byte[] Patch(byte[] b,int w,int h,int dpi,bool alpha) {
 byte[] sig={137,80,78,71,13,10,26,10};
 if(b.Length<33)throw new Exception("Truncated PNG");
 for(int i=0;i<8;i++)if(b[i]!=sig[i])throw new Exception("Not PNG");
 if(U(b,8)!=13 || System.Text.Encoding.ASCII.GetString(b,12,4)!="IHDR" || U(b,16)!=w || U(b,20)!=h)throw new Exception("Unexpected pixel size");
 if(alpha && b[25]!=6 && b[25]!=4)throw new Exception("PNG has no alpha channel");
 using(var s=new MemoryStream()) {
 s.Write(sig,0,8); bool end=false;
 for(int p=8;p<b.Length;) {
 uint len=U(b,p); if(len>int.MaxValue || (long)p+12+len>b.Length)throw new Exception("Invalid chunk");
 int n=(int)len;string t=System.Text.Encoding.ASCII.GetString(b,p+4,4);
 if(Crc(b,p+4,n+4)!=U(b,p+n+8))throw new Exception("CRC mismatch");
 if(t!="pHYs")s.Write(b,p,n+12);
 if(t=="IHDR") {
 uint ppm=(uint)Math.Floor(dpi/0.0254+0.5);
 using(var c=new MemoryStream()) {
 byte[] name=System.Text.Encoding.ASCII.GetBytes("pHYs");c.Write(name,0,4);Put(c,ppm);Put(c,ppm);c.WriteByte(1);byte[] data=c.ToArray();Put(s,9);s.Write(data,0,data.Length);Put(s,Crc(data,0,data.Length));
 }
 }
 p+=n+12; if(t=="IEND"){if(p!=b.Length)throw new Exception("Trailing bytes");end=true;break;}
 }
 if(!end)throw new Exception("Missing IEND");return s.ToArray();
 }
 }
}
"@
$inputBytes=[IO.File]::ReadAllBytes($InputPath)
if($CanvasWidth -gt 0 -or $CanvasHeight -gt 0){
 if($Width -lt 1 -or $Height -lt 1 -or $CropX + $Width -gt $CanvasWidth -or $CropY + $Height -gt $CanvasHeight){throw 'Canvas is smaller than selected region'}
 # Validate actual canvas before cropping; no interpolation or rescaling.
 $null=[PngDensity]::Patch($inputBytes,$CanvasWidth,$CanvasHeight,$Dpi,$false)
 Add-Type -AssemblyName System.Drawing
 $bitmap=$null;$cropped=$null;$memory=$null
 try {
  $bitmap=[Drawing.Bitmap]::new($InputPath)
  $rectangle=[Drawing.Rectangle]::new($CropX,$CropY,$Width,$Height)
  $cropped=$bitmap.Clone($rectangle,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $memory=[IO.MemoryStream]::new()
  $cropped.Save($memory,[Drawing.Imaging.ImageFormat]::Png)
  $inputBytes=$memory.ToArray()
 } finally {
  if($memory){$memory.Dispose()};if($cropped){$cropped.Dispose()};if($bitmap){$bitmap.Dispose()}
 }
}
$bytes=[PngDensity]::Patch($inputBytes,$Width,$Height,$Dpi,[bool]$RequireAlpha)
# Stage in destination directory. Existing output is replaced only after validation.
$full=[IO.Path]::GetFullPath($OutputPath)
$stage=Join-Path ([IO.Path]::GetDirectoryName($full)) ([IO.Path]::GetRandomFileName())
[IO.File]::WriteAllBytes($stage,$bytes)
if([IO.File]::Exists($full)){[IO.File]::Replace($stage,$full,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($stage,$full)}
exit 0
} catch {
 try { [IO.File]::WriteAllText($InputPath+'.error.txt',$_.Exception.ToString()) } catch {}
 exit 1
} finally { if($stage -and [IO.File]::Exists($stage)){[IO.File]::Delete($stage)} }

