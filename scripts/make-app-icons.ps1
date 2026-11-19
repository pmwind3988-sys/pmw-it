# Draws the installed-app icons from src/assets/logo-128.png: the web app's
# (public/icons/) and the Android launcher's (android/app/src/main/res/).
# Re-run after the logo changes:  pwsh scripts/make-app-icons.ps1
#
# The only logo in the repo is 128px wide, so the large icons are upscaled
# and slightly soft. Drop a bigger PNG in as $Source and re-run to fix that.

param([string]$Source = "$PSScriptRoot/../src/assets/logo-128.png")

Add-Type -AssemblyName System.Drawing
$root = Resolve-Path "$PSScriptRoot/.."
$logo = [System.Drawing.Image]::FromFile((Resolve-Path $Source))

# $fill = how much of the icon's width the logo takes. A maskable icon may be
# cropped to a circle, so its logo sits inside the middle 80% (fill 0.62).
function Write-Icon([string]$path, [int]$size, [double]$fill, [bool]$round) {
  $dir = Split-Path $path
  New-Item -ItemType Directory -Force $dir | Out-Null
  $bmp = New-Object System.Drawing.Bitmap $size, $size
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.InterpolationMode = 'HighQualityBicubic'
  $g.PixelOffsetMode = 'HighQuality'
  $g.Clear([System.Drawing.Color]::Transparent)
  $white = [System.Drawing.Brushes]::White
  if ($round) { $g.FillEllipse($white, 0, 0, $size, $size) } else { $g.FillRectangle($white, 0, 0, $size, $size) }
  $w = [int]($size * $fill)
  $h = [int]($w * $logo.Height / $logo.Width)
  $g.DrawImage($logo, [int](($size - $w) / 2), [int](($size - $h) / 2), $w, $h)
  $g.Dispose()
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
}

Write-Icon "$root/public/icons/icon-192.png" 192 0.8 $false
Write-Icon "$root/public/icons/icon-512.png" 512 0.8 $false
Write-Icon "$root/public/icons/maskable-512.png" 512 0.62 $false
Write-Icon "$root/public/icons/apple-touch-icon.png" 180 0.8 $false

$res = "$root/android/app/src/main/res"
foreach ($d in @{ mdpi = 48; hdpi = 72; xhdpi = 96; xxhdpi = 144; xxxhdpi = 192 }.GetEnumerator()) {
  Write-Icon "$res/mipmap-$($d.Key)/ic_launcher.png" $d.Value 0.8 $false
  Write-Icon "$res/mipmap-$($d.Key)/ic_launcher_round.png" $d.Value 0.72 $true
}
$logo.Dispose()
Write-Host "Icons written."
