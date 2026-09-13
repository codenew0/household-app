"""Windows標準OCR。画像は端末内で処理する。"""
import base64
import json
import os
import subprocess


OCR_SCRIPT = r'''
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1
function Await-Result($operation, $type) {
    $task = $asTask.MakeGenericMethod($type).Invoke($null, @($operation))
    $task.GetAwaiter().GetResult()
}
$file = Await-Result ([Windows.Storage.StorageFile]::GetFileFromPathAsync($env:HOUSEHOLD_RECEIPT_PATH)) ([Windows.Storage.StorageFile])
$stream = Await-Result ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
try {
    $decoder = Await-Result ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $limit = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
    $scale = [Math]::Min(1.0, $limit / [double][Math]::Max($decoder.PixelWidth, $decoder.PixelHeight))
    $transform = [Windows.Graphics.Imaging.BitmapTransform]::new()
    $transform.ScaledWidth = [uint32][Math]::Max(1, [Math]::Floor($decoder.PixelWidth * $scale))
    $transform.ScaledHeight = [uint32][Math]::Max(1, [Math]::Floor($decoder.PixelHeight * $scale))
    $bitmap = Await-Result ($decoder.GetSoftwareBitmapAsync(
        [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
        [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore, $transform,
        [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
        [Windows.Graphics.Imaging.ColorManagementMode]::ColorManageToSRgb)) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('ja'))
        if ($null -eq $engine) { throw 'Japanese OCR is unavailable. Install Japanese OCR in Windows language settings.' }
        $result = Await-Result ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @($result.Lines | ForEach-Object {
            $words = @($_.Words | ForEach-Object {
                @{text=$_.Text; x=$_.BoundingRect.X; y=$_.BoundingRect.Y;
                  width=$_.BoundingRect.Width; height=$_.BoundingRect.Height}
            })
            @{text=$_.Text; words=$words}
        })
        $regions = @()
        if ($env:HOUSEHOLD_OCR_REGIONS) {
            foreach ($region in (ConvertFrom-Json $env:HOUSEHOLD_OCR_REGIONS)) {
                $cropTransform = [Windows.Graphics.Imaging.BitmapTransform]::new()
                $cropTransform.ScaledWidth = $transform.ScaledWidth
                $cropTransform.ScaledHeight = $transform.ScaledHeight
                $bounds = New-Object Windows.Graphics.Imaging.BitmapBounds
                $bounds.X = [uint32]$region.x
                $bounds.Y = [uint32]$region.y
                $bounds.Width = [uint32]$region.width
                $bounds.Height = [uint32]$region.height
                $cropTransform.Bounds = $bounds
                $crop = Await-Result ($decoder.GetSoftwareBitmapAsync(
                    [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
                    [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore, $cropTransform,
                    [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
                    [Windows.Graphics.Imaging.ColorManagementMode]::ColorManageToSRgb)) ([Windows.Graphics.Imaging.SoftwareBitmap])
                try {
                    $cropEngine = $engine
                    if ($region.language) {
                        $alternative = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new($region.language))
                        if ($null -ne $alternative) { $cropEngine = $alternative }
                    }
                    $cropResult = Await-Result ($cropEngine.RecognizeAsync($crop)) ([Windows.Media.Ocr.OcrResult])
                    $cropLines = @($cropResult.Lines | ForEach-Object {
                        @{text=$_.Text; words=@($_.Words | ForEach-Object {
                            @{text=$_.Text; x=$_.BoundingRect.X + $bounds.X; y=$_.BoundingRect.Y + $bounds.Y;
                              width=$_.BoundingRect.Width; height=$_.BoundingRect.Height}
                        })}
                    })
                    $regions += ,@{lines=$cropLines}
                } finally { $crop.Dispose() }
            }
        }
        @{lines=$lines; width=$bitmap.PixelWidth; height=$bitmap.PixelHeight;
          angle=$result.TextAngle; regions=$regions} | ConvertTo-Json -Depth 8 -Compress
    } finally { if ($null -ne $bitmap) { $bitmap.Dispose() } }
} finally { $stream.Dispose() }
'''


def recognize_layout(path, regions=None):
    if os.name != 'nt':
        raise RuntimeError('レシートOCRはWindowsに対応しています。')
    environment = os.environ.copy()
    environment['HOUSEHOLD_RECEIPT_PATH'] = os.path.abspath(path)
    environment['HOUSEHOLD_OCR_REGIONS'] = json.dumps(regions or [])
    command = base64.b64encode(OCR_SCRIPT.encode('utf-16-le')).decode('ascii')
    executable = os.path.join(environment.get('SystemRoot', r'C:\Windows'),
                              'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    result = subprocess.run([executable, '-NoProfile', '-NonInteractive', '-EncodedCommand', command],
                            env=environment, capture_output=True, encoding='utf-8', errors='replace',
                            timeout=90, creationflags=subprocess.CREATE_NO_WINDOW)
    if result.returncode:
        raise RuntimeError('文字認識に失敗しました。Windowsの日本語OCR機能と画像サイズを確認してください。\n' + result.stderr[-1800:])
    return json.loads(result.stdout.lstrip('\ufeff'))
