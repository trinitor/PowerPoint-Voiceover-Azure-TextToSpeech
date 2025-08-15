### Define Variables
param(
    [Parameter(Mandatory=$true)][string]$PptxFile,
    [Parameter(Mandatory=$true)][string]$SsmlTemplate,
    [Parameter(Mandatory=$true)][string]$SpeechEndpoint,
    [Parameter(Mandatory=$true)][string]$ResourceKeyPath,
    [string]$AudioFormat = "",
    [ValidateSet("mp3","wav")][string]$AudioExt = "mp3",
    [string]$OutputFolder = "output"
)

$PptxFile        = (Resolve-Path $PptxFile).Path
$SsmlTemplate    = (Resolve-Path $SsmlTemplate).Path
$ResourceKeyPath = (Resolve-Path $ResourceKeyPath).Path
$AudioOutputDir  = Join-Path $OutputFolder "voice"
$DebugOutputDir  = Join-Path $OutputFolder "debug"
New-Item -ItemType Directory -Force -Path $AudioOutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $DebugOutputDir | Out-Null

if ([string]::IsNullOrEmpty($AudioFormat)) {
    if ($AudioExt -eq "wav") {
        $AudioFormat = "riff-16khz-16bit-mono-pcm"
    } else {
        $AudioFormat = "audio-16khz-128kbitrate-mono-mp3"
    }
}

# Read Azure resource key file
if (!(Test-Path $ResourceKeyPath)) { throw "Missing Azure resource key file: $ResourceKeyPath" }
$ResourceKeyPlain = (Get-Content -Path $ResourceKeyPath -Raw).Trim()
$SecureKey = ConvertTo-SecureString $ResourceKeyPlain -AsPlainText -Force
Remove-Variable ResourceKeyPlain

### Functions
function Get-PlainResourceKey {
    param([SecureString]$SecureKey)
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureKey)
    )
}

function Get-SlideNote {
    param([object]$Slide)

    $noteText = ""
    try {
        $notesPage = $Slide.NotesPage
        foreach ($shape in $notesPage.Shapes) {
            if ($shape.Type -eq 14 -and $shape.TextFrame -and $shape.TextFrame.HasText) {
                $noteText = $shape.TextFrame.TextRange.Text.Trim()
                break
            }
        }
    } catch {
        Write-Warning ("Failed to extract notes for slide {0}: {1}" -f $Slide.SlideIndex, $_)
    }
    return $noteText
}

function Convert-TextToSSML {
    param(
        [string]$Text,
        [string]$TemplateFile
    )

    $template = Get-Content -Path $TemplateFile -Raw -Encoding UTF8
    $ssml = $template.Replace('$TEXT', $Text)
    return $ssml
}

function Convert-SSMLToAudio {
    param(
        [string]$SsmlContent,
        [int]$SlideIdx,
        [string]$AudioFormat,
        [string]$AudioExt,
        [string]$OutputDir,
        [SecureString]$SecureKey,
        [string]$Endpoint
    )

    $audioFile = Join-Path $OutputDir ("{0:D3}.{1}" -f $SlideIdx, $AudioExt)
    $PlainKey = Get-PlainResourceKey -SecureKey $SecureKey
    $headers = @{
        "Ocp-Apim-Subscription-Key" = $PlainKey
        "Content-Type"              = "application/ssml+xml"
        "X-Microsoft-OutputFormat"  = $AudioFormat
        "User-Agent"                = "AzureTTSClient"
    }
    $utf8Body = [System.Text.Encoding]::UTF8.GetBytes($SsmlContent)

    try {
        Invoke-WebRequest -Uri $Endpoint -Method Post -Headers $headers -Body $utf8Body -OutFile $audioFile -ErrorAction Stop
    } catch {
        throw ("Failed to generate audio for slide {0}: {1}" -f $SlideIdx, $_)
    } finally {
        if ($PlainKey) { Remove-Variable PlainKey -ErrorAction SilentlyContinue }
    }
    return $audioFile
}

function Get-AudioDuration {
    param([string]$AudioFile)

    $resolvedPath = (Resolve-Path $AudioFile).Path
    $shell = New-Object -ComObject Shell.Application
    $folderPath = Split-Path $resolvedPath
    $fileName = Split-Path $resolvedPath -Leaf 
    $folder = $shell.Namespace($folderPath)
    $file = $folder.ParseName($fileName)
    $durationStr = $folder.GetDetailsOf($file, 27)

    if ($durationStr -match "^(\d{1,2}):(\d{2})(?::(\d{2}))?$") {
        $hours   = if ($matches[3]) { [int]$matches[1] } else { 0 }
        $minutes = if ($matches[3]) { [int]$matches[2] } else { [int]$matches[1] }
        $seconds = if ($matches[3]) { [int]$matches[3] } else { [int]$matches[2] }
        $total = $hours * 3600 + $minutes * 60 + $seconds + 1
        return ($total)
    } else {
        Write-Warning "Could not parse duration string: '$durationStr'"
        return 0
    }
}

function Embed-Audio-To-Slide {
    param(
        [object]$Slide,
        [string]$AudioPath
    )

    if (!(Test-Path $AudioPath)) {
        Write-Warning ("Audio not found for slide {0}: {1}" -f $Slide.SlideIndex, $AudioPath)
        return $null
    }

    $audioPathResolved = (Resolve-Path $AudioPath).Path
    Start-Sleep -Milliseconds 300

    $shape = $Slide.Shapes.AddMediaObject2($audioPathResolved, $false, $true, 50, 50, 40, 40)
    $shape.Visible = $true
    $shape.Left = -50
    $shape.Top  = 50
    $shape.AnimationSettings.Animate = $true
    $shape.AnimationSettings.PlaySettings.PlayOnEntry = -1
    $shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = 0
    $shape.AnimationSettings.PlaySettings.LoopUntilStopped = 0
    $shape.AnimationSettings.PlaySettings.RewindMovie = -1
    $shape.AnimationSettings.AdvanceMode = 2
    $shape.AnimationSettings.AdvanceTime = 0

    Write-Host ("Slide {0}: Audio embedded & set to autoplay." -f $Slide.SlideIndex)
    return $shape
}

function Set-Transition {
    param(
        [object]$Slide,
        [int]$DurationSeconds
    )

    $trans = $Slide.SlideShowTransition
    $trans.AdvanceOnClick = $false
    $trans.AdvanceOnTime  = $true
    $trans.AdvanceTime    = [math]::Max(1, $DurationSeconds)
    Write-Host ("Slide {0}: Transition set to {1}s." -f $Slide.SlideIndex, $trans.AdvanceTime)
}

function Export-PresentationToVideo {
    param(
        [object]$Presentation,
        [string]$OutputFile
    )

    Write-Host "Starting video export to $OutputFile (720p)..."
    $Presentation.CreateVideo($OutputFile, $true, 5, 720, 30, 85)

    while ($Presentation.CreateVideoStatus -ne 3) { 
        Write-Host "Exporting video... "
        Start-Sleep -Seconds 5
    }

    Write-Host "Video export completed: $OutputFile"
}

# Main
Write-Host "### Configuration"
Write-Host "PPTX file: $PptxFile"
Write-Host "Template: $SsmlTemplate"
Write-Host "Audio format: $AudioFormat"
Write-Host "Audio extension: $AudioExt"
Write-Host "Endpoint: $SpeechEndpoint"
Write-Host "Azure resource key file: $ResourceKeyPath"

$ppt = $null
$presentation = $null

$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $ppt.Presentations.Open($PptxFile)

for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
    Write-Host "### Processing slide $i..."
    $slide = $presentation.Slides.Item($i)

    $text = Get-SlideNote -Slide $slide
    Write-Host ("Slide {0} Notes: {1}" -f $i, $text)
    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_notes.txt" -f $i)) -Value $text -Encoding UTF8

    $ssml = Convert-TextToSSML -Text $text -TemplateFile $SsmlTemplate
    Write-Host ("Slide {0} SSML: {1}" -f $i, $ssml)
    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_ssml.xml" -f $i)) -Value $ssml -Encoding UTF8

    $audioFile = Convert-SSMLToAudio -SsmlContent $ssml -SlideIdx $i -AudioFormat $AudioFormat -AudioExt $AudioExt -OutputDir $AudioOutputDir -SecureKey $SecureKey -Endpoint $SpeechEndpoint

    $duration = Get-AudioDuration -AudioFile $audioFile
    Write-Host ("Slide {0} Audio length: {1}" -f $i, $duration)
    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_duration.txt" -f $i)) -Value $duration -Encoding UTF8

    Embed-Audio-To-Slide -Slide $slide -AudioPath $audioFile | Out-Null
    Set-Transition -Slide $slide -DurationSeconds $duration
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($PptxFile)
$dirName  = [System.IO.Path]::GetDirectoryName($PptxFile)
$newFile  = Join-Path $dirName ($baseName + "_with_audio.pptx")
$videoFile= Join-Path $dirName ($baseName + "_with_audio.mp4")

$ppt.Activate()
Start-Sleep -Seconds 2

$presentation.SaveAs($newFile)
Write-Host "Saved new PPTX: $newFile"

Export-PresentationToVideo -Presentation $presentation -OutputFile $videoFile

$presentation.Close()
$ppt.Quit()
