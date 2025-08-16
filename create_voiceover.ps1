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

$AudioOutputDir = Join-Path $OutputFolder "voice"
$DebugOutputDir = Join-Path $OutputFolder "debug"
New-Item -ItemType Directory -Force -Path $AudioOutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $DebugOutputDir | Out-Null

if ([string]::IsNullOrEmpty($AudioFormat)) {
    if ($AudioExt -eq "wav") {
        $AudioFormat = "riff-16khz-16bit-mono-pcm"
    } else {
        $AudioFormat = "audio-16khz-128kbitrate-mono-mp3"
    }
}

# Azure resource key
if (!(Test-Path $ResourceKeyPath)) { throw "Missing Azure resource key file: $ResourceKeyPath" }
$ResourceKeyPlain = (Get-Content -Path $ResourceKeyPath -Raw).Trim()
$SecureKey = ConvertTo-SecureString $ResourceKeyPlain -AsPlainText -Force
Remove-Variable ResourceKeyPlain

Write-Host "### Configuration ###"
Write-Host "PPTX file: $PptxFile"
Write-Host "Template: $SsmlTemplate"
Write-Host "Audio format: $AudioFormat"
Write-Host "Audio extension: $AudioExt"
Write-Host "Endpoint: $SpeechEndpoint"
Write-Host "Azure resource key file: $ResourceKeyPath"
Write-Host "#####################"

### Functions
function Get-PlainResourceKey {
    param([SecureString]$SecureKey)
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureKey)
    )
}

function Get-SlideNotes {
    param([object]$Presentation)
    $notesList = @()

    for ($i = 1; $i -le $Presentation.Slides.Count; $i++) {
        $slide = $Presentation.Slides.Item($i)
        $noteText = ""
        try {
            $notesPage = $slide.NotesPage
            foreach ($shape in $notesPage.Shapes) {
                if ($shape.Type -eq 14 -and $shape.TextFrame.HasText) {  
                    $noteText = $shape.TextFrame.TextRange.Text.Trim()
                    break
                }
            }
        } catch {
            Write-Warning ("Failed to extract notes for slide " + $i + ": " + $_)
        }
        Write-Host "Slide $i Notes: '$noteText'"
        $notesList += ,@($i, $noteText)
    }
    return $notesList
}

function Convert-TextToSSML {
    param([string]$Text, [string]$TemplateFile)
    $template = Get-Content -Path $TemplateFile -Raw -Encoding UTF8
    $ssml = $template.Replace('$TEXT', $Text)
    Write-Host "Slide $i SSML: '$ssml'"
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
        Write-Host "Created audio file for slide ${SlideIdx}"
    } catch {
        Write-Error "Failed to generate audio for slide ${SlideIdx}: $_"
    }
    Remove-Variable PlainKey
    return $audioFile
}

function Get-AudioDuration {
    param([string]$AudioFile)
    $resolvedPath = (Resolve-Path $AudioFile).Path
    $shell = New-Object -ComObject Shell.Application
    $folderPath = Split-Path $resolvedPath
    $fileName = Split-Path $resolvedPath -Leaf 

    $folder = $shell.Namespace($folderPath)
    if ($null -eq $folder) {
        Write-Warning "Shell could not resolve folder: $folderPath"
        return "Duration not found"
    }    

    $file = $folder.ParseName($fileName)
    if ($null -eq $file) {
        Write-Warning "Shell could not resolve file: $fileName"
        return "Duration not found"
    }    

    $durationStr = $folder.GetDetailsOf($file, 27)

    if ($durationStr -match "^(\d{1,2}):(\d{2})(?::(\d{2}))?$") {
        $hours = if ($matches[3]) { [int]$matches[1] } else { 0 }
        $minutes = if ($matches[3]) { [int]$matches[2] } else { [int]$matches[1] }
        $seconds = if ($matches[3]) { [int]$matches[3] } else { [int]$matches[2] }
        $output = $hours * 3600 + $minutes * 60 + $seconds
        $duration = $output + 1  
        return $duration
    } else {
        Write-Warning "Could not parse duration string: '$durationStr'"
        return "Duration not found"
    }
}

function Embed-Audio-And-Transition {
    param(
        [object]$Presentation,
        [string]$AudioFolder,
        [hashtable]$SlideDurations
    )

    for ($i = 1; $i -le $Presentation.Slides.Count; $i++) {
        $slide = $Presentation.Slides.Item($i)
        $audioPath = Join-Path $AudioFolder ("{0:D3}.mp3" -f $i)

        if (Test-Path $audioPath) {
            $audioPath = (Resolve-Path $audioPath).Path
            Start-Sleep -Milliseconds 300

            $shape = $slide.Shapes.AddMediaObject2($audioPath, $false, $true, 50, 50, 40, 40)
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

            $trans = $slide.SlideShowTransition
            $trans.AdvanceOnClick = $false
            $trans.AdvanceOnTime  = $true
            $trans.AdvanceTime    = $SlideDurations[$i]

            Write-Host "Slide ${i}: Audio embedded and set to auto-play with PlaySettings"
        }
    }
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

### run steps
$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $ppt.Presentations.Open($PptxFile)

$notes = Get-SlideNotes -Presentation $presentation

$durations = @{}
foreach ($note in $notes) {
    $idx = $note[0]
    $text = $note[1]

    Write-Host "### Processing slide $idx..."
    $ssml = Convert-TextToSSML -Text $text -TemplateFile $SsmlTemplate
    $audioFile = Convert-SSMLToAudio -SsmlContent $ssml -SlideIdx $idx -AudioFormat $AudioFormat -AudioExt $AudioExt -OutputDir $AudioOutputDir -SecureKey $SecureKey -Endpoint $SpeechEndpoint
    $duration = Get-AudioDuration -AudioFile $audioFile
    $durations[$idx] = $duration

    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_notes.txt" -f $idx)) -Value $text -Encoding UTF8
    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_ssml.xml" -f $idx)) -Value $ssml -Encoding UTF8
    Set-Content -Path (Join-Path $DebugOutputDir ("{0:D3}_duration.txt" -f $idx)) -Value $duration -Encoding UTF8
}

Embed-Audio-And-Transition -Presentation $presentation -AudioFolder $AudioOutputDir -SlideDurations $durations

# Remove extension and trailing dot properly
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($PptxFile)
$dirName  = [System.IO.Path]::GetDirectoryName($PptxFile)

$newFile = Join-Path $dirName ($baseName + "_with_audio.pptx")
$videoFile = Join-Path $dirName ($baseName + "_with_audio.mp4")

$ppt.Activate()
Start-Sleep -Seconds 2

$presentation.SaveAs($newFile)
Write-Host "Saved new PPTX: $newFile"

Export-PresentationToVideo -Presentation $presentation -OutputFile $videoFile

$presentation.Close()
$ppt.Quit()
