Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class HotKeyHelper {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}
"@

# Paths
$docPath = "$env:USERPROFILE\Desktop\AutoScreenshots.docx"
$tempFolder = "$env:TEMP\AutoScreenshots"

New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

# Open/Create Word
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = 0

if (Test-Path $docPath) {
    $doc = $word.Documents.Open($docPath)
}
else {
    $doc = $word.Documents.Add()
    $doc.SaveAs2($docPath, 16)
}

function Take-Screenshot($filePath) {

    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    $bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height

    $graphics = [System.Drawing.Graphics]::FromImage($bmp)

    $graphics.CopyFromScreen(
        $bounds.Location,
        [System.Drawing.Point]::Empty,
        $bounds.Size
    )

    $bmp.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)

    $graphics.Dispose()
    $bmp.Dispose()
}

function Get-Description {

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Screenshot Description"
    $form.Width = 420
    $form.Height = 150
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter heading/description:"
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(350,20)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,35)
    $textBox.Size = New-Object System.Drawing.Size(380,25)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Capture"
    $button.Location = New-Object System.Drawing.Point(10,70)

    $button.Add_Click({
        $form.Tag = $textBox.Text
        $form.Close()
    })

    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($button)

    $form.ShowDialog() | Out-Null

    return $form.Tag
}

Write-Host ""
Write-Host "======================================="
Write-Host " F9  -> Capture Screenshot"
Write-Host " ESC -> Exit"
Write-Host "======================================="
Write-Host ""

while ($true) {

    Start-Sleep -Milliseconds 300

    # F9
    if ([HotKeyHelper]::GetAsyncKeyState(120) -eq -32767) {

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

        $imgPath = "$tempFolder\shot_$timestamp.png"

        Write-Host "Capturing screenshot..."

        Take-Screenshot $imgPath

        $description = Get-Description

        # Add Heading
        $range = $doc.Content
        $range.Collapse(0)

        $range.InsertAfter("`r`n")
        $range.InsertAfter("$description`r`n")

        $range.Font.Bold = 1
        $range.Font.Size = 16

        $range.InsertParagraphAfter()

        # Insert Image
        $range = $doc.Content
        $range.Collapse(0)

        $inlineShape = $range.InlineShapes.AddPicture($imgPath)

        $inlineShape.Width = 500

        $range.InsertParagraphAfter()

        # Save document
        $doc.Save()

        Write-Host "Added to Word document"
    }

    # ESC
    if ([HotKeyHelper]::GetAsyncKeyState(27) -eq -32767) {

        Write-Host "Closing..."

        $doc.Save()

        $doc.Close()

        $word.Quit()

        break
    }
}