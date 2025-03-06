# Open Folder Selection Dialog
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.Description = "Select the folder containing subfolders to zip"
$null = $FolderBrowser.ShowDialog()
$SourcePath = $FolderBrowser.SelectedPath

# If no folder is selected, exit
if (!$SourcePath) {
    Write-Host "No folder selected. Exiting..." -ForegroundColor Yellow
    Exit
}

# Ask the user for the preferred format (ZIP or CBZ)
$zipType = Read-Host "Enter format (zip/cbz)"
if ($zipType -ne "zip" -and $zipType -ne "cbz") {
    Write-Host "Invalid choice. Defaulting to zip format." -ForegroundColor Yellow
    $zipType = "zip"
}

# Set Output Path (Zipped folder inside SourcePath)
$DestinationPath = "$SourcePath\Zipped"

# Ensure the destination folder exists
if (!(Test-Path -Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath | Out-Null
}

# Get only the subfolders (exclude the "Zipped" folder itself)
$folders = Get-ChildItem -Path $SourcePath -Directory | Where-Object { $_.FullName -ne $DestinationPath }

# Ask the user if they want to delete subfolders after zipping
$deleteChoice = Read-Host "Do you want to delete the original subfolders after zipping? (Y/N)"

foreach ($folder in $folders) {
    $zipFile = "$DestinationPath\$($folder.Name).$zipType"

    # Use Compress-Archive to zip each folder
    Compress-Archive -Path "$($folder.FullName)\*" -DestinationPath $zipFile -Force

    Write-Host "✅ Zipped: $($folder.Name) -> $zipFile"

    # Delete the folder if the user chose "Y"
    if ($deleteChoice -eq "Y" -or $deleteChoice -eq "y") {
        Remove-Item -Path $folder.FullName -Recurse -Force
        Write-Host "🗑️ Deleted: $($folder.Name)"
    }
}

Write-Host "🎉 All folders have been processed successfully!"
