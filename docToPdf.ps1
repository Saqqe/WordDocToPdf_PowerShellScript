# Prompt the user for source and destination folders
$source_folder = Read-Host "Enter the path to the source folder containing Word documents"
$destination_folder = Read-Host "Enter the path to the destination folder for PDF files"

# Check if the source folder exists
if (-not (Test-Path -Path $source_folder)) {
    Write-Host "Source folder does not exist. Exiting."
    exit
}

# Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destination_folder)) {
    New-Item -ItemType Directory -Path $destination_folder | Out-Null
    Write-Host "Destination folder created: $destination_folder"
}

# Initialize Word application
$word_app = New-Object -ComObject Word.Application
$word_app.Visible = $false  # Hide Word application

# Error handling: Ensure Word application closes
try {
    # Process Word documents
    Get-ChildItem -Path $source_folder -Filter *.doc? | ForEach-Object {
        try {
            $document = $word_app.Documents.Open($_.FullName)
            $pdf_filename = Join-Path $destination_folder "$($_.BaseName).pdf"
            
            # Save the document as PDF (no [ref] needed)
            $document.SaveAs([string] $pdf_filename, [int] 17) # 17 for PDF format
            
            $document.Close()
            Write-Host "Converted: $($_.Name) to PDF"
        } catch {
            Write-Host "Failed to convert: $($_.Name). Error: $_"
        }
    }
} finally {
    # Close Word application
    $word_app.Quit()
    Write-Host "Word application closed."
}

Write-Host "Conversion complete. PDFs saved in: $destination_folder"
