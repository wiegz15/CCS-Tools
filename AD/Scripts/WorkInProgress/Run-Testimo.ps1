# Generate the date string in the desired format
$dateString = (Get-Date).ToString("MM-dd-yy")

# Define the path for the HTML report file with the date in the name
$htmlFileName = "${dateString}_TestimoSummary.html"
$htmlPath = Join-Path -Path $reportsDir -ChildPath $htmlFileName

# Check if the Testimo module is installed, and install it if it isn't
if (-not (Get-Module -ListAvailable -Name Testimo)) {
    Write-Output "Testimo module not found. Installing..."
    Install-Module -Name Testimo -AllowClobber -Force
} else {
    Write-Output "Testimo module is already installed."
}

# Run the Testimo command
Invoke-Testimo -ReportPath $htmlPath -AlwaysShowSteps
