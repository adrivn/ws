# The path to your Python virtual environment
$venvPath = "~/Envs/pacific"

# The paths to your Python scripts
$scriptPaths = @("./update_offers.py", "./update_stock.py")

# Activate the virtual environment
Write-Host "Activating Python virtual environment..."
& $venvPath\Scripts\Activate.ps1

# Check if the virtual environment is activated
if ($env:VIRTUAL_ENV -eq $null) {
    Write-Host "Failed to activate Python virtual environment. Exiting..."
    exit
}

# Loop through each script and execute it
foreach ($scriptPath in $scriptPaths) {
    Write-Host "Running Python script at $scriptPath..."
    python $scriptPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to execute Python script at $scriptPath. Exiting..."
        exit
    }
}

Write-Host "All scripts executed successfully. Exiting..."

