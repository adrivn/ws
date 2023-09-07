# The path to your Python virtual environment
$venvPath = "$HOME\Envs\dataviz"

# Check if virtual environment exists, and creates one if not
if (Test-Path $venvPath\Scripts\Activate.ps1 -PathType Leaf) {
  Write-Host "Found virtual environment in $venvPath"
  # Activate the virtual environment
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
}
else {
  Write-Host "No virtual environment found. Creating a new one in: $venvPath"
  python -m venv $venvPath
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
  Write-Host "Installing superset requirements..."
  & pip install -r conf/supersetreqs.txt
}

# Check if the virtual environment is activated
if ($env:VIRTUAL_ENV -eq $null) {
    Write-Host "Failed to activate Python virtual environment. Exiting..."
    exit
}

Write-Host "Setting config variables"
set FLASK_APP=superset
cp superset_config.py $venvPath/superset_config.py
Write-Host "Initializing database"
cd $venvPath & cd Scripts & superset db upgrade 
Write-Host "Creating admin user"
superset fab create-admin
Write-Host "Initializing platform"
superset init
