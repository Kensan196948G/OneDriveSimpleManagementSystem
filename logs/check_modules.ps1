try { 
  $ErrorActionPreference = 'Stop' 
  $modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Files') 
  foreach ($module in $modules) { 
    if (Get-Module -ListAvailable $module) { 
      Write-Host "$module - Installed" 
    } else { 
      Write-Host "$module - Not Installed" 
    } 
  } 
} catch { 
  Write-Host "Error checking modules: $_" 
  exit 1 
} 
