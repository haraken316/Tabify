$csprojPath = "Tabify\Tabify.csproj"
[xml]$csproj = Get-Content $csprojPath

# Get current version
$version = $csproj.Project.PropertyGroup.Version
if ([string]::IsNullOrEmpty($version)) {
    $version = "1.0.0"
}
$parts = $version.Split('.')
$major = [int]$parts[0]
$minor = [int]$parts[1]
$patch = if ($parts.Length -gt 2) { [int]$parts[2] } else { 0 }

# Increment patch version
$patch++
$newVersion = "$major.$minor.$patch"
$newAssemblyVersion = "$newVersion.0"

Write-Host "Bumping version to $newVersion..."

# Update XML nodes
if ($csproj.Project.PropertyGroup.SelectSingleNode("Version") -eq $null) {
    $node = $csproj.CreateElement("Version")
    $node.InnerText = $newVersion
    $csproj.Project.PropertyGroup[0].AppendChild($node) > $null
} else {
    $csproj.Project.PropertyGroup.Version = $newVersion
}

if ($csproj.Project.PropertyGroup.SelectSingleNode("AssemblyVersion") -eq $null) {
    $node = $csproj.CreateElement("AssemblyVersion")
    $node.InnerText = $newAssemblyVersion
    $csproj.Project.PropertyGroup[0].AppendChild($node) > $null
} else {
    $csproj.Project.PropertyGroup.AssemblyVersion = $newAssemblyVersion
}

if ($csproj.Project.PropertyGroup.SelectSingleNode("FileVersion") -eq $null) {
    $node = $csproj.CreateElement("FileVersion")
    $node.InnerText = $newAssemblyVersion
    $csproj.Project.PropertyGroup[0].AppendChild($node) > $null
} else {
    $csproj.Project.PropertyGroup.FileVersion = $newAssemblyVersion
}

# Save without adding XML declaration if it didn't have one
$settings = New-Object System.Xml.XmlWriterSettings
$settings.OmitXmlDeclaration = $true
$settings.Indent = $true
$writer = [System.Xml.XmlWriter]::Create($csprojPath, $settings)
$csproj.Save($writer)
$writer.Close()

Write-Host "Publishing single file executable..."
dotnet publish $csprojPath -c Release -r win-x64

Write-Host "Publish complete! New version is $newVersion."
