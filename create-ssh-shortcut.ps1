param(
    [Parameter(Mandatory,Position=0)][String]$Hostname,
    [String]$ConfigFile,
    [String]$IconLocation="%SystemRoot%\System32\cmd.exe,0",
    [String]$LocalForward,
    [String]$Name=$Hostname,
    [Int32]$Port=22,
    [String]$ProxyCommand,
    [String]$ProxyJump,
    [Int32]$WindowStyle=3,
    [String]$WorkingDirectory="%cd%"
)

$arguments = "-p " + $port
if(!([String]::IsNullOrEmpty($ConfigFile))){$arguments += " -F " + $ConfigFile}
if(!([String]::IsNullOrEmpty($LocalForward))){$arguments += " -L " + $LocalForward}
#This is going to need to be done differently
if(!([String]::IsNullOrEmpty($ProxyCommand))){$arguments += " -o ProxyCommand=" + $ProxyCommand}
if(!([String]::IsNullOrEmpty($ProxyJump))){$arguments += " -J " + $ProxyJump}
$arguments += " " + $Hostname

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($name + ".lnk")
$shortcut.Arguments = $arguments
$shortcut.Description = "SSH - " + $Hostname
$shortcut.IconLocation = $IconLocation
$shortcut.TargetPath = "C:/windows/system32/OpenSSH/ssh.exe"
$shortcut.WindowStyle = $WindowStyle
$shortcut.WorkingDirectory = $WorkingDirectory
$shortcut.Save()
