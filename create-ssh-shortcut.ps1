param(
    [Parameter(Mandatory,Position=0)][String]$Hostname,
    [String]$ConfigFile,
    [String]$IconLocation="%SystemRoot%\System32\cmd.exe,0",
    [String]$LocalForward,
    [String]$Name=$Hostname,
    [Int32]$Port=22,
    [String]$ProxyCommand,
    [String]$ProxyJump,
    [String]$User,
    [Int32]$WindowStyle=3,
    [String]$WorkingDirectory="%cd%"
)

#Initialize the arguments variable with the port number.
$arguments = "-p " + $port

#Process posible arguments with command line switches.
if(![String]::IsNullOrEmpty($ConfigFile)){$arguments += " -F " + $ConfigFile}
if(![String]::IsNullOrEmpty($LocalForward)){$arguments += " -L " + $LocalForward}
if(![String]::IsNullOrEmpty($ProxyJump)){$arguments += " -J " + $ProxyJump}
if(![String]::IsNullOrEmpty($User)){$arguments += " -l " + $User}

#Section to process options that use the -o(options) flag.
$optionsArray = @()
$optionsString = ""
if(![String]::IsNullOrEmpty($ProxyCommand)){$optionsArray += "ProxyCommand=" + $ProxyCommand}
if($optionsArray.count -gt 0){
    for ( $index = 0; $index -lt $optionsArray.count; $index++){
        switch($index){
            0 { $optionsString = $optionsArray[$index]; Break }
            Default { $optionsString += "," + $optionsArray[$index]; Break }
        }
    }
    $arguments += " -o " + $optionsString
}

#Always close out with the Hostname to connect to.
$arguments += " " + $Hostname

#Write the shortcut to a file.
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($name + ".lnk")
$shortcut.Arguments = $arguments
$shortcut.Description = "SSH - " + $Hostname
$shortcut.IconLocation = $IconLocation
$shortcut.TargetPath = "C:/windows/system32/OpenSSH/ssh.exe"
$shortcut.WindowStyle = $WindowStyle
$shortcut.WorkingDirectory = $WorkingDirectory
$shortcut.Save()
