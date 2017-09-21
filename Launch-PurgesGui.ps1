
# Original example posted at http://technet.microsoft.com/en-us/library/ff730941.aspx
# All Examples can be found here https://github.com/dlwyatt/WinFormsExampleUpdates 
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$XamlPath
    )
if(!$XamlPath){
    $XamlPath = "$($PWD)" + '\PowershellGuiMain.xaml'
} 
$Global:xmlWPF = Get-Content -Path $XamlPath
#load types
try{
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.Forms
    } 
catch{
    Throw “Failed to load Windows Presentation Framework assemblies.”
    }

[Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089") | Out-Null
#Create the XAML reader using a new XML node reader
[xml]$xaml = 
$reader =(New-Object [System.Xml.XmlNodeReader] $Global:xmlWPF)
$Global:xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $Global:xmlWPF))

#Create hooks to each named object in the XAML
$Global:xmlWPF.SelectNodes(“//*[@Name]”) | %{
Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#show
$result = $form.ShowDialog()


#check results
if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    #store results
    $x = $textBox.Text
    $x
}