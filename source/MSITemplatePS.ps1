[Void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') 
[Void][reflection.assembly]::loadwithpartialname("System.Windows.Forms")
[Void][System.Windows.Forms.Application]::EnableVisualStyles()
$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition
function Invoke-Method ($Object, $MethodName, $ArgumentList) {
 return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
}


#Define Constants
$msiOpenDatabaseModeReadOnly = 0
$msiOpenDatabaseModeTransact = 1
$msiViewModifyUpdate = 2
$msiViewModifyReplace = 4
$msiViewModifyDelete = 6
$msiTransformErrorNone = 0
$msiTransformValidationNone = 0

#Check if XML exists, if not create

$XMLPath = $env:APPDATA + "\ModifyMsiXML.xml"
$text = '<?xml version="1.0" standalone="yes"?>
<xml>
  <SumInfo Name="Title" Value="[ProductName] [ProductVersion]" />
  <SumInfo Name="Subject" Value="Subject" />
  <SumInfo Name="Author" Value="Alex" />
  <SumInfo Name="Keyword" Value="" />
  <SumInfo Name="Comments" Value="Comment" />
  <SumInfo Name="Template" Value="N/A" />
  <SumInfo Name="LastAuthor" Value="N/A" />
  <SumInfo Name="Revision" Value="N/A" />
  <Property Name="ALLUSERS" Value="1" />
  <Property Name="REBOOT" Value="ReallySupress" />
</xml>'

    If(!(Test-Path $XMLPath -PathType leaf)){
        $text | Out-File $XMLPath
    }

#Select File Dialog
function Select-FileDialog
{
	param([string]$Title,[string]$Directory,[string]$Filter="MSI/MST/ISM Files (*.msi,*.mst,*.ism)|*.msi;*.mst;*.ism")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$objForm.ShowHelp = $true
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Return "Operation cancelled by user"
	}
}

#Open File Dialog
Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "XML (*.xml)| *.xml"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
} 

#Save File Dialog
Function Save-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "XML (*.xml)| *.xml"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
} 

#Create Form Functions
function Create-Form {

#	Draw Form
	$form1 = New-Object System.Windows.Forms.Form
	$form1.ClientSize = New-Object System.Drawing.Size(730,415)
	$form1.Text = "MSI Template Editor" 
	$form1.Name = "form1" 
	$form1.DataBindings.DefaultDataSourceUpdateMode = 0 
	$form1.FormBorderStyle = 5

#   Text Configuration
    $Text1 = New-Object System.Windows.Forms.Label   
    $Text1.Text = "Configuration"  
    $Text1.Size = New-Object System.Drawing.Size(150,20)
    $Text1.Location = New-Object System.Drawing.Point(13,8) 
    $form1.Controls.Add($Text1)

#   Text MSI
    $Text2 = New-Object System.Windows.Forms.Label   
    $Text2.Text = "Select MSI or MST"  
    $Text2.Size = New-Object System.Drawing.Size(150,20)
    $Text2.Location = New-Object System.Drawing.Point(378,8) 
    $form1.Controls.Add($Text2)

#   Textbox1
    $TextBox1 = New-Object System.Windows.Forms.TextBox 
    $TextBox1.Location = New-Object System.Drawing.Size(380,35) 
    $TextBox1.Size = New-Object System.Drawing.Size(260,20) 
    $form1.Controls.Add($TextBox1) 
    if ($TextBox1.Text -ne $null) {$TextBox1.Text = "MSI File"}

#   Textbox2
    $TextBox2 = New-Object System.Windows.Forms.TextBox 
    $TextBox2.Location = New-Object System.Drawing.Size(380,90) 
    $TextBox2.Size = New-Object System.Drawing.Size(260,20) 
    $TextBox2.Enabled = $true
    $form1.Controls.Add($TextBox2) 
    $TextBox2.Text = "MST File"

#   Select Button MSI
    $SelectButton1 = New-Object System.Windows.Forms.Button
    $SelectButton1.Location = New-Object System.Drawing.Size(648,35)
    $SelectButton1.Size = New-Object System.Drawing.Size(75,23)
    $SelectButton1.Text = "Select"
    $SelectButton1.Add_Click({$TextBox1.Text = Select-FileDialog})
    $form1.Controls.Add($SelectButton1)

#   Checkbox MST
    $checkbox1 = new-object System.Windows.Forms.checkbox
    $checkbox1.Location = new-object System.Drawing.Size(380,50)
    $checkbox1.Size = new-object System.Drawing.Size(250,50)
    $checkbox1.Text = "Create MST"
    $checkbox1.Checked = $false
    $checkbox1.Add_CheckStateChanged({
        if ($checkBox1.Checked){ 
        $SelectButton2.Enabled = $false
        $TextBox2.Enabled = $false }
        else { 
        $SelectButton2.Enabled = $true
        $TextBox2.Enabled = $true }
        })
    $form1.Controls.Add($checkbox1)  

#   Select Button MST
    $SelectButton2 = New-Object System.Windows.Forms.Button
    $SelectButton2.Location = New-Object System.Drawing.Size(648,90)
    $SelectButton2.Size = New-Object System.Drawing.Size(75,23)
    $SelectButton2.Text = "Select"
    $SelectButton2.Enabled = $true
    $SelectButton2.Add_Click({$TextBox2.Text = Select-FileDialog})
    $form1.Controls.Add($SelectButton2)

#   Modify Button
    $ModifyButton = New-Object System.Windows.Forms.Button
    $ModifyButton.Location = New-Object System.Drawing.Size(490,120)
    $ModifyButton.Size = New-Object System.Drawing.Size(110,110)
    $ModifyButton.Text = "Modify"
    $ModifyButton.Add_Click({
        if ($checkbox1.Checked -eq $true){
            $WKDir = $TextBox1.Text.TrimEnd([System.IO.Path]::GetFileName($TextBox1.Text))
            $MSI2 = [System.IO.Path]::GetFileNameWithoutExtension($TextBox1.Text) + '.mod' + [System.IO.Path]::GetExtension($TextBox1.Text)
            Copy-Item $TextBox1.Text ($WKDir + $MSI2)
            $NewDir = $WKDir + $MSI2
            write-host $NewDir
            Change-MSISysInfo2 $NewDir
            Modify-MSI-Create-MST -MSI_Path $NewDir -MST_Path ischecked
            Remove-Item $NewDir
        }
        ElseIf ($TextBox2.Text -eq "MST File") {
        Change-MSISysInfo2 $TextBox1.Text
        Modify-MSI -MSI_Path $TextBox1.Text -MST_Path ischecked
        }
        Else{
            $WKDir = $TextBox1.Text.TrimEnd([System.IO.Path]::GetFileName($TextBox1.Text))
            $MSI2 = [System.IO.Path]::GetFileNameWithoutExtension($TextBox1.Text) + '.mod' + [System.IO.Path]::GetExtension($TextBox1.Text)
            Copy-Item $TextBox1.Text ($WKDir + $MSI2)
            $NewDir = $WKDir + $MSI2
            write-host $NewDir
            Change-MSISysInfo2 $NewDir
            Modify-MSI-Create-MST -MSI_Path $NewDir -MST_Path $TextBox2.Text
            Remove-Item $NewDir
        }
    })
    $form1.Controls.Add($ModifyButton) 

#	Draw DataGrid
	$dataGrid1 = New-Object System.Windows.Forms.DataGrid
	$dataGrid1.Size = New-Object System.Drawing.Size(350,338)
	$dataGrid1.Location = New-Object System.Drawing.Point (13,35)
	$dataGrid1.DataBindings.DefaultDataSourceUpdateMode = 0 
	$dataGrid1.HeaderForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0) 
	$dataGrid1.Name = "dataGrid1" 
	$dataGrid1.DataMember = "" 
	$dataGrid1.TabIndex = 0
	$form1.Controls.Add($dataGrid1) 

#   Text OutPut
    $Text1 = New-Object System.Windows.Forms.Label   
    $Text1.Text = "Output"  
    $Text1.Size = New-Object System.Drawing.Size(150,17)
    $Text1.Location = New-Object System.Drawing.Point(382,250) 
    $form1.Controls.Add($Text1)

#   Add Output Box for operations
    $MyMultiLineTextBox = New-Object System.Windows.Forms.TextBox 
    $MyMultiLineTextBox.Multiline = $true
    $MyMultiLineTextBox.Width = 300
    $MyMultiLineTextBox.Height = 100
    $MyMultiLineTextBox.Scrollbars = "Vertical"
    $MyMultiLineTextBox.location = new-object system.drawing.point(380,268)
    $MyMultiLineTextBox.Font = "Microsoft Sans Serif,8"
    $form1.controls.Add($MyMultiLineTextBox)

#   Load Configuration File
    $xml_networkschema = $XMLPath
    $ds = New-Object System.Data.Dataset
	$ds.ReadXml($xml_networkschema)
	$dataGrid1.DataSource = $ds
    $MyMultiLineTextBox.AppendText("`nLoaded Configuration XML File: " + $XMLPath)

#	Draw Open xml Configuration Button
	$button_openxml = New-Object System.Windows.Forms.Button
	$button_openxml.Size = New-Object System.Drawing.Size(150,25)
	$button_openxml.Location = New-Object System.Drawing.Point(13,365)
	$button_openxml.Text = "Open XML Document"
	$button_openxml.Add_Click({
		$xml_networkschema = Get-FileName
#		Bind Data to DataGrid

		$ds = New-Object System.Data.Dataset
		$ds.ReadXml($xml_networkschema)
		$dataGrid1.DataSource = $ds
	})
	#$form1.Controls.Add($button_openxml)

#	Save xml Configuration Button
	$button_savexml = New-Object System.Windows.Forms.Button
	$button_savexml.Size = New-Object System.Drawing.Size(150,25)
	$button_savexml.Location = New-Object System.Drawing.Point(13,380)
	$button_savexml.Text = "Save Template"
	$button_savexml.enabled = "false"
	$button_savexml.Add_Click({
		#$dbm_savenetworkschema = Save-FileName
        $dbm_savenetworkschema = $XMLPath
		$dataGrid1.DataSource.writexml($dbm_savenetworkschema)
        $MyMultiLineTextBox.AppendText("`r`nNew Configuration Saved to: " + $XMLPath)
	})
	$form1.Controls.Add($button_savexml)
	$form1.ShowDialog()| Out-Null 
}

#Create MSI/MST Modify Functions
function Change-MSIProperties {
    param (
            $database2,
            [string]$str_propertyname,
            [string]$str_propertyvalue
          )

    #$view = Invoke-Method $database2 OpenView @("SELECT * FROM Property WHERE Property='$str_propertyname'")
    $view =  $database2.GetType().InvokeMember("Openview", "InvokeMethod", $Null, $database2, @("SELECT * FROM Property WHERE Property='$str_propertyname'"))
    Invoke-Method $view Execute
    $record = Invoke-Method $view Fetch @()
    $view = $null
    if ($record -ne $Null) {
      #$view = Invoke-Method $database2 OpenView @("UPDATE Property SET Value='$str_propertyvalue' WHERE Property='$str_propertyname'")
      $view =  $database2.GetType().InvokeMember("Openview", "InvokeMethod", $Null, $database2, @("UPDATE Property SET Value='$str_propertyvalue' WHERE Property='$str_propertyname'"))
    }
    Else {
      #$view = Invoke-Method $database2 OpenView @("INSERT INTO Property (Property, Value) VALUES ('$str_propertyname','$str_propertyvalue')")
      $view =  $database2.GetType().InvokeMember("Openview", "InvokeMethod", $Null, $database2, @("INSERT INTO Property (Property, Value) VALUES ('$str_propertyname','$str_propertyvalue')"))
    }

     Invoke-Method $view Execute
     $view.GetType().InvokeMember("Close", "InvokeMethod", $Null, $view, $Null)
     Invoke-Method $view Close @()
     $view = $Null
} 

Function Change-MSISysInfo2{
            param (
                        $MSILOC
                      )
    start-job {
        param (
                [string]$MSI
            )
      $sc = New-Object -ComObject MSScriptControl.ScriptControl.1
      $sc.Language = 'VBScript'
      $sc.AddCode('
      Function MyFunction(a,b,c,d,e,f)
      MSI = f 

    Set installer = Nothing
    Set installer = CreateObject("WindowsInstaller.Installer")
    Set suminfo = installer.SummaryInformation(MSI, 20)

    If (Not IsNull(a)) Then
    suminfo.Property(2) = a   
    End If     
    If (Not IsNull(b)) Then  
    suminfo.Property(3) = b     
    End If
    If (Not IsNull(c)) Then      
    suminfo.Property(4) = c    
    End IF
    If (Not IsNull(d)) Then                
    suminfo.Property(5) = d  
    End If
    If (Not IsNull(e)) Then                  
    suminfo.Property(6) = e   
    End If
    suminfo.Persist 

    MyFunction = a & ";" & b & ";" & c & ";" & d & ";" & e

    End Function
      ')

        $XMLPath = $env:APPDATA + "\ModifyMsiXML.xml"
        [xml]$xml = Get-Content $XMLPath
        foreach ($summ in $xml.xml.SumInfo){
                 If ($summ.Name -eq "Title"){If ($summ.Value -ne "N/A"){$a = $summ.Value}}
                If ($summ.Name -eq "Subject"){If ($summ.Value -ne "N/A"){$b = $summ.Value}}
                If ($summ.Name -eq "Author"){If ($summ.Value -ne "N/A"){$c = $summ.Value}}
                If ($summ.Name -eq "Keyword"){If ($summ.Value -ne "N/A"){$d = $summ.Value}}
                If ($summ.Name -eq "Comments"){If ($summ.Value -ne "N/A"){$e = $summ.Value}}
            }
  
          $sc.codeobject.MyFunction([string]$a,[string]$b,[string]$c,[string]$d,[string]$e,[string]$MSI)
        } -runas32 -ArgumentList $MSILOC | wait-job | receive-job -OutVariable outputs

            foreach ($a in $outputs.split(";")){
                If ($a -ne ""){$MyMultiLineTextBox.AppendText("`r`nAdded " + $a + " in Summary Information")}
            }
}

function Modify-MSI
{
 [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
		[ValidateScript({$_ | Test-Path -PathType Leaf})]
        [string]$MSI_Path,
		[Parameter(Mandatory=$false)]
		[string]$MST_Path
    )
    try {


    $installer = New-Object -ComObject WindowsInstaller.Installer
    $WorkingDir = $MSI_Path.TrimEnd([System.IO.Path]::GetFileName($MSI_Path))

    #if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".msi")
    #if (Test-Path $MSI_Path )
    #{
     #$MSI_file2 = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.001' + [System.IO.Path]::GetExtension($MSI_Path)
     #Copy-Item $MSI_Path ($WorkingDir + $MSI_file2)
    #}
$MSI_file2 = $MSI_Path

    #$database1 = Invoke-Method $installer OpenDatabase  @($MSI_Path, $msiOpenDatabaseModeReadOnly)
   # $database1 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @($MSI_Path, $msiOpenDatabaseModeReadOnly))
    #$database2 = Invoke-Method $installer OpenDatabase  @(($WorkingDir + $MSI_file2), $msiViewModifyUpdate)
    #$database2 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @(($WorkingDir + $MSI_file2), $msiViewModifyUpdate))
    $database2 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @($MSI_file2, $msiViewModifyUpdate))

    if (Test-Path $MST_Path )
    {
     $transform = [System.IO.Path]::GetFileNameWithoutExtension($MST_Path) + '.new' + [System.IO.Path]::GetExtension($MST_Path)
     $database2.GetType().InvokeMember("ApplyTransform", "InvokeMethod", $Null, $database2, @($MST_Path, 63))
    }
    Else {
         if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".msi")
            {
                $transform = $WorkingDir + [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.mst'
            }
         if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".ism")
            {
                $transform = $Null
            }
    }
    [xml]$xml = Get-Content $XMLPath
    
    foreach ($prop in $xml.xml.Property){
        Change-MSIProperties $database2 $prop.Name $prop.Value
        $MyMultiLineTextBox.AppendText("`r`nAdded " + $prop.Name + " = " + $prop.Value + " Property")
    }
    
        
     Invoke-Method $database2 Commit
     $database2 = $Null


     if ($transform -ne $Null) {
         if (Test-Path ($WorkingDir + $MSI_file2) )
         {
          $MSI_file3 = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.002' + [System.IO.Path]::GetExtension($MSI_Path)
          Copy-Item ($WorkingDir + $MSI_file2) ($WorkingDir + $MSI_file3)
         }

         $database3 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @(($WorkingDir + $MSI_file3), $msiOpenDatabaseModeReadOnly))

         #Invoke-Method $database3 GenerateTransform $database1 $transform
         $database3.GetType().InvokeMember("GenerateTransform", "InvokeMethod", $Null, $database3, @($database1,$transform))
         $transformSummarySuccess = $database3.GetType().InvokeMember("CreateTransformSummaryInfo", "InvokeMethod", $Null, $database3, @($database1,$transform, $msiTransformErrorNone, $msiTransformValidationNone))
     }
     
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($installer)| Out-Null
    $database1 = $Null
    $database3 = $Null
        
    if ($transform -ne $Null) {
        $MyMultiLineTextBox.AppendText("`r`nResult: Created transform: $transform")
    }
    else {
     $transform = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.001' + [System.IO.Path]::GetExtension($MSI_Path)
      $MyMultiLineTextBox.AppendText("`r`nResult: Created ISM: $transform")
    }
    } 
    catch 
    {
      $MyMultiLineTextBox.AppendText("`r`nResult: Error creating Transform: {0}." -f $_)
    }
    
    $transform = $Null
}


function Modify-MSI-Create-MST
{
 [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
		[ValidateScript({$_ | Test-Path -PathType Leaf})]
        [string]$MSI_Path,
		[Parameter(Mandatory=$false)]
		[string]$MST_Path
    )
    try {


   $installer = New-Object -ComObject WindowsInstaller.Installer
    $WorkingDir = $MSI_Path.TrimEnd([System.IO.Path]::GetFileName($MSI_Path))

    #if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".msi")
    if (Test-Path $MSI_Path )
    {
     $MSI_file2 = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.001' + [System.IO.Path]::GetExtension($MSI_Path)
     Copy-Item $MSI_Path ($WorkingDir + $MSI_file2)
    }

    #$database1 = Invoke-Method $installer OpenDatabase  @($MSI_Path, $msiOpenDatabaseModeReadOnly)
    $database1 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @($MSI_Path, $msiOpenDatabaseModeReadOnly))
    #$database2 = Invoke-Method $installer OpenDatabase  @(($WorkingDir + $MSI_file2), $msiViewModifyUpdate)
    $database2 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @(($WorkingDir + $MSI_file2), $msiViewModifyUpdate))


    if (Test-Path $MST_Path )
    {
     $transform = [System.IO.Path]::GetFileNameWithoutExtension($MST_Path) + '.new' + [System.IO.Path]::GetExtension($MST_Path)
     $database2.GetType().InvokeMember("ApplyTransform", "InvokeMethod", $Null, $database2, @($MST_Path, 63))
    }
    Else {
         if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".msi")
            {
                $transform = $WorkingDir + [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.mst'
            }
         if ([System.IO.Path]::GetExtension($MSI_Path) -eq ".ism")
            {
                $transform = $Null
            }
    }
    [xml]$xml = Get-Content $XMLPath
    
    foreach ($prop in $xml.xml.Property){
        Change-MSIProperties $database2 $prop.Name $prop.Value
        $MyMultiLineTextBox.AppendText("`r`nAdded " + $prop.Name + " = " + $prop.Value + " Property")
    }
    
        
     Invoke-Method $database2 Commit
     $database2 = $Null


     if ($transform -ne $Null) {
         if (Test-Path ($WorkingDir + $MSI_file2) )
         {
          $MSI_file3 = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.002' + [System.IO.Path]::GetExtension($MSI_Path)
          Copy-Item ($WorkingDir + $MSI_file2) ($WorkingDir + $MSI_file3)
         }

         $database3 = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @(($WorkingDir + $MSI_file3), $msiOpenDatabaseModeReadOnly))

         #Invoke-Method $database3 GenerateTransform $database1 $transform
         $database3.GetType().InvokeMember("GenerateTransform", "InvokeMethod", $Null, $database3, @($database1,$transform))
         $transformSummarySuccess = $database3.GetType().InvokeMember("CreateTransformSummaryInfo", "InvokeMethod", $Null, $database3, @($database1,$transform, $msiTransformErrorNone, $msiTransformValidationNone))
     }
     
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($installer)
    $database1 = $Null
    $database3 = $Null
    $DelDB2 = $WorkingDir + $MSI_file2
    $DelDB3 = $WorkingDir + $MSI_file3
        
    if ($transform -ne $Null) {
        $MyMultiLineTextBox.AppendText("`r`nResult: Created transform: $transform")
    }
    else {
     $transform = [System.IO.Path]::GetFileNameWithoutExtension($MSI_Path) + '.001' + [System.IO.Path]::GetExtension($MSI_Path)
      $MyMultiLineTextBox.AppendText("`r`nResult: Created ISM: $transform")
    }
    } 
    catch 
    {
      $MyMultiLineTextBox.AppendText("`r`nResult: Error creating Transform: {0}." -f $_)
    }
    
    $transform = $Null
}

Create-Form