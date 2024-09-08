#
#	(c) ___, ___ (2024)
#
#	Version:
#		2024-09-08:	Erstellt.
#	Original:
#		
#	Verweise:
#		-/-
#	Verwendet:
#		-/-
#
param([String] $weldlog = $null, [String] $stylesheet = 'weldlog.xslt', [Boolean] $verbose = $false, [Boolean] $debug = $false) 
#
Clear-Host
#
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Xml\v4.0_4.0.0.0__b77a5c561934e089\System.Xml.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.IO\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.IO.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Drawing\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Drawing.dll"
Add-Type -Path "C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Windows.Forms\v4.0_4.0.0.0__b77a5c561934e089\System.Windows.Forms.dll"
#
#	To allow transformations including the 'document()' function.
#
[AppContext]::SetSwitch('Switch.System.Xml.AllowDefaultResolver', $true)
#
# -----------------------------------------------------------------------------------------------
#	Vielleicht kommt hier noch was rein ...
# -----------------------------------------------------------------------------------------------
#
function script:msxml([Object] $xsl, [Object] $xml, [String] $out = '', [System.Collections.Hashtable] $param = @{}, [Boolean] $verbose = $false, [Boolean] $debug = $false) {
	#
	# -----------------------------------------------------------------------------------------------
	#
	#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
	#	und entsprechend den Parametern $verbose und $debug neu setzen.
	#
	[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
	[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
	#
	if($verbose -eq $true -or  $debug -eq $true) {
		#
		Write-Verbose @"
	msxml(...): switched to verbose mode.
"@
		#
		$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
		#
		if($debug -eq $true) {
			Write-Verbose @"
	msxml(...): switched to debug mode.
"@
			$DebugPreference = [System.Management.Automation.ActionPreference]::Inquire
		} else {
			$DebugPreference = [System.Management.Automation.ActionPreference]::Continue
		}
	} else {
		$VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
		$DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
	}
	#
	# -----------------------------------------------------------------------------------------------
	#
	#
	[System.Xml.XmlUrlResolver] $local:xres = New-Object -typeName System.Xml.XmlUrlResolver
	[System.Xml.Xsl.XsltSettings] $local:xset = New-Object -typeName System.Xml.Xsl.XsltSettings
	#
	$null = $xset.set_EnableDocumentFunction($true)
	$null = $xset.set_EnableScript($true)
	#
	[System.Xml.Xsl.XslCompiledTransform] $script:xslt = New-Object -typeName System.Xml.Xsl.XslCompiledTransform
	[System.Xml.Xsl.XsltArgumentList] $script:xal = New-Object -typename System.Xml.Xsl.XsltArgumentList
	#
	if($verbose -eq $true) {
		$script:xal.AddParam("verbose", "", 1)
	} else {
		$script:xal.AddParam("verbose", "", 0)
	}
	if($debug -eq $true) {
		$script:xal.AddParam("debug", "", 1)
	} else {
		$script:xal.AddParam("debug", "", 0)
	}
	#
	if($null -ne $param) {
		[String[]] $local:val = $null
		foreach($val in $param.get_Keys()) {
			Write-Debug @"
[DEBUG] msxml::msxml(...): Adding XSLT parameter $val = $($param[$val])
"@
			$script:xal.AddParam($val, "", $($param[$val]))
		}
	}
	#
	if($xsl.GetType().Name -eq 'XmlDocument') {
		#
		$xslt.Load([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xsl.PsBase.InnerXml)), $local:xset, $local:xres)
		#
	} else {
		#
		[String] $local:fullXsl = $(Get-Item $xsl | Select-Object FullPath -ExpandProperty FullName)
		#
		if ([System.IO.File]::Exists($local:fullXsl) -eq $true) {
		#
		$xslt.Load($local:fullXsl, $local:xset, $local:xres)
		#
		} else {
			#
		Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (1).
"@
			#
		}
		#
	}
	#
	if ($out -ne '') {
		#
		[System.IO.FileStream] $local:fsm = New-Object System.IO.FileStream -ArgumentList @($out, 2)
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $script:xal, $local:fsm)
			#
		} else {
			#
			[String] $local:fullXml = $(Get-Item $xml | Select-Object FullPath -ExpandProperty FullName)
			#
			if ([System.IO.File]::Exists($local:fullXml) -eq $true) {
			#
			$xslt.Transform($local:fullXml, $script:xal, $local:fsm)
			#
			} else {
				Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (2).
"@
			}
			#
		}
		#
		$local:fsm.Close()
		$local:fsm.Dispose()
		#
		# -----------------------------------------------------------------------------------------------
		#
		#	Werte vor Start des Skripts wiederherstellen
		#
		$VerbosePreference = $script:saveVerbosePref
		$DebugPreference = $script:saveDebugPref
		#
		# -----------------------------------------------------------------------------------------------
		#
		return $null
		#
	} else {
		#	
		[System.IO.StringWriter] $local:osw = New-Object System.IO.StringWriter
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $script:xal, $local:osw)
			#
		} else {
			#
			[String] $local:fullXml = $(Get-Item $xml | Select-Object FullPath -ExpandProperty FullName)
			#
			if ([System.IO.File]::Exists($local:fullXml) -eq $true) {
			#
			$xslt.Transform($local:fullXml, $script:xal, $local:osw)
			#
			} else {
				#
				Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed (3).
"@
			}
			#
		}
		$local:osw.close()
		[String] $local:res = $local:osw.ToString()
		$local:osw.Dispose()
		#
		# -----------------------------------------------------------------------------------------------
		#
		#	Werte vor Start des Skripts wiederherstellen
		#
		$VerbosePreference = $script:saveVerbosePref
		$DebugPreference = $script:saveDebugPref
		#
		# -----------------------------------------------------------------------------------------------
		#
		return $local:res
	}
}
#
# -----------------------------------------------------------------------------------------------
#
function script:OpenFileDialog([String] $title = "Oups.", [String] $type = "all", [String] $defpath = "$pwd") {
	#
	[System.Windows.Forms.OpenFileDialog] $openFileDialog1 = new-object -typeName System.Windows.Forms.OpenFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'pptx'='Powerpoint Slideshow (*.pptx)|*.pptx'; 'pdf'='Portable Data Format (*.pdf)|*.pdf'; 'fo'='Formatting Objects (*.fo)|*.fo';'xmlcsv'='Excel 2003 XML or Excel CSV file (*.xml)|*.xml;*csv';'csv'='Comma separated values spread sheet (*.csv)|*.csv'; 'htm'='HTML Source (*.htm, *.html)|*.htm;*.html'; 'xslt'= 'XSLT Stylesheet (*.xsl, *.xslt)|*.xsl;*.xslt'; 'i6z'='IUCLID (*.i6z)|*.i6z'; 'all'='All files (*.*)|*.*'}
	#
	[String] $local:res = $null
	#
	$openFileDialog1.Reset()
	#
	if([System.IO.Directory]::Exists($defpath)) {
		Write-Debug @"
[Debug] framework::OpenFileDialog: Default path set to $defpath
"@
		$openFileDialog1.set_InitialDirectory($defpath)
	} else {
		Write-Debug @"
[Debug] framework::OpenFileDialog: Default path not set.
"@
	}
	$openFileDialog1.set_Filter($filters[$type])
	$openFileDialog1.set_FilterIndex(2)
	$openFileDialog1.set_Title($title)
	$openFileDialog1.set_AddExtension($true)
	$openFileDialog1.set_AutoUpgradeEnabled($true)
	$openFileDialog1.set_CheckFileExists($true)
	$openFileDialog1.set_CheckPathExists($true)
	$openFileDialog1.set_ShowHelp($false)
	$openFileDialog1.set_RestoreDirectory($true)
	$openFileDialog1.set_Multiselect($false)
	$openFileDialog1.set_ShowReadOnly($false)
	#
	if($openFileDialog1.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$res = $openFileDialog1.get_FileName()
        } else {
		$res = $null
	}
	#
	$openFileDialog1.Dispose()
	#
	return $res
}
#
# -----------------------------------------------------------------------------------------------
#
#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
#	und entsprechend den Parametern $verbose und $debug neu setzen.
#
[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
#
if($verbose -eq $true -or  $debug -eq $true) {
	#
	Write-Verbose @"
framework.ps1::main(...): switched to verbose mode.
"@
	#
	$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
	#
	if($debug -eq $true) {
		Write-Verbose @"
framework.ps1::main(...): switched to debug mode.
"@
		$DebugPreference = [System.Management.Automation.ActionPreference]::Inquire
		#
	} else {
		#
		$DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
		#
	}
} else {
	#
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
	#
}
#
# -----------------------------------------------------------------------------------------------
#
if (($verbose -or $debug) -ne $true) { Clear-Host }
#
if ($true) {
	#
    #
	if(-not $local:weldlog) { 
        #
        $local:weldlog = $(. script:OpenFileDialog -title "Log Datei zum Bearbeiten" -defpath $pwd -type 'txt')
        #
    }
	#
	if (Test-Path -Path $local:weldlog) {
		#
        # Set-Location -Path $DataDir
        #
		[System.Byte []] $local:raw = [System.IO.File]::ReadAllBytes("$local:weldlog")
        #
		[System.Text.StringBuilder] $local:clean = New-Object -TypeName System.Text.StringBuilder -ArgumentList 50
		#
		foreach($i in $local:raw) { 
			#
            if($i -eq 0x00) { $null = $local:clean.append(' ') }
            #
			if($i -eq 0x09 -or $i -eq 0x0A -or $i -eq 0x0D -or ($i -ge 0x20 -and $i -le 0x7E)) { $null = $local:clean.append([System.Text.Encoding]::UTF8.GetString($i)) }
			#
		}
		#
		# $local:clean.ToString() # | ConvertFrom-Csv -Delimiter ";" -$Header | Format-Table
		#
		[System.IO.StringReader] $local:srd = New-Object -TypeName System.IO.StringReader -ArgumentList $local:clean.ToString() 
		#
		$local:srd.ReadLine() | ConvertFrom-Csv -Delimiter ';' -Header 'Machine Type','Serial No.','Software Version','Unknown' | Format-Table
		#
		$local:srd.ReadToEnd() | ConvertFrom-Csv -Delimiter ';' -Header ('Tag','Number','Date','Time','Welder','Construction Site','Temperature','D5','Fitting','D7','Voltage Max','Vmin','Energy','D11','D12','Next Maintenance','Software Version','D15','D16','D17') | Out-GridView -Title "RODATA"
		#
	}
	#
}
#
#
# -----------------------------------------------------------------------------------------------
#
#	Werte vor Start des Skripts wiederherstellen
#
$VerbosePreference = $script:saveVerbosePref
$DebugPreference = $script:saveDebugPref
#
# -----------------------------------------------------------------------------------------------
#
