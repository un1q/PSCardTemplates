$InkscapePath = 'C:\Program Files\Inkscape\bin\inkscape.com'

#####change working directory
function here { pwd | % { [IO.Directory]::SetCurrentDirectory($_.path) } }

####generate png from svg with Inkscape
Function Generate-Png-From-Svg-Content {
	Param(
	[parameter(Mandatory=$true)][String]$Content,
		[parameter(Mandatory=$true)][Int32]$CardWidth,
		[parameter(Mandatory=$true)][String]$OutputFilePath,
		[parameter(Mandatory=$false)][String]$TempFile = "temp.svg"
	)
	Process {
		Write-Host "Creating file $OutputFilePath"
		$Content | Out-File $TempFile
		& $InkscapePath --export-filename=$OutputFilePath --export-width=$CardWidth $TempFile
	}
}

####Create-Deck, but for one json definition (json contains more then one definition), for example front only.
Function Run-Def {
	Param (
		$Width,
		$Def,
		[switch]$ClearOutputDir,
		[switch]$SaveSVG
	)
	Process {
		if ($ClearOutputDir) {
			Clear-Def $Def
		}
		New-Item -ItemType Directory -Force -Path $Def.outputDir
		Write-Host "Template : $($Def.template)"
		$Template = get-content $Def.template -raw
		if ($Def.CSV) {
			if (Test-Path $Def.CSV) {
				$CSVPath = $Def.CSV
			} else {
				$CSVPath = (Split-Path $JsonPath) + "\" + $Def.CSV
			}
			Write-Host "CSVPath : $CSVPath"
			$CSV = Get-Content $CSVPath |  ConvertFrom-Csv -Delimiter "`t"
			$Id = 0
			$CSV | % {
				$Row = $_
				$Count = 1
				$SVGContent = Create-Card -Replaces $Def.replace -Template $Template -Row $Row
				if ($Def.countColumn -ne $null) {
					if ($Row.($Def.countColumn)) {
						$Count = $Row.($Def.countColumn)
					} else {
						Write-Error "Count column defined in json, but not in CSV"
					}
				}
				$OutputFilePath = Generate-Card-Path -Def $Def -Count $Count -Id $Id
				Write-Host "OutputFilePath: $OutputFilePath"
				$Id = $Id+1
				Generate-Png-From-Svg-Content -Content $SvgContent -CardWidth $Width -OutputFilePath $OutputFilePath
				if ($SaveSVG) {
					$SvgContent | Out-File ($OutputFilePath+".svg")
				}
			}
		} else {
			Write-Host "No CSV defined"
			$SVGContent = Create-Card -Replaces $Def.replace -Template $Template
			$OutputFilePath = Generate-Card-Path -Def $Def -Count "1" -Id ""
			Generate-Png-From-Svg-Content -Content $SvgContent -CardWidth $Width -OutputFilePath $OutputFilePath
			if ($SaveSVG) {
				$SvgContent | Out-File ($OutputFilePath+".svg")
			}
		}
		if ($Def.atlasPath -ne $null) {
			Concatenate-Cards -OutputPath $Def.atlasPath -DirectoryPath $Def.outputDir
		}
	}
}

#### Remove directories with old generated files
Function Clear-Def {
	Param($Def)
	Process {
		here 
		if (Test-Path $Def.outputDir) {
			$Path = Join-Path -Path $Def.outputDir -ChildPath "*.png"
			Write-Host "Removing everything from $Path"
			Remove-Item $Path -Confirm
		}
		if ($Def.atlasPath -ne $null) {
			if (Test-Path ($Def.atlasPath+"*.png")) {
				Write-Host "Removing everything from $($Def.atlasPath)*.png"
				Remove-Item ($Def.atlasPath+"*.png") -Confirm
			}
		}
	}
}

####Create Deck (atlas compatible with Tabletop Simulator) from json and csv
Function Create-Deck {
	Param (
		[parameter(Mandatory=$true)][string]$JsonPath,
		[switch]$SaveSVG
	)
	Process {
		$Json = Get-Content $JsonPath -raw | ConvertFrom-Json
		$Width = $Json.width
		if ($Json.front) { Clear-Def $Json.front }
		if ($Json.back) { Clear-Def $Json.back }
		if ($Json.multi) { $Json.multi | %{Clear-Def $_} }
		if ($Json.front) {
			Write-Host "**FRONT**"
			Run-Def -Def $Json.front $Width Width -SaveSVG:$SaveSVG
		}
		if ($Json.back) {
			Write-Host "**BACK**"
			Run-Def -Def $Json.back $Width Width -SaveSVG:$SaveSVG
		}
		if ($Json.multi) {
			Write-Host "**MULTI**"
			$Json.multi | % { Run-Def -Def $_ $Width Width -SaveSVG:$SaveSVG }
		}
	}
}

####Generate path for png with one card
Function Generate-Card-Path {
	Param(
		$Def,
		[string]$Count,
		[string]$Id
	)
	Process {
		$OutputDirectory = $Def.outputDir
		$Name = $Def.template -replace '(.*\\)?([^\\+\.]+)(\.[^\\]+)?','$2'
		if (-not (Test-Path $OutputDirectory)) {
			New-Item -ItemType Directory -Force $OutputDirectory
		}
		return "${OutputDirectory}\x${Count}_${Name}_${Id}.png"
	}
}

####Create png with one card from template and set of replace definitions (see json examples)
Function Create-Card {
	Param(
		[parameter(Mandatory=$true)]$Replaces,
		[parameter(Mandatory=$false)]$Row,
		[parameter(Mandatory=$true)][string]$Template
	)
	Process {
		$Result = $Template
		$Replaces | % {
			$Replace = $_
			if ($Replace.value) {
				$Value = $Replace.value
			} elseif ($Replace.column -and $Row) {
				$Value = $Row.($Replace.column)
			} else {
				Write-Error "Replace should have specified value (.replace[i].value) or both CSV path and csv column name (.replace[i].column): $Replacewers"
				return;
			}
			if ($Replace.dictionary) {
				if ($Replace.dictionary.$Value -eq $null -and -not $Replace.dictionaryNoWarningFlag) {
					Write-Warning "Value not found in dictionary (Value: $Value) (Dictionary: $($Replace.dictionary))"
				}
				$Value = $Replace.dictionary.$Value
			}
			if ($Value -is [array]) {
				if ((-not ($Replace.search -is [array])) -or ($Replace.search.length -ne $Value.length)) {
					Write-Error "Value is array, so search also needs to be an array with the same length($Replace)"
					return
				}
				0..($Value.Length-1) | %{
					$Result = Replace $Result $Replace.search[$_] $Value[$_]
				}
			} else {
				$Replace.search | %{ $Result = Replace $Result $_ $Value }
			}
		}
		return $Result
	}
}

####Replace $Pattern in $Text with $Value. | is replaced to new line.
Function Replace ([string]$Text, [string] $Pattern, [string]$Value) {
	#Write-Host "Replacing", $Pattern, $Value
	if ($Value -ne $null) {
		return $Text -replace $Pattern, ($Value -replace "\|","`n")
	}
}

####Merges all cards into one atlas (or more, if there is more then 70 cards)
Function Concatenate-Cards {
	Param(
		[String]$DirectoryPath = ".",
		[String]$OutputPath = ".",
		[Int]$AtlasWidth = 4096,
		[Int]$AtlasHeight = 4096,
		[Int]$MaxCountX = 10,
		[Int]$MaxCountY = 7
	)
	Begin {
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		here
	}
	Process {
		Write-Host "Making atlas from all png in $DirectoryPath"
		$files = ls "${DirectoryPath}\x*.png"
		$bmp = [System.Drawing.Bitmap]::new($AtlasWidth,$AtlasHeight)
		$graph = [System.Drawing.Graphics]::FromImage($bmp)
		$x = 0
		$y = 0
		$Sum = 0
		$AtlasId = 1;
		$CountX = 0;
		$CountY = 0;
		$files | %{
			$FileName = $_
			$Count = 1
			if ($FileName -match '.*\\x([0-9]+).*') {
				$Count = $Matches[1]
			}
			$CardImg = [System.Drawing.Image]::FromFile($FileName)
			$CardImg.SetResolution($graph.DpiX, $graph.DpiY)
			for ($i = 0; $i -lt $Count; $i++) {
				if ($CountY -eq $MaxCountY) {
					$Path = "$($OutputPath)_$($Sum)_$($AtlasId).png"
					Write-Host "Saving part of atlas to $Path"
					$bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
					$graph.Dispose()
					$bmp.Dispose()
					$bmp = [System.Drawing.Bitmap]::new($AtlasWidth,$AtlasHeight)
					$graph = [System.Drawing.Graphics]::FromImage($bmp)
					$AtlasId = $AtlasId + 1
					$x = 0
					$y = 0
					$Sum = 0
					$CountX = 0;
					$CountY = 0;
				}
				Write-Host "$x x $y"
				$graph.DrawImage($CardImg, $x, $y)
				$x = $x + $AtlasWidth/$MaxCountX #$CardImg.Width
				$CountX = $CountX + 1
				if ($CountX -eq $MaxCountX) {
					$x = 0
					$y = $y + $AtlasHeight/$MaxCountY #$CardImg.Height
					$CountX = 0
					$CountY = $CountY + 1
				}
				$Sum = $Sum + 1
			}
			$CardImg.Dispose()
		}
		$Path = "$($OutputPath)_$($Sum)_$($AtlasId).png"
		Write-Host "Saving atlas to $Path"
		$bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
		$graph.Dispose()
		$bmp.Dispose()
	}
}