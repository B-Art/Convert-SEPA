#!/usr/bin/env powershell
# Set-StrictMode -Version Latest
function Convert-Sepa {
  <#
      .SYNOPSIS
      The purpose of "Convert-Sepa" is to sanitise a downloaded CSV-file from your bank-account's website.

      .DESCRIPTION
      The fields Omschrijving is reused for SEPA input, which messes up a lot. This Script will create new columns according to the extra fields used for SEPA. ('Naam: ', 'Omschrijving: ', 'IBAN: ', 'BIC: ', 'ID begunstigde: ', 'SEPA ID machtiging: ', 'Kenmerk: ', 'Machtiging ID: ', 'Incassant ID: ')

      .PARAMETER filename
      Parameter -filename will be used as a sourcename and .csv will be replace by .iban.csv for destination filename.

      .EXAMPLE
      Convert-Sepa -filename uses a csv file as import and creates a iban.csv as export
 
      .NOTES
 
      .LINK
 
      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>

  [cmdletbinding()]
  param
  (
    [Parameter(
      Mandatory = $true,
      Position = 1,
      ParameterSetName = "ParameterSetName",
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true,
      HelpMessage = 'Give a csv filename which you downloaded from your bank.')
    ]
      [ValidateNotNullOrEmpty()]
      [SupportsWildcards()]
      [string[]]
      $filename
    )

  $sepa = @(
    'Naam'
    'Omschrijving'
    'IBAN'
    'BIC'
    'ID begunstigde'
    'SEPA ID machtiging'
    'Kenmerk'
    'Machtiging ID'
    'Incassant ID'
  )

  $sepatxt = 'Kenmerk'

  $filedest = $filename.Replace('.csv', '.iban.csv')
  $SepaFile = Get-Content -Path $filename -TotalCount 1
  Write-Verbose($SepaFile)
  if ($SepaFile.Split(';').Count -gt 1)
  {$split = ';'}
  else
  {$split = ','}
  Write-Verbose('Delimiter found: {0}' -f $split)
  $SepaFile = Import-Csv -Path $filename -Delimiter $split
  $SepaFile | Add-Member -MemberType 'NoteProperty' -Name 'Omschrijving' -Value ''
  $SepaFile | Add-Member -MemberType 'NoteProperty' -Name 'Kenmerk' -Value ''
  # for ($j =0; $j -lt 1; $j++)
  $SepaFile.foreach{
    $Mededelingen = $_.Mededelingen
    # Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
    # $host.PrivateData.VerboseForegroundColor = 'Magenta'
    Write-Verbose("[ORG]:`r`n`tMededelingen: {0}`n`r`tOmschrijving: {1}" -f $_.Mededelingen, $_.Omschrijving)
    If ($Mededelingen.Contains($sepa[0] + ': ')) {
      foreach ($element in $sepa) {
        if (($Mededelingen -split $element + ': ').Count -gt 1) {
          $Mededelingen = $Mededelingen -Replace $element, ('{}' + $element)
        }
      }
      $Mededelingen = '{}' + $Mededelingen.Replace(' {}', '{}') + '{}'
      foreach ($element in $sepa) {
        # Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
        $host.PrivateData.VerboseForegroundColor = 'White'
        if (($Mededelingen -split '{}' + $element + ': ').Count -gt 1) {
          if ($element -in $sepatxt) {
                $mark = '#'
            $host.PrivateData.VerboseForegroundColor = 'Green'
            #Write-Verbose $mark$element
          } else {
            $mark = ''
            $host.PrivateData.VerboseForegroundColor = 'Cyan'
            #Write-Verbose $element
          }
          Add-Member -MemberType 'NoteProperty' -InputObject $_ `
            -Name $element -Value ($mark + (($Mededelingen -split '{}' + $element + ': ')[1] -split '{}')[0]) -Force
          $Mededelingen = $Mededelingen.Replace('{}' + $element + ': ' + $mark + $_.($element),'')
        }
      }
      $_.Mededelingen = ($_.Omschrijving.replace('{}','')).Trim()
      $_.Omschrijving = ($Mededelingen.replace('{}','') -replace 'Kenmerk:',' ' -replace 'Omschrijving:',' ' -replace 'Kenmerk','' -replace 'Omschrijving','' -replace '  ',' ').Trim()
      if (($_.Kenmerk).length -gt 1) {
        $_.Omschrijving = ($_.Omschrijving -replace ($_.Kenmerk).substring(1), '').Trim()
      }
    }
    Write-Verbose("[CHG]:`n`r`tMededelingen: {0}`n`r`tOmschrijving: {1}`n`r`tKenmerk: {2}" -f $_.Mededelingen, $_.Omschrijving, $_.Kenmerk)
    $host.PrivateData.VerboseForegroundColor = 'Yellow'
  }
  $SepaFile | Export-Csv -Path $filedest -Delimiter $split -NoTypeInformation
}
#   [string](0..33|%{[char][int](46+("686552495351636652556262185355647068516270555358646562655775 0645570").substring(($_*2),2))})-replace " "
# Convert-Sepa