#Column Count Helper Function for ranges
function Get-ExcelColLetter {
    param (
        [int]$columnNumber
    )

    [char]$columnNumHelper = [char](([int][char]("A"))+$columnNumber)

    Return $columnNumHelper
}
    

<#Get function that will get the first row of headers of a given excel file
This can be a component for a larger tool that reduce data input errors#>

function Get-ExcelHeaderRows {
    param (


         # Validate that filename ends with .[extension]
        [Parameter(Position = 0,Mandatory)][ValidatePattern(".+\.\w{1,8}")][string]$fileName,

        #Provide the sheet Name that you want to interact with. Default is "Sheet1".
        [Parameter(Position = 1)][string]$sheetName = "Sheet1",

        #Set Starting Cell to A1 by default
        [Parameter(Position = 2)][string]$startingCell = "A1",

        #Explicitly validate on string types
        [Parameter(Position = 3)][string]$folderPath


    )
    #Import the ReadLine Module
    Import-Module PSReadLine

    #Remove double quotes if read as literals
    $fileName = $fileName.Replace("`"","")
    Write-Host ("Your inputs are `n FolderPath: {0}`n fileName: {1} " -f $folderPath, $fileName )

<#Welcome Statement:
This portion begins the main part of the code that will interact with excel and output the list of headers#>
   
    #Create the excel com object
    $excelObject = New-Object -ComObject Excel.Application

    #if the folderPath Var is not used only use the fileName var
   if([string]::IsNullOrEmpty($folderPath))
   {$excelWB = $excelObject.workbooks.open($fileName)}
   else 
   {$excelWB = $excelObject.workbooks.open("$folderPath$fileName")}

    #init excel worksheet
    $workingSheet = $excelWB.worksheets($sheetName)
    #Get the range column count from the worksheet
    $columnCount = $workingSheet.usedRange.Columns.count
    #Convert Column count to Alphabetic reference
    [char]$columnAlpahbeticRef = [char](([int][char]("A")) + $columnCount-1)
    
    <#if a alternate starting cell was provided use split on the alpha char
     to get the digit of the row that is provided#>
     if ([string]::IsNullOrEmpty($startingCell) -or $startingCell -eq "A1") 
    {
        $rowAlpha = "A"
        $rowNum = "1"

    }
     else
     {
        $startingCell -match "\D+"
        [string]$rowAlpha = $Matches[0]
        $startingCell -match "\d+"
        [string]$rowNum = $Matches[0]
    }
    #Get the excel Range
    $excelHeaderRange = $workingSheet.usedRange.Range("$rowAlpha$rowNum","$columnAlpahbeticRef$rowNum")

    $excelHeaderRange = $excelHeaderRange.value2

    $headerArray = $excelHeaderRange -split "`n"
    $headerArrayOld = $headerArray
    
    $reservedCharacters = "<>%&)?/"
    $loopcounter = 0
    foreach ($header in $headerArray) {
        <# Validate if any of the headers contain reserved chars #>
        if($header -match "[$reservedCharacters]")
        {
            $loopcounter += 1
            $newHeaderName = $header -replace $Matches[0],"_"
            Write-Host "$header contains one of the reserved characters($reservedCharacters)`
            `nIf you would like to replace it with char('_') type yes and hit enter."
            $userResponseHeaders = Read-Host "Enter 'yes' without quotes to update special char in col header
             $header to $newHeaderName"
            if($userResponseHeaders -eq "yes")
                {
                    #Update the Value in the excel document
                    $workingSheet.Range([string](Get-ExcelColLetter -columnNumber `
                     $HeaderArray.indexof($header)+1)+[string]($rowNum)).value2 = $newHeaderName

                    $headerArray = $headerArray.replace($header,$newHeaderName)
                    Write-Host "The name has been updated in excel, new name = $($headerArray[$headerArray.indexof($newHeaderName)])"

                    # Save and close the workbook
                    $excelWB.Save()
                    $excelWB.Close()
                }
             
             
            
        }
    }

     if ($loopcounter -eq 0) 
     { 
        Write-Host "None of the columns has any characters that are reserved"
        $outputList=$excelHeaderRange
    } 
    else 
    {
        Write-Host "One or more column headers has/had reserved characters in their names`
         Old values are $(foreach($col in $headerArrayOld) {"$col`n"}) New values are`
         $(foreach($col in $headerArray) {"$col`n"})"

        $outputList=$headerArray

    }
    Return $outputList
# Quit the Excel application
$excelWB.Close()
$excelObject.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workingsheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelWB) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObject) | Out-Null

# Force garbage collection to finalize cleanup
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
Stop-Process -Name *Excel*

Start-Sleep -Seconds 120
}

#Call the Main Function to Run the script
Get-ExcelHeaderRows
Read-Host "Hit Enter to close the window"