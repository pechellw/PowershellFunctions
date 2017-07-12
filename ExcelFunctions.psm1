function Import-Excel
{
    <#
    .SYNOPSIS
    Function to import an excel sheet into either a hash table or an array. Requires Microsoft Excel to be installed
    on the machine.
    .DESCRIPTION
    Using Excel APIs this function can import a worksheet from any file that can be opened in Excel. It assumes that the 
    sheet is formatted as a data table with unique column headers. Any columns without headers will be skipped from the import.
    If importing as a hashtable the function assumes that there is no header row. Duplicate column names or keys will be skipped 
    and only the first will be imported as fields must be unique.

    .PARAMETER Path
    Path to file name.

    .PARAMETER Password
    For password protected files.

    .PARAMETER SheetName
    The specific name of the sheet to be imported

    .PARAMETER Type
    Data format to import. Either HashTable (Key,Value pairs) or Array.

    .PARAMETER SkipRows
    Rows to be skipped. Can be entered as a comma separated list e.g. 1,2,(4..6). Ranges entered in brackets. Useful
    if worksheets have title rows that spoil the import.

    .PARAMETER SkipColumns
    Columns to be skipped. Can be entered as a comma separated list e.g. 1,2,(4..6). Ranges entered in brackets.

    .EXAMPLE
        #Import into an array using the first sheet of data.
        Import-Excel -Type Array -Path [Path to File Name]

    .EXAMPLE
        #Import into an array from a specific sheet.
        Import-Excel -Type Array -Path [Path to File Name] -SheetName [SheetName]

    .EXAMPLE
        #Import into a hashtable
        Import-Excel -Type HashTable -Path [Path to File Name] -SheetName [SheetName] -KeyColumn [Key Column number] -ValueColumn [Value Column number]

    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True)]
        [string]$Path=$(throw "-Path is required!"),
        [switch]$PasswordProtected,
        [string]$SheetName,
        [Parameter(Mandatory=$True,
        HelpMessage="Please choose a data type (HashTable or Array)")]
        [ValidateSet('Array','HashTable')]
        [string]$Type=$(throw "-Type is required"),
        [int]$KeyColumn = 1,
        [int]$ValueColumn = 2,
        [int[]]$SkipRows,
        [int[]]$SkipColumns,
        [int]$TopRow
        )


    begin
    {
        try
        {
            $xl = New-Object -ComObject "Excel.Application"
        }
        catch
        {
            throw "-Unable to initialize Excel, check the application is installed correctly and try again."
        }
        $xl.Visible = $false
        if($Path -match '.csv$')
        {
            $SheetName = $null
        }
        if($PasswordProtected)
        {   
            $response = Read-Host "Password:" -AsSecureString
            $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($response))
            $wb = $xl.Workbooks.Open($Path,0,0,5,$Password)
        }
        else
        {
        $wb = $xl.Workbooks.Open($Path)
        }
        if($SheetName)
        {
            try
            {
                $sheet = $wb.Sheets($SheetName)
            }
            catch
            {
                $xl.Application.Quit()
                throw "-Invalid Sheet Name: `"$SheetName`""
            }
        }
        else
        {
            $sheet = $wb.ActiveSheet
        }
        Write-Host "Sheet: $($sheet.Name)"
        $rows = $sheet.UsedRange.Rows.Count
        Write-Host "Total Rows: $rows"
        $columns = $sheet.UsedRange.Columns.Count
        Write-Host "Total Columns: $columns"
    }

    process
    {
        switch($Type)
        {
            'HashTable'
            {
                [HashTable]$XlData = @{}
                if(!($topRow))
                {
                    $topRow = 1
                    while($topRow -in $SkipRows)
                    {
                        $topRow++
                    }
                }
                
                Write-Host "TopRow is $topRow"
                for($row = $topRow+1;$row -le $rows;$row++)
                {
                    if(($sheet.Cells($row,$KeyColumn).Value()) -and ($XlData.Keys -notcontains $sheet.Cells($row,$KeyColumn).Value()) -and ($row -notin $SkipRows))
                    {
                        $XlData.Add($sheet.Cells($row,$KeyColumn).Value(),$sheet.Cells($row,$ValueColumn).Value())
                    }
                }
            }
            'Array'
            {
                [Array]$XlData = @()
                if(!($topRow))
                {
                    $topRow = 1
                    while($topRow -in $SkipRows)
                    {
                        $topRow++
                    }
                }
                Write-Host "TopRow is $topRow"
                for($row = $topRow+1;$row -le $rows;$row++)
                {
                    $item = New-Object psobject
                    for($column=1;$column -le $columns;$column++)
                    {
                        if(($sheet.Cells($topRow,$column).Value()) -and (!($item.($sheet.Cells($topRow,$column).Value()))) -and ($row -notin $SkipRows) -and ($column -notin $SkipColumns))
                        {
                            $item|Add-Member -NotePropertyName $sheet.Cells($topRow,$column).Value() -NotePropertyValue $sheet.Cells($row,$column).Value()
                        }
                    }
                    $XlData += $item
                }
            }
        }
    }

    end
    {
        $wb.Close()|Out-Null
        $xl.Application.Quit()|Out-Null
        $XlData
    }
}