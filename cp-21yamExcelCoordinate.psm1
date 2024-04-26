function Invoke-ExcelCellCoordinate([string]$inpout) {
    $match = splitString($inpout)

    $column = $match.Groups["Column"].Value
    $row = $match.Groups["Row"].Value

    $CharArray = reverseCharArray($column)

    $finalVal = convertToInt($CharArray)

    Write-Output "X: $finalVal      Y: $row"
}

function splitString([string]$inpout){
    $inpout = $inpout.ToUpper()

    $pattern = [regex]"(?<Column>[A-Z]+)(?<Row>[0-9]+)"
    $match = $pattern.Match($inpout)
    return $match
}

function reverseCharArray([string]$column){
    $charArray = $column.ToCharArray()
    [array]::Reverse($charArray)
    return $charArray
}

function convertToInt([array]$charArray)
{
    $pow = 0
    $finalVal = 0

    foreach ($char in $charArray) {
        $ascii = ([int][char]$char) - 64
        $finalVal += $ascii * [Math]::Pow(26, $pow)
        $pow++
    }
    return $finalVal
}
