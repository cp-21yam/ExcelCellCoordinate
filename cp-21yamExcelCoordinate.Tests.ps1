Describe 'Excel Cell Coordinate Functions' {
    BeforeAll {
        . $PSCommandPath.Replace('.Tests.ps1', '.psd1')
    }

    # Test the main function: Invoke-ExcelCellCoordinate
    Describe 'Invoke-ExcelCellCoordinate' {
        It 'Correctly parses and computes the Excel coordinate for "A1"' {
            Mock splitString { return @{ Groups = @{ "Column" = "A"; "Row" = 1 } } }
            Mock reverseCharArray { return @('A') }
            Mock convertToInt { return 1 }

            Invoke-ExcelCellCoordinate -inpout 'A1' | Should -BeExactly 'X: 1      Y: 1'
        }

        It 'Correctly parses and computes the Excel coordinate for "CZ100"' {
            Mock splitString { return @{ Groups = @{ "Column" = "CZ"; "Row" = 100 } } }
            Mock reverseCharArray { return @('Z', 'C') }
            Mock convertToInt { return 104 }

            Invoke-ExcelCellCoordinate -inpout 'CZ100' | Should -BeExactly 'X: 104      Y: 100'
        }
    }

    # Test the splitString function
    Describe 'splitString' {
        It 'Correctly splits input string "AB12" into components' {
            $result = splitString -inpout 'AB12'
            $result.Groups["Column"].Value | Should -BeExactly 'AB'
            $result.Groups["Row"].Value | Should -Be 12
        }

        It 'Handles lowercase input by converting to uppercase' {
            $result = splitString -inpout 'ab12'
            $result.Groups["Column"].Value | Should -BeExactly 'AB'
        }

        It 'Throws an error for invalid input' {
            { splitString -inpout '123' } | Should -Throw -ExceptionType System.Management.Automation.RuntimeException
        }
    }

    # Test the reverseCharArray function
    Describe 'reverseCharArray' {
        It 'Correctly reverses the character array from "ABC"' {
            $result = reverseCharArray -column 'ABC'
            $result | Should -Be @('C', 'B', 'A')
        }
    }

    # Test the convertToInt function
    Describe 'convertToInt' {
        It 'Converts column "A" to integer 1' {
            $result = convertToInt -charArray @('A')
            $result | Should -Be 1
        }

        It 'Converts column "BA" to integer 53' {
            $result = convertToInt -charArray @('A', 'B')
            $result | Should -Be 53
        }
    }
}