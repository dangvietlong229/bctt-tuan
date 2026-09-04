on replaceText(findText, replacementText, sourceText)
    set AppleScript's text item delimiters to findText
    set textItems to every text item of sourceText
    set AppleScript's text item delimiters to replacementText
    set resultText to textItems as text
    set AppleScript's text item delimiters to ""
    return resultText
end replaceText

on cleanText(sourceValue)
    if sourceValue is missing value then return ""
    set cleaned to sourceValue as text
    set cleaned to my replaceText(return, " ", cleaned)
    set cleaned to my replaceText(linefeed, " ", cleaned)
    set cleaned to my replaceText(tab, " ", cleaned)
    set cleaned to my replaceText((character id 30), " ", cleaned)
    set cleaned to my replaceText((character id 31), " ", cleaned)
    return cleaned
end cleanText

on run argv
    set workbookPath to item 1 of argv
    set worksheetName to item 2 of argv
    set startRow to (item 3 of argv) as integer
    set startColumn to (item 4 of argv) as integer
    set rowCount to (item 5 of argv) as integer
    set columnCount to (item 6 of argv) as integer
    set outputMode to item 7 of argv
    set unitSeparator to character id 31
    set recordSeparator to character id 30
    set outputText to ""

    tell application "Microsoft Excel"
        open workbook workbook file name workbookPath update links never read only true
        set wb to active workbook
        calculate full rebuild
        set ws to worksheet worksheetName of wb
        activate object ws

        repeat with rowOffset from 0 to (rowCount - 1)
            repeat with columnOffset from 0 to (columnCount - 1)
                set targetCell to cell (startColumn + columnOffset) of row (startRow + rowOffset) of ws
                try
                    if outputMode is "raw" then
                        set cellValue to value2 of targetCell
                    else
                        set cellValue to string value of targetCell
                    end if
                on error
                    set cellValue to ""
                end try
                set outputText to outputText & my cleanText(cellValue)
                if columnOffset < (columnCount - 1) then set outputText to outputText & unitSeparator
            end repeat
            if rowOffset < (rowCount - 1) then set outputText to outputText & recordSeparator
        end repeat

        close wb saving no
    end tell

    return outputText
end run
