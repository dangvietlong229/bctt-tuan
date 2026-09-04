on writeTable(targetTable, tableData, startRow, startColumn)
    tell application "Microsoft PowerPoint"
        repeat with rowOffset from 1 to count of tableData
            set rowData to item rowOffset of tableData
            repeat with columnOffset from 1 to count of rowData
                set targetCell to get cell from targetTable row (startRow + rowOffset - 1) column (startColumn + columnOffset - 1)
                set content of text range of text frame of shape of targetCell to item columnOffset of rowData
            end repeat
        end repeat
    end tell
end writeTable

on setParagraphText(targetShape, paragraphNumber, newText)
    tell application "Microsoft PowerPoint"
        set targetRange to text range of text frame of targetShape
        set content of paragraph paragraphNumber of targetRange to newText
    end tell
end setParagraphText

on run argv
    tell application "Microsoft PowerPoint"
        open "/Users/longmac/Library/CloudStorage/GoogleDrive-dangvietlong229@gmail.com/My Drive/Documents/Python Utilities/BCTT Tuan/output_reports/2026-07-31/MBS Dau Tu - BC Thi truong Tuan - 03.08.2026.pptx"
        set pres to active presentation
        
        save pres
        close pres saving yes
    end tell
end run
