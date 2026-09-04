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
        open "/Users/longmac/Library/CloudStorage/GoogleDrive-dangvietlong229@gmail.com/My Drive/Documents/Python Utilities/BCTT Tuan/output_reports/2026-07-31/MBS Dau Tu - BC Thi truong Tuan - 03.08.2026 - DRAFT.pptx"
        set pres to active presentation
        my setParagraphText(shape "TextBox 4" of slide 1 of pres, 2, "NHẬN ĐỊNH THỊ TRƯỜNG TUẦN 03/08/2026 – 07/08/2026")

        my setParagraphText(shape "TextBox 5" of slide 2 of pres, 1, "DIỄN BIẾN THỊ TRƯỜNG TUẦN 27/07 – 31/07/2026")

        my setParagraphText(shape "TextBox 5" of slide 2 of pres, 6, "CÁC SỰ KIỆN DIỄN RA TRONG TUẦN 03/08/2026 – 07/08/2026")

        my setParagraphText(shape "TextBox 5" of slide 2 of pres, 8, "NHẬN ĐỊNH XU HƯỚNG THỊ TRƯỜNG TUẦN 03/08/2026 – 07/08/2026")

        set content of text range of text frame of shape "Title 5" of slide 3 of pres to "DIỄN BIẾN THỊ TRƯỜNG TUẦN 27/07 – 31/07/2026"

        set content of text range of text frame of shape "Title 5" of slide 4 of pres to "DIỄN BIẾN THỊ TRƯỜNG TUẦN 27/07 – 31/07/2026"

        set content of text range of text frame of shape "Title 5" of slide 5 of pres to "DIỄN BIẾN THỊ TRƯỜNG TUẦN 27/07 – 31/07/2026"

        set content of text range of text frame of shape "Title 5" of slide 6 of pres to "DIỄN BIẾN DANH MỤC THEO DÕI TUẦN 27/07 – 31/07/2026"

        set content of text range of text frame of shape "Title 5" of slide 7 of pres to "CÁC SỰ KIỆN DIỄN RA TRONG TUẦN 03/08/2026 – 07/08/2026"

        set content of text range of text frame of shape "Title 5" of slide 8 of pres to "NHẬN ĐỊNH THỊ TRƯỜNG TUẦN 03/08/2026 – 07/08/2026"

        set targetTable to table object of shape "Table 1" of slide 4 of pres
my writeTable(targetTable, {{"Công nghệ Thông tin", "0,16%", "5,88%", "-4,69%"}, {"Tài nguyên Cơ bản", "-0,50%", "5,87%", "-3,88%"}, {"Hóa chất", "-0,24%", "4,11%", "-11,14%"}, {"Bất động sản", "-1,71%", "3,40%", "-4,14%"}, {"Bảo hiểm", "-1,33%", "3,18%", "-5,24%"}, {"Du lịch và Giải trí", "-0,05%", "2,83%", "-9,93%"}, {"Ngân hàng", "0,63%", "2,63%", "-6,84%"}, {"Dịch vụ tài chính", "-0,57%", "2,38%", "-8,99%"}, {"Xây dựng và Vật liệu", "0,32%", "2,22%", "-2,95%"}, {"Dầu khí", "-1,66%", "1,98%", "-0,99%"}, {"Hàng & Dịch vụ Công nghiệp", "0,82%", "1,61%", "-8,51%"}, {"Hàng cá nhân & Gia dụng", "-2,23%", "1,55%", "-5,18%"}, {"Điện, nước & xăng dầu khí đốt", "0,17%", "1,30%", "-6,60%"}, {"Thực phẩm và đồ uống", "-1,08%", "0,30%", "1,99%"}, {"Ô tô và phụ tùng", "0,13%", "-0,63%", "-3,05%"}, {"Y tế", "-0,75%", "-0,65%", "-4,29%"}, {"Bán lẻ", "0,70%", "-3,01%", "-5,60%"}, {"Truyền thông", "-1,41%", "-3,18%", "-10,24%"}}, 3, 1)

        set targetTable to table object of shape "Table 24" of slide 5 of pres
my writeTable(targetTable, {{"VIC", "214.100", "453.206.561.200", "3,16%"}, {"VNM", "60.900", "360.929.511.400", "49,48%"}, {"HPG", "21.700", "334.575.454.150", "21,55%"}, {"FPT", "67.100", "308.275.896.900", "27,06%"}, {"VCB", "59.300", "262.715.524.800", "20,06%"}, {"MSN", "66.100", "253.465.305.100", "24,10%"}, {"PNJ", "31.000", "249.209.847.100", "42,64%"}, {"MCH", "141.000", "168.397.080.500", "15,57%"}, {"LPB", "51.800", "167.693.860.000", "1,24%"}, {"BSR", "25.350", "116.760.717.850", "1,10%"}}, 2, 1)

        set targetTable to table object of shape "Table 25" of slide 5 of pres
my writeTable(targetTable, {{"VHM", "148.100", "-3.445.478.678.100", "6,93%"}, {"TCB", "28.950", "-376.569.390.550", "21,13%"}, {"VPB", "24.800", "-278.784.238.950", "23,35%"}, {"VIX", "13.000", "-243.482.009.400", "6,70%"}, {"ACB", "21.900", "-209.489.006.300", "24,34%"}, {"NVL", "13.000", "-207.581.893.150", "4,92%"}, {"GEX", "22.500", "-138.865.045.700", "6,50%"}, {"SSI", "23.550", "-113.275.752.300", "29,90%"}, {"HDB", "25.200", "-101.722.392.350", "21,56%"}, {"EIB", "17.800", "-51.607.960.000", "1,52%"}}, 2, 1)

        set targetTable to table object of shape "Table 16" of slide 5 of pres
my writeTable(targetTable, {{"Mã", "Giá", "Bán ròng (GT)", "Tỷ trọng bán của Khối Tự doanh"}, {"POW", "13.600", "-133.837.690.000", "8,32%"}, {"VIC", "214.100", "-115.044.020.000", "3,97%"}, {"LPB", "51.800", "-79.504.650.000", "9,96%"}, {"VNM", "60.900", "-78.434.940.000", "12,89%"}, {"FPT", "67.100", "-71.242.630.000", "9,85%"}, {"CTG", "30.800", "-51.252.425.000", "9,73%"}, {"VJC", "125.000", "-47.052.790.000", "6,68%"}, {"VPB", "24.800", "-39.589.604.000", "9,46%"}, {"FUEVFVND", "32.440", "-38.107.090.000", "36,40%"}, {"E1VFVN30", "33.700", "-35.818.329.000", "53,01%"}}, 1, 1)

        set targetTable to table object of shape "Table 23" of slide 5 of pres
my writeTable(targetTable, {{"Mã", "Giá", "Mua ròng (GT)", "Tỷ trọng mua của Khối Tự doanh"}, {"VPX", "25.400", "577.171.900.000", "45,55%"}, {"GMD", "77.000", "110.050.100.000", "12,27%"}, {"SSI", "23.550", "56.268.401.000", "4,99%"}, {"VHM", "148.100", "49.617.640.000", "2,80%"}, {"REE", "46.500", "38.159.400.000", "14,75%"}, {"VCB", "59.300", "33.983.990.000", "8,54%"}, {"NLG", "22.200", "32.946.787.000", "6,39%"}, {"FUEMAV30", "23.190", "27.172.518.000", "49,98%"}, {"HPG", "21.700", "21.850.054.000", "15,60%"}, {"TCX", "41.000", "19.679.225.000", "3,74%"}}, 1, 1)

        set targetTable to table object of shape "Table 2" of slide 6 of pres
my writeTable(targetTable, {{"31/07/2026", "ACB", "ACB: Báo cáo tình hình Quản trị công ty 6 tháng đầu năm 2026"}, {"30/07/2026", "ACB", "ACB: Nghị quyết HĐQT số 3411 ngày 28/07/2026"}, {"29/07/2026", "ACB", "ACB: Điều lệ công ty sửa đổi"}, {"", "", ""}}, 2, 1)

        set targetTable to table object of shape "Table 6" of slide 6 of pres
my writeTable(targetTable, {{"Tháng 5", "-17.370.861", "-17.860.342", "-41.923.991", "-38.085.620"}, {"Tháng 6", "4.293.512", "-16.017.491", "-32.182.660", "-32.762.639"}, {"Tháng 7", "-20.044.268", "-25.052.808", "-39.496.132", "-43.747.657"}, {"Từ đầu năm", "-228.820.850", "-53.116.009", "-103.383.042", "-108.872.152"}, {"Room NN còn lại", "6,4%", "1,5%", "6,7%", "1,4%"}}, 2, 1)

        set targetTable to table object of shape "Table 2" of slide 7 of pres
my writeTable(targetTable, {{"07/08/2026", "[United States] Non Farm Payrolls"}, {"09/08/2026", "[China] Inflation Rate YoY"}, {"", ""}, {"", ""}, {"", ""}}, 2, 1)

        set targetTable to table object of shape "Table 3" of slide 9 of pres
my writeTable(targetTable, {{"30", "FRT", "", "", "", "", "", "", "", "", "", ""}, {"49", "VHM", "", "", "", "", "", "", "", "", "", ""}, {"60", "VRE", "", "", "", "", "", "", "", "", "", ""}, {"63", "CTD", "", "", "", "", "", "", "", "", "", ""}, {"41", "VCB", "", "", "", "", "", "", "", "", "", ""}, {"36", "DGW", "", "", "", "", "", "", "", "", "", ""}, {"46", "PVT", "", "", "", "", "", "", "", "", "", ""}, {"70", "VCI", "", "", "", "", "", "", "", "", "", ""}, {"28", "VGC", "", "", "", "", "", "", "", "", "", ""}, {"55", "FPT", "", "", "", "", "", "", "", "", "", ""}, {"54", "CTG", "", "", "", "", "", "", "", "", "", ""}, {"39", "DGC", "", "", "", "", "", "", "", "", "", ""}, {"35", "BSR", "", "", "", "", "", "", "", "", "", ""}, {"47", "BID", "", "", "", "", "", "", "", "", "", ""}, {"6", "GMD", "", "", "", "", "", "", "", "", "", ""}, {"38", "GVR", "", "", "", "", "", "", "", "", "", ""}, {"29", "REE", "", "", "", "", "", "", "", "", "", ""}, {"9", "MCH", "", "", "", "", "", "", "", "", "", ""}, {"69", "VIX", "", "", "", "", "", "", "", "", "", ""}, {"16", "MWG", "", "", "", "", "", "", "", "", "", ""}, {"51", "HPG", "", "", "", "", "", "", "", "", "", ""}, {"59", "NKG", "", "", "", "", "", "", "", "", "", ""}, {"14", "DPM", "", "", "", "", "", "", "", "", "", ""}, {"21", "NLG", "", "", "", "", "", "", "", "", "", ""}, {"42", "DCM", "", "", "", "", "", "", "", "", "", ""}}, 4, 1)

        set targetTable to table object of shape "Table 8" of slide 10 of pres
my writeTable(targetTable, {{"23", "MSN", "", "", "", "", "", "", "", "", "", ""}, {"5", "VNM", "", "", "", "", "", "", "", "", "", ""}, {"66", "SSI", "", "", "", "", "", "", "", "", "", ""}, {"57", "GEX", "", "", "", "", "", "", "", "", "", ""}, {"18", "HSG", "", "", "", "", "", "", "", "", "", ""}, {"10", "GAS", "", "", "", "", "", "", "", "", "", ""}, {"11", "DXG", "", "", "", "", "", "", "", "", "", ""}, {"26", "POW", "", "", "", "", "", "", "", "", "", ""}, {"65", "IJC", "", "", "", "", "", "", "", "", "", ""}, {"34", "ACV", "", "", "", "", "", "", "", "", "", ""}, {"45", "EIB", "", "", "", "", "", "", "", "", "", ""}, {"7", "PC1", "", "", "", "", "", "", "", "", "", ""}, {"13", "DBC", "", "", "", "", "", "", "", "", "", ""}, {"37", "TCB", "", "", "", "", "", "", "", "", "", ""}, {"62", "ANV", "", "", "", "", "", "", "", "", "", ""}, {"64", "PNJ", "", "", "", "", "", "", "", "", "", ""}, {"43", "VIB", "", "", "", "", "", "", "", "", "", ""}, {"22", "VND", "", "", "", "", "", "", "", "", "", ""}, {"58", "HAH", "", "", "", "", "", "", "", "", "", ""}, {"24", "VIC", "", "", "", "", "", "", "", "", "", ""}, {"8", "NVL", "", "", "", "", "", "", "", "", "", ""}, {"67", "KDH", "", "", "", "", "", "", "", "", "", ""}, {"20", "SAB", "", "", "", "", "", "", "", "", "", ""}, {"44", "HHV", "", "", "", "", "", "", "", "", "", ""}, {"31", "KBC", "", "", "", "", "", "", "", "", "", ""}}, 4, 1)

        set targetTable to table object of shape "Table 4" of slide 11 of pres
my writeTable(targetTable, {{"2", "PVS", "", "", "", "", "", "", "", "", "", ""}, {"61", "VCG", "", "", "", "", "", "", "", "", "", ""}, {"25", "OCB", "", "", "", "", "", "", "", "", "", ""}, {"40", "VPB", "", "", "", "", "", "", "", "", "", ""}, {"50", "TPB", "", "", "", "", "", "", "", "", "", ""}, {"56", "HDG", "", "", "", "", "", "", "", "", "", ""}, {"33", "PLX", "", "", "", "", "", "", "", "", "", ""}, {"17", "BVB", "", "", "", "", "", "", "", "", "", ""}, {"52", "IDC", "", "", "", "", "", "", "", "", "", ""}, {"1", "HCM", "", "", "", "", "", "", "", "", "", ""}, {"19", "VHC", "", "", "", "", "", "", "", "", "", ""}, {"32", "DIG", "", "", "", "", "", "", "", "", "", ""}, {"12", "HDB", "", "", "", "", "", "", "", "", "", ""}, {"27", "TCH", "", "", "", "", "", "", "", "", "", ""}, {"4", "ACB", "", "", "", "", "", "", "", "", "", ""}, {"3", "STB", "", "", "", "", "", "", "", "", "", ""}, {"53", "CII", "", "", "", "", "", "", "", "", "", ""}, {"48", "PDR", "", "", "", "", "", "", "", "", "", ""}, {"15", "PVD", "", "", "", "", "", "", "", "", "", ""}, {"68", "BCM", "", "", "", "", "", "", "", "", "", ""}}, 4, 1)

        set targetTable to table object of shape "Table 3" of slide 9 of pres
my writeTable(targetTable, {{"31/07/2026"}}, 2, 2)

        set targetTable to table object of shape "Table 8" of slide 10 of pres
my writeTable(targetTable, {{"31/07/2026"}}, 2, 2)

        set targetTable to table object of shape "Table 4" of slide 11 of pres
my writeTable(targetTable, {{"31/07/2026"}}, 2, 2)
        save pres
        close pres saving yes
    end tell
end run
