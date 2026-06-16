row = 2

Do While objSheet.Cells(row, 2).Value <> ""

    ' Read block
    plantCode = objSheet.Cells(row + 1, 2).Value
    fileName  = objSheet.Cells(row + 2, 2).Value
    fromDate  = objSheet.Cells(row + 3, 2).Value
    toDate    = objSheet.Cells(row + 4, 2).Value
    material  = objSheet.Cells(row + 5, 2).Value

    ' Run SAP
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = plantCode
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = material

    session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").text = fromDate
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").text = toDate

    session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Export
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Projects\Gypsum QC Report\SAP Dump\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName & ".xls"

    session.findById("wnd[1]/tbar[0]/btn[0]").press

    session.findById("wnd[0]/tbar[0]/btn[3]").press

    ' Move to next block (7 rows per plant)
    row = row + 7

Loop