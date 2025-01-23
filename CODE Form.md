**BỘ CODE ACCESS 2003**

**CODE Form**

`COUNT, LOOKUP, MIN, MAX
DLookup("NAMDT", "MACDINH", "ID=1")
DLookup("NAMCS", "MACDINH", "ID=1")
DCount("MATV", "DS_TV", "DS_TV.SOSO ='" & SOSO & "'")
DLookUp("TUOILD","TUOILD","KY.ID=1")
DMax("STT", "DS_TV", "DS_TV.SOSO like '" & Form_frm_NHAP.SOSO & "'")
DLookup("NHAPBS", "MACDINH", "ID=1") = -1
DLookup("LOAINHAP", "MACDINH", "ID=1") = 1`
