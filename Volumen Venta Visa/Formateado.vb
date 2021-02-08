Sub copiar_Formato()

  Dim wb As Workbook
  Set wb = ThisWorkbook

  wb.Activate
         
  Set ws_visaclasica = wb.Sheets("VISA_CCLS-TxT")
  Set ws_visaclasicaN = wb.Sheets("VISA_CCLS")
  Set ws_visagold = wb.Sheets("VISA_GOLD-TxT")
  Set ws_visagoldN = wb.Sheets("VISA_GOLD)"
  Set ws_visacbsn = wb.Sheets("VISA_CBSN-TxT")
  Set ws_visacbsnN = wb.Sheets("VISA_CBSN")
  Set ws_visacorp = wb.Sheets("VISA_CORP-TxT")
  Set ws_visacorpN = wb.Sheets("VISA_CORP")
  Set ws_visaplat = wb.Sheets("VISA_PLAT-TxT")
  Set ws_visaplatN = wb.Sheets("VISA_PLAT")
  Set ws_visasig = wb.Sheets("VISA_SIG-TxT")
  Set ws_visasigN = wb.Sheets("VISA_SIG")
  Set ws_visacta = wb.Sheets("VISA_CTA-TxT")
  Set ws_visactaN = wb.Sheets("VISA_CTA")
  Set ws_visainf = wb.Sheets("VISA_INF-TxT")
  Set ws_visainfN = wb.Sheets("VISA_INF")
  Set ws_visadebemp = wb.Sheets("VISA_DBSN-TxT")
  Set ws_visadebempN = wb.Sheets("VISA_DBSN")
  Set ws_visauclsd = wb.Sheets("VISA_UCLSD-TxT")
  Set ws_visauclsdN = wb.Sheets("VISA_UCLSD")
  Set ws_visagift = wb.Sheets("VISA_GIFT-TxT")
  Set ws_visagiftN = wb.Sheets("VISA_GIFT")

 Call Formatear_Valores(ws_visaclasica,ws_visaclasicaN)
 Call Formatear_Valores(ws_visagold,ws_visagoldN)
 Call Formatear_Valores(ws_visacbsn,ws_visacbsnN)
 Call Formatear_Valores(ws_visacorp,ws_visacorpN)
 Call Formatear_Valores(ws_visaplat,ws_visaplatN)
 Call Formatear_Valores(ws_visasig,ws_visasigN)
 Call Formatear_Valores(ws_visacta,ws_visactaN)
 Call Formatear_Valores(ws_visainf,ws_visainfN)
 Call Formatear_Valores(ws_visadebemp,ws_visadebempN)
 Call Formatear_Valores(ws_visauclsd,ws_visauclsdN)
 Call Formatear_Valores(ws_visagift,ws_visagiftN)



End Sub