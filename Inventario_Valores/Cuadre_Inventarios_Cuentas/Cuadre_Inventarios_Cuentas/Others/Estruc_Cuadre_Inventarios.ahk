; #INDEX# =======================================================================================================================
; Title .........: Estructura de Cuadre de Inventario - Reporte CCC
; Version : 1.0
; Description ...: Guarda los datos en memoria
; Author(s) .....: LUIS ORLANDO LUNA CRUZ
; ===============================================================================================================================

class Cuadre_Inventario{

    Cuenta := ""
    Moneda := ""
    ;Oficina := ""
    SAltamira := ""
    SContable := ""
    Diferencia := ""
    

    Init(Cuenta, Moneda, SAltamira,SContable,Diferencia){
        
        this.Cuenta := Cuenta
        this.Moneda := Moneda  
        ;this.Oficina := Oficina
        this.SAltamira := SAltamira
        this.SContable := SContable
        this.Diferencia := Diferencia
        
    }
}


class Cuadre_Cuentas{

    Cuenta := ""
    Propiedad := ""
    Moneda := ""
    Descripcion := ""
    Tipo := ""
    Fisico := ""
    Desmate :=""
    FisicoInterna := ""
    DesmateInterna :=""
    TipoCuenta	:= ""
    

    Init(Cuenta, Propiedad,Moneda,Descripcion,Tipo,Fisico,Desmate,FisicoInterna,DesmateInterna,TipoCuenta){
        
        this.Cuenta := Cuenta
        this.Propiedad := Propiedad
        this.Moneda := Moneda  
        this.Descripcion := Descripcion
        this.Tipo := Tipo
        this.Fisico := Fisico
        this.Desmate := Desmate
        this.FisicoInterna := FisicoInterna
        this.DesmateInterna := DesmateInterna
        this.TipoCuenta := TipoCuenta

    }
}


class Carga_Contable{

    Cuenta := ""
    Titular := ""
    Fecha := ""
    Concepto := ""
    Origen := ""
    Destino := ""
    Divisa := ""
    Debe := ""
    Haber := ""

    

    Init(Cuenta,Titular,Fecha,Concepto,Origen,Destino,Divisa,Debe,Haber){
        
        this.Cuenta := Cuenta
        this.Titular := Titular
        this.Fecha := Fecha
        this.Concepto := Concepto
        this.Origen := Origen
        this.Destino := Destino
        this.Divisa := Divisa
        this.Debe := Debe
        this.Haber := Haber

    }
}


class Mov_Contable{

    Fecha := ""
    Cuenta := ""
    Divisa := ""
    Importe := ""
    Contrato := ""
    Origen := ""
    Destino := ""

    Init(Fecha,Cuenta,Divisa,Importe,Contrato,Origen,Destino){
        
        This.Fecha := Fecha
        This.Cuenta := Cuenta
        This.Divisa := Divisa
        This.Importe := Importe
        This.Contrato := Contrato
        This.Origen := Origen
        This.Destino := Destino

    }
}



ToPrint(vSheet, arrayReporte, xl){

    for i, vDataReporte in arrayReporte {

        vSheet.Cells(i+1,1) := vDataReporte.Cuenta
        vSheet.Cells(i+1,2) := vDataReporte.Moneda
        vSheet.Cells(i+1,3) := vDataReporte.SAltamira
        vSheet.Cells(i+1,4) := vDataReporte.SContable		
        vSheet.Cells(i+1,5) := vDataReporte.Diferencia

    }

    vSheet.columns("A:E").autoFit
    XL_Filter_Turn_On(xl,"A:E")
    XL_Format_Cell_Shading(xl,RG:="A1:E1",Color:=20)
}

ToPrint2(vSheet, arrayReporte, xl){

    for i, vDataReporte in arrayReporte {

        vSheet.Cells(i+1,1) := vDataReporte.Cuenta
        vSheet.Cells(i+1,2) := vDataReporte.Propiedad
        vSheet.Cells(i+1,3) := vDataReporte.Moneda
        vSheet.Cells(i+1,4) := vDataReporte.Descripcion
        vSheet.Cells(i+1,5) := vDataReporte.Tipo
        vSheet.Cells(i+1,6) := vDataReporte.Fisico
        vSheet.Cells(i+1,7) := vDataReporte.Desmate			
        vSheet.Cells(i+1,8) := vDataReporte.FisicoInterna
        vSheet.Cells(i+1,9) := vDataReporte.DesmateInterna
        vSheet.Cells(i+1,10) := vDataReporte.TipoCuenta

    }

    vSheet.columns("A:J").autoFit
    XL_Filter_Turn_On(xl,"A:J")
    XL_Format_Cell_Shading(xl,RG:="A1:J1",Color:=20)
}


ToPrint3(vSheet, arrayReporte, xl){

    for i, vDataReporte in arrayReporte {

        vSheet.Cells(i+1,1) := vDataReporte.Cuenta
        vSheet.Cells(i+1,2) := vDataReporte.Titular
        vSheet.Cells(i+1,3) := vDataReporte.Fecha
        vSheet.Cells(i+1,4) := vDataReporte.Concepto
        vSheet.Cells(i+1,5) := vDataReporte.Origen
        vSheet.Cells(i+1,6) := vDataReporte.Destino
        vSheet.Cells(i+1,7) := vDataReporte.Divisa			
        vSheet.Cells(i+1,8) := vDataReporte.Debe
        vSheet.Cells(i+1,9) := vDataReporte.Haber

    }

    vSheet.columns("A:I").autoFit
    XL_Filter_Turn_On(xl,"A:I")
    XL_Format_Cell_Shading(xl,RG:="A1:I1",Color:=20)
    ;vSheet.Sheets("Hoja1").Delete
}


ToPrint4(vSheet, arrayReporte, xl){

    for i, vDataReporte in arrayReporte {

        vSheet.Cells(i+1,1) := vDataReporte.Fecha
        vSheet.Cells(i+1,2) := vDataReporte.Cuenta
        vSheet.Cells(i+1,3) := vDataReporte.Divisa
        vSheet.Cells(i+1,4) := vDataReporte.Importe
        vSheet.Cells(i+1,5) := vDataReporte.Contrato
        vSheet.Cells(i+1,6) := vDataReporte.Origen
        vSheet.Cells(i+1,7) := vDataReporte.Destino			

    }

    vSheet.columns("A:G").autoFit
    XL_Filter_Turn_On(xl,"A:G")
    XL_Format_Cell_Shading(xl,RG:="A1:G1",Color:=20)
   
}






