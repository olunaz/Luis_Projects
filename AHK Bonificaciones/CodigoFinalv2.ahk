; #INDEX# =======================================================================================================================
; Title .........: Bonificaciones_Operativa
; Version : 1.1
; Description ...: Hace la operativa de bonificaciones de tarjetas
; Author(s) .....: Franco Tejero | Sandra Alvarado
; ===============================================================================================================================
; #ENTORNO#
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
FileEncoding UTF-8
#SingleInstance, Ignore
; ===============================================================================================================================
; #LIBRERIAS#
#Include ../../lib/ECSLib.ahk
#Include ../../lib/PI_3270.ahk
#Include ../../lib/PI_XLHandle.ahk
#Include ../../Process_Improvement/Library_Dev/Dev_Mode.ahk
#Include ../../Process_Improvement/Library_Dev/Cryption.ahk
#Include ../../Process_Improvement/Library_Dev/Developers.ahk
#Include ./Others/Class_ImageButton.ahk
; ===============================================================================================================================
; #HOTKEY#
; ===============================================================================================================================
; #TRIGGER#
Script_Start()
;MainForm()
; ===============================================================================================================================
; #HEADER#
;limpiamos el clipboard para que no afecte a la ejecuci�n

global RepDate, startDate, endDate, MyProgress, BT1
global IniCoin, fjda:=A_ScriptDir "\Config.jda", fini:=A_ScriptDir "\Config.ini", PcWinStat, PcDiro, PcDiro2, PcDiri, PcDiri2, PcDiri3
Script_Start()
{
	Progress, m w280 y520, Verificando conexi`�n..., BONIFICACIONES OPERATIVA, [PI] - BONIFICACIONES OPERATIVA
	Progress, 5
	Clipboard =
	Inicoin := IniSd()
	;timestamp de inicio de ejecuci�n
	startDate := A_Now
	;Checking connection and Updating local repository
	configurarAmbiente()
	conn := CheckConnection()
	;ImprimirAmbiente()

	Progress, 50
	If (conn = 1)
	{
		p = 50
		;Calculando tiempo de conexi�n
		connDate := A_Now
		connTime := calcProgress(startDate, connDate)

		;Connection timestamp
		Loop 45
		{
			Progress, %p%
			p++
		}	
		Progress, 100, Conexi`�n verificada
		sMsgConn := "Tiempo de conexi�n: " connTime " segundos"
		log_Info(sMsgConn)
		;escribimos la primera l�nea en el log
		sMessageInicio := "Comienzo ejecuci�n"
		log_Info(sMessageInicio)		
		;Comienzo ejecuci�n
		Sleep, 450
		Progress, Hide
		MainForm()
	}	
	Else
	{
		Progress, Hide
		MsgBox, 48 , Sin Conexion, No existe conexi`�n con Atenea. `nConectar e intentar otra vez. , 5
		ExitApp
	}
}
; ===============================================================================================================================
; #CODIGO SCRIPT#
MainForm()
{
	Static User, Pass

	yest = %A_Now%
	
	if A_WDay = 2
		yest += -3, d
	else
		yest += -1, d

	global nh := Chr(200)
    Gui, Main:Add, Progress, x0 y0 w310 h45 c072146 vMyProgress, 100
    Gui, Main:Add, Picture, x110 y5 w90 h35 BackgroundTrans, ./Files/bbva-logo-captura.png
    Gui, Main:Font, s12
    Gui, Main:Add, Picture, x30 y60 w250 h50, ./Files/textbox01.png
    Gui, Main:Color,ffffff, f4f4f4
    Gui, Main:Font, s8
    Gui, Main:Add, Text, c666666 xp+25 yp+5, Usuario
    Gui, Main:Font, s11
    Gui, Main:Add, edit, x54 y80 w200 -VScroll Limit7 Uppercase vUser -E0x200 ;#f4f4f4
    Gui, Main:Font, s8
    Gui, Main:Add, Picture, x30 yp+50 w250 h50, ./Files/textbox01.png
    Gui, Main:Font, s11
    Gui, Main:Add, edit, wp-50 yp21 xp22 -VScroll Limit8 Uppercase Password vPass -E0x200 ;#f4f4f4
    Gui, Main:Font, s8
    Gui, Main:Add, Text, c666666 xp+1 yp-17, Contrase`�a

	Gui, Main:Add, Button, vBT1 yp+60 w120 h35 xm x95 hwndHBT1, INGRESAR
	Opt1 := [0, 0X1d73b2, , "White", "H", , 0X004481, 4] ; normal flat background & text color
	Opt2 := [ , 0X004481]                                          ; hot flat background color
	Opt5 := [ , , ,0X004481]                                      ; defaulted text color -> animation
	If !ImageButton.Create(HBT1, Opt1, Opt2, , , Opt5)
	MsgBox, 0, ImageButton Error Btn1, % ImageButton.LastError
	Gui, Main:Font, cGray
    Gui, Main:Add, Text, xp+180 yp+25, v1.7
	Gui, Main:Show, w310 h250
	Gui, Main:Show,, [PI] - BONIFICACIONES
    Hotkey, !#F12, DevMode
    return 


	DevMode:
		Input, command, L2 
		if (command = "pi")
		{
			Gui, Main:Hide
			MsgBox,1, CAUTION, You will get into Developer mode

			IfMsgBox, Ok
			{
				DevForm()
				Return
			}

			IfMsgBox, Cancel
				MsgBox, You will be back to main form
				Gui, Main:Show
		}
	Return

	MainButtonINGRESAR:
	{
		Gui, Main:Submit                 
		if (User = "" or Pass = "")
		{
			log_Warn("Datos Incompletos")
			MsgBox,48, Datos Incompletos, Por favor`, ingrese registro y contrase`�a.,5
			Gui, Main:Show 
		}
		Else if not(RegExMatch(User, "^([p|P]{1}[0-9]{6})$"))
		{
			log_Warn("Registro errado")
			MsgBox,48, Error de Registro, Por favor`, ingrese registro correcto.,5
			Gui, Main:Show            
			GuiControl, Gui, Main:Focus, User     
		}
		Else
		{ 
			Progress, m2 w280 y520, Ejecutando Host..., [PI] - BONIFICACIONES OPERATIVA
			Progress, 1

			;Leyendo archivo Config
			DecryptIni(fjda,Inicoin)
			IniRead, team, .\Config.ini, OPERATIONS, Team
			IniRead, PcWinStat, .\Config.ini, PC3270, PcWinState
			IniRead, PcDiro, .\Config.ini, OUTPUT, OutDir
			IniRead, PcDiro2, .\Config.ini, OUTPUT, OutDir2
			IniRead, PcDiri, .\Config.ini, INPUT, InDir,
			IniRead, PcDiri2, .\Config.ini, INPUT, InDir2,
			IniRead, PcDiri3, .\Config.ini, INPUT, InDir3,
			EncryptIni(fini,Inicoin)
			log_Info("Datos Config obtenidos")

			;Ejecuntado 3270
			c3270 := new PI_z3270()                                        ;Instancia de la clase z3270 que es una libreria importada z3270.ahk
			Try opnhst := c3270.OpenHost(PcWinStat)
			log_Info("Ejecutando PC3270")

			if (opnhst != "")
			{    
				log_Info("Conectando PC3270")
				Progress, 5, Conectando Host...
 
				Try conx := c3270.Connect()                                  ;Se usa la funci�n Connect(session) del objeto instanciado c3270

				if (conx != "")
				{
					Hotkey, !#F12, DevMode, Off
					log_Info("Comenzando operativa ")
					Connect3270(c3270,User,Pass)
					ExitApp
				}
				else
				{
					log_Error("No se pudo conectar al host")
					Progress, Hide
					Gui, Main:Show
					MsgBox, 48 , Sin Conexion , No se pudo conectar al host. `nIntentar otra vez. , 5
					Try c3270.CloseSess()
				}
			}
			else
			{
				log_Error("No se pudo abrir host")
				Progress, Hide
				Gui, Main:Show
				MsgBox, 48 , Sin Conexion , No se pudo abrir host. `nIntentar otra vez. , 5
				Try c3270.CloseSess()
			}
		}
		Return          
	}
	MainGuiClose:
	MainGuiEscape:
	MainButtonCancelar:
		Script_End(c3270)
		;ExitApp
	Return
}

Connect3270(c3270,User,Pass){
	EstadosTarjetas := []
	inicio3270 := A_Now

	Progress, 5
	;clogin := CesnLogin(c3270, User, Pass)
	;log_Info("clogin = " clogin)

	;--------------------------------------------AMBIENTE DE PRODUCCI�N-----------------------------------------------------------------------------
	;clogin := CesnLogin(c3270, User, Pass)
	;log_Info("clogin = " clogin)

	;If (clogin){
	;--------------------------------------------AMBIENTE DE CALIDAD--------------------------------------------------------------------------------

	slogin := CtrolSLogin(c3270,User,Pass)

	If (slogin){

		qlogin := CicsAmbienteLogin(c3270, User, Pass)

		If (qlogin) {


		;Validar centro 0445
		Progress, 10, Revisando facultades...
		CheckPermissions(c3270)

		Progress, 15, Revisando centro...
		center0445 := ValidateCenter0445(c3270)
		log_Info("center0445 = " center0445)

	If (qlogin And center0445)
	{
		
		Progress, 25, Abriendo Archivo Excel...
		xl := ComObjCreate("Excel.Application")
		wb := XL_Open(xl, PcDiri, 1)
		ws := wb.Sheets("DATOS")
		
		PXL_Run_Macro(wb, PcDiri, "Principal")

		If (ws.Cells(2,"A").Value = "" or ws.Cells(2,"B").Value = "" or ws.Cells(2,"C").Value = "")
		{
			MsgBox No se encontraron registros de bonificaciones v�lidos.`nEl robot se detendra.
			Progress, 100, Operativa interrumpida
			wb.Application.Quit
			Script_End(c3270)
		}

		filaInicial := 2
		filaFinal := XL_Last_Row(wb, 1, "D") 
		n := filaFinal - filaInicial + 1

		Loop, %n%
		{
			fexcel := A_Index + filaInicial - 1
			;MsgBox Prueba consultando la fila %fexcel%
			Progress, 20, Revisando fila %fexcel% de %filaFinal%

		    ;Revisando
			;Progress % sobre longitud de registro% 

			c3270.BlackScreen()
			c3270.sTxt("MPZ0",1,1)
			c3270.sKey("e")

			if c3270.gTxt(1,2,7) = "BXA0147"{
				c3270.sKey("c",1,400)
			} 

			contrato :=StrReplace( ws.Cells(A_Index + filaInicial - 1, "F").Value, "'","" )

			StringLeft, primerosdigitos, contrato, 4
			StringMid, tipocontrato, contrato, 11 , 2

			codbonificacion := ws.Cells(A_Index + filaInicial - 1, "H").Value
			fechafin := ws.Cells(A_Index + filaInicial - 1, "I").Value 
			;MsgBox fechafin %fechafin%

			if ((StrLen(contrato) != 20) or (primerosdigitos != "0011") or (tipocontrato != "50"))
			{
				;MsgBox Prueba contrato %contrato% invalido
				Estado := "ERROR: CONTRATO INVALIDO (debe ser una l�nea de cr�dito y tener 20 d�gitos)"
			}
			else if (StrLen(codbonificacion) = 0)
			{
				;MsgBox Prueba codigo %codbonificacion% invalido
				Estado := "ERROR: CODIGO DE BONIFICACION VACIO (debe tener dos d�gitos)"
			}
			else if (StrLen(codbonificacion) != 2)
			{
				;MsgBox Prueba codigo %codbonificacion% invalido
				Estado := "ERROR: CODIGO DE BONIFICACION INVALIDO (debe tener dos d�gitos)"
			}
			else if (!RegExMatch(fechafin, "^([0-9]{2}/[0-9]{2}/[0-9]{4})$")) 
			{
				;MsgBox Prueba fecha %fechafin% no valida
				Estado := "ERROR: FECHA INVALIDA (debe ser de la forma DD/MM/YYYY)"
			}
			else
			{
				c3270.sTxt("1", 17, 32)
				c3270.sTxt(contrato, 17, 56)
				c3270.sKey("e")
				Sleep,400

				;MsgBox Prueba Se introdujo contrato en MPZ0

				if (Trim(c3270.gTxt(23,2,7)) = "MPE0007")
				{
					;Contrato no existe
					;MsgBox Prueba contrato no existe
					Estado := "ERROR: CONTRATO NO EXISTE"
				}
				else if (Trim(c3270.gTxt(23,2,7)) = "MPE1031")
				{
					;MsgBox Prueba contrato inactivo
					;Contrato inactivo
					Estado := "ERROR: CONTRATO INACTIVO"
				}
				else if (Trim(c3270.gTxt(23,2,7)) = "MPE8152")
				{
					;MsgBox Prueba centro no autorizado
					;Centro actual no permitido para bonificar
					MsgBox No se puede bonificar con el centro actual. Use otro registro.`nEl robot se detendra.
					Progress, 100, Operativa interrumpida
					wb.Application.Quit
					Script_End(c3270)
				}
				;else if (Trim(c3270.gTxt(23,2,7)) = "MPE2068")
				else if (Trim(c3270.gTxt(10,6,2)) = "")
				{ 
					MsgBox Prueba se va a ir a MP10
					;Contrato valido y sin bonificacion previa
					c3270.BlackScreen()

					Estado := ConsultMP10(c3270,contrato,codbonificacion,fechafin)

					MsgBox  EstadoFinal : %Estado% 
				} 
				;BUCLE EXITOSO
				else if (Trim(c3270.gTxt(2,30,24)) = "CONSULTA Y MANTENIMIENTO")
				{
					;MsgBox Prueba contrato valido
					;Contrato valido y con bonificacion previa
					c3270.sKey("f3")
					Sleep, 400

					if (Trim(c3270.gTxt(2,30,24)) = "VINCULACION DE CONTRATOS")
					{

						c3270.sTxt(codbonificacion,7,18)
						c3270.sTxt(fechafin,9,18) ;DD/MM/AAAA
						c3270.sKey("f3")
						Sleep, 400
						;MsgBox Prueba dio de alta a la bonificacion

						if (Trim(c3270.gTxt(3,2,7)) = "MPA0182")
						{
							;C�digo de confirmaci�n, operaci�n exitosa
							c3270.sKey("f7")
							Sleep, 400
							if (Trim(c3270.gTxt(3,2,7)) = "MPA0017")
							{
								;MsgBox Prueba boni correcta
								Estado := "BONIFICACION PROCESADA MPZ0"
							}
										
							else if ((Trim(c3270.gTxt(3,2,7)) ) = "")
							{
								;MsgBox Prueba boni correcta
								Estado := "ERROR SE INGRESO UN CODIGO O FECHA YA EXISTENTE"
							}
					
							else
							{
								;MsgBox Prueba boni incorrecta
								Estado := "ERROR EN MPZ0 LUEGO DE CONFIRMAR"
							}
						}
						else if (Trim(c3270.gTxt(23,2,7)) = "MPE1032")
						{
							;C�digo de bonificacion invalido
							;MsgBox Prueba cod boni invalido
							Estado := "ERROR: CODIGO BONIFICACION INVALIDO"
						}
						else if (Trim(c3270.gTxt(23,2,7)) = "MPE6000")
						{
							;Bonificacion principal existente
							;MsgBox Prueba boni principal existente
							Estado := "ERROR: BONIFICACION PRINCIPAL EXISTENTE"
						}
						else if (Trim(c3270.gTxt(23,2,7)) = "MPE2068")
						{
							;Codigo de bonificacion no esta registrado
							;MsgBox Prueba boni principal existente
							Estado := "ERROR: COD. DE BONIFICACION NO ESTA REGISTRADO"
						}
						else if (Trim(c3270.gTxt(23,2,7)) = "MPE0944")
						{
							;Codigo de bonificacion no esta registrado
							;MsgBox Prueba boni principal existente
							Estado := "ERROR: USUARIO NO FACULTADO PARA REALIZAR OPERACION"
						}

						else
						{
							;MsgBox Prueba boni no se pudo por motivo no mapeado 1
							Estado := "ERROR: MOTIVO NO IDENTIFICADO"
						}
					}
					else
					{
						;Usuario no facultado para dar altas de modificaciones
						MsgBox No se puede bonificar con el usuario actual en el MPZ0. Use otro registro.`nEl robot se detendra.
						Progress, 100, Operativa interrumpida
						wb.Application.Quit
						Script_End(c3270)
					}
				}else
				{
					;MsgBox Prueba boni no se pudo por motivo no mapeado 1
					Estado := "ERROR: MOTIVO NO IDENTIFICADO"
				}
			}

			MsgBox %Estado% 
			;EstadosTarjetas.push(Estado)
			MsgBox Estado por Imprimir
			ws.Cells(fexcel,"M").Value := Estado
			MsgBox Estado Imprimido

			Estado = ""
			contrato := ""
			codbonificacion := ""
			fechafin := ""
			tipocontrato := ""
			primerosdigitos := ""

			;MsgBox Prueba siguiente contrato
		}

		/*
		n := EstadosTarjetas.MaxIndex()
		fFinal := n + 1
		COMArray_EstadosTarjetas := ComObjArray(12, n, 1)
		
		for i, estado in EstadosTarjetas
		{
			COMArray_EstadosTarjetas[i-1, 0] := estado
		}
		ws.Range("M17:M" . fFinal).Value := estado
		*/

		varDate := A_Now
		vHora := SubStr(varDate, 9, 4)
		StringLeft, varHoraL, vHora, 2
		StringRight, varHoraR, vHora, 2
		varHora := varHoraL . "h" . varHoraR . "'"
		FormatTime, vFecha , %varDate% , dd.MM.yyyy
		vDate := vFecha . " - " . varHora

		FullPathName := PcDiro . "\BONIFICACIONES PROCESADAS - " . vDate . ".xlsm"

		wb.SaveAs(FullPathName)
		Progress, 95, Archivo guardado
		;wb.Close()
		;wb.Application.Quit
		
		fin3270 := A_Now
		tiempo3270 := calcProgress(inicio3270,fin3270)

		log_Info("Tiempo 3270: " tiempo3270 " segundos")
		log_Info("Operativa Realizada")

		Progress, 100, Operativa Realizada
		Progress, Hide
		try c3270.CloseSess()
		
		Script_End(c3270)
	}
}

	}

}

CheckPermissions(c3270){
	;Revisar si tiene permiso a todas las transacciones
	;donde no pueda entrar a una, el robot se debe detener
	Transacciones := ["MPZ0","MP10"]

	for i, transaccion in Transacciones
	{
		c3270.BlackScreen()
		c3270.sTxt(transaccion,1,1)
		c3270.sKey("e")

		if c3270.gTxt(1,2,7) = "BXA0147"{
			c3270.sKey("c",1,400)
		}

		if (Trim(c3270.gTxt(23,21,31)) = "CICSPTOR You are not authorized"){
			log_Info("Operativa Interrumpida, usuario sin facultades")
			
			MsgBox Usted no tiene facultades para la transaccion %transaccion%.`nEl robot se detendra.
			Progress, 100, Operativa Interrumpida
			
			Script_End(c3270)
		}
	}
}

ValidateCenter0445(c3270){
	;Si no estuviera con centro 0445
	;cambiar con QGTR al centro 0445 y parar
	;si no se pudiera
	validate := 1
	c3270.BlackScreen()
	c3270.sTxt("PE29",1,1)
	c3270.sKey("e")
	
	if c3270.gTxt(1,2,7) = "BXA0147"{
		;MsgBox Entro al primer IF Que no debio entrar
		c3270.sKey("c",1,400)
	} 
	
	if c3270.gTxt(2,2,4) != "0445"{
		;MsgBox Entro al segundo IF Que no debio entrar
		c3270.BlackScreen()
		c3270.sTxt("QGTR",1,1)
		c3270.sKey("e")

		if c3270.gTxt(1,2,7) = "BXA0147"{
			;MsgBox Entro al tercer IF Que no debio entrar
			c3270.sKey("c",1,400)
		} 

		c3270.sTxt("0445",4,34)
		c3270.sKey("e")

		if (trim(c3270.gTxt(2,2,18)) != "DIARIO ELECTRONICO"){
			MsgBox No se puede cambiar al centro 0445. Ingrese con un usuario facultado.`nEl robot se detendra.
			Progress, 100, Operativa interrumpida
			Script_End(c3270)
		}
	} 

	c3270.BlackScreen()
	Return validate
}


ConsultMP10(c3270,contrato,codbonificacion,fechafin){
	c3270.sTxt("MP10",1,1)
	c3270.sKey("e")

	if c3270.gTxt(1,2,7) = "BXA0147"{
		c3270.sKey("c",1,400)
	}

	c3270.sTxt(contrato,5,20)
	c3270.sKey("e")
	Sleep, 300
	c3270.sKey("e")
	Sleep, 300

	;FormatTime, fechafin2, fechafin, MM-yyyy
	StringMid, ffmes, fechafin, 4, 2
	StringRight,  ffanho, fechafin, 4
	fechafin2 := ffmes . "-" . ffanho
	ffmes :=""
	ffanho := ""

	MsgBox Prueba la fecha fin %fechafin% ahora es %fechafin2%

	c3270.sTxt(codbonificacion,21,22)
	c3270.sTxt(fechafin2,21,49) ;MM-AAAA
	fechafin2 := ""
	MsgBox Prueba escribio codigo de boni y fecha fin

	c3270.sKey("f2")
	MsgBox Prueba dio de alta
	Sleep, 400

	if (Trim(c3270.gTxt(3,2,7)) = "MPA0012"){
		;Bonificacion correcta
		MsgBox Prueba boni MP10 bien hecha
		Estado := "BONIFICACION PROCESADA MP10" 
		Return Estado
		MsgBox %Estado% 
	}
	else if (Trim(c3270.gTxt(23,2,7)) = "MPE0005"){
		;C�digo de bonificacion invalido
		MsgBox Prueba error codigo de boni malo
		Estado := "ERROR: CODIGO BONIFICACION INVALIDO"
		Return Estado
	}

	else if (Trim(c3270.gTxt(23,2,7)) = "MPE2068")
	{
		;Codigo de bonificacion no esta registrado
		MsgBox Prueba boni principal existente
		Estado := "ERROR: COD. DE BONIFICACION NO ESTA REGISTRADO"
		Return Estado
	}
	else if (Trim(c3270.gTxt(23,2,7)) = "MPE0093")
	{
		;Codigo de bonificacion no esta registrado
		MsgBox Prueba boni principal existente
		Estado := "ERROR: MODIFICACION NO PERMITIDA"
		Return Estado
	}

	else{
		MsgBox Prueba otro tipo de error 
		Estado := "ERROR: MOTIVO NO IDENTIFICADO"
		Return Estado
	}
    
	MsgBox %Estado%

	c3270.sKey("c")
	Sleep, 400

	MsgBox %Estado%
}


CtrolSLogin(c3270,User,Pass){
	
	res := 0
	c3270.sTxt("S", 24, 24)
    c3270.sKey("e",,320)

    c3270.sTxt(User,15,31)
	c3270.sKey("e",,320) ;OBLIGATORIO EN EL ENTORNO SUPER --- SOLO AQU�
	c3270.sTxt(Pass,16,31)
	c3270.sKey("e",,320)
	
	sleep 3000

	viol := trim(c3270.gTxt(23, 12, 7))
    pctittle := trim(c3270.gTxt(3, 23, 17))

    If (pctittle = "BANCO CONTINENTAL"){

		res := 1	
	}	
	Else 
    {
		Try c3270.CloseSess()
        Try Progress, Hide

		If (viol = "INVALID")
			MsgBox, 64, ERROR DE CREDENCIALES, Usuario y/o contrase%nh%a incorrecta.,5
		Else
			MsgBox, 64, LOGIN ERROR, Error no clasificado. Por favor`, contacte con 'Process Improvement',5

		Try Gui, Show   
	}

	Return res

}

CicsAmbienteLogin(c3270, User, Pass){

	res := 0
	fila:=9
	col:=2

	estado:=Trim(c3270.gTxt(9, 51, 13))

	While (estado <> "Disponible"){
		
		c3270.sTxt("t", fila, col)
		c3270.sKey("e",,320)
		estado:=Trim(c3270.gTxt(9, 51, 13))
	
	}
		sleep 2000
		c3270.sTxt("s", fila, col)
		c3270.sKey("e",,320)

	if	(Trim(c3270.gTxt(1, 2, 11))="BIENVENIDOS"){

		c3270.sKey("e",,320)
		c3270.sKey("c",,320)
		c3270.sTxt("cesn",1,1)
		c3270.sKey("e",,320)

	}
	if	(Trim(c3270.gTxt(1, 29, 6))="Signon"){
    
		c3270.sTxt(User,10,26)
		;c3270.sTxt(Pass,11,26)
		c3270.sTxt("agua0001",11,26)
		c3270.sKey("e",,320)mpz
		c3270.sKey("c",,320)

		res := 1
	}
	Else 
	{
		Try c3270.CloseSess()
		Try Progress, Hide

		If (invalid = "password is invalid")
			MsgBox, 64, ERROR DE CREDENCIALES, Usuario y/o contrase`�a incorrecta.,5
		Else If (InStr(invalid,"expired"))
			MsgBox, 48, ERROR DE CREDENCIALES, Contrase`�a expirada.,5
		Else
			MsgBox, LOGIN ERROR, Error no clasificado. Por favor`, contacte con 'Process Improvement',5

		Try Gui, Show
	}

	Return res     

}

;====================================================================================================
; #CIERRE SCRIPT#
Script_End(c3270)
{
	Try Gui, Main:Destroy
	Try c3270.CloseSess()
	;Subiendo archivo Config a repositorio

	log_Info("Config File uploaded" )

	;calculamos el tiempo de ejecuci�n
	endDate := A_Now
	totalTime := calcProgress(startDate, endDate)
	sMessageFin := "Fin ejecuci�n. Tiempo de ejecuci�n: " . totalTime . " segundos"

	;Subiendo logs a repositorio
	log_End(sMessageFin)

	;limpiamos el clipboard para que no afecte a la ejecuci�n
	Clipboard =
	ExitApp
}
; ===============================================================================================================================
