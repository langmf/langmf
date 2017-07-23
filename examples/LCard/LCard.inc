Declare Function CreateInstance Lib "*lcomp.dll" (ByVal slot As Long) As Long

<#Module=mdlConst>
Const IID_ILDEV = "{32bb8320-b41b-11cf-a6bb-0080c7b2d682}"


Const card_L791   = 19
Const card_E440   = 30
Const card_E140   = 31
Const card_E2010  = 32
Const card_E270   = 33
Const card_E154   = 38
Const card_E2010B = 39

Const L_SUCCESS = 0
Const L_NOTSUPPORTED = 1
Const L_ERROR = 2
Const L_ERROR_NOBOARD = 3
Const L_ERROR_INUSE = 4

Const L_ADC_PARAM = 1
Const L_DAC_PARAM = 2

Const L_ASYNC_ADC_CFG = 3
Const L_ASYNC_TTL_CFG = 4
Const L_ASYNC_DAC_CFG = 5

Const L_ASYNC_ADC_INP = 6
Const L_ASYNC_TTL_INP = 7

Const L_ASYNC_TTL_OUT = 8
Const L_ASYNC_DAC_OUT = 9

Const L_STREAM_ADC = 1
Const L_STREAM_DAC = 2
Const L_STREAM_TTLIN = 3
Const L_STREAM_TTLOUT = 4

Const L_EVENT_ADC_BUF = 1
Const L_EVENT_DAC_BUF = 2

Const L_EVENT_ADC_OVF = 3
Const L_EVENT_ADC_FIFO = 4
Const L_EVENT_DAC_USER = 5
Const L_EVENT_DAC_UNF = 6
Const L_EVENT_PWR_OVR = 7

Interface IFC_LCard [{},{},0] = QueryInterface(2),AddRef,Release,inbyte(4),inword(4),indword(4),outbyte(4),outword(4),outdword(4),inmbyte(4),inmword(4),inmdword(4),outmbyte(4),outmword(4),outmdword(4),GetWord_DM(2),PutWord_DM(2),PutWord_PM(2),GetWord_PM(2),GetArray_DM(3),PutArray_DM(3),PutArray_PM(3),GetArray_PM(3),SendCommand(1),PlataTest,GetSlotParam(1),OpenLDevice,CloseLDevice,SetParametersStream(5),RequestBufferStream(2),FillDAQparameters(1),InitStartLDevice,StartLDevice,StopLDevice,LoadBios(1),IoAsync(1),ReadPlataDescr(1),WritePlataDescr(2),ReadFlashWord(2),WriteFlashWord(2),EnableFlashWrite(1),EnableCorrection(1),GetParameter(2),SetParameter(2),SetLDeviceEvent(2),InitStartLDeviceEx(1),StartLDeviceEx(1),StopLDeviceEx(1)

Type SLOT_PARAM
	Base As Long
	BaseL As Long
	Base1 As Long
	BaseL1 As Long
	Mem As Long
	MemL As Long
	Mem1 As Long
	MemL1 As Long
	Irq As Long
	BoardType As Long
	DspType As Long
	Dma As Long
	DmaDac As Long
	DTA_REG As Long
	IDMA_REG As Long
	CMD_REG As Long
	IRQ_RST As Long
	DTA_ARRAY As Long
	RDY_REG As Long
	CFG_REG As Long
End Type

Type PLATA_DESCR_E140_E154
	SerNum As String * -9
	BrdName As String * -11
	Rev As String * 1
	DspType As String * -11
	IsDacPresent As Byte
	Quartz As Long
	Reserv2(2) As Integer
	KoefADC(7) As Single
	KoefDAC(3) As Single
	Custom(19) As Integer
End Type

Type ASYNC_PAR
	'DAQ_PAR
	s_Type As Long
	FIFO As Long
	IrqStep As Long
	Pages As Long
	'ASYNC_PAR
	dRate As Double
	Rate As Long
	NCh As Long
	Chn(127) As Long
	Data(127) As Long
	Mode As Long
End Type

Class LCard
	Dim ldev, pLDev, pIUnknown, slot, descr, pp
	
	Sub Class_Initialize()
		Dim slot as new SLOT_PARAM
		Dim descr as new PLATA_DESCR_E140_E154
		Dim pp as new ASYNC_PAR
		Dim ldev as new IFC_LCard
	End Sub

	Sub Class_Terminate()
		if pIUnknown then
			ldev.CloseLDevice
			Set ldev = Nothing
		end if
	End Sub
		
	Function Create(numSLot)
		pIUnknown = CreateInstance(numSLot)
		
		If pIUnknown Then
			Call sys.com.QueryInterface(pIUnknown, IID_ILDEV, pLDev)
			Call sys.com.Release(pIUnknown)
			
			if pLDev then
				ldev.ifc_Obj = pLDev
				ldev.ifc_FCP.ptrIfc = pLDev
				
				ldev.OpenLDevice
				ldev.GetSlotParam slot
				ldev.ReadPlataDescr descr
			else
				pIUnknown = 0
			end if			
		End If
		
		Create = pIUnknown
	End Function
	
	Function ADC_IN(param) 
	    Dim value, tmp
	    
	    if not isarray(param) then param = array(param) 

	    Select Case slot.BoardType
	    	case card_E140
	    		'Channel, Ku, Mode32, Kalib
	    		sys.prm param, 0, 1, 1, 0
					    		
			    value = param(2) * 32 + param(3) * 16 + param(0)
			    
				Select Case param(1)
			        Case 1
			            ADC_IN = ADC_IN_FAST(value) * 0.00125
			        Case 4
			            ADC_IN = ADC_IN_FAST(value + 64) * 0.0003125
			        Case 16
			            ADC_IN = ADC_IN_FAST(value + 128) * 0.000078125
			        Case 64
			            ADC_IN = ADC_IN_FAST(value + 192) * 0.00001953125
			    End Select
			
	    	case card_E154
	    		'Channel, Ku
	    		sys.prm param, 0, 1
					    		
			    value = param(0)
			    
				Select Case param(1)
			        Case 1
			            ADC_IN = ADC_IN_FAST(value) * 0.0024414
			        Case 3
			            ADC_IN = ADC_IN_FAST(value + 64) * 0.00078125
			        Case 10
			            ADC_IN = ADC_IN_FAST(value + 128) * 0.00024414
			        Case 30
			            ADC_IN = ADC_IN_FAST(value + 192) * 0.000078125
			    End Select	    	
	    End Select
	End Function
	
	Property Get ADC_IN_FAST(value)
	    pp.Chn(0) = value
		ADC_IN_FAST = Sys.Bytes.Long2Integer(Data(L_ASYNC_ADC_INP))
	End Property
	
	Property Get TTL_IN()
	    TTL_IN = Data(L_ASYNC_TTL_INP)
	End Property
	
	Property Let TTL_OUT(value)
	    Data(L_ASYNC_TTL_OUT) = value
	End Property
	
	Property Let DAC_OUT(numChannel, value)
	    pp.Mode = numChannel
		Data(L_ASYNC_DAC_OUT) = value
	End Property
	
	Property Let TTL_CFG(value)
	    Data(L_ASYNC_TTL_CFG) = value
	End Property

	Property Get Data(cfg)
	    pp.s_Type = cfg
	    Call ldev.IoAsync(pp)
	    Data = pp.Data(0)
	End Property

	Property Let Data(cfg, value)
	    pp.s_Type = cfg
	    pp.Data(0) = value
	    Call ldev.IoAsync(pp)
	End Property
End Class
<#Module>