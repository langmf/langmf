<#Module=mdl_SIO>
Const sio_CONFIG_CONTROL_REG = &H2
Const sio_DEVICE_SELECT_REG = &H7
Const sio_CHIP_ID_REG = &H20
Const sio_CHIP_REV_REG = &H21
Const sio_BASE_ADDRESS_REG = &H60
sio_STANDART_PORT = Array(&H2E, &H4E)

Dim sioRegRW, sioDataRW


'----------------------------------
Function FindSensor()
    Dim tmpObj, i

    Set FindSensOr = Nothing

    For Each tmpObj In Array(New IT87x, New F718x, New NCT677x)
        If tmpObj.Find Then Set FindSensOr = tmpObj : Exit For : End If
    Next
End Function


'----------------------------------
Function SIO_Read(reg)
    PortIO(sioRegRW) = reg
    SIO_Read = PortIO(sioDataRW)
End Function

Function SIO_Write(reg, value)
    PortIO(sioRegRW) = reg
    PortIO(sioDataRW) = value
End Function

Function SIO_ReadWord(reg)
    SIO_ReadWord = SIO_Read(reg) * 256 + SIO_Read(reg + 1)
End Function

Function SIO_SelectLDN(num)
    PortIO(sioRegRW) = sio_DEVICE_SELECT_REG
    PortIO(sioDataRW) = num
End Function

Function SIO_GetLDA(reg)
  Dim tmpAddr
    SIO_GetLDA = 0
    tmpAddr = SIO_ReadWord(reg)
    If tmpAddr < &H100 Then Exit Function
    If (tmpAddr And &HF007&) = &HF007& Then Exit Function
    Sleep 250
    If tmpAddr <> SIO_ReadWord(reg) Then Exit Function
    SIO_GetLDA = tmpAddr
End Function

Function SIO_SelectRW(index)
    sioRegRW = sio_STANDART_PORT(index) 
    sioDataRW = sioRegRW + 1
End Function


'----------------------------------
Class SIO_CommonFunc
    Dim cf_Model, cf_Address, cf_Reg, cf_Data

    Dim cf_VENDOR_CHECK, cf_ADDRESS_OFFSET, cf_DATA_OFFSET, cf_MONITOR_LDN
    Dim cf_CHIPS_ID, cf_CHIPS_NAME, cf_ENTER, cf_EXIT
    Dim cf_CHIPS_MASK, cf_BANK_REG

    Dim cf_FAN_SPEED_MSB, cf_FAN_SPEED_LSB, cf_FAN_SPEED_MIN, cf_FAN_SPEED_KOEF    
    Dim cf_VOLTAGE_GAIN, cf_VOLTAGE_MVLSB
    Dim cf_VOLT_COUNT, cf_FAN_COUNT, cf_TEMP_COUNT
    Dim cf_VOLTAGE_BASE, cf_TEMP_BASE

    Private Sub Class_Initialize
        cf_ADDRESS_OFFSET = 5
        cf_DATA_OFFSET = 6
        cf_MONITOR_LDN = Array(Array(0, 4))
        cf_FAN_SPEED_KOEF = 0
        cf_FAN_SPEED_MIN = 40
        cf_VOLTAGE_MVLSB = 8
    End Sub


    '----------------------------------
    Function PortLXM(value)
        Dim a, tmp, lxm, lxm2
        For each tmp in value
        lxm = split(tmp, "=") : lxm(0) = trim(lxm(0)) : lxm(1) = trim(lxm(1))
        select case ucase(lxm(0))
            case "R" : PortIO(sioRegRW) = lxm(1)
            case "D" : PortIO(sioDataRW) = lxm(1)
            case "W"
                lxm2 = split(lxm(1), ",")
                lxm2(0) = val(trim(lxm2(0)))
                lxm2(1) = val(trim(lxm2(1)))
                Call SIO_Write(lxm2(0), lxm2(1))
            case "IF"
                lxm2 = split(lxm(1), ",")
                lxm2(0) = val(trim(lxm2(0)))
                lxm2(1) = val(trim(lxm2(1)))
                lxm2(2) = val(trim(lxm2(2)))
                If sioRegRW = lxm2(0) Then PortIO(sioRegRW) = lxm2(1) else PortIO(sioRegRW) = lxm2(2)
        end select
        next
    End Function

    Function Probe()
        Dim id, a
        Probe = 0
        id = SIO_ReadWord(sio_CHIP_ID_REG)
        If id = 0 Or id = &HFFFF& Then Exit Function
        If Not IsEmpty(cf_CHIPS_MASK) Then id = id And cf_CHIPS_MASK
        For a = 0 to ubound(cf_CHIPS_ID)
            If id = cf_CHIPS_ID(a) Then Probe = id : Exit For : end if
        next
    End Function

    Function Name()
        Dim a
        Name = "unknown"
        For a = 0 to ubound(cf_CHIPS_ID)
            If cf_Model = cf_CHIPS_ID(a) Then Name = cf_CHIPS_NAME(a) : exit For : end if
        next
    End Function

    Function Find(obj)
        Dim a, tmp
        Find = false
        For a = 0 to Ubound(sio_STANDART_PORT)
            SIO_SelectRW(a)
            PortLXM cf_ENTER
            cf_Model = Probe
            If cf_Model > 0 then
                For each tmp in cf_MONITOR_LDN 
                    If tmp(0) = 0 Or tmp(0) = cf_Model Then SIO_SelectLDN(tmp(1)) : Exit For : end if
                next
                cf_Address = SIO_GetLDA(sio_BASE_ADDRESS_REG)
                If cf_Address > 0 then
                    cf_Reg = cf_Address + cf_ADDRESS_OFFSET
                    cf_Data = cf_Address + cf_DATA_OFFSET
                    If Eval(cf_VENDOR_CHECK) then
                        PortLXM cf_EXIT
                        Find = true
                        obj.Init
                    end if
                end if
                Exit Function
            end if
            PortLXM cf_EXIT
        next
    End Function

    Property Get Port(ByVal index)
        If Not IsEmpty(cf_BANK_REG) then
            PortIO(cf_Reg) = cf_BANK_REG
            PortIO(cf_Data) = sys.bytes.rshift(index, 8)
            index = index And &HFF
        end if
        PortIO(cf_Reg) = index
        Port = PortIO(cf_Data)
    End Property

    Property Let Port(ByVal index, value)
        If Not IsEmpty(cf_BANK_REG) then
            PortIO(cf_Reg) = cf_BANK_REG
            PortIO(cf_Data) = sys.bytes.rshift(index, 8)
            index = index And &HFF
        end If    
        PortIO(cf_Reg) = index
        PortIO(cf_Data) = value
    End Property

    Property Get Bit(index, num)
        Bit = Sys.Bytes.GetBit(Port(index), num)
    End Property

    Property Let Bit(index, num, value)
        Port(index) = sys.bytes.SetBit(Port(index), num, value)
    End Property


    '----------------------------------
    Function GetFans()
        Dim a, tmp
        If IsEmpty(cf_FAN_COUNT) Then Exit Function
        Redim tmp(cf_FAN_COUNT - 1)
        For a = 0 to ubound(tmp)
            If IsArray(cf_FAN_SPEED_MSB) then
                tmp(a) = Port(cf_FAN_SPEED_MSB(a)) * 256 + Port(cf_FAN_SPEED_LSB(a))
            else
                tmp(a) = Port(cf_FAN_SPEED_MSB + a * 2) * 256 + Port(cf_FAN_SPEED_MSB + a * 2 + 1)        
            end if
            If tmp(a) > 0 And cf_FAN_SPEED_KOEF > 0 Then tmp(a) = int(cf_FAN_SPEED_KOEF / tmp(a))
            If tmp(a) < cf_FAN_SPEED_MIN Then tmp(a) = 0
        next
        GetFans = tmp
    End Function

    Function GetVolts()
        Dim a, tmp
        If IsEmpty(cf_VOLT_COUNT) Then Exit Function
        Redim tmp(cf_VOLT_COUNT - 1)
        For a = 0 to ubound(tmp)
            If IsArray(cf_VOLTAGE_BASE) Then tmp(a) = Port(cf_VOLTAGE_BASE(a)) else tmp(a) = Port(cf_VOLTAGE_BASE + a)       
            tmp(a) = Round((tmp(a) * cf_VOLTAGE_MVLSB * cf_VOLTAGE_GAIN(a)) / 1000, 2)
        next
        GetVolts = tmp
    End Function

    Function GetTemps()
        Dim a, tmp
        If IsEmpty(cf_TEMP_COUNT) Then Exit Function
        Redim tmp(cf_TEMP_COUNT - 1) 
        For a = 0 to ubound(tmp)
            If IsArray(cf_TEMP_BASE) Then tmp(a) = Port(cf_TEMP_BASE(a)) else tmp(a) = Port(cf_TEMP_BASE + a)
        next
        GetTemps = tmp
    End Function    
End Class
<#Module>


<#Module=mdl_NCT677x>
Const NCT6771F = &HB470&, NCT6776F = &HC330&, NCT6779D = &HC560&

NCT_FAN_PWM_OUT_REG = Array(&H1, &H3, &H11, &H13, &H15)
NCT_FAN_PWM_COMMAND_REG = Array(&H109&, &H209&, &H309&, &H809&, &H909&)
NCT_FAN_CONTROL_MODE_REG = Array(&H102&, &H202&, &H302&, &H802&, &H902&) 
 

Class NCT677x
    Dim CF, LabelVolts, LabelFans, LabelTemps, FanCPU, TempCPU

    Sub Class_Initialize
        Set CF = New SIO_CommonFunc

        With CF
            .cf_CHIPS_ID        = Array(NCT6771F, NCT6776F, NCT6779D)
            .cf_CHIPS_NAME      = Array("NCT6771F", "NCT6776F", "NCT6779D")
            .cf_ENTER           = Array("R=&H87", "R=&H87")
            .cf_EXIT            = Array("R=&HAA")
            .cf_VENDOR_CHECK    = "((Me.Port(&H804F&) * 256) Or Me.Port(&H004F&)) = &H5CA3&"
            .cf_MONITOR_LDN     = Array(Array(0, &HB))

            .cf_VOLTAGE_GAIN    = Array(1, 12, 2, 2, 5, 1, 1, 2, 2)
            .cf_VOLTAGE_BASE    = Array(&H20, &H21, &H22, &H23, &H24, &H25, &H26, &H550&, &H551&)
            .cf_TEMP_BASE       = Array(&H27, &H73, &H250&)
            .cf_VOLT_COUNT      = 9
            .cf_FAN_COUNT       = 5
            .cf_TEMP_COUNT      = 3

            .cf_CHIPS_MASK         = &HFFF0&
            .cf_BANK_REG        = &H4E
        End With

        FanCPU = Array(0,2) : TempCPU = 1

        LabelVolts = Array("CPU VCore", "+12V", "Analog +3.3V", "+3.3V", "+5V", "Vin5", "Vin6", "Standby +3.3V", "VBAT")
        LabelFans = Array("Fan1", "Fan2", "Fan3", "Fan4", "Fan5")
        LabelTemps = Array("TMPIN1", "TMPIN2", "TMPIN3") 
    End Sub

    '============ Base Functions =============
    Property Get Family : Family = "NCT677x" : End Property

    Sub Init()
        With CF
            Select Case .cf_Model
                Case NCT6771F, NCT6776F
                    If .cf_Model = NCT6771F Then 
                        .cf_FAN_COUNT = 4
                    else
                        .cf_FAN_SPEED_MIN = 165
                    end if
                    .cf_FAN_SPEED_MSB = &H656&    
                                            
                Case NCT6779D
                    .cf_FAN_SPEED_MIN = 165
                    .cf_FAN_SPEED_MSB = &H4C0&
                    .cf_VOLTAGE_BASE = &H480&
            End Select
        End With     
    End Sub

    '========== Control Functions ============
    Property Get FanPWM(index)
        FanPWM = CF.Port(NCT_FAN_PWM_OUT_REG(index)) / 2.55
    End Property

    Property Let FanPWM(index, value)
        Dim vpwm
        vpwm = round(value * 2.55, 0)
        If vpwm < 0 Then vpwm = 0
        If vpwm > 255 Then vpwm = 255
        CF.Port(NCT_FAN_PWM_COMMAND_REG(index)) = vpwm
    End Property
    
    Property Get FanAuto(index)
        FanAuto = Abs(Not (CF.Port(NCT_FAN_CONTROL_MODE_REG(index)) = 0))
    End Property

    Property Let FanAuto(index, value)
        If value = 0 then
            CF.Port(NCT_FAN_CONTROL_MODE_REG(index)) = 0    'Manual
        else
            CF.Port(NCT_FAN_CONTROL_MODE_REG(index)) = 64    'Smart Fan IV
        end if
    End Property

    Property Let FanALL(bFull)
        For a = 0 to CF.cf_FAN_COUNT - 1
            FanAuto(a) = 0
            If bFull Then FanPWM(a) = 255 else FanPWM(a) = 5
        Next
    End Property

    '============ Param Functions ===========    
    Property Get GetFans  : GetFans = CF.GetFans   : End Property
    Property Get GetVolts : GetVolts = CF.GetVolts : End Property
    Property Get GetTemps : GetTemps = CF.GetTemps : End Property

    '============ SIO Functions =============
    Property Get Name : Name = CF.Name     : End Property   
    Property Get Find : Find = CF.Find(Me) : End Property 
End Class
<#Module>


<#Module=mdl_IT87x>
const ITE_CHIP_VERSION_REG = &H22
const ITE_FAN_TACHOMETER_16_BIT_ENABLE_REG = &HC
const ITE_SMARTGUARD_MAIN_CONTROL = &H13
const ITE_FAN_CTL_REG = &H14
const ITE_BEEP_REG = &H5C

ITE_SMARTGUARD_PWM_CONTROL = Array(&H15, &H16, &H17, &H88, &H89)
ITE_SMARTGUARD_TEMP_STOP = Array(&H60, &H68, &H70, &H90, &H98)
ITE_SMARTGUARD_TEMP_START = Array(&H61, &H69, &H71, &H91, &H99)
ITE_SMARTGUARD_TEMP_FULL_ON = Array(&H62, &H6A, &H72, &H92, &H9A)
ITE_SMARTGUARD_TEMP_FULL_OFF = Array(&H65, &H6D, &H75, &H95, &H9D)
ITE_SMARTGUARD_START_PWM = Array(&H63, &H6B, &H73, &H93, &H9B)
ITE_SMARTGUARD_CONTROL = Array(&H64, &H6C, &H74, &H94, &H9C)

const IT8512F = &H8512&, IT8712F = &H8712&, IT8716F = &H8716&, IT8718F = &H8718&, IT8720F = &H8720&
const IT8721F = &H8721&, IT8726F = &H8726&, IT8728F = &H8728&, IT8752F = &H8752&, IT8771E = &H8771&, IT8772E = &H8772& 


Class IT87x
    Dim CF, LabelVolts, LabelFans, LabelTemps, FanCPU, TempCPU

    Sub Class_Initialize
        Set CF = New SIO_CommonFunc

        With CF
            .cf_CHIPS_ID        = Array(IT8512F, IT8712F, IT8716F, IT8718F, IT8720F, IT8721F, IT8726F, IT8728F, IT8752F, IT8771E, IT8772E)
            .cf_CHIPS_NAME      = Array("IT8512F", "IT8712F", "IT8716F", "IT8718F", "IT8720F", "IT8721F", "IT8726F", "IT8728F", "IT8752F", "IT8771E", "IT8772E")
            .cf_ENTER           = Array("R=&H87", "R=1", "R=&H55", "IF=&H4E,&HAA,&H55")
            .cf_EXIT            = Array("W=" & sio_CONFIG_CONTROL_REG & ",2")
            .cf_VENDOR_CHECK    = "Me.Port(&H58) = &H90"

            .cf_VOLTAGE_GAIN    = Array(1, 1, 1, 1.68, 4, 1, 1, 1, 1)
            .cf_FAN_SPEED_MSB   = Array(&H18, &H19, &H1A, &H81, &H83)
            .cf_FAN_SPEED_LSB   = Array(&HD, &HE, &HF, &H80, &H82)
            .cf_FAN_SPEED_KOEF  = 675000
            .cf_VOLTAGE_BASE    = &H20
            .cf_TEMP_BASE       = &H29
            .cf_VOLT_COUNT      = 9
            .cf_FAN_COUNT       = 5
            .cf_TEMP_COUNT      = 3
        End With

        FanCPU = Array(0,2) : TempCPU = 1

        LabelVolts = Array("VCore1", "VCore2", "+3.3V", "+5V", "+12V", "Vin5", "Vin6", "Vin7", "VBat")
        LabelFans = Array("Fan1", "Fan2", "Fan3", "Fan4", "Fan5")
        LabelTemps = Array("TMPIN1", "TMPIN2", "TMPIN3") 
    End Sub

    '============ Base Functions =============
    Property Get Family : Family = "IT87x" : End Property

    Sub Init()
        CF.cf_VOLTAGE_MVLSB = 16
        If (CF.cf_Model = IT8721F Or CF.cf_Model = IT8728F Or CF.cf_Model = IT8771E Or CF.cf_Model = IT8772E) Then CF.cf_VOLTAGE_MVLSB = 12
    End Sub

    '========== Control Functions ============
    Property Get FanPWM(index)
        FanPWM = (CF.Port(ITE_SMARTGUARD_PWM_CONTROL(index)) And &H7F) / 1.27
    End Property

    Property Let FanPWM(index, value)
        Dim vpwm
        vpwm = round(value * 1.27, 0)
        If vpwm < 0 Then vpwm = 0
        If vpwm > 127 Then vpwm = 127
        CF.Bit(ITE_SMARTGUARD_MAIN_CONTROL, index) = 1
        CF.Port(ITE_SMARTGUARD_PWM_CONTROL(index)) = vpwm
    End Property
    
    Property Get FanAuto(index)
        FanAuto = Abs(Not (CF.Bit(ITE_SMARTGUARD_MAIN_CONTROL, index) = 1 And CF.Bit(ITE_SMARTGUARD_PWM_CONTROL(index), 7) = 0))
    End Property

    Property Let FanAuto(index, value)
        CF.Bit(ITE_SMARTGUARD_MAIN_CONTROL, index) = 1
        If value = 0 then
            CF.Bit(ITE_SMARTGUARD_PWM_CONTROL(index), 7) = 0
        else
            CF.Port(ITE_SMARTGUARD_PWM_CONTROL(index)) = Array(129, 128, 129)(index)    
        end if
    End Property

    Property Let FanALL(bFull)
        If bFull then
            CF.Port(ITE_SMARTGUARD_MAIN_CONTROL) = 112
            CF.Port(ITE_FAN_CTL_REG) = 215
        else
            CF.Port(ITE_SMARTGUARD_MAIN_CONTROL) = 112
            CF.Port(ITE_FAN_CTL_REG) = 208
        end if
    End Property

    '============ Param Functions ===========    
    Property Get GetFans  : GetFans = CF.GetFans   : End Property
    Property Get GetVolts : GetVolts = CF.GetVolts : End Property
    Property Get GetTemps : GetTemps = CF.GetTemps : End Property

    '============ SIO Functions =============
    Property Get Name : Name = CF.Name     : End Property   
    Property Get Find : Find = CF.Find(Me) : End Property 
End Class
<#Module>


<#Module=mdl_F718x>
FIN_FAN_CONTROL = &H96
FIN_FAN_PWM = Array(&HA3, &HB3, &HC3, &HD3)

const F71858 = &H0507&, F71862 = &H0601&, F71869 = &H0814&, F71882 = &H0541&
const F71889ED = &H0909&, F71889F = &H0723&, F71808 = &H0901&


Class F718x
    Dim CF, LabelVolts, LabelFans, LabelTemps, FanCPU, TempCPU

    Sub Class_Initialize
        Set CF = New SIO_CommonFunc

        With CF
            .cf_CHIPS_ID        = Array(F71858, F71862, F71869, F71882, F71889ED, F71889F, F71808)
            .cf_CHIPS_NAME      = Array("F71858", "F71862", "F71869", "F71882", "F71889ED", "F71889F", "F71808")
            .cf_ENTER           = Array("R=&H87", "R=&H87")
            .cf_EXIT            = Array("R=&HAA", "W=" & sio_CONFIG_CONTROL_REG & ",2")
            .cf_VENDOR_CHECK    = "SIO_ReadWord(&H23) = &H1934&"
            .cf_MONITOR_LDN     = Array(Array(F71858, 2), Array(0, 4))

            .cf_VOLTAGE_GAIN    = Array(1, 1, 1, 1, 1, 1, 1, 1)
            .cf_FAN_SPEED_MSB   = Array(&HA0, &HB0, &HC0, &HD0)
            .cf_FAN_SPEED_LSB   = Array(&HA1, &HB1, &HC1, &HD1)
            .cf_FAN_SPEED_KOEF  = 1500000
            .cf_VOLTAGE_BASE    = &H20
            .cf_TEMP_BASE       = &H70
            .cf_VOLT_COUNT      = 8
            .cf_TEMP_COUNT      = 3
        End With

        FanCPU = 0 : TempCPU = 0

        LabelVolts = Array("Vin1", "Vin2", "Vin3", "Vin4", "Vin5", "Vin6", "+3.3V", "VBat")
        LabelFans = Array("Fan1", "Fan2", "Fan3", "Fan4")
        LabelTemps = Array("T1", "T2", "T3")
    End Sub

    '============ Base Functions =============
    Property Get Family : Family = "F718x" : End Property

    Sub Init()
        CF.cf_FAN_COUNT = IIF(CF.cf_Model = F71882, 4, 3)
    End Sub

    '========== Control Functions ============
    Property Get FanPWM(index)
        FanPWM = CF.Port(FIN_FAN_PWM(index)) / 2.55
    End Property

    Property Let FanPWM(index, value)
        Dim vpwm
        vpwm = round(value * 2.55, 0)
        If vpwm < 0 Then vpwm = 0
        If vpwm > 255 Then vpwm = 255
        CF.Bit(FIN_FAN_CONTROL, index * 2) = 1
        CF.Bit(FIN_FAN_CONTROL, index * 2 + 1) = 1
        CF.Port(FIN_FAN_PWM(index)) = vpwm
    End Property
    
    Property Get FanAuto(index)
        FanAuto = Abs(CF.Bit(FIN_FAN_CONTROL, index * 2)=0 Or CF.Bit(FIN_FAN_CONTROL, index * 2 + 1)=0)
    End Property

    Property Let FanAuto(index, value)
        If value = 0 then
            CF.Bit(FIN_FAN_CONTROL, index * 2) = 1
            CF.Bit(FIN_FAN_CONTROL, index * 2 + 1) = 1
        else
            CF.Bit(FIN_FAN_CONTROL, index * 2) = 0
            CF.Bit(FIN_FAN_CONTROL, index * 2 + 1) = 0
        end if
    End Property

    Property Let FanALL(bFull)
        Dim a
        For a = 0 to CF.cf_FAN_COUNT - 1
            If bFull Then FanPWM(a) = 255 else FanPWM(a) = 0
        next
    End Property

    '============ Param Functions ===========    
    Property Get GetFans  : GetFans = CF.GetFans   : End Property
    Property Get GetVolts : GetVolts = CF.GetVolts : End Property

    Property Get GetTemps
        Dim a, tmp
        If IsEmpty(CF.cf_TEMP_COUNT) Then Exit Property
        Redim tmp(CF.cf_TEMP_COUNT - 1) 
        For a = 0 to ubound(tmp)
            tmp(a) = CF.Port(CF.cf_TEMP_BASE + 2 * (a + 1))
        next
        GetTemps = tmp
    End Property

    '============ SIO Functions =============
    Property Get Name : Name = CF.Name     : End Property   
    Property Get Find : Find = CF.Find(Me) : End Property 
End Class
<#Module>