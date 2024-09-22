' 计算湿空气相关物性参数的函数
Function AirProperties(propertyType As String, param1 As String, value1 As Double, param2 As String, value2 As Double) As Double
    Dim Tdb As Double ' 干球温度
    Dim Twb As Double ' 湿球温度
    Dim RH As Double ' 相对湿度
    Dim W As Double ' 含湿量 (kg 水蒸气 / kg 干空气)
    Dim result As Double ' 结果变量

    ' 根据输入参数确定温度和相对湿度
    If param1 = "T_dry" And param2 = "T_wet" Then
        Tdb = value1
        Twb = value2
        RH = AirRelativeHumidity(Tdb, Twb)
    ElseIf param1 = "T_dry" And param2 = "RH" Then
        Tdb = value1
        RH = value2
        Twb = AirWetBulbTemperature(Tdb, RH)
    ElseIf param1 = "T_wet" And param2 = "RH" Then
        Twb = value1
        RH = value2
        Tdb = AirDryBulbTemperature(Twb, RH)
    Else
        AirProperties = CVErr(xlErrValue) ' 返回错误值
        Exit Function
    End If

    ' 计算含湿量
    W = AirHumidityRatio(Tdb, Twb)

    ' 根据用户请求计算不同的属性
    Select Case propertyType
        Case "Enthalpy"
            ' 计算比焓 (kJ/kg 干空气)
            result = 1.006 * Tdb + W * (2501 + 1.805 * Tdb)
            
        Case "Density"
            ' 计算湿空气密度 (kg/m3)
            result = AirDensity(Tdb, Twb)
            
        Case "SpecificHeat"
            ' 计算比热容 (kJ/kg·K)
            result = AirSpecificHeat(Tdb, Twb)
            
        Case "SpecificVolume"
            ' 计算比容 (m3/kg)
            result = AirSpecificVolume(Tdb, Twb)
            
        Case "DewPoint"
            ' 计算露点温度 (°C)
            result = AirDewPoint(Tdb, Twb)
            
        Case "WetBulbTemperature"
            ' 输出湿球温度 (°C)
            result = Twb
            
        Case "HumidityRatio"
            ' 输出含湿量 (kg 水蒸气 / kg 干空气)
            result = W
            
        Case "RH"
            ' 输出相对湿度 (%)
            result = RH
            
        Case Else
            ' 无效的属性类型
            result = CVErr(xlErrValue)
    End Select
    
    ' 返回计算结果
    AirProperties = result
End Function

' 计算湿空气密度的函数
Function AirDensity(Tdb As Double, Twb As Double) As Double
    Dim W As Double ' 含湿量
    Dim T_k As Double ' 温度 (K)
    Dim p As Double ' 大气压 (Pa)
    Dim R_da As Double ' 干空气气体常数 (J/kg·K)
    Dim R_v As Double ' 水蒸气气体常数 (J/kg·K)
    
    p = 101325 ' 标准大气压 (Pa)
    R_da = 287.05 ' 干空气气体常数
    R_v = 461.5 ' 水蒸气气体常数
    T_k = Tdb + 273.15 ' 转换为绝对温度 (K)
    
    ' 计算含湿量
    W = AirHumidityRatio(Tdb, Twb)
    
    ' 计算湿空气密度 (kg/m3)
    AirDensity = p / (R_da * T_k * (1 + 1.6078 * W))
End Function

' 计算湿空气比热容的函数
Function AirSpecificHeat(Tdb As Double, Twb As Double) As Double
    Dim W As Double ' 含湿量
    
    ' 计算含湿量
    W = AirHumidityRatio(Tdb, Twb)
    
    ' 计算比热容 (kJ/kg·K)
    AirSpecificHeat = 1.006 + W * 1.805
End Function

' 计算湿空气比容的函数
Function AirSpecificVolume(Tdb As Double, Twb As Double) As Double
    Dim rho As Double ' 密度 (kg/m3)
    
    ' 调用密度计算函数
    rho = AirDensity(Tdb, Twb)
    
    ' 计算比容 (m3/kg)
    AirSpecificVolume = 1 / rho
End Function

' 计算露点温度的函数
Function AirDewPoint(Tdb As Double, Twb As Double) As Double
    Dim pv As Double ' 水蒸气分压力 (Pa)
    
    ' 调用函数计算水蒸气分压力
    pv = AirWaterVaporPressure(Tdb, Twb)
    
    ' 计算露点温度 (°C)
    AirDewPoint = (243.04 * Log(pv / 610.78)) / (17.625 - Log(pv / 610.78))
End Function

' 计算水蒸气分压力的函数
Function AirWaterVaporPressure(Tdb As Double, Twb As Double) As Double
    Dim p As Double ' 大气压 (Pa)
    Dim pvs_wb As Double ' 湿球温度下的饱和水蒸气压 (Pa)
    p = 101325 ' 标准大气压 (Pa)
    
    ' 计算湿球温度下的饱和水蒸气压 (Tetens公式)
    pvs_wb = 610.78 * Exp((17.27 * Twb) / (Twb + 237.3))
    
    ' 计算水蒸气分压力
    AirWaterVaporPressure = pvs_wb - (p - pvs_wb) * (Tdb - Twb) / (1544 - 1.44 * Twb)
End Function

' 计算相对湿度的函数
Function AirRelativeHumidity(Tdb As Double, Twb As Double) As Double
    Dim pvs_db As Double ' 干球温度下的饱和水蒸气压 (Pa)
    Dim pv As Double ' 水蒸气分压力 (Pa)
    
    ' 计算干球温度下的饱和水蒸气压 (Tetens公式)
    pvs_db = 610.78 * Exp((17.27 * Tdb) / (Tdb + 237.3))
    
    ' 调用函数计算水蒸气分压力
    pv = AirWaterVaporPressure(Tdb, Twb)
    
    ' 计算相对湿度 (%)
    AirRelativeHumidity = (pv / pvs_db) * 100
End Function

' 计算含湿量的函数
Function AirHumidityRatio(Tdb As Double, Twb As Double) As Double
    Dim pv As Double ' 水蒸气分压力 (Pa)
    Dim p As Double ' 大气压 (Pa)
    p = 101325 ' 标准大气压 (Pa)
    
    ' 调用函数计算水蒸气分压力
    pv = AirWaterVaporPressure(Tdb, Twb)
    
    ' 计算含湿量 (kg 水蒸气 / kg 干空气)
    AirHumidityRatio = 0.622 * (pv / (p - pv))
End Function

' 计算湿球温度的函数
Function AirWetBulbTemperature(Tdb As Double, RH As Double) As Double
    Dim pv As Double ' 水蒸气分压力 (Pa)
    Dim Twb As Double ' 湿球温度 (°C)
    
    ' 通过相对湿度计算水蒸气分压力
    pv = (RH / 100) * (610.78 * Exp((17.27 * Tdb) / (Tdb + 237.3)))
    
    ' 初始湿球温度猜测
    Twb = Tdb
    Do While AirRelativeHumidity(Tdb, Twb) <> RH
        Twb = Twb - 0.1 ' 逐步逼近
    Loop
    
    AirWetBulbTemperature = Twb
End Function

' 计算干球温度的函数（根据湿球温度和相对湿度）
Function AirDryBulbTemperature(Twb As Double, RH As Double) As Double
    Dim Tdb As Double
    Dim error As Double
    Dim iteration As Integer
    
    ' 初始猜测干球温度等于湿球温度
    Tdb = Twb
    error = 0.01
    iteration = 0
    
    ' 通过迭代逼近干球温度
    Do While Abs(AirRelativeHumidity(Tdb, Twb) - RH) > error And iteration < 100
        Tdb = Tdb + 0.1 ' 每次增加0.1度
        iteration = iteration + 1
    Loop
    
    AirDryBulbTemperature = Tdb
End Function

