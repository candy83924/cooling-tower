"""
冷卻水塔水蒸發量計算模組
包含：簡易經驗公式、Merkel 焓差法、濕空氣性質計算
"""
import math

# ===== 濕空氣性質計算模組 =====

def saturation_pressure(T):
    """飽和水蒸氣壓力 (kPa)，Antoine 方程式，T 為溫度 (°C)"""
    if T < 0:
        # 冰面飽和壓力
        return 0.61115 * math.exp((23.036 - T / 333.7) * T / (279.82 + T))
    return 0.61078 * math.exp(17.27 * T / (T + 237.3))


def humidity_ratio(Tdb, Twb, P=101.325):
    """
    比濕度 (kg水/kg乾空氣)
    Tdb: 乾球溫度 °C, Twb: 濕球溫度 °C, P: 大氣壓 kPa
    """
    Pws_wb = saturation_pressure(Twb)
    Ws_wb = 0.622 * Pws_wb / (P - Pws_wb)
    W = ((2501 - 2.326 * Twb) * Ws_wb - 1.006 * (Tdb - Twb)) / \
        (2501 + 1.86 * Tdb - 4.186 * Twb)
    return max(W, 0)


def humidity_ratio_saturated(T, P=101.325):
    """飽和狀態下的比濕度"""
    Ps = saturation_pressure(T)
    if Ps >= P:
        return 0.622  # 防止除零
    return 0.622 * Ps / (P - Ps)


def enthalpy_moist_air(T, W):
    """濕空氣焓值 (kJ/kg乾空氣)，T: 溫度°C，W: 比濕度"""
    return 1.006 * T + W * (2501 + 1.86 * T)


def enthalpy_saturated(T, P=101.325):
    """飽和空氣焓值 (kJ/kg乾空氣)"""
    Ws = humidity_ratio_saturated(T, P)
    return enthalpy_moist_air(T, Ws)


def relative_humidity(Tdb, Twb, P=101.325):
    """相對濕度 (%)"""
    W = humidity_ratio(Tdb, Twb, P)
    Ps = saturation_pressure(Tdb)
    Pw = W * P / (0.622 + W)
    return min(max(Pw / Ps * 100, 0), 100)


def wet_bulb_from_rh(Tdb, RH, P=101.325):
    """從乾球溫度和相對濕度反算濕球溫度（迭代法）"""
    Twb_low, Twb_high = -10, Tdb
    for _ in range(100):
        Twb_mid = (Twb_low + Twb_high) / 2
        rh_calc = relative_humidity(Tdb, Twb_mid, P)
        if abs(rh_calc - RH) < 0.01:
            return Twb_mid
        if rh_calc < RH:
            Twb_low = Twb_mid
        else:
            Twb_high = Twb_mid
    return (Twb_low + Twb_high) / 2


# ===== 簡易經驗公式 =====

def simple_evaporation(Q, T_in, T_out, season='summer'):
    """
    簡易經驗公式：E = C × ΔT × Q / 600
    Q: 循環水量 m³/h
    T_in: 進水溫度 °C
    T_out: 出水溫度 °C
    season: 'summer' or 'winter'
    回傳：蒸發水量 m³/h
    """
    C = 1.0 if season == 'summer' else 0.8
    delta_T = T_in - T_out
    E = C * delta_T * Q / 600
    return E


# ===== Merkel 焓差法 =====

def merkel_integral(T_out, T_in, Q, G, KaV_L=1.5, P=101.325, n_steps=50):
    """
    Merkel 焓差法計算
    T_out: 出水溫度 °C (冷水)
    T_in: 進水溫度 °C (熱水)
    Q: 循環水量 m³/h
    G: 空氣流量 m³/h
    KaV_L: 填料特性係數 (無因次)
    P: 大氣壓力 kPa
    n_steps: 積分步數

    回傳 dict: 包含各項計算結果
    """
    # 水的密度與比熱
    rho_w = 1000  # kg/m³
    Cp_w = 4.186  # kJ/(kg·°C)

    # 空氣密度（近似）
    rho_a = 1.2  # kg/m³

    # 質量流率
    L = Q * rho_w / 3600  # 水的質量流率 kg/s
    G_mass = G * rho_a / 3600  # 空氣質量流率 kg/s

    # 氣水比
    LG_ratio = L / G_mass if G_mass > 0 else float('inf')

    delta_T = T_in - T_out
    dT = delta_T / n_steps

    # Merkel 數值積分 (辛普森法)
    merkel_number = 0
    temperatures = []
    enthalpies_sat = []
    enthalpies_air = []

    # 入口空氣狀態假設（使用濕球溫度接近出水溫度）
    T_wb_in = T_out - 5 if T_out > 10 else T_out - 2
    W_in = humidity_ratio_saturated(T_wb_in, P) * 0.6
    h_air = enthalpy_moist_air(T_wb_in, W_in)

    for i in range(n_steps + 1):
        T_w = T_out + i * dT
        h_sat = enthalpy_saturated(T_w, P)

        temperatures.append(T_w)
        enthalpies_sat.append(h_sat)
        enthalpies_air.append(h_air)

        diff = h_sat - h_air
        if diff <= 0:
            diff = 0.1  # 防止除零

        if i == 0 or i == n_steps:
            merkel_number += Cp_w / diff
        elif i % 2 == 1:
            merkel_number += 4 * Cp_w / diff
        else:
            merkel_number += 2 * Cp_w / diff

        # 更新空氣焓值
        if i < n_steps:
            dh_air = (L * Cp_w * dT) / G_mass if G_mass > 0 else 0
            h_air += dh_air

    merkel_number *= dT / 3  # 辛普森法係數

    # 蒸發水量計算
    W_out_approx = humidity_ratio_saturated((T_in + T_out) / 2, P) * 0.85
    delta_W = W_out_approx - W_in
    E_mass = G_mass * max(delta_W, 0)  # kg/s
    E_volume = E_mass / rho_w * 3600  # m³/h

    # 冷卻效率
    T_approach = T_out - T_wb_in
    T_range = T_in - T_out
    efficiency = T_range / (T_range + T_approach) * 100 if (T_range + T_approach) > 0 else 0

    # 出口空氣狀態
    T_air_out = T_out + (T_in - T_out) * 0.7
    W_air_out = W_in + delta_W
    h_air_out = enthalpy_moist_air(T_air_out, W_air_out)

    return {
        'evaporation_rate': E_volume,           # 蒸發水量 m³/h
        'evaporation_mass': E_mass * 3600,      # 蒸發水量 kg/h
        'LG_ratio': LG_ratio,                   # 氣水比
        'merkel_number': merkel_number,          # Merkel 數
        'efficiency': efficiency,                # 冷卻效率 %
        'T_approach': T_approach,                # 逼近溫度 °C
        'T_range': T_range,                      # 冷卻範圍 °C
        'T_air_out': T_air_out,                  # 出口空氣溫度 °C
        'W_air_out': W_air_out,                  # 出口空氣比濕度
        'h_air_out': h_air_out,                  # 出口空氣焓值
        'temperatures': temperatures,            # 溫度分布
        'enthalpies_sat': enthalpies_sat,         # 飽和焓值分布
        'enthalpies_air': enthalpies_air,         # 空氣焓值分布
    }


# ===== 水損失計算 =====

def calculate_water_losses(Q, T_in, T_out, G, tower_type='counter', season='summer',
                           COC=3.0, KaV_L=1.5, Tdb=35, Twb=28, P=101.325):
    """
    綜合水損失計算
    Q: 循環水量 m³/h
    T_in: 進水溫度 °C
    T_out: 出水溫度 °C
    G: 空氣流量 m³/h
    tower_type: 'counter'(逆流) / 'cross'(橫流)
    season: 'summer' / 'winter'
    COC: 濃縮倍數
    KaV_L: 填料特性係數
    Tdb: 乾球溫度 °C
    Twb: 濕球溫度 °C
    P: 大氣壓力 kPa
    """
    # 簡易公式蒸發量
    E_simple = simple_evaporation(Q, T_in, T_out, season)

    # Merkel 法蒸發量
    merkel_result = merkel_integral(T_out, T_in, Q, G, KaV_L, P)
    E_merkel = merkel_result['evaporation_rate']

    # 飛濺損失
    splash_rate = 0.001 if tower_type == 'counter' else 0.003
    E_splash = Q * splash_rate

    # 排污損失（基於 Merkel 蒸發量）
    if COC > 1:
        E_blowdown = E_merkel / (COC - 1)
    else:
        E_blowdown = E_merkel

    # 總補給水量
    E_total = E_merkel + E_splash + E_blowdown

    # 濕空氣性質
    W_in = humidity_ratio(Tdb, Twb, P)
    h_in = enthalpy_moist_air(Tdb, W_in)
    RH = relative_humidity(Tdb, Twb, P)

    return {
        # 蒸發量
        'E_simple': E_simple,
        'E_merkel': E_merkel,
        'E_splash': E_splash,
        'E_blowdown': E_blowdown,
        'E_total': E_total,

        # Merkel 計算詳細結果
        'merkel_details': merkel_result,

        # 環境空氣性質
        'W_inlet': W_in,
        'h_inlet': h_in,
        'RH': RH,

        # 比例
        'splash_rate_pct': splash_rate * 100,
        'evap_pct': E_merkel / E_total * 100 if E_total > 0 else 0,
        'splash_pct': E_splash / E_total * 100 if E_total > 0 else 0,
        'blowdown_pct': E_blowdown / E_total * 100 if E_total > 0 else 0,
    }


# ===== 敏感度分析 =====

def sensitivity_analysis(base_params, variable='delta_T', values=None):
    """
    敏感度分析：改變單一變數，觀察蒸發量變化
    """
    if values is None:
        values = []

    results = []
    for val in values:
        params = base_params.copy()
        if variable == 'delta_T':
            params['T_in'] = params['T_out'] + val
        elif variable == 'G':
            params['G'] = val
        elif variable == 'COC':
            params['COC'] = val
        elif variable == 'Q':
            params['Q'] = val

        r = calculate_water_losses(**params)
        results.append({
            'variable_value': val,
            'E_merkel': r['E_merkel'],
            'E_total': r['E_total'],
            'E_blowdown': r['E_blowdown'],
            'efficiency': r['merkel_details']['efficiency'],
        })
    return results
