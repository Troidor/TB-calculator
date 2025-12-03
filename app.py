import streamlit as st
import math
from copy import deepcopy

# Embedded workbook constants
VALUES = {"A1": "Target Leadership", "B1": null, "C1": null, "D1": "HEALTH RATIO (MAX/MIN)", "E1": null, "F1": null, "G1": null, "H1": null, "I1": null, "J1": null, "K1": null, "L1": null, "M1": null, "A2": "Target Leadership:", "B2": 46800, "C2": null, "D2": 1.0310722561617864, "E2": null, "F2": null, "G2": null, "H2": null, "I2": null, "J2": null, "K2": null, "L2": null, "M2": null, "A3": "Calculated Leadership:", "B3": 46520, "C3": 280, "D3": null, "E3": null, "F3": null, "G3": null, "H3": null, "I3": null, "J3": null, "K3": null, "L3": null, "M3": null, "A4": "Ratio", "B4": 0.161, "C4": null, "D4": null, "E4": null, "F4": null, "G4": null, "H4": null, "I4": null, "J4": null, "K4": null, "L4": null, "M4": null, "A5": "Target Max Health", "B5": 11754288, "C5": null, "D5": null, "E5": null, "F5": null, "G5": null, "H5": null, "I5": null, "J5": null, "K5": null, "L5": null, "M5": null, "A6": null, "B6": null, "C6": null, "D6": null, "E6": null, "F6": null, "G6": null, "H6": null, "I6": null, "J6": null, "K6": null, "L6": null, "M6": null, "A7": "Normal Troops", "B7": "Army", "C7": "Leadership", "D7": "Rank", "E7": "Leadership per Active Troop", "F7": "Troop Base Health", "G7": "INPUT Troop Health Bonus", "H7": "Total Health Bonus", "I7": "Target Health per Stack", "J7": "Actual Health per Stack", "K7": "Troop Level", "L7": null, "M7": null, "A8": "G6 Flyings", "B8": 200, "C8": 4000, "D8": 7, "E8": 20, "F8": 57000, "G8": 0, "H8": 57000, "I8": 11626837.2, "J8": 11400000, "K8": 6, "L8": null, "M8": null, "A9": "G5 Flyings", "B9": 390, "C9": 7800, "D9": 3, "E9": 20, "F9": 30000, "G9": 0, "H9": 30000, "I9": 11716828.2, "J9": 11700000, "K9": 5, "L9": null, "M9": null, "A10": "G6 Mounted", "B10": 2040, "C10": 4080, "D10": 6, "E10": 2, "F10": 5700, "G10": 0, "H10": 5700, "I10": 11673632.52, "J10": 11628000, "K10": 6, "L10": null, "M10": null, "A11": "G5 Mounted", "B11": 3720, "C11": 7440, "D11": 2, "E11": 2, "F11": 3150, "G11": 0, "H11": 3150, "I11": 11745625.32, "J11": 11718000, "K11": 5, "L11": null, "M11": null, "A12": "G6 Ranged", "B12": 4040, "C12": 4040, "D12": 8, "E12": 1, "F12": 2820, "G12": 0, "H12": 2820, "I12": 11398860, "J12": 11392800, "K12": 6, "L12": null, "M12": null, "A13": "G6 Melee", "B13": 4140, "C13": 4140, "D13": 5, "E13": 1, "F13": 2820, "G13": 0, "H13": 2820, "I13": 11683231.56, "J13": 11674800, "K13": 6, "L13": null, "M13": null, "A14": "G5 Ranged", "B14": 7490, "C14": 7490, "D14": 4, "E14": 1, "F14": 1560, "G14": 0, "H14": 1560, "I14": 11698830, "J14": 11684400, "K14": 5, "L14": null, "M14": null, "A15": "G5 Melee", "B15": 7530, "C15": 7530, "D15": 1, "E15": 1, "F15": 1560, "G15": 0, "H15": 1560, "I15": 11754288, "J15": 11746800, "K15": 5, "L15": null, "M15": null, "A16": null, "B16": 29550, "C16": 46520, "D16": null, "E16": null, "F16": null, "G16": null, "H16": null, "I16": null, "J16": null, "K16": null, "L16": null, "M16": null, "A17": null, "B17": null, "C17": null, "D17": null, "E17": null, "F17": null, "G17": null, "H17": null, "I17": null, "J17": null, "K17": null, "L17": null, "M17": null, "A18": null, "B18": "Army", "C18": "Leadership", "D18": "Rank", "E18": "Leadership per Active Troop", "F18": "Troop Base Health", "G18": "INPUT Troop Health Bonus", "H18": "Total Health Bonus", "I18": "Target Health per Stack", "J18": "Actual Health per Stack", "K18": "Troop Level", "L18": null, "M18": null, "A19": "M6 Jungle Melee", "B19": 20, "C19": 7800000, "D19": null, "E19": 34, "F19": 390000, "G19": 0, "H19": 390000, "I19": 11673632.52, "J19": 7800000, "K19": 6, "L19": null, "M19": null, "A20": "M5 Ettin Melee", "B20": 50, "C20": 7200000, "D20": null, "E20": 23, "F20": 144000, "G20": 0, "H20": 144000, "I20": 7799220, "J20": 7200000, "K20": 5, "L20": null, "M20": null, "A21": "M4 Phoenix Flyings", "B21": 140, "C21": 7140000, "D21": null, "E21": 15, "F21": 51000, "G21": 0, "H21": 51000, "I21": 7199280, "J21": 7140000, "K21": 4, "L21": null, "M21": null, "A22": null, "B22": null, "C22": null, "D22": null, "E22": null, "F22": null, "G22": null, "H22": null, "I22": null, "J22": null, "K22": null, "L22": null, "M22": null, "A23": "Mercs", "B23": "Army", "C23": "Authority", "D23": "Rank", "E23": "Leadership per Active Troop", "F23": "Troop Base Health", "G23": "INPUT Troop Health Bonus", "H23": "Total Health Bonus", "I23": "Target Health per Stack", "J23": "Actual Health per Stack", "K23": "Troop Level", "L23": null, "M23": null, "A24": "Epic Monster Hunter (Green)", "B24": 100, "C24": null, "D24": null, "E24": 1, "F24": 75000, "G24": 0, "H24": 75000, "I24": 7799220, "J24": 7500000, "K24": null, "L24": null, "M24": null, "A25": "Epic Monster Hunter (Yellow)", "B25": 660, "C25": null, "D25": null, "E25": 1, "F25": 11250, "G25": 0, "H25": 11250, "I25": 7499250, "J25": 7425000, "K25": null, "L25": null, "M25": null}
LABELS = {"A1": "Target Leadership", "A2": "Target Leadership:", "A3": "Calculated Leadership:", "A4": "Ratio", "A5": "Target Max Health", "A6": null, "A7": "Normal Troops", "A8": "G6 Flyings", "A9": "G5 Flyings", "A10": "G6 Mounted", "A11": "G5 Mounted", "A12": "G6 Ranged", "A13": "G6 Melee", "A14": "G5 Ranged", "A15": "G5 Melee", "A16": null, "A17": null, "A18": null, "A19": "M6 Jungle Melee", "A20": "M5 Ettin Melee", "A21": "M4 Phoenix Flyings", "A22": null, "A23": "Mercs", "A24": "Epic Monster Hunter (Green)", "A25": "Epic Monster Hunter (Yellow)"}

NORMAL_ROWS = [8, 9, 10, 11, 12, 13, 14, 15]
MERC_ROWS = [19, 20, 21, 22, 23, 24, 25, 26]

def safe_float(v, d=0.0):
    try:
        if v is None: return d
        return float(v)
    except:
        return d

def rounddown_to10(x):
    try:
        return math.floor(float(x)/10.0)*10
    except:
        return 0.0

st.set_page_config(page_title='TB Calculator Final', layout='wide')
st.title('TB Calculator — Final (locked formulas)')

st.markdown('Enter B2 and percentages for G8–G15 and G19–G26. Enter percentages like 100 = 100% (they will be converted to decimals). Hidden constants from the workbook are embedded.')

# Inputs
inputs = {}
default_b2 = safe_float(VALUES.get('B2', 0))
inputs['B2'] = st.number_input(f"{LABELS.get('A2', 'B2')} (B2)", value=default_b2, format='%.6g')

for r in NORMAL_ROWS + MERC_ROWS:
    coordG = f'G{r}'
    label = LABELS.get(f'A{r}', coordG)
    default = safe_float(VALUES.get(coordG, 0))
    v = st.number_input(f"{label} ({coordG}) - %", value=default, format='%.6g')
    inputs[coordG] = v/100.0

if st.button('Calculate'):
    vm = deepcopy(VALUES)
    # inject inputs
    for k,v in inputs.items():
        vm[k] = v

    # helper to compute B4 and B5
    def compute_B4(vm):
        G14 = safe_float(vm.get('G14',0))
        G15 = safe_float(vm.get('G15',0))
        vals = [safe_float(vm.get(f'G25',0)) for r in range(8,15)]
        avg = sum(vals)/len(vals) if vals else 0
        base = 0.161  # 16.1%
        if G15 <= G14:
            return base
        else:
            if avg == 0 or G15 == 0:
                return base
            return base / (G15 / avg)

    # initial H = F*(1+G)
    for r in NORMAL_ROWS + MERC_ROWS:
        vm[f'H25'] = safe_float(vm.get(f'F25',0)) * (1 + safe_float(vm.get(f'G25',0)))

    # compute B4 and B5 initial
    vm['B4'] = compute_B4(vm)
    vm['B5'] = vm['B4'] * safe_float(vm.get('B2',0)) * safe_float(vm.get('H15',0)) / (safe_float(vm.get('E15',1)) or 1)

    # Ensure I mapping keys exist
    for key in ['I8','I9','I10','I11','I12','I13','I14','I15','I19','I20','I21','I24','I25']:
        vm.setdefault(key, safe_float(vm.get(key, 0)))

    # Iterative solver because of circular dependencies
    mapping = {
        'I8':'J10','I9':'J11','I10':'J13','I11':'J15',
        'I12':'J8','I13':'J14','I14':'J9','I15':'B5',
        'I19':'J12','I20':'J12','I21':'J12','I24':'J12','I25':'J12'
    }

    for _ in range(200):
        changed = False

        # compute B from current I and H
        for r in NORMAL_ROWS + MERC_ROWS:
            I = safe_float(vm.get(f'I25',0))
            H = safe_float(vm.get(f'H25',1))
            if H == 0:
                B = 0.0
            else:
                B = rounddown_to10(I / H)
            if abs(safe_float(vm.get(f'B25',0)) - B) > 1e-9:
                vm[f'B25'] = B
                changed = True

        # compute J = H * B
        for r in NORMAL_ROWS + MERC_ROWS:
            H = safe_float(vm.get(f'H25',0))
            Bval = safe_float(vm.get(f'B25',0))
            J = H * Bval
            if abs(safe_float(vm.get(f'J25',0)) - J) > 1e-9:
                vm[f'J25'] = J
                changed = True

        # update I according to mapping (J-based multiplied by 0.9999)
        for i_coord, ref in mapping.items():
            if ref.startswith('J'):
                refval = safe_float(vm.get(ref,0))
                newI = refval * 0.9999
            else:
                newI = safe_float(vm.get(ref,0))
            if abs(safe_float(vm.get(i_coord,0)) - newI) > 1e-6:
                vm[i_coord] = newI
                changed = True

        # recompute H (depends on F and G)
        for r in NORMAL_ROWS + MERC_ROWS:
            vm[f'H25'] = safe_float(vm.get(f'F25',0)) * (1 + safe_float(vm.get(f'G25',0)))

        # recompute B4 and B5
        vm['B4'] = compute_B4(vm)
        vm['B5'] = vm['B4'] * safe_float(vm.get('B2',0)) * safe_float(vm.get('H15',0)) / (safe_float(vm.get('E15',1)) or 1)

        if not changed:
            break

    # compute C = B * E for all rows
    for r in NORMAL_ROWS + MERC_ROWS:
        vm[f'C25'] = safe_float(vm.get(f'B25',0)) * safe_float(vm.get(f'E25',0))

    # display tables
    def table_rows(start, end):
        rows = []
        for r in range(start, end+1):
            rows.append([vm.get(f'A25',''), vm.get(f'B25',''), vm.get(f'C25','')])
        return rows

    st.header('Results — A8:C16')
    st.table(table_rows(8,16))
    st.header('Results — A19:C26')
    st.table(table_rows(19,26))
