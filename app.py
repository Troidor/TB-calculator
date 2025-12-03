import streamlit as st
import math
from copy import deepcopy

# Embedded workbook constants
VALUES = {"A1": "Target Leadership", "B1": None, "C1": None, "D1": "HEALTH RATIO (MAX/MIN)", "E1": None, "F1": None, "G1": None, "H1": None, "I1": None, "J1": None, "K1": None, "L1": None, "M1": None, "A2": "Target Leadership:", "B2": 46800, "C2": None, "D2": 1.0310722561617864, "E2": None, "F2": None, "G2": None, "H2": None, "I2": None, "J2": None, "K2": None, "L2": None, "M2": None, "A3": "Calculated Leadership:", "B3": 46520, "C3": 280, "D3": None, "E3": None, "F3": None, "G3": None, "H3": None, "I3": None, "J3": None, "K3": None, "L3": None, "M3": None, "A4": "Ratio", "B4": 0.161, "C4": None, "D4": None, "E4": None, "F4": None, "G4": None, "H4": None, "I4": None, "J4": None, "K4": None, "L4": None, "M4": None, "A5": "Target Max Health", "B5": 11754288, "C5": None, "D5": None, "E5": None, "F5": None, "G5": None, "H5": None, "I5": None, "J5": None, "K5": None, "L5": None, "M5": None, "A6": None, "B6": None, "C6": None, "D6": None, "E6": None, "F6": None, "G6": None, "H6": None, "I6": None, "J6": None, "K6": None, "L6": None, "M6": None, "A7": "Normal Troops", "B7": "Army", "C7": "Leadership", "D7": "Rank", "E7": "Leadership per Active Troop", "F7": "Troop Base Health", "G7": "INPUT Troop Health Bonus", "H7": "Total Health Bonus", "I7": "Target Health per Stack", "J7": "Actual Health per Stack", "K7": "Troop Level", "L7": None, "M7": None, "A8": "G6 Flyings", "B8": 200, "C8": 4000, "D8": 7, "E8": 20, "F8": 57000, "G8": 0, "H8": 57000, "I8": 11626837.2, "J8": 11400000, "K8": 6, "L8": None, "M8": None, "A9": "G5 Flyings", "B9": 390, "C9": 7800, "D9": 3, "E9": 20, "F9": 30000, "G9": 0, "H9": 30000, "I9": 11716828.2, "J9": 11700000, "K9": 5, "L9": None, "M9": None, "A10": "G6 Mounted", "B10": 2040, "C10": 4080, "D10": 6, "E10": 2, "F10": 5700, "G10": 0, "H10": 5700, "I10": 11673632.52, "J10": 11628000, "K10": 6, "L10": None, "M10": None, "A11": "G5 Mounted", "B11": 3720, "C11": 7440, "D11": 2, "E11": 2, "F11": 3150, "G11": 0, "H11": 3150, "I11": 11745625.32, "J11": 11718000, "K11": 5, "L11": None, "M11": None, "A12": "G6 Ranged", "B12": 4040, "C12": 4040, "D12": 8, "E12": 1, "F12": 2820, "G12": 0, "H12": 2820, "I12": 11398860, "J12": 11392800, "K12": 6, "L12": None, "M12": None, "A13": "G6 Melee", "B13": 4140, "C13": 4140, "D13": 5, "E13": 1, "F13": 2820, "G13": 0, "H13": 2820, "I13": 11683231.56, "J13": 11674800, "K13": 6, "L13": None, "M13": None, "A14": "G5 Ranged", "B14": 7490, "C14": 7490, "D14": 4, "E14": 1, "F14": 1560, "G14": 0, "H14": 1560, "I14": 11698830, "J14": 11684400, "K14": 5, "L14": None, "M14": None, "A15": "G5 Melee", "B15": 7530, "C15": 7530, "D15": 1, "E15": 1, "F15": 1560, "G15": 0, "H15": 1560, "I15": 11754288, "J15": 11746800, "K15": 5, "L15": None, "M15": None, "A16": None, "B16": 29550, "C16": 46520, "D16": None, "E16": None, "F16": None, "G16": None, "H16": None, "I16": None, "J16": None, "K16": None, "L16": None, "M16": None, "A17": None, "B17": None, "C17": None, "D17": None, "E17": None, "F17": None, "G17": None, "H17": None, "I17": None, "J17": None, "K17": None, "L17": None, "M17": None, "A18": None, "B18": "Army", "C18": "Leadership", "D18": "Rank", "E18": "Leadership per Active Troop", "F18": "Troop Base Health", "G18": "INPUT Troop Health Bonus", "H18": "Total Health Bonus", "I18": "Target Health per Stack", "J18": "Actual Health per Stack", "K18": "Troop Level", "L18": None, "M18": None, "A19": "M6 Jungle Melee", "B19": 20, "C19": 7800000, "D19": None, "E19": 34, "F19": 390000, "G19": 0, "H19": 390000, "I19": 11673632.52, "J19": 7800000, "K19": 6, "L19": None, "M19": None, "A20": "M5 Ettin Melee", "B20": 50, "C20": 7200000, "D20": None, "E20": 23, "F20": 144000, "G20": 0, "H20": 144000, "I20": 7799220, "J20": 7200000, "K20": 5, "L20": None, "M20": None, "A21": "M4 Phoenix Flyings", "B21": 140, "C21": 7140000, "D21": None, "E21": 15, "F21": 51000, "G21": 0, "H21": 51000, "I21": 7199280, "J21": 7140000, "K21": 4, "L21": None, "M21": None, "A22": None, "B22": None, "C22": None, "D22": None, "E22": None, "F22": None, "G22": None, "H22": None, "I22": None, "J22": None, "K22": None, "L22": None, "M22": None, "A23": "Mercs", "B23": "Army", "C23": "Authority", "D23": "Rank", "E23": "Leadership per Active Troop", "F23": "Troop Base Health", "G23": "INPUT Troop Health Bonus", "H23": "Total Health Bonus", "I23": "Target Health per Stack", "J23": "Actual Health per Stack", "K23": "Troop Level", "L23": None, "M23": None, "A24": "Epic Monster Hunter (Green)", "B24": 100, "C24": None, "D24": None, "E24": 1, "F24": 75000, "G24": 0, "H24": 75000, "I24": 7799220, "J24": 7500000, "K24": None, "L24": None, "M24": None, "A25": "Epic Monster Hunter (Yellow)", "B25": 660, "C25": None, "D25": None, "E25": 1, "F25": 11250, "G25": 0, "H25": 11250, "I25": 7499250, "J25": 7425000, "K25": None, "L25": None, "M25": None}
LABELS = {"A1": "Target Leadership", "A2": "Target Leadership:", "A3": "Calculated Leadership:", "A4": "Ratio", "A5": "Target Max Health", "A6": None, "A7": "Normal Troops", "A8": "G6 Flyings", "A9": "G5 Flyings", "A10": "G6 Mounted", "A11": "G5 Mounted", "A12": "G6 Ranged", "A13": "G6 Melee", "A14": "G5 Ranged", "A15": "G5 Melee", "A16": None, "A17": None, "A18": None, "A19": "M6 Jungle Melee", "A20": "M5 Ettin Melee", "A21": "M4 Phoenix Flyings", "A22": None, "A23": "Mercs", "A24": "Epic Monster Hunter (Green)", "A25": "Epic Monster Hunter (Yellow)"}

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
