"""
Parser for Walton Compressor Calorimeter Evaluation Program PDF reports.
Extracts key performance parameters from each report's text content.
"""
import re
import pdfplumber

# Regex patterns for each field, matched against the flattened page text
PATTERNS = {
    'sample_id':        r'Sample ID\s+(ID-[\d.]+)',
    'model':            r'Model\s+(\S+)',
    'test_point':       r'Test Point Nr\.\s+(\S+)',
    'rpm_rated':        r'Speed \[RPM\]\s+([\d.]+)\s+[\d.]+',
    'rpm_actual':       r'Speed \[RPM\]\s+[\d.]+\s+([\d.]+)',
    'cooling_capacity': r'Cooling Capacity\s+([\d.]+)\s*\(W\)',
    'input_power':      r'Primary input power P\s*\(W\)\s+([\d.]+)',
    'cop':              r'COP\s+([\d.]+)\s*\(W/W\)',
    'current':          r'Primary input Current \(A\)\s+([\d.]+)',
    'inv_eff':          r'Inverter efficiency \(%\)\s+([\d.]+)',
    'inv_loss':         r'Inverter loss \(W\)\s+([\d.]+)',
    'power_factor':     r'Primary power factor \(-\)\s+([\d.\-]+)',
    'shell_top_temp':   r'Shell top temp[^\d\-]*([\d.\-]+)',
    'discharge_temp':   r'Discharge temp[^\d\-]*([\d.\-]+)',
    'mass_flow':        r'Mass Flow refrigerant \(g/h\)\s+([\d.\-]+)',
    'vol_eff':          r'Volumetric efficiency\s+([\d.\-]+)',
    'isen_eff':         r'Isentropic efficiency\s+([\d.\-]+)',
    'start_datetime':   r'Start Date / Time\s+([\d\-]+\s[\d:]+)',
}

NUMERIC_FIELDS = [
    'rpm_rated', 'rpm_actual', 'cooling_capacity', 'input_power', 'cop',
    'current', 'inv_eff', 'inv_loss', 'power_factor', 'shell_top_temp',
    'discharge_temp', 'mass_flow', 'vol_eff', 'isen_eff'
]

# Maps internal keys -> display metadata used across UI, xlsx, pdf export
PARAM_META = {
    'rpm_actual':       {'label': 'Actual RPM',          'fmt': '{:.0f}'},
    'cooling_capacity': {'label': 'Cooling Cap. (W)',    'fmt': '{:.1f}'},
    'input_power':      {'label': 'Input Power (W)',     'fmt': '{:.2f}'},
    'cop':              {'label': 'COP (W/W)',           'fmt': '{:.3f}'},
    'current':          {'label': 'Current (A)',         'fmt': '{:.3f}'},
    'inv_eff':          {'label': 'Inv. Efficiency (%)', 'fmt': '{:.2f}'},
    'inv_loss':         {'label': 'Inv. Loss (W)',       'fmt': '{:.2f}'},
    'power_factor':     {'label': 'Power Factor',        'fmt': '{:.3f}'},
    'shell_top_temp':   {'label': 'Shell Top Temp (°C)', 'fmt': '{:.1f}'},
    'discharge_temp':   {'label': 'Discharge Temp (°C)', 'fmt': '{:.1f}'},
    'mass_flow':        {'label': 'Mass Flow (g/h)',     'fmt': '{:.0f}'},
    'vol_eff':          {'label': 'Volumetric Eff.',     'fmt': '{:.3f}'},
    'isen_eff':         {'label': 'Isentropic Eff.',     'fmt': '{:.3f}'},
}

PARAM_ORDER = [
    'rpm_actual', 'cooling_capacity', 'input_power', 'cop', 'current',
    'inv_eff', 'inv_loss', 'power_factor', 'shell_top_temp',
    'discharge_temp', 'mass_flow', 'vol_eff', 'isen_eff'
]


def parse_pdf_bytes(file_bytes, fallback_id=None, source_name=""):
    """Parse a single calorimeter PDF report (bytes) and return a dict of extracted fields."""
    import io
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text = pdf.pages[0].extract_text() or ""

    row = {'source_file': source_name}
    for key, pattern in PATTERNS.items():
        m = re.search(pattern, text)
        row[key] = m.group(1) if m else None

    # numeric coercion
    for key in NUMERIC_FIELDS:
        val = row.get(key)
        if val is not None:
            try:
                row[key] = float(val)
            except ValueError:
                row[key] = None

    if not row.get('sample_id'):
        row['sample_id'] = fallback_id or "UNKNOWN"

    # bucket by rated RPM (nominal test point speed), rounded to nearest 10
    if row.get('rpm_rated') is not None:
        row['rpm_tier'] = int(round(row['rpm_rated'] / 10.0) * 10)
    else:
        row['rpm_tier'] = None

    return row
