"""
Walton Compressor Calorimeter Comparison Dashboard
Flask backend: ingest a folder of PDF calorimeter reports (organized as
MainFolder/ID-xxxx.xx/*.pdf), extract key parameters, and present RPM-wise
comparison tables across compressor IDs, with XLSX / PDF export.
"""
import io
import os
import re
import gc
from collections import defaultdict

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from parser import parse_pdf_bytes, PARAM_META, PARAM_ORDER

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB

# Simple in-memory store (single-user local tool)
STORE = {
    'rows': []   # list of parsed row dicts
}


def _folder_id_from_path(path):
    """Extract the compressor ID folder name from a relative path like
    'MainFolder/ID-7338.01/report.pdf' -> 'ID-7338.01'"""
    parts = [p for p in re.split(r'[\\/]', path) if p]
    if len(parts) >= 2:
        return parts[-2]
    return None


@app.route('/')
def index():
    return render_template('index.html', param_meta=PARAM_META, param_order=PARAM_ORDER)


@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files received'}), 400

    rows = []
    errors = []
    for f in files:
        filename = f.filename or ""
        if not filename.lower().endswith('.pdf'):
            continue
        try:
            data = f.read()
            folder_id = _folder_id_from_path(filename)
            row = parse_pdf_bytes(data, fallback_id=folder_id, source_name=os.path.basename(filename))
            rows.append(row)
            
            # Free up memory immediately after processing each file
            f.close()
            del data
            gc.collect()
            
        except Exception as e:
            errors.append({'file': filename, 'error': str(e)})

    STORE['rows'] = rows

    return jsonify({
        'count': len(rows),
        'errors': errors,
        'data': build_comparison_payload(rows)
    })


def build_comparison_payload(rows):
    """Group parsed rows by rpm_tier -> compressor_id -> row, plus summary stats."""
    tiers = defaultdict(dict)
    ids_seen = set()
    for r in rows:
        tier = r.get('rpm_tier')
        sid = r.get('sample_id')
        if tier is None or sid is None:
            continue
        tiers[tier][sid] = r
        ids_seen.add(sid)

    sorted_tiers = sorted(tiers.keys())
    payload_tiers = []
    for t in sorted_tiers:
        id_rows = tiers[t]
        payload_tiers.append({
            'tier': t,
            'ids': sorted(id_rows.keys()),
            'rows': id_rows
        })

    # summary KPIs
    cops = [r['cop'] for r in rows if r.get('cop') is not None]
    summary = {
        'total_reports': len(rows),
        'unique_ids': len(ids_seen),
        'rpm_tiers': len(sorted_tiers),
        'avg_cop': round(sum(cops) / len(cops), 3) if cops else None,
        'max_cop': round(max(cops), 3) if cops else None,
        'min_cop': round(min(cops), 3) if cops else None,
    }

    return {
        'tiers': payload_tiers,
        'ids': sorted(ids_seen),
        'summary': summary
    }


@app.route('/data')
def data():
    return jsonify(build_comparison_payload(STORE['rows']))


@app.route('/export/xlsx', methods=['POST'])
def export_xlsx():
    selected = request.json.get('params', PARAM_ORDER)
    buf = build_xlsx(STORE['rows'], selected)
    return send_file(buf, as_attachment=True,
                      download_name='calorimeter_comparison.xlsx',
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/export/pdf', methods=['POST'])
def export_pdf():
    selected = request.json.get('params', PARAM_ORDER)
    buf = build_pdf(STORE['rows'], selected)
    return send_file(buf, as_attachment=True,
                      download_name='calorimeter_comparison.pdf',
                      mimetype='application/pdf')


def build_xlsx(rows, selected_params):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    payload = build_comparison_payload(rows)
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(bold=True, size=14, color="1F3864")
    best_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    thin = Side(style='thin', color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')

    for tier_block in payload['tiers']:
        tier = tier_block['tier']
        ids = tier_block['ids']
        id_rows = tier_block['rows']
        ws = wb.create_sheet(title=f"{tier} RPM")

        ws.cell(row=1, column=1, value=f"Calorimeter Comparison — {tier} RPM Test Point").font = title_font
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(selected_params) + 1)

        header_row = 3
        ws.cell(row=header_row, column=1, value="Compressor ID").font = header_font
        ws.cell(row=header_row, column=1).fill = header_fill
        ws.cell(row=header_row, column=1).border = border
        for j, pkey in enumerate(selected_params, start=2):
            c = ws.cell(row=header_row, column=j, value=PARAM_META[pkey]['label'])
            c.font = header_font
            c.fill = header_fill
            c.alignment = center
            c.border = border

        # determine best COP row for highlight if cop is among selected
        best_id = None
        if 'cop' in selected_params:
            cop_vals = {sid: id_rows[sid].get('cop') for sid in ids if id_rows[sid].get('cop') is not None}
            if cop_vals:
                best_id = max(cop_vals, key=cop_vals.get)

        for i, sid in enumerate(ids, start=header_row + 1):
            r = id_rows[sid]
            ws.cell(row=i, column=1, value=sid).border = border
            for j, pkey in enumerate(selected_params, start=2):
                val = r.get(pkey)
                cell = ws.cell(row=i, column=j, value=val)
                cell.border = border
                cell.alignment = center
                if sid == best_id and pkey == 'cop':
                    cell.fill = best_fill

        # column widths
        ws.column_dimensions['A'].width = 16
        for j in range(2, len(selected_params) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 20

    # Summary sheet
    ws = wb.create_sheet(title="Summary", index=0)
    ws.cell(row=1, column=1, value="Compressor Calorimeter Comparison Summary").font = title_font
    s = payload['summary']
    lines = [
        ("Total Reports Analyzed", s['total_reports']),
        ("Unique Compressor IDs", s['unique_ids']),
        ("RPM Test Tiers", s['rpm_tiers']),
        ("Average COP (all)", s['avg_cop']),
        ("Max COP (all)", s['max_cop']),
        ("Min COP (all)", s['min_cop']),
    ]
    for i, (label, val) in enumerate(lines, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        ws.cell(row=i, column=2, value=val)
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_pdf(rows, selected_params):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    payload = build_comparison_payload(rows)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                             leftMargin=14*mm, rightMargin=14*mm,
                             topMargin=14*mm, bottomMargin=14*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleX', parent=styles['Title'], textColor=colors.HexColor('#1F3864'), fontSize=18)
    sub_style = ParagraphStyle('SubX', parent=styles['Normal'], textColor=colors.HexColor('#444444'), fontSize=10)
    tier_style = ParagraphStyle('TierX', parent=styles['Heading2'], textColor=colors.HexColor('#1F3864'))

    elements = []
    elements.append(Paragraph("Walton Compressor Calorimeter Comparison Report", title_style))
    s = payload['summary']
    elements.append(Paragraph(
        f"Reports analyzed: {s['total_reports']} &nbsp;|&nbsp; Compressor IDs: {s['unique_ids']} "
        f"&nbsp;|&nbsp; RPM tiers: {s['rpm_tiers']} &nbsp;|&nbsp; Avg COP: {s['avg_cop']}",
        sub_style))
    elements.append(Spacer(1, 10))

    for tier_block in payload['tiers']:
        tier = tier_block['tier']
        ids = tier_block['ids']
        id_rows = tier_block['rows']

        elements.append(Paragraph(f"{tier} RPM Test Point", tier_style))

        header = ["Compressor ID"] + [PARAM_META[p]['label'] for p in selected_params]
        table_data = [header]
        for sid in ids:
            r = id_rows[sid]
            row_vals = [sid]
            for pkey in selected_params:
                val = r.get(pkey)
                if val is None:
                    row_vals.append("-")
                else:
                    row_vals.append(PARAM_META[pkey]['fmt'].format(val))
            table_data.append(row_vals)

        t = Table(table_data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F3864')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BFBFBF')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F6FC')]),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 14))

    doc.build(elements)
    buf.seek(0)
    return buf


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)