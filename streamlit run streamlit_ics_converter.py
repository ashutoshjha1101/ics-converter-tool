# streamlit_ics_converter.py
import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import zipfile
import json
from datetime import datetime
import re

st.set_page_config(page_title="ICS Converter — Nive Solutions", layout="centered")

# --- CSS / UI styling (mirror-like white cards) ---
st.markdown("""
<style>
body { background: linear-gradient(180deg, #f7f9fc 0%, #ffffff 100%); }
header .decoration { display:none; }
.appview-container .main .block-container{ padding-top: 1rem; padding-bottom: 2rem; }
.card {
  background: #ffffff;
  border-radius: 14px;
  padding: 16px;
  box-shadow: 0 8px 30px rgba(15,20,30,0.06);
  border: 1px solid rgba(15,20,30,0.04);
}
.logo-row { display:flex; align-items:center; gap:12px; margin-bottom:8px; }
.logo-row img { height:40px; border-radius:6px; }
.h1 { font-size:20px; font-weight:700; margin:0; }
.h1-sub { font-size:12px; color:#6b7280; margin:0; }
.upload-box { border:2px dashed rgba(15,20,30,0.04); padding:12px; border-radius:10px; text-align:center;}
.btn-primary {
  background: linear-gradient(90deg, #0ea5ff, #7c3aed);
  color:white;
  padding:8px 12px;
  border-radius:10px;
  border:none;
}
.small-muted { color:#6b7280; font-size:13px; }
.table-wrap { max-height:360px; overflow:auto; }
</style>
""", unsafe_allow_html=True)

# --- Header ---
with st.container():
    cols = st.columns([0.14, 0.86])
    with cols[0]:
        # Logo: expects 'nive_logo.png' in the same directory
        st.image("nive_logo.png", width=56)
    with cols[1]:
        st.markdown('<div class="logo-row"><div><div class="h1">ICS Converter</div><div class="h1-sub">by Nive Solutions — Convert .ics to CSV, Excel, JSON, ZIP</div></div></div>', unsafe_allow_html=True)

st.markdown("")

# --- Main card ---
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    st.markdown("#### Upload ICS files")
    st.markdown('<div class="small-muted">Select one or more .ics calendar files. Default limit: 20 files.</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader("Upload .ics files", type=["ics"], accept_multiple_files=True, help="You can upload multiple .ics files", key="ics_uploader")

    max_files = 20
    if uploaded_files and len(uploaded_files) > max_files:
        st.warning(f"You uploaded {len(uploaded_files)} files. Only the first {max_files} will be processed.")
        uploaded_files = uploaded_files[:max_files]

    # Options
    col1, col2, col3 = st.columns(3)
    with col1:
        expand_recurrences = st.checkbox("Expand simple RRULE?", value=False, help="(Simple RRULE expansion is limited)")
    with col2:
        separate_zips = st.checkbox("Export separate CSVs (zipped)", value=True)
    with col3:
        combined_sheet = st.checkbox("Export single Excel workbook", value=True)

    st.markdown("---")

    # Basic ICS parser functions
    def unfold_ics(text):
        # Unfold lines folded per RFC: lines that start with space or tab are continuations
        return re.sub(r'\\r?\\n[ \\t]+', '', text)

    def parse_props(block):
        # Build dict of properties: property name -> list of values
        props = {}
        for line in block.splitlines():
            if not line.strip(): continue
            if ':' not in line:
                # fallback: skip
                continue
            left, val = line.split(':', 1)
            # property may have parameters: e.g., DTSTART;TZID=Asia/Kolkata
            prop = left.split(';')[0].upper()
            val = val.strip()
            props.setdefault(prop, []).append(val)
        return props

    def parse_ics_text(text):
        # returns list of event dicts
        text = unfold_ics(text)
        vevents = re.split(r'BEGIN:VEVENT', text, flags=re.IGNORECASE)[1:]
        events = []
        for v in vevents:
            # stop at END:VEVENT
            v = v.split('END:VEVENT')[0]
            props = parse_props(v)
            # pick common fields, take first value if multiple
            ev = {
                'UID': props.get('UID', [''])[0],
                'SUMMARY': props.get('SUMMARY', [''])[0],
                'DESCRIPTION': props.get('DESCRIPTION', [''])[0],
                'LOCATION': props.get('LOCATION', [''])[0],
                'DTSTART': props.get('DTSTART', [''])[0],
                'DTEND': props.get('DTEND', [''])[0],
                'RRULE': props.get('RRULE', [''])[0],
                'ORGANIZER': props.get('ORGANIZER', [''])[0],
                'ATTENDEE': ';'.join(props.get('ATTENDEE', [])) if props.get('ATTENDEE') else ''
            }
            events.append(ev)
        return events

    def normalize_dt(dt_str):
        if not dt_str: return ''
        # simple normalizer: remove timezone params if present and parse common forms
        # handle forms: 20250917T153000Z or 20250917T153000 or 2025-09-17T15:30:00
        dt = dt_str
        # remove extra params like "TZID=..."
        if dt_str.upper().startswith('TZID='):
            # value may be like: TZID=Asia/Kolkata:20250917T153000
            parts = dt_str.split(':', 1)
            if len(parts) == 2:
                dt = parts[1]
        # strip trailing Z
        dt = dt.rstrip('Z')
        # try parse known formats
        fmt_candidates = ['%Y%m%dT%H%M%S', '%Y%m%dT%H%M', '%Y-%m-%dT%H:%M:%S', '%Y%m%d']
        for fmt in fmt_candidates:
            try:
                return datetime.strptime(dt, fmt).isoformat()
            except Exception:
                continue
        # fallback: return raw
        return dt_str

    if uploaded_files:
        all_parsed = []  # list of (filename, events)
        total_events = 0
        errors = []
        for f in uploaded_files:
            try:
                raw = f.read().decode('utf-8', errors='ignore')
                events = parse_ics_text(raw)
                # normalize dates
                for ev in events:
                    ev['DTSTART_ISO'] = normalize_dt(ev.get('DTSTART',''))
                    ev['DTEND_ISO'] = normalize_dt(ev.get('DTEND',''))
                all_parsed.append((f.name, events))
                total_events += len(events)
            except Exception as e:
                errors.append((f.name, str(e)))

        st.markdown(f"**Files processed:** {len(all_parsed)}  •  **Total events:** {total_events}")
        if errors:
            st.error("There were parse errors in some files. See details below.")
            for fn, msg in errors:
                st.write(f"- {fn}: {msg}")

        # Preview combined table
        preview_df = pd.DataFrame()
        preview_rows = []
        for filename, events in all_parsed:
            for ev in events:
                row = {
                    'file': filename,
                    'uid': ev.get('UID',''),
                    'summary': ev.get('SUMMARY',''),
                    'start': ev.get('DTSTART_ISO',''),
                    'end': ev.get('DTEND_ISO',''),
                    'location': ev.get('LOCATION',''),
                    'description': ev.get('DESCRIPTION',''),
                    'rrule': ev.get('RRULE',''),
                }
                preview_rows.append(row)
        if preview_rows:
            preview_df = pd.DataFrame(preview_rows)
            st.markdown("**Preview (first 200 rows):**")
            st.markdown('<div class="table-wrap">', unsafe_allow_html=True)
            st.dataframe(preview_df.head(200))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No events found in uploaded files.")

        st.markdown("---")
        st.markdown("### Export options")

        # Helper exporters
        def to_csv_bytes(df):
            return df.to_csv(index=False).encode('utf-8')

        def generate_combined_csv():
            if preview_df.empty:
                return None
            return to_csv_bytes(preview_df)

        def generate_separate_csvs_zip():
            mem_zip = BytesIO()
            with zipfile.ZipFile(mem_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
                for filename, events in all_parsed:
                    rows = []
                    for ev in events:
                        rows.append({
                            'uid': ev.get('UID',''),
                            'summary': ev.get('SUMMARY',''),
                            'start': ev.get('DTSTART_ISO',''),
                            'end': ev.get('DTEND_ISO',''),
                            'location': ev.get('LOCATION',''),
                            'description': ev.get('DESCRIPTION',''),
                            'rrule': ev.get('RRULE',''),
                        })
                    df = pd.DataFrame(rows)
                    safe_name = re.sub(r'[^0-9A-Za-z_.-]', '_', filename)
                    csv_bytes = to_csv_bytes(df)
                    zf.writestr(safe_name + '.csv', csv_bytes)
            mem_zip.seek(0)
            return mem_zip.read()

        def generate_excel_bytes():
            out = BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for filename, events in all_parsed:
                    rows = []
                    for ev in events:
                        rows.append({
                            'uid': ev.get('UID',''),
                            'summary': ev.get('SUMMARY',''),
                            'start': ev.get('DTSTART_ISO',''),
                            'end': ev.get('DTEND_ISO',''),
                            'location': ev.get('LOCATION',''),
                            'description': ev.get('DESCRIPTION',''),
                            'rrule': ev.get('RRULE',''),
                        })
                    df = pd.DataFrame(rows)
                    sheet_name = filename[:31] if filename else 'sheet'
                    safe_name = re.sub(r'[^0-9A-Za-z_]', '_', sheet_name)
                    try:
                        df.to_excel(writer, sheet_name=safe_name[:31], index=False)
                    except Exception:
                        # fallback: write to a default sheet if name fails
                        df.to_excel(writer, sheet_name='sheet_'+str(hash(filename))[:10], index=False)
            out.seek(0)
            return out.read()

        def generate_json_bytes(separate=False):
            if separate:
                payload = { fn: events for fn, events in all_parsed }
            else:
                payload = []
                for fn, events in all_parsed:
                    for ev in events:
                        o = ev.copy()
                        o['file'] = fn
                        payload.append(o)
            return json.dumps(payload, indent=2).encode('utf-8')

        # Buttons and downloads
        colA, colB, colC, colD = st.columns(4)
        with colA:
            csv_bytes = generate_combined_csv()
            if csv_bytes is not None:
                st.download_button("Download combined CSV", data=csv_bytes, file_name="events_combined.csv", mime="text/csv", key="dl_combined_csv")
            else:
                st.button("Download combined CSV", disabled=True)
        with colB:
            zip_bytes = generate_separate_csvs_zip()
            st.download_button("Download separate CSVs (ZIP)", data=zip_bytes, file_name="events_individual_csvs.zip", mime="application/zip", key="dl_zip")
        with colC:
            if combined_sheet:
                excel_bytes = generate_excel_bytes()
                st.download_button("Download Excel workbook", data=excel_bytes, file_name="events_workbook.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_excel")
            else:
                st.button("Download Excel workbook", disabled=True)
        with colD:
            json_bytes = generate_json_bytes(separate=False)
            st.download_button("Download JSON (combined)", data=json_bytes, file_name="events.json", mime="application/json", key="dl_json")

        st.markdown("---")
        st.markdown("Small note: this parser handles standard/event fields and simple line-folding. Complex Microsoft/Exchange-specific properties or deep RRULE expansion may require additional logic.")
    else:
        st.markdown('<div class="upload-box">No files uploaded yet. Drag and drop .ics files here or click to select.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("<div style='padding-top:12px;color:#6b7280;font-size:13px;'>Nive Solutions — ICS Converter. Built with Python & Streamlit. Contact: support@nivesolutions.example</div>", unsafe_allow_html=True)
