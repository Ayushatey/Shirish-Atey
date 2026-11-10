import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import streamlit as st
import openpyxl
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Coil Engineering Calculator", page_icon="⚙️", layout="wide")

BASE_DIR = Path(__file__).parent
WORKBOOK_PATH = BASE_DIR / "Shirish_kaka_source.xlsx"
SHEET_NAME = "Calculations"

# Input mapping follows the labels in column B and values in column C.
INPUTS = [
    (8, "Stator Cavity OD (b1)", 83.49983346727426, "mm"),
    (9, "Stator Cavity ID (b2)", 69.50014062714472, "mm"),
    (10, "Stator Thickness at stator resting face (t1)", 2.5, "mm"),
    (11, "Stator width (a1)", 17.59129837673297, "mm"),
    (14, "Assembly Space at OD (b3)", 0.6, "mm"),
    (15, "Assembly Space at ID (b4)", 0.2, "mm"),
    (16, "Tape Layer - radial (b5)", 0.3, "mm"),
    (17, "Axial gap between coil and stator (a2)", 0.2, "mm"),
    (18, "Minimum Epoxy Gap (a4)", 1.0, "mm"),
    (21, "Conductor Size / Diameter (d)", 0.5722029843614138, "mm"),
    (22, "Insulation Thickness (ti)", 0.02861014921807069, "mm"),
    (23, "Copper Resistivity at 20°C (ρ0)", 1.724e-8, "Ω·m"),
    (24, "Temperature Coefficient of Resistance (α)", 0.004041, "1/°C"),
    (25, "M/wire & L/wire Solder Size (Dl)", 3.0, "mm"),
    (28, "Minimum Operating Temperature (Tmin)", -40.0, "°C"),
]

def recalculate_with_libreoffice(source_path: Path, target_dir: Path) -> Path:
    exe = shutil.which("libreoffice") or shutil.which("soffice")
    if not exe:
        raise RuntimeError("LibreOffice is not installed. Install LibreOffice to calculate the workbook formulas.")
    target_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [exe, "--headless", "--convert-to", "xlsx", "--outdir", str(target_dir), str(source_path)],
        capture_output=True, text=True, timeout=120
    )
    calculated = target_dir / source_path.name
    if not calculated.exists():
        raise RuntimeError("LibreOffice could not recalculate the workbook. " + (proc.stderr or proc.stdout))
    return calculated

def as_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

st.title("⚙️ Coil Engineering Calculator")
st.caption("Interactive frontend powered by the existing Excel calculation workbook")
st.info("Enter the numerical design inputs, then calculate. The workbook's existing formulas are recalculated by LibreOffice; calculated results are read back from Excel cells and estimation tables.")

with st.sidebar:
    st.header("Design inputs")
    st.caption("Values are written to the matching cells in Calculations!C8:C28. Blank separator rows in the workbook are not editable.")
    with st.form("coil_inputs"):
        values = {}
        cols = st.columns(1)
        for row, label, default, unit in INPUTS:
            values[row] = st.number_input(
                f"{label} — {unit}",
                value=float(default),
                format="%.10g",
                key=f"input_{row}",
                help=f"Excel cell Calculations!C{row}"
            )
        submitted = st.form_submit_button("Calculate", type="primary", use_container_width=True)

if "result" not in st.session_state:
    st.session_state.result = None

if submitted:
    try:
        with st.spinner("Updating Excel inputs and recalculating formulas..."):
            with tempfile.TemporaryDirectory(prefix="coil_calc_") as tmp:
                tmp_path = Path(tmp)
                working_xlsx = tmp_path / "user_inputs.xlsx"
                shutil.copy2(WORKBOOK_PATH, working_xlsx)

                book = openpyxl.load_workbook(working_xlsx)
                sheet = book[SHEET_NAME]
                for row, value in values.items():
                    sheet[f"C{row}"] = float(value)
                book.calculation.fullCalcOnLoad = True
                book.calculation.forceFullCalc = True
                book.calculation.calcMode = "auto"
                book.save(working_xlsx)

                recalculated = recalculate_with_libreoffice(working_xlsx, tmp_path / "recalculated")
                formulas_book = openpyxl.load_workbook(recalculated, data_only=False)
                values_book = openpyxl.load_workbook(recalculated, data_only=True)
                formulas_sheet = formulas_book[SHEET_NAME]
                data_sheet = values_book[SHEET_NAME]

                def getv(cell):
                    return as_float(data_sheet[cell].value)

                # Read outputs from the workbook, not from Python-side engineering formulas.
                output = {
                    "amp_turns_9v_150c": getv("C115"),
                    "lead_wire_turns": getv("C92"),
                    "coil_turns_without_lead": getv("C91"),
                    "equivalent_lead_wire_turns": getv("C84"),
                    "current_9v_150c": getv("C114"),
                    "maximum_current": getv("C116"),
                    "coil_thickness": getv("C50"),
                    "overall_conductor_diameter": getv("C53"),
                    "actual_coil_id": getv("C54"),
                    "actual_coil_od": getv("C55"),
                }

                # Table 3: row 141 = 9 V; row 146 = 16 V in the source workbook.
                temperatures, at_9v, at_16v = [], [], []
                for col in range(3, 16):
                    temp = getv(f"{openpyxl.utils.get_column_letter(col)}122")
                    a9 = getv(f"{openpyxl.utils.get_column_letter(col)}141")
                    a16 = getv(f"{openpyxl.utils.get_column_letter(col)}146")
                    if temp is not None and -40 <= temp <= 150 and a9 is not None and a16 is not None:
                        temperatures.append(temp)
                        at_9v.append(a9)
                        at_16v.append(a16)

                # Sort temperature points because the source workbook's temperature row is not strictly ascending.
                points = sorted(zip(temperatures, at_9v, at_16v), key=lambda x: x[0])
                temperatures = [p[0] for p in points]
                at_9v = [p[1] for p in points]
                at_16v = [p[2] for p in points]

                # Capture all current and ampere-turn table combinations for the dashboard.
                table_rows = []
                for row in range(130, 136):
                    voltage = as_float(data_sheet[f"B{row}"].value)
                    for col in range(3, 16):
                        letter = openpyxl.utils.get_column_letter(col)
                        temp = getv(f"{letter}122")
                        current = getv(f"{letter}{row}")
                        amp_turns = getv(f"{letter}{row+11}")
                        if temp is not None and current is not None and amp_turns is not None:
                            table_rows.append({
                                "Voltage (V)": voltage,
                                "Temperature (°C)": temp,
                                "Current (A)": current,
                                "Ampere-turns": amp_turns
                            })

                # Keep a copy of the recalculated workbook for download.
                download_path = Path("/tmp") / "coil_engineering_calculated.xlsx"
                shutil.copy2(recalculated, download_path)

                st.session_state.result = {
                    "outputs": output,
                    "temperatures": temperatures,
                    "at_9v": at_9v,
                    "at_16v": at_16v,
                    "table": pd.DataFrame(table_rows),
                    "workbook_bytes": download_path.read_bytes(),
                }
        st.success("Calculation completed using the Excel workbook.")
    except Exception as exc:
        st.error(f"Calculation failed: {exc}")

result = st.session_state.result
if result:
    o = result["outputs"]
    st.subheader("Calculation results")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Ampere-turns at 9 V / 150°C", f"{o['amp_turns_9v_150c']:.3f}" if o["amp_turns_9v_150c"] is not None else "—")
    k2.metric("Turns in lead wire", f"{o['lead_wire_turns']:.0f}" if o["lead_wire_turns"] is not None else "—")
    k3.metric("Current at 9 V / 150°C", f"{o['current_9v_150c']:.3f} A" if o["current_9v_150c"] is not None else "—")
    k4.metric("Maximum current", f"{o['maximum_current']:.3f} A" if o["maximum_current"] is not None else "—")

    with st.expander("Additional calculated geometry and winding outputs", expanded=False):
        extra = pd.DataFrame([
            ["Turns without lead-wire allowance", o["coil_turns_without_lead"]],
            ["Equivalent turns consumed by lead wire", o["equivalent_lead_wire_turns"]],
            ["Coil thickness", o["coil_thickness"]],
            ["Overall conductor diameter", o["overall_conductor_diameter"]],
            ["Actual coil ID", o["actual_coil_id"]],
            ["Actual coil OD", o["actual_coil_od"]],
        ], columns=["Parameter", "Calculated value"])
        st.dataframe(extra, hide_index=True, use_container_width=True)

    st.subheader("Ampere-turns vs temperature")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=result["temperatures"], y=result["at_9v"], mode="lines+markers", name="9 V"))
    fig.add_trace(go.Scatter(x=result["temperatures"], y=result["at_16v"], mode="lines+markers", name="16 V"))
    fig.update_layout(
        title="Ampere-Turns vs Temperature at 9 V and 16 V",
        xaxis_title="Temperature (°C)",
        yaxis_title="Ampere-turns (A-turns)",
        xaxis=dict(range=[-40, 150]),
        legend_title="Voltage",
        template="plotly_white",
        height=480,
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Current and ampere-turn estimation table")
    st.dataframe(result["table"].sort_values(["Voltage (V)", "Temperature (°C)"]), hide_index=True, use_container_width=True)

    st.download_button(
        "Download recalculated Excel workbook",
        data=result["workbook_bytes"],
        file_name="coil_engineering_calculated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.subheader("Calculation results")
    st.write("Submit the design inputs to calculate results from the workbook.")
    st.subheader("Ampere-turns vs temperature")
    st.write("The 9 V and 16 V curves will appear here after calculation.")

st.divider()
st.caption("Engineering note: verify the workbook's conductor-length assumptions and all design constraints before using outputs for manufacturing or safety-critical decisions.")
