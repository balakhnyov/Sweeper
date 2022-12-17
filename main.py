from io import BytesIO
import os
import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pyvisa

st.set_page_config(page_title="Sweeper")
test = False
default_folder = "C:\\Users\\Lab-Nano\\Desktop\\Sweeper_meas"
inst_address = "USB0::0x05E6::0x2450::04505744::INSTR"


@st.cache
def convert_df(df):
    return df.to_csv(index=False).encode("utf-8")


@st.cache
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    format1 = workbook.add_format({"num_format": "0.00"})
    worksheet.set_column("A:A", None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def get_measures(test=False, filename=None):
    if test:
        measures = pd.read_csv(filename)
        measures.loc[:, "Voltage (V)"] = measures.loc[:, "Voltage (V)"].values
    else:
        buffer = inst.query("triggerprocess()")
        buffer = np.fromstring(buffer, dtype=float, sep=",").reshape(-1, 2)
        measures = pd.DataFrame(columns=["Voltage (V)", "Current (A)"], data=buffer)
    return measures


def measure_on(df):
    df["Light"] = "Disabled"
    df2 = get_measures(test=test, filename="4ir.csv")
    df2["Light"] = "Enabled"
    df = pd.concat([df, df2], axis=0)
    st.session_state["df"] = df


def measure_off(df):
    df["Light"] = "Enabled"
    df2 = get_measures(test=test, filename="4ir.csv")
    df2["Light"] = "Disabled"
    df = pd.concat([df2, df], axis=0)
    st.session_state["df"] = df


def calculate_difference(df):
    difference = pd.DataFrame()
    difference["Voltage (V)"] = df.loc[df["Light"] == "Enabled"]["Voltage (V)"]
    difference["Current (A)"] = (
        df.loc[df["Light"] == "Enabled"]["Current (A)"] - df.loc[df["Light"] == "Disabled"]["Current (A)"]
    )
    return difference


def plot_difference(df, name):
    difference = calculate_difference(df)
    values = calculate_values(difference)
    i_box, v_box = eval_plot_ranges(difference, float(values["Isc"].values), float(values["Voc"].values))
    fig = px.line(
        difference,
        x="Voltage (V)",
        y="Current (A)",
        line_shape="spline",
        render_mode="svg",
        markers=True,
        range_x=v_box,
        range_y=i_box,
    )
    fig.add_trace(
        go.Scatter(
            x=[0.0, float(values["Vmax"])],
            y=[0.0, float(values["Imax"])],
            mode="lines",
            line=go.scatter.Line(color="red"),
            showlegend=False,
        )
    )
    fig.update_layout(showlegend=False, title_text="Difference " + name, title_x=0.5)
    st.session_state["fig_difference"] = fig
    st.plotly_chart(fig, use_container_width=True)


def plot_log(df, name, vmax):
    fig = px.line(
        df,
        x="Voltage (V)",
        y="Current (A)",
        color="Light",
        markers=True,
        line_shape="spline",
        render_mode="svg",
        log_y=True,
        range_x=[0, vmax],
    )
    fig.update_layout(title_text="Logarithmic " + name, title_x=0.5)
    st.session_state["fig_log"] = fig
    st.plotly_chart(fig, use_container_width=True)


def plot_single(df, filename):
    fig = px.line(df, x="Voltage (V)", y="Current (A)", markers=True, line_shape="spline", render_mode="svg",)
    fig.update_layout(title_text=filename + " I-V", title_x=0.5)
    st.session_state["fig"] = fig
    st.plotly_chart(fig, use_container_width=True)


def plot_duo(df, filename):
    fig = px.line(
        df,
        x="Voltage (V)",
        y="Current (A)",
        color="Light",
        markers=True,
        line_shape="spline",
        render_mode="svg",
    )
    fig.update_layout(title_text=filename + " I-V", title_x=0.5)
    st.session_state["fig"] = fig
    st.plotly_chart(fig, use_container_width=True)


def eval_plot_ranges(df, isc, voc):
    _, i_box, v_box = eval_box(df, voc)
    ratio = 1.05
    i_plot_box = [ratio * i_box[0], ratio * i_box[1]]
    v_plot_box = [ratio * v_box[0], ratio * v_box[1]]
    return i_plot_box, v_plot_box


def eval_box(df, voc):
    v_box = [min(0, voc), max(0, voc)]
    boxed = df.loc[(df["Voltage (V)"] >= v_box[0]) & (df["Voltage (V)"] <= v_box[1])].copy()
    i_abs = float(np.max(np.abs(boxed["Current (A)"])))
    i_max = float(boxed.loc[np.abs(boxed["Current (A)"]) == i_abs]["Current (A)"])
    i_box = [min(0.0, i_max), max(0.0, i_max)]
    return boxed, i_box, v_box


# def df_box(df, voc):
#     i_box, v_box = eval_box(df, voc)
#     boxed = df.loc[
#         (df['Voltage (V)'] >= v_box[0]) &
#         (df['Voltage (V)'] <= v_box[1])
#         ].copy()
#     return boxed


def calculate_values(df, idx="0"):
    isc = float(df.loc[np.abs(df["Voltage (V)"]) == np.min(np.abs(df["Voltage (V)"]))]["Current (A)"])
    voc = float(df.loc[np.abs(df["Current (A)"]) == np.min(np.abs(df["Current (A)"]))]["Voltage (V)"])
    box, _, _ = eval_box(df, voc)
    box.loc[:, "p"] = np.abs(box.loc[:, "Voltage (V)"].values * box.loc[:, "Current (A)"].values)
    pint = np.abs(np.trapz(box["Current (A)"], x=box["Voltage (V)"]))
    pmax = np.max(box["p"])
    imax = float(box.loc[box["p"] == pmax]["Current (A)"])
    vmax = float(box.loc[box["p"] == pmax]["Voltage (V)"])

    values = pd.DataFrame(
        [[pint, pmax, imax, vmax, isc, voc]],
        columns=["Pint", "Pmax", "Imax", "Vmax", "Isc", "Voc"],
        index=[idx],
    )
    return values


def get_values(df):
    if not ("calculated" in st.session_state):
        enabled = df.loc[df["Light"] == "Enabled"].copy()
        enabled_values = calculate_values(enabled, idx="Enabled")
        difference_values = calculate_values(calculate_difference(df), idx="Difference")
        values = pd.concat([enabled_values, difference_values], axis=0)
        values.columns = [
            "Pint, W",
            "Pmax, W",
            "Imax, A",
            "Vmax, V",
            "Isc, A",
            "Voc, V",
        ]
        st.session_state["calculated"] = np.abs(values)
    values = st.session_state.calculated
    st.write(values)


def calculate_efficiency(beam_power):
    values = st.session_state.calculated
    values.loc[:, "η int, %"] = 100 * values.loc[:, "Pint, W"] / beam_power
    values.loc[:, "η max, %"] = 100 * values.loc[:, "Pmax, W"] / beam_power
    st.session_state["calculated"] = values


def wipe_state(keys):
    for key in keys:
        if key in st.session_state:
            del st.session_state[key]


def save_existing(path):
    path_interact = os.path.join(path, "figures_interact")
    if not os.path.exists(path_interact):
        os.makedirs(path_interact)
    if "fig" in st.session_state:
        fig = st.session_state.fig
        fig.write_image(os.path.join(path, "Figure Meas.png"))
        fig.write_html(os.path.join(path_interact, "Figure Meas.html"))
    if "fig_difference" in st.session_state:
        fig = st.session_state.fig_difference
        fig.write_image(os.path.join(path, "Difference.png"))
        fig.write_html(os.path.join(path_interact, "Difference.html"))
    if "fig_log" in st.session_state:
        fig = st.session_state.fig_log
        fig.write_image(os.path.join(path, "Log_I.png"))
        fig.write_html(os.path.join(path_interact, "Log_I.html"))
    if "df" in st.session_state:
        measurements = st.session_state.df
        measurements.to_csv(os.path.join(path, "measurements.csv"), index=False)
        measurements.to_excel(os.path.join(path, "measurements.xlsx"), index=False)
    if "calculated" in st.session_state:
        results = st.session_state.calculated
        results.to_csv(os.path.join(path, "results.csv"))
        results.to_excel(os.path.join(path, "results.xlsx"))


# configure instrument
if not test:
    rm = pyvisa.ResourceManager()
    inst = rm.open_resource(inst_address)
    inst.timeout = 100000
    # For 4-wire testing
    inst.write("CellTest()")
    # inst.write("CellTest2Wire()")


# UI & script body
with st.sidebar.form(key="Params"):
    sample_name = st.text_input("Sample Name", placeholder="Untitled")
    num = st.number_input("Number of Points", min_value=3, max_value=1000, value=10, step=1)
    min_voltage = st.number_input("Min Voltage (V)", min_value=-100.0, max_value=0.0, value=0.0, step=0.1)
    max_voltage = st.number_input("Max Voltage (V)", min_value=0.0, max_value=100.0, value=3.0, step=0.1)
    measure = st.form_submit_button("Measure I-V")

if measure:
    wipe_state(["fig", "fig_difference", "fig_log", "calculated", "df"])
    range_v = max(abs(min_voltage), abs(max_voltage))
    if not test:
        inst.write(
            "sweepconfig("
            + str(range_v)
            + ", "
            + str(min_voltage)
            + ", "
            + str(max_voltage)
            + ", "
            + str(num)
            + ")"
        )
    st.session_state["df"] = get_measures(test=test, filename="4.csv")


if "df" in st.session_state:
    filename = "Untitled"
    if sample_name != "":
        filename = sample_name
    df = st.session_state.df
    if not ("Light" in df.columns):
        plot_single(df, filename)
        st.button(label="+Measure Light Enabled", on_click=measure_on, args=(df,))
        st.button(label="+Measure Light Disabled", on_click=measure_off, args=(df,))
    else:
        plot_duo(df, filename)
        with st.sidebar.form(key="BeamPower"):
            beam_p = st.number_input(
                "Beam Power, W", min_value=0.0, max_value=50.0, value=0.5, step=0.01, format="%.4f",
            )
            submitted = st.form_submit_button("Calculate Efficiency")
        if submitted:
            calculate_efficiency(beam_p)
            get_values(df)
        else:
            get_values(df)

    with st.form(key="Saving"):
        path = st.text_input("Path:", value=os.path.join(default_folder, filename))
        save = st.form_submit_button("Save all Figures and Data")
    if save:
        save_existing(path)
    csv = convert_df(df)
    st.download_button(
        label="Download measurements as CSV", data=csv, file_name=filename + ".csv", mime="text/csv",
    )

    xlsx = to_excel(df)
    st.download_button(label="Download measurements as XLSX", data=xlsx, file_name=filename + ".xlsx")

    if "Light" in df.columns:
        plot_difference(df, filename)
        # plot_log(df, filename, 1.05 * max_voltage)

st.button("Re-run")
