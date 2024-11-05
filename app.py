import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit
from fpdf import FPDF
import base64
import math

st.set_page_config(page_title="Energy Balance Software", page_icon=":bar_chart:", layout="wide")
report_title = "Zone-Circle-Division wise Import for September-2024"
# ---- READ EXCEL ----
@st.cache
def get_data_from_excel():
    df = pd.read_excel(
        io="EB.xlsx",
        engine="openpyxl",
        sheet_name="Linked_11KV",
        skiprows=0,
        usecols="B:Z",
        nrows=1226,
    )

    return df

df_selection = get_data_from_excel()
# month_list = ['Latest']+list(df_selection.columns)[13:]
# st.markdown("""---""")
# month_choice = st.selectbox("Please Select Month",month_list)
# st.markdown("""---""")
# if month_choice == 'Latest':
#     df_selection = df_selection
# else:
#     df_selection["Consumption"] = df_selection[month_choice]
#     df_selection["Corrected_Consumption"] = df_selection[month_choice]


# ---- MAINPAGE ----
st.title(":bar_chart: Energy Balance Dashboard")
st.markdown("##")

# TOP KPI's
# consumption = int(df_selection["Consumption"].sum())
consumption_corrected = round(float(df_selection["Corrected_Consumption"].sum()))

left_column,right_column = st.columns(2)
with left_column:
    st.subheader("Total Import at 33 KV Level (All NOCS):- ")
with right_column:
    st.subheader(f"  {consumption_corrected} KWH")

st.markdown("""----""")
#-----Data Preprocessing for Summary Report-----------------
consumption_by_nocs = (
    df_selection.groupby(by=["NOCS"])["Corrected_Consumption"].sum().reset_index()
)
# Assuming consumption_by_nocs is a DataFrame
consumption_by_nocs["Corrected_Consumption"] = consumption_by_nocs["Corrected_Consumption"].round()

# Round the values to two decimal points
consumption_by_nocs["Corrected_Consumption"] = consumption_by_nocs["Corrected_Consumption"].apply(lambda x: round(x, 2))

### Reporting Engine Creation
def create_download_link(val, filename):
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}.pdf">Download Report</a>'

def output_df_to_pdf(pdf, df):
    # A cell is a rectangular area, possibly framed, which contains some text
    # Set the width and height of cell
    table_cell_width = 35
    table_cell_height = 6
    # Select a font as Arial, bold, 8
    pdf.set_font('Arial', 'B', 8)
    
    # Loop over to print column names
    cols = df.columns
    for col in cols:
        pdf.cell(table_cell_width, table_cell_height, col, align='C', border=1)
        
    # Line break
    pdf.ln(table_cell_height)
    # Select a font as Arial, regular, 7
    pdf.set_font('Arial', '', 7)
    # Loop over to print each data in the table
    for row in df.itertuples():
        for col in cols:
            value = str(getattr(row, col))
            pdf.cell(table_cell_width, table_cell_height, value, align='C', border=1)
        pdf.ln(table_cell_height)
    pdf.ln()
    pdf.set_font('Arial', 'B', 10)
    try:
        pdf.cell(0,10, txt="Total Consumption:- "+str(df["Consumption"].sum()))
    except KeyError:
        print("")
    pdf.ln()
    pdf.cell(0,10, txt="Total Corrected Consumption:- "+str(df["Corrected_Consumption"].sum()))

def export_as_pdf(report_text,data,report_type):
    if report_type == 'table':
        pdf = FPDF('landscape','mm',"A4")
    elif report_type == 'summary':
        pdf = FPDF('Portrait','mm',"A4")
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    if report_type == 'table':
        pdf.cell(0, 10, txt=report_text, align="C")
    elif report_type == 'summary':
        pdf.cell(0, 10, txt=report_text, align="L")
    pdf.ln()
    output_df_to_pdf(pdf,data)
    
    html = create_download_link(pdf.output(dest="S").encode("latin-1"), "Report")

    return(st.markdown(html, unsafe_allow_html=True))

#-----------------------NOCS-Wise Summary TreeMap-------------------#
summary_tree_nw = px.treemap(consumption_by_nocs,
                 path=consumption_by_nocs.columns,
                 values=consumption_by_nocs["Corrected_Consumption"],
                 color =consumption_by_nocs["NOCS"],
                 color_continuous_scale = ['red','yellow','green'],
                 title='NOCS-Wise Summary of Import',
                 width = 1000,
                 height = 700,
                 )

summary_tree_nw.update_layout(
    font_size = 15,
    title_font_size = 30, 
    title_font_family ='Arial',
)
st.plotly_chart(summary_tree_nw, use_container_width=True)
#----------------Report-Download---------------------
st.write("---")
st.caption("Instruction: You can download the NOCS-wise Import Summary by Clicking the following Link.")
export_as_pdf("Summary of NOCS-Wise Import",consumption_by_nocs,'summary')
st.write("---")
# Create a dictionary to map "NOCS" to "Circle" and "Zone"
nocs_mapping = {
    "Tejgaon": ("Tejgaon", "North"),
    "Kakrail": ("Tejgaon", "North"),
    "Moghbazar": ("Moghbazar", "North"),
    "Khilgaon": ("Moghbazar", "North"),
    "Satmasjid": ("Satmasjid", "North"),
    "Shere b.nagar": ("Satmasjid", "North"),
    "Dhanmondi": ("Dhanmondi", "North"),
    "Jigatola": ("Dhanmondi", "North"),
    "Azimpur": ("Azimpur", "North"),
    "Paribag": ("Azimpur", "North"),
    "Shyamoli": ("Shyamoli", "North"),
    "Adabor": ("Shyamoli", "North"),
    "Lalbag": ("Lalbag", "Central"),
    "Kamrangirchar": ("Lalbag", "Central"),
    "Ramna": ("Ramna", "Central"),
    "Rajarbag": ("Ramna", "Central"),
    "Bashabo": ("Bashabo", "Central"),
    "Banosree": ("Bashabo", "Central"),
    "Motijheel": ("Motijheel", "Central"),
    "Mugdapara": ("Motijheel", "Central"),
    "Banglabazar": ("Banglabazar", "Central"),
    "Bangshal": ("Banglabazar", "Central"),
    "Narinda": ("Narinda", "Central"),
    "Swamibag": ("Narinda", "Central"),
    "Kazla": ("Kazla", "South"),
    "Maniknagar": ("Kazla", "South"),
    "Shyampur": ("Shyampur", "South"),
    "Matuail": ("Shyampur", "South"),
    "Postogola": ("Postogola", "South"),
    "Jurain": ("Postogola", "South"),
    "N.Gonj (West)": ("N.Gonj (West)", "South"),
    "N.Gonj (East)": ("N.Gonj (West)", "South"),
    "Demra": ("Demra", "South"),
    "Siddirgonj": ("Demra", "South"),
    "Fatullah": ("Fatullah", "South"),
    "Sytalakhya": ("Fatullah", "South"),
}

# Function to map "NOCS" to "Circle" and "Zone"
def map_nocs(row):
    nocs = row["NOCS"]
    if nocs in nocs_mapping:
        circle, zone = nocs_mapping[nocs]
        return pd.Series([circle, zone], index=["Circle", "Zone"])
    else:
        return pd.Series(["Unknown", "Unknown"], index=["Circle", "Zone"])

# Apply the mapping function to add "Circle" and "Zone" columns
consumption_by_nocs[["Circle", "Zone"]] = consumption_by_nocs.apply(map_nocs, axis=1)

# -----------Templace Creation------------------
st.title(report_title)
html_table =f"""
<center>
<table class="tg" bgcolor="#063970">
<thead>
<tr>
<th class="tg-op08">Zone</th>
<th class="tg-pl3c">Zone Total</th>
<th class="tg-pl3c">Circle</th>
<th class="tg-pl3c">Circle Total</th>
<th class="tg-op08">NOCS</th>
<th class="tg-pl3c">NOCS Total Import</th>
</tr>
</thead>
<tbody>
<tr>
<td class="tg-e23d" rowspan="12">Central</td>
<td class="Z_Central" rowspan="12">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banglabazar"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bangshal"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bashabo"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banosree"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Lalbag"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kamrangirchar"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Motijheel"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Mugdapara"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Narinda"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Swamibag"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Ramna"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Rajarbag"].item()}</td>
<td class="tg-e23d" rowspan="2">Banglabazar</td>
<td class="C_Banglabazar" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banglabazar"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bangshal"].item()}</td>
<td class="tg-ncgp">Banglabazar</td>
<td class="N_Banglabazar">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banglabazar"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Bangshal</td>
<td class="N_Bangshal">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bangshal"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Bashabo</td>
<td class="C_Bashabo" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bashabo"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banosree"].item()}</td>
<td class="tg-ncgp">Banosree</td>
<td class="N_Banosree">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Banosree"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Bashabo</td>
<td class="N_Bashabo">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Bashabo"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Lalbag</td>
<td class="C_Lalbag" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Lalbag"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kamrangirchar"].item()}</td>
<td class="tg-ncgp">Kamrangirchar</td>
<td class="N_Kamrangirchar">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kamrangirchar"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Lalbag</td>
<td class="N_Lalbag">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Lalbag"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Motijheel</td>
<td class="C_Motijheel" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Motijheel"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Mugdapara"].item()}</td>
<td class="tg-ncgp">Motijheel</td>
<td class="N_Motijheel">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Motijheel"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Mugdapara</td>
<td class="N_Mugdapara">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Mugdapara"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Narinda</td>
<td class="C_Narinda" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Narinda"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Swamibag"].item()}</td>
<td class="tg-ncgp">Narinda</td>
<td class="N_Narinda">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Narinda"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Swamibag</td>
<td class="N_Swamibag">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Swamibag"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Ramna</td>
<td class="C_Ramna" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Ramna"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Rajarbag"].item()}</td>
<td class="tg-ncgp">Rajarbag</td>
<td class="N_Rajarbag">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Rajarbag"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Ramna</td>
<td class="N_Ramna">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Ramna"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="12">North</td>
<td class="Z_North" rowspan="12">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Azimpur"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Paribag"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Dhanmondi"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jigatola"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Moghbazar"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Khilgaon"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Satmasjid"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shere b.nagar"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyamoli"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Adabor"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Tejgaon"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kakrail"].item()}</td>
<td class="tg-e23d" rowspan="2">Azimpur</td>
<td class="C_Azimpur" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Azimpur"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Paribag"].item()}</td>
<td class="tg-ncgp">Azimpur</td>
<td class="N_Azimpur">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Azimpur"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Paribag</td>
<td class="N_Paribag">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Paribag"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Dhanmondi</td>
<td class="C_Dhanmondi" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Dhanmondi"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jigatola"].item()}</td>
<td class="tg-ncgp">Dhanmondi</td>
<td class="N_Dhanmondi">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Dhanmondi"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Jigatola</td>
<td class="N_Jigatola">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jigatola"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Moghbazar</td>
<td class="C_Moghbazar" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Moghbazar"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Khilgaon"].item()}</td>
<td class="tg-ncgp">Khilgaon</td>
<td class="N_Khilgaon">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Khilgaon"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Moghbazar</td>
<td class="N_Moghbazar">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Moghbazar"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Satmasjid</td>
<td class="C_Satmasjid" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Satmasjid"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shere b.nagar"].item()}</td>
<td class="tg-ncgp">Satmasjid</td>
<td class="N_Satmasjid">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Satmasjid"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Shere b.nagar</td>
<td class="N_Shere b.nagar">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shere b.nagar"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Shyamoli</td>
<td class="C_Shyamoli" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyamoli"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Adabor"].item()}</td>
<td class="tg-ncgp">Adabor</td>
<td class="N_Adabor">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Adabor"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Shyamoli</td>
<td class="N_Shyamoli">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyamoli"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Tejgaon</td>
<td class="C_Tejgaon" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Tejgaon"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kakrail"].item()}</td>
<td class="tg-ncgp">Kakrail</td>
<td class="N_Kakrail">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kakrail"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Tejgaon</td>
<td class="N_Tejgaon">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Tejgaon"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="12">South</td>
<td class="Z_South" rowspan="12">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Fatullah"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Sytalakhya"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Fatullah"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Sytalakhya"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kazla"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Maniknagar"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (East)"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (West)"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyampur"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Matuail"].item()+
consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Postogola"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jurain"].item()}</td>
<td class="tg-e23d" rowspan="2">Demra</td>
<td class="C_Demra" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Fatullah"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Sytalakhya"].item()}</td>
<td class="tg-ncgp">Demra</td>
<td class="N_Demra">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Demra"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Siddirgonj</td>
<td class="N_Siddirgonj">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Siddirgonj"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Fatullah</td>
<td class="C_Fatullah" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Fatullah"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Sytalakhya"].item()}</td>
<td class="tg-ncgp">Fatullah</td>
<td class="N_Fatullah">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Fatullah"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Sytalakhya</td>
<td class="N_Sytalakhya">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Sytalakhya"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Kazla</td>
<td class="C_Kazla" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kazla"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Maniknagar"].item()}</td>
<td class="tg-ncgp">Kazla</td>
<td class="N_Kazla">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Kazla"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Maniknagar</td>
<td class="N_Maniknagar">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Maniknagar"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">N.Gonj (West)</td>
<td class="C_N.Gonj (West)" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (East)"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (West)"].item()}</td>
<td class="tg-ncgp">N.Gonj (East)</td>
<td class="N_N.Gonj (East)">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (East)"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">N.Gonj (West)</td>
<td class="N_N.Gonj (West)">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="N.Gonj (West)"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Postogola</td>
<td class="C_Postogola" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Postogola"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jurain"].item()}</td>
<td class="tg-ncgp">Jurain</td>
<td class="N_Jurain">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Jurain"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Postogola</td>
<td class="N_Postogola">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Postogola"].item()}</td>
</tr>
<tr>
<td class="tg-e23d" rowspan="2">Shyampur</td>
<td class="C_Shyampur" rowspan="2">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyampur"].item()+consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Matuail"].item()}</td>
<td class="tg-ncgp">Matuail</td>
<td class="N_Matuail">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Matuail"].item()}</td>
</tr>
<tr>
<td class="tg-ncgp">Shyampur</td>
<td class="N_Shyampur">{consumption_by_nocs["Corrected_Consumption"][consumption_by_nocs["NOCS"]=="Shyampur"].item()}</td>
</tr>
</tbody>
</table>
</center>
"""

# Render the HTML template
st.markdown(html_table, unsafe_allow_html=True)
st.markdown("---")
# Display the updated DataFrame
consumption_by_nocs.groupby(['Zone','Circle','NOCS'])['Corrected_Consumption'].sum().reset_index()
#-----------------------Cirlce, Zone and NOCS-wise Summary TreeMap-------------------#
summary_tree_zcn = px.treemap(consumption_by_nocs,
                 path=['Zone','Circle','NOCS','Corrected_Consumption'],
                 values=consumption_by_nocs["Corrected_Consumption"],
                 color =consumption_by_nocs["Circle"],
                 color_continuous_scale = ['red','yellow','green'],
                 title='Zone, Circle and NOCS-Wise Summary of Import',
                 width = 1000,
                 height = 700,
                 )

summary_tree_zcn.update_layout(
    font_size = 15,
    title_font_size = 30, 
    title_font_family ='Arial',
)

st.plotly_chart(summary_tree_zcn, use_container_width=True)
#----------------Report-Download---------------------
st.write("---")
st.caption("Instruction: You can download the Zone, Circle and NOCS-Wise Import Summary by Clicking the following Link.")
export_as_pdf("Summary of Zone, Circle and NOCS-Wise Import",consumption_by_nocs.sort_values(by=["Zone","Circle"])[["Zone","Circle","NOCS","Corrected_Consumption"]],'summary')
st.write("---")
#--------------------------------------------#

fig_nocs_consumption = px.bar(
    consumption_by_nocs,
    y=consumption_by_nocs["Corrected_Consumption"],
    x=consumption_by_nocs["NOCS"],
    labels = consumption_by_nocs["Corrected_Consumption"],
    orientation = "v",
    title="<b>Consumption by NOCS</b>",
    color="Corrected_Consumption",
    template="plotly_dark",
    text_auto = ".4s",
    height=600
)

fig_nocs_consumption.update_layout(
    font_size = 15,
    title_font_size = 30, 
    title_font_family ='Arial',
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

fig_nocs_consumption.update_traces(textfont_size=45, textangle=-90, textposition="inside", cliponaxis=False)


st.plotly_chart(fig_nocs_consumption, use_container_width=True)

# ss_wise = df_selection.groupby(['Substation_Name','NOCS'])['Corrected_Consumption'].sum().reset_index()
# ss_wise = ss_wise[ss_wise['Corrected_Consumption']!=0]
# ss_wise['Corrected_Consumption']=ss_wise['Corrected_Consumption'].astype(int)
# ss_wise['Corrected_Consumption']=ss_wise['Corrected_Consumption'].abs()
# summary_sb = px.treemap(ss_wise,
#     path=['NOCS','Substation_Name','Corrected_Consumption'],
#     values=ss_wise["Corrected_Consumption"],
#     color =ss_wise["Corrected_Consumption"] ,
#     color_continuous_scale=['Green','Violet','Yellow','Red'],
#     title='NOCS-Wise and Substation-wise Detailed Import',
#     width = 800,
#     height = 1000
# )
# summary_sb.update_layout(
#     title_font_size = 30, 
#     title_font_family ='Arial'
# )

# st.plotly_chart(summary_sb, use_container_width=True)
 
# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

ss_list=['Moghbazar 132/33/11KV S/S','Moghbazar 33/11KV S/S','Green Road 33/11KV S/S','Lalmatia  33/11KV S/S',
         'Tejgoan 33/11KV S/S','T&T 33/11 KV','Dhanmondi 132/33/11KV S/S','Dhanmondi 33/11KV S/S','Kawranbazar  33/11KV S/S',
         'New Ramna   33/11KV S/S','Ullon 132/33/11KV S/S','Ullon local 33/11KV S/S','Kakrail  33/11KV S/S','Khillgaon  33/11KV S/S',
         'Goran  33/11KV S/S','Taltola  33/11KV S/S','Satmasjid 33/11KV S/S','Jigatola 33/11KV S/S','Kallyanpur 33/11KV S/S',
         'Kamrangirchar 132/33/11KV S/S','Kamrangirchar 33/11KV S/S','Lalbagh old 33/11KV S/S','Banshal 33/11KV S/S','Japan Garden 33/11 KV',
         'Azimpur 33/11 KV','SHERE BANGLA NAGAR  33/11 KV S/S ','LALBAGH   132/33 KV S/S ','LALBAGH   33/11 KV S/S ','Asad Gate 33/11 KV S/S',
         'Shatmasjid 132/33 KV S/S','Banasree 33/11 SS','Mugdhapara Hospital 33/11 KV SS','DMC 33/11 KV SS','Green Road Dormatory 33/11 SS',
         'BSMMU 33/11KV S/S','Dhaka Uddyan 33/11 KV S/S','Dhaka University 132/33 KV S/S','Dhaka University 33/11 KV S/S','Monipuripara 33/11 KV S/S',
         'BGB 33/11 KV S/S','BB Aveneu 33/11 KV S/S','Jigatola 132/33 KV S/S','Jigatola New 33/11 KV S/S','Ispahani 33/11 KV SS',
         'Bangabhaban 132/11KV S/S','Narinda 132/33KV S/S','Narinda 33/11KV S/S','Kumertuly  33/11KV S/S','Maniknagar 132/33 S/S',
         'Maniknagar 33/11 KV SS','Madarteck 132/33KV S/S','Madarteck 33/11KV S/S','Kazla  33/11KV S/S','Shyampur  132/33KV S/S',
         'Shyampur  33/11KV S/S','Shyampur BISIC  33/11KV S/S','Postogola 33/11KV S/S','Fatullah 33/11KV S/S','Sitalakhya  132/33KV S/S',
         'Sitalakhya  33/11KV S/S','Narayangonj (west) BSCIC33/11KV S/S','Char Syedpur 33/11KV S/S','Siddhirganj 132/33/11KV S/S',
         'Siddhirganj  33/11KV S/S','Demra 33/11KV S/S','Mondalpara 33/11KV S/S','Khanpur 33/11KV S/S','Matuail 33/11KV S/S','Matuail 132/33 KV S/S',
         'Sarulia  33/11KV S/S','Maniknagar 33/11KV S/S','IG Gate GIS 33/11 kV','Motijheel old 33/11 kV','Mitford 33/11 kV','Biddyut Bhaban 33/11 KV',
         'Nandalalpur 33/11 kV','Dapa 33/11 kV','Laxmi Narayan Cotton Mill 33/11 kV','Amulia 33/11 kv','New Fatullah 132/33 KV SS',
         'New Fatullah 33/11 KV SS','P & T 33/11 KV SS','Motijheel 132/33 KV SS','Motijheel 33/11 KV SS (new)','Kazla 132/133 KV SS',
         'Kamalapur Railway 33/11 KV SS','Char Syedpur 132/33KV S/S','Char Syedpur 33/11 KV S/S New','Postogola 132/33 KV S/S'
]

substation_choice = st.selectbox("Pick one Substation from Below",ss_list)
st.markdown("""---""")
col1, col2, col3, col4 = st.columns(4)
with col1:
    tableview = st.button("Click to Show Substation-wise Table")
with col2:
    tablehide = st.button("Click to Hide Substation-wise Table")
with col3:
    graphview = st.button("Click to Show Substation-wise Graph")
with col4:
    graphhide = st.button("Click to Hide Substation-wise Graph")
if(tableview):
    df_show=df_selection.query("Substation_Name==@substation_choice")
    df_show["Consumption"]=df_show["Consumption"].astype(int)
    df_show["Corrected_Consumption"]=df_show["Corrected_Consumption"].astype(int)
    st.write(df_show[["Substation_Name","Feeder_Name","CF","Opening_Reading","Closing_Reading","Difference","OMF","Consumption","Corrected_Consumption","NOCS"]])
    col1, col2, col3= st.columns(3)
    col1.write("Consumption : " + str(df_show["Consumption"].sum()))
    col2.write("Corrected Consumption: " +str(df_show["Corrected_Consumption"].sum()))
    col3.write("Substation Loss: "+str(((df_show["Corrected_Consumption"].sum())-(df_show["Consumption"].sum()))/(df_show["Consumption"].sum())*100)+"%")
    export_as_pdf("Substation Name: "+substation_choice,df_show[["Feeder_Name","CF","Opening_Reading","Closing_Reading","OMF","Consumption","Corrected_Consumption","NOCS"]],'table')

elif(tablehide): st.markdown("---")
elif(graphview):
    consumption_by_substation=df_selection.query("Substation_Name==@substation_choice")[["Feeder_Name","Corrected_Consumption","NOCS"]]
    temp_pt =consumption_by_substation[consumption_by_substation['Corrected_Consumption']!=0]
    temp_pt['Corrected_Consumption']=temp_pt['Corrected_Consumption'].astype(int)
    temp_pt['Corrected_Consumption']=temp_pt['Corrected_Consumption'].abs()
    summary_sb = px.sunburst(temp_pt,
        path=['NOCS','Feeder_Name','Corrected_Consumption'],
        values=temp_pt["Corrected_Consumption"],
        color =temp_pt["NOCS"],
        color_continuous_scale = ['red','yellow','green'],
        title='Feeder-wise Consumption',
        width=1500,
        height= 800
    )
    summary_sb.update_layout(
        title_font_size = 20, 
        title_font_family ='Arial'
    )

    st.plotly_chart(summary_sb, use_container_width=True)
    
    fig_nocs_ss = px.bar(
    consumption_by_substation,
    y="Corrected_Consumption",
    x="Feeder_Name",
    labels= consumption_by_substation["Corrected_Consumption"],
    orientation = "v",
    title="<b>Feeder Wise Consumption</b>",
    color="NOCS",
    text_auto = "0.2s",
    template="plotly_dark",
    height=600
    )

    fig_nocs_ss.update_layout(
        font_size = 15,
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=(dict(showgrid=False))
    )


    st.plotly_chart(fig_nocs_ss, use_container_width=True)
    

elif(graphhide): st.markdown("---")

st.markdown("""---""")
nocs_list=['Motijheel','Khilgaon','Lalbag','Kazla','Postogola','Banglabazar','N.Gonj (West)','Siddirgonj','Bashabo','Narinda',
           'Maniknagar','Jurain','Shyampur','Swamibag','Bangshal','N.Gonj (East)','Fatullah','Mugdapara','Tejgaon','Satmasjid',
           'Paribag','Kakrail','Moghbazar','Dhanmondi','Ramna','Shyamoli','Shere b.nagar','Rajarbag','Jigatola','Azimpur','Demra',
           'Matuail','Sytalakhya','Kamrangirchar','Banosree','Adabor'
]
nocs_choice = st.selectbox("Pick one NOCS from Below",nocs_list)
col1, col2, col3, col4 = st.columns(4)
with col1:
    tableview2 = st.button("Click to Show NOCS-wise Table")
with col2:
    tablehide2 = st.button("Click to Hide NOCS-wise Table ")
with col3:
    graphview2 = st.button("Click to Show NOCS-wise Graph ")
with col4:
    graphhide2 = st.button("Click to Hide NOCS-wise Graph ")
if(tableview2):
    df_show=df_selection.query("NOCS==@nocs_choice")
    df_show["Consumption"]=df_show["Consumption"].astype(int)
    df_show["Corrected_Consumption"]=df_show["Corrected_Consumption"].astype(int)
    st.write(df_show[["NOCS","Substation_Name","Feeder_Name","Consumption","Corrected_Consumption"]])
    col1, col2= st.columns(2)
    col1.write("Consumption : " + str(df_show["Consumption"].sum()))
    col2.write("Corrected Consumption: " +str(df_show["Corrected_Consumption"].sum()))
    export_as_pdf("NOCS Name: "+nocs_choice,df_show[["Substation_Name","Feeder_Name","CF","Opening_Reading","Closing_Reading","OMF","Consumption","Corrected_Consumption"]],'table')
elif(tablehide2): st.markdown("---")
elif(graphview2):
    consumption_by_feeder=df_selection.query("NOCS==@nocs_choice")[["Substation_Name","Feeder_Name","Corrected_Consumption"]]
    temp_pt =consumption_by_feeder[consumption_by_feeder['Corrected_Consumption']!=0]
    temp_pt['Corrected_Consumption']=temp_pt['Corrected_Consumption'].astype(int)
    temp_pt['Corrected_Consumption']=temp_pt['Corrected_Consumption'].abs()
    summary_sb = px.sunburst(temp_pt,
        path=['Substation_Name','Feeder_Name','Corrected_Consumption'],
        values=temp_pt["Corrected_Consumption"],
        color =temp_pt["Substation_Name"],
        color_continuous_scale = ['red','yellow','green'],
        title='Feeder-wise Consumption',
        width=1500,
        height= 800
    )
    summary_sb.update_layout(
        title_font_size = 20, 
        title_font_family ='Arial'
    )

    st.plotly_chart(summary_sb, use_container_width=True)

    fig_nocs_feeder = px.bar(
    consumption_by_feeder,
    y=consumption_by_feeder["Corrected_Consumption"],
    x=consumption_by_feeder["Feeder_Name"],
    labels= consumption_by_feeder["Corrected_Consumption"],
    orientation = "v",
    title="<b>Feeder Wise Consumption</b>",
    color="Substation_Name",
    text_auto = "0.2s",
    template="plotly_dark",
    height=800
    )

    fig_nocs_feeder.update_layout(
        font_size= 15,
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=(dict(showgrid=False))
    )
    
    st.plotly_chart(fig_nocs_feeder, use_container_width=True)

elif(graphhide2): st.markdown("---")


html_about ="""
      <br>
      <br>
      <h3><center>Developed By</center></h3>
      <p>This Web-Application has been developed by Abu Md. Raihan, Sub-Divisional Engineer, Tariff & Energy Audit, Dhaka Power Distribution Company (Ltd.) </p>"""
st.markdown(html_about, unsafe_allow_html=True)






