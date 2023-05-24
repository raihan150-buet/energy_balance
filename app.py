import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit


st.set_page_config(page_title="Energy Balance Software", page_icon=":bar_chart:", layout="wide")

# ---- READ EXCEL ----
@st.cache
def get_data_from_excel():
    df = pd.read_excel(
        io="EB.xlsx",
        engine="openpyxl",
        sheet_name="Linked_11KV",
        skiprows=0,
        usecols="B:O",
        nrows=1117,
    )

    return df

df_selection = get_data_from_excel()

# # ---- SIDEBAR ----
# st.sidebar.header("Please Filter Here:")
# nocs = st.sidebar.multiselect(
#     "Select the NOCS:",
#     options=df["NOCS"].unique(),
#     default=df["NOCS"].unique()
# )

# substation = st.sidebar.multiselect(
#     "Select the Substation Name:",
#     options=df["Substation_Name"].unique(),
#     default=df["Substation_Name"].unique(),
# )



# ---- MAINPAGE ----
st.title(":bar_chart: Energy Balance Dashboard")
st.markdown("##")

# TOP KPI's
df_selection["Consumption"]=df_selection["Consumption"].astype(int)
df_selection["Corrected_Consumption"]=df_selection["Corrected_Consumption"].astype(int)
consumption = df_selection["Consumption"].sum()
consumption_corrected = df_selection["Corrected_Consumption"].sum()

left_column, right_column = st.columns(2)
with left_column:
    st.subheader("Total Consumption:")
    st.subheader(f"Unit {consumption:,}")
with right_column:
    st.subheader("Total Corrected Consumption:")
    st.subheader(f"Unit {consumption_corrected}")

st.markdown("""---""")

consumption_by_nocs = (
    df_selection.groupby(by=["NOCS"])["Corrected_Consumption"].sum()
)
fig_nocs_consumption = px.bar(
    consumption_by_nocs,
    x="Corrected_Consumption",
    y=consumption_by_nocs.index,
    orientation="h",
    title="<b>Consumption by NOCS</b>",
    color_discrete_sequence=["#0083B8"] * len(consumption_by_nocs),
    template="plotly_white",
)
fig_nocs_consumption.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

st.divider()
st.plotly_chart(fig_nocs_consumption, use_container_width=True)
st.divider()

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

st.dataframe(df_selection.pivot_table(index=["NOCS"], values=["Consumption","Corrected_Consumption"]))
st.dataframe(df_selection.pivot_table(index=["Substation_Name"], values=["Consumption","Corrected_Consumption"]))
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
st.divider()
substation_choice = st.selectbox("Pick one Substation from Below",ss_list)
st.divider()
st.markdown("""---""")
df_show=df.query("Substation_Name==@substation_choice")
st.divider()
st.write(df_show[["Substation_Name","Feeder_Name","CF","Opening_Reading","Closing_Reading","Difference","OMF","Consumption","Corrected_Consumption","NOCS"]].astype(str))
st.divider()
col1, col2, col3= st.columns(3)
st.divider()
col1.write("Consumption : " + str(df_show["Consumption"].sum()))
col2.write("Corrected Consumption: " +str(df_show["Corrected_Consumption"].sum()))
col3.write("Substation Loss: "+str(((df_show["Corrected_Consumption"].sum())-(df_show["Consumption"].sum()))/(df_show["Consumption"].sum())*100)+"%")
st.divider()
st.markdown("""---""")
nocs_list=['Motijheel','Khilgaon','Lalbag','Kazla','Postogola','Banglabazar','N.Gonj (West)','Siddirgonj','Bashabo','Narinda',
           'Maniknagar','Jurain','Shyampur','Swamibag','Bangshal','N.Gonj (East)','Fatullah','Mugdapara','Tejgaon','Satmasjid',
           'Paribag','Kakrail','Moghbazar','Dhanmondi','Ramna','Shyamoli','Shere b.nagar','Rajarbag','Jigatola','Azimpur','Demra',
           'Matuail','Sytalakhya','Kamrangirchar','Banosree','Adabor'
]
st.divider()
nocs_choice = st.selectbox("Pick one NOCS from Below",nocs_list)
df_show=df.query("NOCS==@nocs_choice")
st.write(df_show[["NOCS","Substation_Name","Feeder_Name","Consumption","Corrected_Consumption"]].astype(str))
st.divider()
col1, col2= st.columns(2)
st.divider()
col1.write("Consumption : " + str(df_show["Consumption"].sum()))
col2.write("Corrected Consumption: " +str(df_show["Corrected_Consumption"].sum()))
st.divider()

