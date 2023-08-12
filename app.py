import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit
from fpdf import FPDF
import base64

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
        nrows=1180,
    )

    return df

df_selection = get_data_from_excel()



# ---- MAINPAGE ----
st.title(":bar_chart: Energy Balance Dashboard")
st.markdown("##")

# TOP KPI's
# consumption = int(df_selection["Consumption"].sum())
consumption_corrected = int(df_selection["Corrected_Consumption"].sum())

left_column,right_column = st.columns(2)
with left_column:
    st.subheader("Total Import at 33 KV Level (All NOCS):- ")
with right_column:
    st.subheader(f"  {consumption_corrected} KWH")

st.markdown("""----""")
consumption_by_nocs = (
    df_selection.groupby(by=["NOCS"])["Corrected_Consumption"].sum().reset_index()[1:36]
)
consumption_by_nocs["Corrected_Consumption"]=consumption_by_nocs['Corrected_Consumption'].astype(int)
#-----------------------TreeMap-------------------#
summary_tree = px.treemap(consumption_by_nocs,
                 path=consumption_by_nocs.columns,
                 values=consumption_by_nocs["Corrected_Consumption"],
                 color =consumption_by_nocs["NOCS"],
                 color_continuous_scale = ['red','yellow','green'],
                 title='Summary of Import',
                 width = 1000,
                 height = 700,
                 )

summary_tree.update_layout(
    font_size = 15,
    title_font_size = 50, 
    title_font_family ='Arial',
)

st.plotly_chart(summary_tree, use_container_width=True)

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
    title_font_size = 50, 
    title_font_family ='Arial',
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

fig_nocs_consumption.update_traces(textfont_size=45, textangle=-90, textposition="inside", cliponaxis=False)


st.plotly_chart(fig_nocs_consumption, use_container_width=True)
ss_wise = df_selection.groupby(['Substation_Name','NOCS'])['Corrected_Consumption'].sum().reset_index()
ss_wise = ss_wise[ss_wise['Corrected_Consumption']!=0]
ss_wise['Corrected_Consumption']=ss_wise['Corrected_Consumption'].astype(int)
summary_sb = px.treemap(ss_wise,
    path=['Substation_Name','NOCS','Corrected_Consumption'],
    values=ss_wise["Corrected_Consumption"],
    color =ss_wise["Corrected_Consumption"] ,
    color_continuous_scale=['Green','Violet','Yellow','Red'],
    title='Summary of Import',
    width = 800,
    height = 1000
)
summary_sb.update_layout(
    title_font_size = 50, 
    title_font_family ='Arial'
)

st.plotly_chart(summary_sb, use_container_width=True)
 
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
    # Select a font as Arial, regular, 10
    pdf.set_font('Arial', '', 7)
    # Loop over to print each data in the table
    for row in df.itertuples():
        for col in cols:
            value = str(getattr(row, col))
            pdf.cell(table_cell_width, table_cell_height, value, align='C', border=1)
        pdf.ln(table_cell_height)
    pdf.ln()
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0,10, txt="Total Consumption:- "+str(df["Consumption"].sum()))
    pdf.ln()
    pdf.cell(0,10, txt="Total Corrected Consumption:- "+str(df["Corrected_Consumption"].sum()))

def export_as_pdf(report_text,data):
    pdf = FPDF('landscape','mm',"A4")
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, txt=report_text, align="C")
    pdf.ln()
    output_df_to_pdf(pdf,data)
    
    html = create_download_link(pdf.output(dest="S").encode("latin-1"), "Report")

    return(st.markdown(html, unsafe_allow_html=True))

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
    export_as_pdf("Substation Name: "+substation_choice,df_show[["Feeder_Name","CF","Opening_Reading","Closing_Reading","OMF","Consumption","Corrected_Consumption","NOCS"]])

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
    labels = "Corrected_Consumption",
    orientation = "v",
    title="<b>Feeder Wise Consumption</b>",
    color="NOCS",
    template="plotly_dark",
    height=600
    )

    fig_nocs_ss.update_layout(
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
    export_as_pdf("NOCS Name: "+nocs_choice,df_show[["Substation_Name","Feeder_Name","CF","Opening_Reading","Closing_Reading","OMF","Consumption","Corrected_Consumption"]])
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
    y="Corrected_Consumption",
    x="Feeder_Name",
    labels = "Corrected_Consumption",
    orientation = "v",
    title="<b>Feeder Wise Consumption</b>",
    color="Corrected_Consumption",
    template="plotly_dark",
    height=600
    )

    fig_nocs_feeder.update_layout(
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
