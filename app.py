st.set_page_config(page_title='Skill Profiling Automator', layout='wide')
st.title('Skill Profiling — Master Updater')


# Sidebar: upload master optionally
master_file = st.sidebar.file_uploader('Upload existing Master.xlsx (optional)', type=['xlsx'])
if master_file:
master_df = pd.read_excel(master_file, sheet_name=0)
else:
master_df = pd.DataFrame(columns=['Employee ID','Employee Name','Skill','Level','Last Updated','Cycle'])


uploaded = st.file_uploader('Upload employee files (single .xlsx or a .zip of many)', type=['xlsx','xls','zip'], accept_multiple_files=False)
cycle_label = st.text_input('Cycle label (e.g., 2025-Cycle2)', value='2025-Cycle2')


if uploaded is not None:
bytes_data = uploaded.read()
employee_dfs = []
if uploaded.name.lower().endswith('.zip'):
employee_dfs = parse_zip_of_excels(BytesIO(bytes_data))
else:
try:
employee_dfs = [parse_employee_excel(BytesIO(bytes_data).getvalue())]
except Exception as e:
# try reading as many sheets
try:
x = pd.read_excel(BytesIO(bytes_data), sheet_name=None)
for k,v in x.items():
employee_dfs.append(v)
except Exception as e2:
st.error('Could not parse uploaded file: ' + str(e))
st.write
# Download XLSX after merge
from io import BytesIO
o
