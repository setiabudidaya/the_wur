import streamlit as st
import pandas as pd
import os
import altair as alt

# --- THE-WUR SUBJECTS ---
the_wur_subjects = [
    "Arts & Humanities",
    "Business & Economics",
    "Clinical, Pre-Clinical & Health",
    "Computer Science",
    "Education",
    "Engineering",
    "Law",
    "Life Sciences",
    "Physical Sciences",
    "Psychology",
    "Social Sciences"
]

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Dasbor THE-WUR",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state for selected row
if 'selected_row_data' not in st.session_state:
    st.session_state.selected_row_data = None

# --- HEADER ---
st.title("Dasbor Rekapitulasi Data THE-WUR 📊")
st.markdown("Dasbor ini dirancang untuk merekapitulasi data dari berbagai program studi ke dalam 11 subjek perankingan THE-WUR.")
st.markdown("Data di aplikasi ini disimpan secara lokal dalam file Excel.")
st.markdown("---")

# --- FILE PATH INPUT ---
st.sidebar.subheader("Pengaturan File Excel")
excel_file_path_input = st.sidebar.text_input(
    "Masukkan path file Excel",
    value='the_wur_data.xlsx', # Default value
    help="Masukkan path lengkap atau relatif ke file Excel untuk penyimpanan data."
)

# Update the global variable (or use session state)
excel_file_path = excel_file_path_input
st.sidebar.info(f"File Excel yang digunakan: `{excel_file_path}`")


# --- FORMULIR INPUT DATA ---
with st.form(key='data_input_form'):
    st.subheader("Input Data per Program Studi")

    # Input dasar - Populate with selected row data if available
    program_studi_default = st.session_state.selected_row_data.get("program_studi", "") if st.session_state.selected_row_data else ""
    subject_pilihan_default = st.session_state.selected_row_data.get("subject", the_wur_subjects[0]) if st.session_state.selected_row_data else the_wur_subjects[0]

    program_studi = st.text_input("Nama Program Studi (contoh: Arsitektur)", value=program_studi_default)
    # Handle case where default subject might not be in the list
    subject_index = the_wur_subjects.index(subject_pilihan_default) if subject_pilihan_default in the_wur_subjects else 0
    subject_pilihan = st.selectbox("Pilih Subjek THE-WUR yang Sesuai", options=the_wur_subjects, index=subject_index)


    st.markdown("---")
    st.subheader("Data Dosen dan Mahasiswa")
    col1, col2 = st.columns(2)

    with col1:
        st.write("### Data Dosen")
        academic_staff_default = st.session_state.selected_row_data.get("academic_staff", 0) if st.session_state.selected_row_data else 0
        academic_staff_overseas_default = st.session_state.selected_row_data.get("academic_staff_overseas", 0) if st.session_state.selected_row_data else 0
        academic_staff_female_default = st.session_state.selected_row_data.get("academic_staff_female", 0) if st.session_state.selected_row_data else 0
        research_staff_default = st.session_state.selected_row_data.get("research_staff", 0) if st.session_state.selected_row_data else 0

        academic_staff = st.number_input("Jumlah Dosen", min_value=0, step=1, value=academic_staff_default)
        academic_staff_overseas = st.number_input("Jumlah Dosen (Dari Luar Negeri)", min_value=0, step=1, value=academic_staff_overseas_default)
        academic_staff_female = st.number_input("Jumlah Dosen (Perempuan)", min_value=0, step=1, value=academic_staff_female_default)
        research_staff = st.number_input("Jumlah Peneliti", min_value=0, step=1, value=research_staff_default)

    with col2:
        st.write("### Data Mahasiswa")
        total_students_default = st.session_state.selected_row_data.get("total_students", 0) if st.session_state.selected_row_data else 0
        students_overseas_default = st.session_state.selected_row_data.get("students_overseas", 0) if st.session_state.selected_row_data else 0
        students_female_default = st.session_state.selected_row_data.get("students_female", 0) if st.session_state.selected_row_data else 0
        bachelors_students_default = st.session_state.selected_row_data.get("bachelors_students", 0) if st.session_state.selected_row_data else 0
        masters_students_default = st.session_state.selected_row_data.get("masters_students", 0) if st.session_state.selected_row_data else 0
        doctorate_students_default = st.session_state.selected_row_data.get("doctorate_students", 0) if st.session_state.selected_row_data else 0
        exchange_students_abroad_default = st.session_state.selected_row_data.get("exchange_students_abroad", 0) if st.session_state.selected_row_data else 0

        total_students = st.number_input("Total Mahasiswa", min_value=0, step=1, value=total_students_default)
        students_overseas = st.number_input("Jumlah Mahasiswa (Dari Luar Negeri)", min_value=0, step=1, value=students_overseas_default)
        students_female = st.number_input("Jumlah Mahasiswa (Perempuan)", min_value=0, step=1, value=students_female_default)
        bachelors_students = st.number_input("Jumlah Mahasiswa Sarjana", min_value=0, step=1, value=bachelors_students_default)
        masters_students = st.number_input("Jumlah Mahasiswa Magister", min_value=0, step=1, value=masters_students_default)
        doctorate_students = st.number_input("Jumlah Mahasiswa Doktor", min_value=0, step=1, value=doctorate_students_default)
        exchange_students_abroad = st.number_input("Jumlah Mahasiswa Pertukaran (Keluar Negeri)", min_value=0, step=1, value=exchange_students_abroad_default)

    st.markdown("---")
    st.subheader("Data Lulusan dan Dana Yang Dikelola")
    col3, col4 = st.columns(2)

    with col3:
        st.write("### Data Lulusan")
        undergrad_degrees_awarded_default = st.session_state.selected_row_data.get("undergrad_degrees_awarded", 0) if st.session_state.selected_row_data else 0
        doctorates_awarded_default = st.session_state.selected_row_data.get("doctorates_awarded", 0) if st.session_state.selected_row_data else 0

        undergrad_degrees_awarded = st.number_input("Jumlah Lulusan Sarjana", min_value=0, step=1, value=undergrad_degrees_awarded_default)
        doctorates_awarded = st.number_input("Jumlah Lulusan Doktor", min_value=0, step=1, value=doctorates_awarded_default)

    with col4:
        st.write("### Data Dana Yang Dikelola")
        total_institutional_income_default = st.session_state.selected_row_data.get("total_institutional_income", 0) if st.session_state.selected_row_data else 0
        research_income_default = st.session_state.selected_row_data.get("research_income", 0) if st.session_state.selected_row_data else 0
        research_income_industry_default = st.session_state.selected_row_data.get("research_income_industry", 0) if st.session_state.selected_row_data else 0

        total_institutional_income = st.number_input("Total Dana Yang Dikelola (Rupiah)", min_value=0, step=1, value=total_institutional_income_default)
        research_income = st.number_input("Dana Yang Dikelola Untuk Penelitian (Rupiah)", min_value=0, step=1, value=research_income_default)
        research_income_industry = st.number_input("Dana Penelitian Yang Berasal Dari Industri (Rupiah)", min_value=0, step=1, value=research_income_industry_default)

    # Change button label based on selected row data
    submit_button_label = "Update Data" if st.session_state.selected_row_data else "Tambahkan Data"
    submitted = st.form_submit_button(submit_button_label)

    if submitted:
        if not program_studi:
            st.error("Nama Program Studi tidak boleh kosong.")
        else:
            data_to_save = {
                "program_studi": program_studi,
                "subject": subject_pilihan,
                "academic_staff": academic_staff,
                "academic_staff_overseas": academic_staff_overseas,
                "academic_staff_female": academic_staff_female,
                "research_staff": research_staff,
                "total_students": total_students,
                "students_overseas": students_overseas,
                "students_female": students_female,
                "bachelors_students": bachelors_students,
                "masters_students": masters_students,
                "doctorate_students": doctorate_students,
                "exchange_students_abroad": exchange_students_abroad,
                "undergrad_degrees_awarded": undergrad_degrees_awarded,
                "doctorates_awarded": doctorates_awarded,
                "total_institutional_income": total_institutional_income,
                "research_income": research_income,
                "research_income_industry": research_income_industry
            }

            try:
                df_existing = pd.DataFrame() # Initialize empty DataFrame
                if os.path.exists(excel_file_path):
                    try:
                        df_existing = pd.read_excel(excel_file_path)
                    except Exception as e:
                        st.error(f"Error reading existing Excel file: {e}")
                        df_existing = pd.DataFrame() # Ensure df_existing is an empty DataFrame on read error

                # Check if necessary columns exist before attempting to access them
                if not df_existing.empty and 'program_studi' in df_existing.columns and 'subject' in df_existing.columns:
                    # Convert relevant columns to string to avoid potential dtype issues during comparison
                    df_existing['program_studi'] = df_existing['program_studi'].astype(str)
                    df_existing['subject'] = df_existing['subject'].astype(str)

                    if st.session_state.selected_row_data:
                        # Update existing data
                        # Find the index of the selected row using the original data from session state
                        # Ensure comparison is done with string types to avoid potential mismatches
                        selected_index_in_df = df_existing[
                            (df_existing['program_studi'] == str(st.session_state.selected_row_data['program_studi'])) &
                            (df_existing['subject'] == str(st.session_state.selected_row_data['subject']))
                        ].index

                        if not selected_index_in_df.empty:
                            # Use iloc with the found index to update the row
                            df_existing.iloc[selected_index_in_df[0]] = data_to_save
                            st.success(f"Data untuk **{program_studi}** berhasil diperbarui di Excel!")
                        else:
                            st.error("Could not find the selected row in the Excel file for update.")
                    else:
                         # Check if a row with the same Program Studi and Subject already exists before adding
                        existing_row_index = df_existing[
                            (df_existing['program_studi'].astype(str) == str(program_studi)) &
                            (df_existing['subject'].astype(str) == str(subject_pilihan))
                        ].index

                        if not existing_row_index.empty:
                             # Update the existing row instead of adding a new one
                             df_existing.loc[existing_row_index[0]] = data_to_save
                             st.success(f"Data untuk **{program_studi}** berhasil diperbarui di Excel (mencegah duplikasi)!")
                        else:
                            # Add new data if no existing row is found
                            df_new = pd.DataFrame([data_to_save])
                            df_existing = pd.concat([df_existing, df_new], ignore_index=True)
                            st.success(f"Data untuk **{program_studi}** berhasil disimpan ke Excel!")
                else:
                     # If DataFrame is empty or columns are missing, just add the new data
                     df_new = pd.DataFrame([data_to_save])
                     df_existing = pd.concat([df_existing, df_new], ignore_index=True)
                     st.success(f"Data untuk **{program_studi}** berhasil disimpan ke Excel!")


                # Save the updated DataFrame (overwrite the file)
                df_existing.to_excel(excel_file_path, index=False)

                # Clear the selection after saving/updating
                st.session_state.selected_row_data = None
                st.rerun()

            except FileNotFoundError:
                 st.error(f"Error saving/updating data: The file '{excel_file_path}' was not found.")
            except PermissionError:
                 st.error(f"Error saving/updating data: Permission denied to write to '{excel_file_path}'. Please check file permissions.")
            except Exception as e:
                st.error(f"Gagal menyimpan/memperbarui data di Excel: {e}")


st.markdown("---")

# --- REKAPITULASI DATA ---
st.header("Rekapitulasi Data")
st.info(f"Data rekapitulasi diambil dari file Excel: `{excel_file_path}`")


# Mengambil data from Excel
@st.cache_data(ttl=60)
def get_all_data():
    if os.path.exists(excel_file_path):
        try:
            df = pd.read_excel(excel_file_path)
            return df
        except FileNotFoundError:
            st.error(f"Error reading data: The file '{excel_file_path}' was not found.")
            return pd.DataFrame() # Return empty dataframe on error
        except PermissionError:
            st.error(f"Error reading data: Permission denied to read from '{excel_file_path}'. Please check file permissions.")
            return pd.DataFrame() # Return empty dataframe on error
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return pd.DataFrame() # Return empty dataframe on error
    else:
        return pd.DataFrame()

df = get_all_data()

if df.empty:
    st.info("Belum ada data yang tersimpan.")
else:
    st.subheader("Ringkasan Total Data per Subjek THE-WUR")

    # Mengelompokkan data berdasarkan subjek THE-WUR
    rekapitulasi_data = df.groupby('subject').agg(
        Total_Staf_Akademik=('academic_staff', 'sum'),
        Total_Staf_Akademik_Luar_Negeri=('academic_staff_overseas', 'sum'),
        Total_Staf_Akademik_Perempuan=('academic_staff_female', 'sum'),
        Total_Staf_Riset=('research_staff', 'sum'),
        Total_Mahasiswa=('total_students', 'sum'),
        Total_Mahasiswa_Luar_Negeri=('students_overseas', 'sum'),
        Total_Mahasiswa_Perempuan=('students_female', 'sum'),
        Total_Mahasiswa_Sarjana=('bachelors_students', 'sum'),
        Total_Mahasiswa_Magister=('masters_students', 'sum'),
        Total_Mahasiswa_Doktor=('doctorate_students', 'sum'),
        Total_Mahasiswa_Pertukaran=('exchange_students_abroad', 'sum'),
        Total_Gelar_Sarjana=('undergrad_degrees_awarded', 'sum'),
        Total_Gelar_Doktor=('doctorates_awarded', 'sum'),
        Total_Pendapatan_Institusi=('total_institutional_income', 'sum'),
        Total_Pendapatan_Riset=('research_income', 'sum'),
        Total_Pendapatan_Riset_Industri=('research_income_industry', 'sum')
    ).reset_index()

    # Menampilkan rekapitulasi dalam tabel
    st.dataframe(rekapitulasi_data, use_container_width=True)

    st.markdown("---")
    st.subheader("Data Mentah Program Studi yang Disimpan")
    # Display the dataframe without direct selection handling here
    st.dataframe(df.sort_values("program_studi"), use_container_width=True)

    # Add a text input for the user to specify the row index to edit
    # Only show the number input if there is data to select
    if not df.empty:
        row_index_to_edit = st.number_input("Masukkan nomor baris untuk diedit (dimulai dari 0)", min_value=0, max_value=len(df)-1 if not df.empty else 0, step=1, key='row_index_input')

        # Add a button to load the selected row data into the form
        if st.button("Pilih Baris untuk Diedit"):
            if not df.empty and 0 <= row_index_to_edit < len(df):
                st.session_state.selected_row_data = df.iloc[row_index_to_edit].to_dict()
                st.rerun()
            else:
                st.error("Nomor baris tidak valid.")


    # Add a button to clear the selection
    if st.session_state.selected_row_data:
        if st.button("Clear Selection"):
            st.session_state.selected_row_data = None
            st.rerun()


    st.markdown("---")
    st.subheader("Visualisasi Data Rekapitulasi per Subjek")

    # Select all aggregated numerical columns for visualization
    columns_to_plot_by_subject = rekapitulasi_data.select_dtypes(include='number').columns.tolist()

    if not rekapitulasi_data.empty:
        for column in columns_to_plot_by_subject:
            st.markdown(f"#### {column} per Subjek THE-WUR")
            # Create a bar chart using Altair
            chart = alt.Chart(rekapitulasi_data).mark_bar().encode(
                x=alt.X('subject', title='Subjek THE-WUR', sort='-y'), # Sort by y-axis value descending
                y=alt.Y(column, title=column),
                tooltip=['subject', column],
                color='subject' # Add color encoding by subject
            ).properties(
    #            title=f'{column} per Subjek THE-WUR'
            ).interactive() # Make the chart interactive

            # Display the chart
            st.altair_chart(chart, use_container_width=True)
    else:
        st.info("Tidak ada data untuk divisualisasikan per subjek.")

    st.markdown("---")
    st.subheader("Visualisasi Total Akumulasi Data dari Semua Program Studi")

    if not df.empty:
        # Calculate the total sum of numerical columns across all program studies
        total_akumulasi = df.select_dtypes(include='number').sum().reset_index()
        total_akumulasi.columns = ['Metric', 'Total']


        if not total_akumulasi.empty:
            # Define metric categories based on user's latest specification and debugging output
            income_metrics = ['total_institutional_income', 'research_income', 'research_income_industry']
            staff_metrics = ['academic_staff', 'academic_staff_overseas', 'academic_staff_female', 'research_staff']
            # Corrected student metrics based on previous aggregation logic and user input
            student_metrics = ['total_students', 'students_overseas', 'students_female', 'bachelors_students', 'masters_students', 'doctorate_students']
            # Corrected graduate metrics based on previous aggregation logic and user input
            graduate_metrics = ['undergrad_degrees_awarded', 'doctorates_awarded']


            # Create and display charts for each category
            chart_categories = {
                "Pendapatan": income_metrics,
                "Staf": staff_metrics,
                "Mahasiswa": student_metrics,
                "Lulusan": graduate_metrics
            }

            for category, metrics in chart_categories.items():
                # Filter total_akumulasi based on the metrics in the current category
                category_data = total_akumulasi[total_akumulasi['Metric'].isin(metrics)]
                if not category_data.empty:
                    st.markdown(f"#### Total Akumulasi Data: {category}")
                    chart = alt.Chart(category_data).mark_bar().encode(
                        x=alt.X('Metric', title='Metrik Data', sort='-y'),
                        y=alt.Y('Total', title='Total Akumulasi'),
                        tooltip=['Metric', 'Total']
                    ).properties(
     #                   title=f'Total Akumulasi Data: {category}'
                    ).interactive()
                    st.altair_chart(chart, use_container_width=True)

    else:
        st.info("Tidak ada data untuk diakumulasi dan divisualisasikan.")

