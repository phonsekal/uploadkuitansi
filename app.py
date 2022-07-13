import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from num2words import num2words
from datetime import datetime
from docx import Document
from docxcompose.composer import Composer
#from PIL import Image
import numpy as np
from streamlit_lottie import st_lottie
import json
import requests
from streamlit_option_menu import option_menu
st.set_page_config(page_title='Bukan Web Personal || dedesaputra@2022', page_icon = ":coffee:", layout = 'centered', initial_sidebar_state = 'auto')
@st.experimental_memo
def kalender_indo(value):
    a = (value).strftime("%d %B %Y")
    kal = a.replace("January", "Januari").replace("February", "Februari").replace("March", "Maret").replace("May", "Mei").replace("June", "Juni").replace("July", "Juli").replace("August", "Agustus").replace("October", "Oktober").replace("December", "Desember")
    return kal
def bulan_indo(value):
    a = (value)
    kal = a.replace("January", "Januari").replace("February", "Februari").replace("March", "Maret").replace("May", "Mei").replace("June", "Juni").replace("July", "Juli").replace("August", "Agustus").replace("October", "Oktober").replace("December", "Desember")
    return kal

def transform_to_rupiah_format(value):
    str_value = str(value)
    separate_decimal = str_value.split(".")
    after_decimal = separate_decimal[0]
    before_decimal = separate_decimal[1]
    reverse = after_decimal[::-1]
    temp_reverse_value = ""
    for index, val in enumerate(reverse):
        if (index + 1) % 3 == 0 and index + 1 != len(reverse):
            temp_reverse_value = temp_reverse_value + val + "."
        else:
            temp_reverse_value = temp_reverse_value + val
    temp_result = temp_reverse_value[::-1]
    return "Rp" + temp_result + ",00" 

def rupiah_strip(value):
    a = str(transform_to_rupiah_format(float(value)))
    ubah = a.replace("Rp0,00", "-")
    return ubah
        
def load_lottiefile(filepath: str):
    with open(filepath, "r") as f:
        return json.load(f)
st.markdown("""
<style>
#MainMenu {
    visibility:hidden;
}
.css-1q1n0ol.egzxvld0 {
    visibility:hidden;
}
footer:after{
    visibility:visible;
    content:'Created by Artupas eD using Streamlit @ 2022';
    display:blok;
    position:relative;
    color:#FECD45;
}
.css-79elbk.e1fqkh3o8 {
  margin: 0;
  padding: 0;
  box-sizing:birder-box;
  width:21rem ;
  #background-color: #f1f1f1;
  position: fixed;
  font-family:"Helvetica Neue", Helvetica;
  font-size: 30px;
}
.css-rytr0c {
    visibility:hidden;
}

.css-1siy2j7.e1fqkh3o3 {
    background-color:#e8f9fd;
}
.css-10trblm.e16nr0p30{
    color:#2568FB;
    font-family: averta_stdregular,Helvetica Neue,Helvetica;
    font-size:2.15rem;
}
.css-znku1x.e16nr0p33 {
color:#000c32;
font-family: averta_stdregular,Helvetica Neue,Helvetica;
}
.css-183lzff.eyqtai90{
color:#3AB0FF;
font-family: averta_stdregular,Helvetica Neue,Helvetica;
font-size:1.75rem;
}
.css-znku1x.p {
    text-align: justify;
}
 </style>


""", unsafe_allow_html=True )
with st.sidebar:
    selected = option_menu("Pilih Menu", ["Home", 'Kuitansi Monev', 'Kuitansi Kegiatan', 'Kuitansi Translok'], 
        icons=['house-fill', 'wallet-fill', 'currency-exchange', 'arrow-right-square-fill'], 
        menu_icon="app-indicator", default_index=0,
        styles={
        "container": {"padding": "5!important", "background-color": "#f0f2f6"},
        "icon": {"color": "#FECD45", "font-size": "28px"}, 
        "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
        "nav-link-selected": {"background-color": "#2568FB"},
    }
        )

if selected == "Home":
    st.write("""
    """)
    col5, col4 = st.columns([3, 1])
    with col5:
        st.write(
            """
            # Selamat Datang
            """
        )
    with col4:
        lottie_gambar = load_lottiefile("pages/templates/b.json")
        st_lottie(
            lottie_gambar,
            speed=1,
            reverse=False,
            loop=True,
            quality="low",
            height="50%",
            width="50%",
        )

    st.write("""
        Selamat datang di tampilan muka aplikasi Kuitansi! :wave:
        Untuk menggunakan aplikasi ini, silahkan download format nominatif pada link yang telah disediakan ==> [Format Nominatif](https://drive.google.com/drive/folders/1dAb_hMT04oq6pWGPMMoyiE3QxFffYqQ8?usp=sharing)
        """
    )
    st.info(
        """
        ✏️ **NOTE:** Aplikasi masih dalam tahap pengembangan, jika terdapat error pada aplikasi bisa menghubungi kami!
        """
    )

    st.markdown("***")

    st.text(""" Peraturan tentang perjalanan dinas.

            """)

    st.write("""
        **Peraturan Menteri Keuangan (PMK) Nomor 113/PMK.05/2012**  
        Peraturan Menteri ini mengatur mengenai pelaksanaan dan pertanggungjawaban Perjalanan Dinas bagi Pejabat Negara, Pegawai Negeri, dan Pegawai Tidak Tetap yang dibebankan pada Anggaran Pendapatan dan Belanja Negara... [Selengkapnya](https://jdih.kemenkeu.go.id/fulltext/2012/113~PMK.05~2012PerLamp.pdf)
    """)
    st.write("""
    **Peraturan Menteri Keuangan (PMK) Nomor 60/PMK.02/2021**  
    Peraturan Menteri ini mengatur mengenai Standar Biaya Masukan (SBM) tahun 2022. SBM 2022 Standar Biaya Masukan Tahun Anggaran 2022 adalah satuan biaya berupa harga satuan, tarif, dan indeks yang ditetapkan untuk menghasilkan biaya komponen keluaran dalam penyusunan rencana kerja dan anggaran
kementerian negara/lembaga Tahun Anggaran 2022... [Selengkapnya](https://jdih.kemenkeu.go.id/download/a73998d2-c308-4451-a907-35438a028e80/60~PMK.02~2021Per.pdf)
    """)
    st.write("""
        **Peraturan Menteri Keuangan (PMK) Nomor 190/PMK.05/2012**  
        Peraturan Menteri ini mengatur mengenai tata cara pembayaran dalam rangka pelaksanaan APBN selain tata cara pembayaran dalam rangka pelaksanaan APBN untuk Perwakilan Republik Indonesia di Luar Negeri dan Kementerian Pertahanan dan Tentara Nasional Indonesia... [Selengkapnya](https://peraturan.bpk.go.id/Home/Download/118021/2012%20PMK%20190.pdf)
    """)
    st.write("""
        **Peraturan Direktur Jenderal Perbendaharaan Nomor PER-22/PB/2013**  
        Peraturan Menteri ini mengatur tentang Ketentuan Lebih Lanjut Pelaksanaan Perjalanan Dinas Dalam Negeri Bagi Pejabat Negara, Pegawai Negeri, dan Pegawai Tidak Tetap... [Selengkapnya](https://www.dropbox.com/s/t7lcwms3idiqptx/per_22_pb_2013.pdf?dl=0)
    """)
    col1, col2, col3 = st.columns(3)
    with col1 :
        ""
    with col2 :
        ""
    with col3:
        ""
if selected == "Kuitansi Monev":
    


    st.write("""
    """)
    col5, col4 = st.columns([3, 1])
    with col5:
        st.write(
            """
            # Format Kuitansi Perjadin ke Daerah
            """
        )
    with col4:
        lottie_gambar = load_lottiefile("pages/templates/a.json")
        st_lottie(
            lottie_gambar,
            speed=1,
            reverse=False,
            loop=True,
            quality="low",
            height="54%",
            width="54%",
        )

    st.write("""
        Selamat datang! :wave:
        Silahkan unduh format nominatif berikut, pastikan untuk tidak mengganti nama Sheet dan menghapus Column dan Row. ==> [Format Nominatif](https://drive.google.com/drive/folders/1dAb_hMT04oq6pWGPMMoyiE3QxFffYqQ8)
        """
    )
    st.info(
        """
        ✏️ **NOTE:** Gunakan format nominatif yang sudah disiapkan!
        """
    )
    st.markdown("***")


    form = st.form("Upload file")
    with form:
        excelfile = st.file_uploader("Unggah file nominatif")
        submit = st.form_submit_button("Proses")
    if submit:
        try:
            df1 = pd.read_excel(excelfile, skiprows=1, sheet_name="kuitansi")
            df2 = pd.read_excel(excelfile, sheet_name="kuitansi2")
            mak = (df2['MAK'][0])
            output = (df2['Output'][0])
            sub_output = (df2['Sub Output'][0])
            akun = (df2['Akun'][0])
            kegiatan = (df2['Nama Kegiatan'][0])
            layanan = df2['Layanan'][0] 
            tempat = df2['UPT Tujuan'][0]
            tanggal = bulan_indo(df2['tgl_pergi_pulang'][0])
            nomor_st = df2['no_st'][0]
            tanggal_st = kalender_indo(df2['tgl_st'][0])
            bulan = bulan_indo(df2['bulan'][0])
            tgl_berangkat = kalender_indo(df2['Tgl Berangkat'][0])
            tgl_pulang = kalender_indo(df2['Tgl Pulang'][0])
            for r_idx, r_val in df1.iterrows():
                if (r_val['Kota Tujuan'] == 'Banten'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Serang'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Bogor'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Bandung'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Bekasi'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Jawa Barat'):
                    doctemp = f"pages/templates/monevdekat.docx"
                elif (r_val['Kota Tujuan'] == 'Tangerang'):
                    doctemp = f"pages/templates/monevdekat.docx"
                else:
                    doctemp = f"pages/templates/monevjauh.docx"
                doc = DocxTemplate(doctemp)
                context = {
                    "mak" : mak,
                    "total" : rupiah_strip(r_val['Tiket'] + r_val['Taksi Asal'] + r_val['Taksi Tujuan'] + r_val['Total Uang Harian'] + r_val['Total Hotel']),
                    "output" : output,
                    "sub_output" : sub_output,
                    "akun" : akun,
                    "kegiatan" : kegiatan,
                    "layanan" : layanan,
                    "tempat" : tempat,
                    "tanggal" : tanggal,
                    "nomor_st" : nomor_st,
                    "tanggal_st" : tanggal_st,
                    "asal" : r_val['Kota Asal'],
                    "tujuan" : r_val['Kota Tujuan'],
                    "asal_tujuan" : r_val['Kota Asal'] + "--" + r_val['Kota Tujuan'],
                    "terbilang" : num2words(int(r_val['Tiket'] + r_val['Taksi Asal'] + r_val['Taksi Tujuan'] + r_val['Total Uang Harian']) + r_val['Total Hotel'], lang='id').title() + " Rupiah",
                    "nama" : r_val['Nama'],
                    "nip" : r_val['NIP'],
                    "jabatan" : r_val['Jabatan'],
                    "tiket" : rupiah_strip(r_val['Tiket']),
                    "taksi_jakarta" : rupiah_strip(r_val['Taksi Asal']),
                    "taksi_daerah" : rupiah_strip(r_val['Taksi Tujuan']),
                    "hari" : r_val['Lama Perjadin'],
                    "hari_rp" : rupiah_strip(r_val['Total Uang Harian'] / r_val['Lama Perjadin']),
                    "total_hari" : rupiah_strip(r_val['Total Uang Harian']),
                    "malam" : int(r_val['Lama Perjadin']) - 1,
                    "malam_rp" : rupiah_strip(r_val['Total Hotel'] / (r_val['Lama Perjadin'] - 1)),
                    "total_malam" : rupiah_strip(r_val['Total Hotel']),
                    "bulan" : bulan,
                    "total_spd" : rupiah_strip(r_val['Taksi Asal'] + r_val['Taksi Tujuan']),
                    "hari_ter" : num2words(int(r_val['Lama Perjadin']), lang="id"),
                    "pangkat" :r_val['Pangkat/Gol'],
                    "tgl_berangkat" : tgl_berangkat,
                    "tgl_pulang" : tgl_pulang,
                    }
                doc.render(context)
                output_path = f"pages/OUTPUT/{context['nama']}.docx"
                doc.save(output_path)
            a = st.success("🎉 File kuitansi telah selesai dibuat, silahkan unduh")
            b = st.success("")
            with b:
                np_array = df1["Nama"].to_numpy()
                files2 = list("pages/OUTPUT/" + (np_array) + ".docx")
                composed = f"pages/gabung.docx"
                result = Document(files2[0])
                result.add_page_break()
                composer = Composer(result)
                for i in range(1, len(files2)):
                    doc2 = Document(files2[i])
                    if i != len(files2) -1:
                        doc2.add_page_break()
                    composer.append(doc2)
                composer.save(composed)
                with open(composed, "rb") as file:
                    st.success("🎉 File kuitansi telah selesai dibuat")
                    st.download_button(
                        label = "⬇️ Download File",
                        data=file,
                        file_name="kuitansi.docx",
                        mime="application/octet-stream",
                        key="10000009"
                    )
               
        except:
            st.error("Silahkan upload nominatif terlebih dahulu/gunakan format nominatif yang telah digunakan")
if selected == 'Kuitansi Translok':
    


    st.write("""
    """)
    col5, col4 = st.columns([3, 1])
    with col5:
        st.write(
            """
            # Format Kuitansi Transpor Lokal
            """
        )
    with col4:
        lottie_gambar = load_lottiefile("pages/templates/c.json")
        st_lottie(
            lottie_gambar,
            speed=1,
            reverse=False,
            loop=True,
            quality="low",
            height="54%",
            width="54%",
        )

    st.write("""
        Selamat datang! :wave:
        Untuk menggunakan aplikasi ini, silahkan isi form berikut. Pastikan kolom nominal hanya diisi menggunakan angka.
        """
    )
    st.info(
        """
        ✏️ **NOTE:** Silahkan isi form di bawah ini!
        """
    )
    form = st.form("template kuitansi")
    with form:
        col1, col2 = st.columns(2)
        nama = col1.text_input('Nama')
        nip = col2.text_input("NIP")
        layanan = col1.text_input("Layanan")
        mak = col2.text_input("MAK")
        tanggal = col1.date_input("Tanggal Kegiatan")
        kegiatan = col1.text_area("Nama Kegiatan")
        lokasi = col2.text_input("Lokasi Kegiatan")
        nilai = col2.text_input("Nominal")
        submit = form.form_submit_button("Kirim")
        tgl = tanggal.strftime("%d %B %Y")
        tgl2 = tgl.replace("January", "Januari").replace("February", "Februari").replace("March", "Maret").replace("May", "Mei").replace("June", "Juni").replace("July", "Juli").replace("August", "Agustus").replace("October", "Oktober").replace("December", "Desember")
    if submit:
        try:
            thousands_separator = "."
            fractional_separator = ","
            currency = rupiah_strip(nilai)
            doc = DocxTemplate("pages/templates/kuitansiperjadinapp.docx")
            context = {
                "nama": nama,
                "nip": nip,
                "mak": mak,
                "kegiatan": kegiatan,
                "lokasi": lokasi,
                "tanggal": tgl2,
                "layanan": layanan,
                "terbilang": num2words(int(nilai), lang='id').title() + " Rupiah",
                "uang": currency
            }
            output_name = f'download/{context["nama"]}.docx'
            doc.render(context)   
            doc.save(output_name)
            with open(output_name, "rb") as file:
                st.success("🎉 File kuitansi telah selesai dibuat")
                st.download_button(
                    "⬇️ Download File",
                    data=file,
                    file_name="Hasil.docx",
                    mime="application/octet-stream",
                )
        except:
            st.error("Silahkan isi form terlebih dahulu, kolom nominal hanya boleh diisi angka")

if selected == 'Kuitansi Kegiatan':

    


    st.write("""
    """)
    col5, col4 = st.columns([3, 1])
    with col5:
        st.write(
            """
            # Format Kuitansi Kegiatan
            """
        )
    with col4:
        lottie_gambar = load_lottiefile("pages/templates/k.json")
        st_lottie(
            lottie_gambar,
            speed=1,
            reverse=False,
            loop=True,
            quality="low",
            height="70%",
            width="70%",
        )

    st.write("""
        Selamat datang di tampilan muka aplikasi Kuitansi! :wave:
        Untuk menggunakan aplikasi ini, silahkan download format nominatif pada link yang telah disediakan ==> [Format Kuitansi](https://drive.google.com/drive/folders/1dAb_hMT04oq6pWGPMMoyiE3QxFffYqQ8?usp=sharing)
        """
    )
    st.error(
        """
        ✏️ **NOTE:** Silahkan pilih PPK terlebih dahulu!
        """
    )
    ppk = st.selectbox(
    '',
     ('PPK 01 (Akik Takjudin)', 'PPK 02 (Syihabudin)')
    )
    if ppk == 'PPK 02 (Syihabudin)':
        form = st.form("Upload file")
        with form:
            excelfile = st.file_uploader("Unggah file nominatif")
            submit = st.form_submit_button("Proses")
        if submit:
            try:
                df2 = pd.read_excel(excelfile, sheet_name="kerja")
                amplop = f"pages/templates/amplop.docx"
                spd1 = f"pages/templates/spdkeg.docx"

                for r_idx, r_val in df2.iterrows():
                    if (r_val['asal'] == 'Jakarta'):
                        doctemp = f"pages/templates/jakartakuitansikegiatan.docx"
                    elif (r_val['asal'] == 'Bogor'):
                        doctemp = f"pages/templates/jakartakuitansikegiatan.docx"
                    elif (r_val['asal'] == 'Banten'):
                        doctemp = f"pages/templates/jakartakuitansikegiatan.docx"
                    elif (r_val['asal'] == 'Bandung'):
                        doctemp = f"pages/templates/jakartakuitansikegiatan.docx"
                    elif (r_val['asal'] == 'Serang'):
                        doctemp = f"pages/templates/jakartakuitansikegiatan.docx"
                    else:
                        doctemp = f"pages/templates/daerahkuitansikegiatan.docx"
                    amplop2 = DocxTemplate(amplop)
                    doc = DocxTemplate(doctemp)
                    spd2 = DocxTemplate(spd1)
                    context = {
                        'total' : rupiah_strip(r_val['tiket'] + r_val['taksi_jakarta'] + r_val['taksi_daerah'] + r_val['total_hari']),
                        'output' : r_val['output'],
                        'sub_output' : r_val['sub_output'],
                        'akun' : r_val['akun'],
                        'kegiatan' : r_val['kegiatan'],
                        'layanan' : r_val['layanan'],
                        'tempat' : r_val['tempat'],
                        'tanggal' : bulan_indo(r_val['tanggal']),
                        'nomor_st' : r_val['nomor_st'],
                        'tanggal_st' : kalender_indo(r_val['tanggal_st']),
                        'asal_tujuan' : r_val['asal'] + "-" + r_val['lokasi'],
                        'terbilang' : num2words(int(r_val['tiket'] + r_val['taksi_jakarta'] + r_val['taksi_daerah'] + r_val['total_hari']), lang='id').title() + " Rupiah",
                        'nama' : r_val['nama'],
                        'nip' : r_val['nip'],
                        'tiket' : rupiah_strip(r_val['tiket']),
                        'taksi_jakarta' : rupiah_strip(r_val['taksi_jakarta']),
                        'taksi_daerah' : rupiah_strip(r_val['taksi_daerah']),
                        'hari' : str(r_val['hari']),
                        'hari_rp' : rupiah_strip(float(r_val['total_hari'] / r_val['hari'])),
                        'total_hari' : rupiah_strip(r_val['total_hari']),
                        'lokasi' : r_val["lokasi"],
                        'bulan' : kalender_indo(r_val['tanggal_akhir']),
                        'jabatan' : r_val['jabatan'],
                        'dpr_jakarta' : rupiah_strip(r_val['dpr_jakarta']),
                        'dpr_daerah' : rupiah_strip(r_val['dpr_daerah']),
                        'total_spd' : rupiah_strip(r_val['dpr_jakarta'] + r_val['dpr_daerah']),
                        'mak' : r_val['mak'],                        
                        }
                    doc.render(context)
                    output_path = f"pages/OUTPUT/{context['nama']}.docx"
                    doc.save(output_path)
                    amplop2.render(context)
                    amplop_path = f"pages/AMPLOP/{context['nama']}.docx"
                    amplop2.save(amplop_path)
                    spd2.render(context)
                    spd2_path = f"pages/spd/spdkeg.docx"
                    spd2.save(spd2_path)
                a = st.success("🎉 File kuitansi telah selesai dibuat, silahkan unduh")
                with a:
                    np_array = df2["nama"].to_numpy()
                    files2 = list("pages/OUTPUT/" + (np_array) + ".docx")
                    composed = f"pages/gabung.docx"
                    result = Document(files2[0])
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files2)):
                        doc2 = Document(files2[i])
                        if i != len(files2) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                    composer.save(composed)

                    files3 = list("pages/AMPLOP/" + (np_array) + ".docx")
                    composed = f"pages/amplopgabung.docx"
                    result = Document(files3[0])
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files3)):
                        doc2 = Document(files3[i])
                        if i != len(files3) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                        composer.save(composed)

                    files4 = "pages/spd/spdkeg.docx"
                    composed = f"pages/spdkeg.docx"
                    result = Document(files4)
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files4)):
                        doc2 = Document(files4)
                        if i != len(files4) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                    composer.save(composed)


            except:
                st.error("Silahkan upload nominatif terlebih dahulu/gunakan format nominatif yang telah digunakan ") 
            
    if ppk == 'PPK 01 (Akik Takjudin)':
        form = st.form("Upload file")
        with form:
            excelfile = st.file_uploader("Unggah file nominatif")
            submit = st.form_submit_button("Proses")
        if submit:
            try:
                df2 = pd.read_excel(excelfile, sheet_name="kerja")
                amplop = f"pages/templates/amplop.docx"
                spd = f"pages/templates/spdkegppk.docx"
                for r_idx, r_val in df2.iterrows():
                    if (r_val['asal'] == 'Jakarta'):
                        doctemp = f"pages/templates/jakartakuitansikegiatanppk.docx"
                    elif (r_val['asal'] == 'Bogor'):
                        doctemp = f"pages/templates/jakartakuitansikegiatanppk.docx"
                    elif (r_val['asal'] == 'Banten'):
                        doctemp = f"pages/templates/jakartakuitansikegiatanppk.docx"
                    elif (r_val['asal'] == 'Bandung'):
                        doctemp = f"pages/templates/jakartakuitansikegiatanppk.docx"
                    elif (r_val['asal'] == 'Serang'):
                        doctemp = f"pages/templates/jakartakuitansikegiatanppk.docx"
                    else:
                        doctemp = f"pages/templates/daerahkuitansikegiatanppk.docx"

                    amplop2 = DocxTemplate(amplop)
                    spd2 = DocxTemplate(spd)
                    doc = DocxTemplate(doctemp)
                    context = {
                        'total' : rupiah_strip(r_val['tiket'] + r_val['taksi_jakarta'] + r_val['taksi_daerah'] + r_val['total_hari']),
                        'output' : r_val['output'],
                        'sub_output' : r_val['sub_output'],
                        'akun' : r_val['akun'],
                        'kegiatan' : r_val['kegiatan'],
                        'layanan' : r_val['layanan'],
                        'tempat' : r_val['tempat'],
                        'tanggal' : bulan_indo(r_val['tanggal']),
                        'nomor_st' : r_val['nomor_st'],
                        'tanggal_st' : kalender_indo(r_val['tanggal_st']),
                        'asal_tujuan' : r_val['asal'] + "-" + r_val['lokasi'],
                        'terbilang' : num2words(int(r_val['tiket'] + r_val['taksi_jakarta'] + r_val['taksi_daerah'] + r_val['total_hari']), lang='id').title() + " Rupiah",
                        'nama' : r_val['nama'],
                        'nip' : r_val['nip'],
                        'tiket' : rupiah_strip(r_val['tiket']),
                        'taksi_jakarta' : rupiah_strip(r_val['taksi_jakarta']),
                        'taksi_daerah' : rupiah_strip(r_val['taksi_daerah']),
                        'hari' : str(r_val['hari']),
                        'hari_rp' : rupiah_strip(float(r_val['total_hari'] / r_val['hari'])),
                        'total_hari' : rupiah_strip(r_val['total_hari']),
                        'lokasi' : r_val["lokasi"],
                        'bulan' : kalender_indo(r_val['tanggal_akhir']),
                        'jabatan' : r_val['jabatan'],
                        'dpr_jakarta' : rupiah_strip(r_val['dpr_jakarta']),
                        'dpr_daerah' : rupiah_strip(r_val['dpr_daerah']),
                        'total_spd' : rupiah_strip(r_val['dpr_jakarta'] + r_val['dpr_daerah']),
                        'mak' : r_val['mak'],                        
                        }

                    doc.render(context)
                    output_path = f"pages/OUTPUT/{context['nama']}.docx"
                    doc.save(output_path)
                    amplop2.render(context)
                    amplop_path = f"pages/AMPLOP/{context['nama']}.docx"
                    amplop2.save(amplop_path)
                    spd2.render(context)
                    spd_path = f"pages/spd/spdkeg.docx"
                    spd2.save(spd_path)

                    np_array = df2["nama"].to_numpy()

                    files2 = list("pages/OUTPUT/" + (np_array) + ".docx")
                    composed = f"pages/gabung.docx"
                    result = Document(files2[0])
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files2)):
                        doc2 = Document(files2[i])
                        if i != len(files2) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                    composer.save(composed)

                    files3 = list("pages/AMPLOP/" + (np_array) + ".docx")
                    composed = f"pages/amplopgabung.docx"
                    result = Document(files3[0])
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files3)):
                        doc2 = Document(files3[i])
                        if i != len(files3) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                    composer.save(composed)

                    files4 = "pages/spd/spdkeg.docx"
                    composed = f"pages/spdkeg.docx"
                    result = Document(files4)
                    result.add_page_break()
                    composer = Composer(result)
                    for i in range(1, len(files4)):
                        doc2 = Document(files4)
                        if i != len(files4) -1:
                            doc2.add_page_break()
                        composer.append(doc2)
                    composer.save(composed)
                
            
            except:
                st.error("Silahkan upload nominatif terlebih dahulu/gunakan format nominatif yang telah digunakan")
    with st.expander("Lihat Hasil"):
        kol1, kol2, kol3 = st.columns(3)
        with kol1:
            with open("pages/amplopgabung.docx", "rb") as file:
                btn = st.download_button(
                    label="⬇️ Download Amplop",
                    data=file,
                    file_name="amplop.docx",
                    mime="application/octet-stream"           
                )

            with open("pages/gabung.docx", "rb") as file:
                btn = st.download_button(
                label="⬇️ Download Kuitansi",
                    data=file,
                    file_name="kuitansi.docx",
                    mime="application/octet-stream"
                )
            with open("pages/spd/spdkeg.docx", "rb") as file:
                btn = st.download_button(
                label="⬇️ Download SPD",
                    data=file,
                    file_name="spd.docx",
                    mime="application/octet-stream"
                )

                    

       
                    