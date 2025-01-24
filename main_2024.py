import multiprocessing
import pandas as pd
import xlsxwriter
import openpyxl
import sys

# Membaca data dari file Excel
df = pd.read_excel('DATA PENYULANG_SAMPLE.xlsx')
df2 = pd.read_excel('No. Tiang PDPJ (1).xlsx')

# Menghitung jumlah baris
jumlah_baris = df.shape[0]

# Mendefinisikan fungsi untuk menghitung total panjang kabel
def hitung_total_panjang(dari, sampai, taruh):
  try:
    # Mencari indeks baris untuk nomor tiang "AWAL" dan "AKHIR"
    index1 = df2[df2['NO.TIANG'] == f'{dari}'].index[0]
    index2 = df2[df2['NO.TIANG'] == f'{sampai}'].index[0]

    # Menghitung total panjang kabel
    if index1 < index2:
      total = df2['PANJANG (M)'][index1+1:index2+1].sum()
    else:
      total = df2['PANJANG (M)'][index2+1:index1+1].sum()

    # Membulatkan hasil total panjang kabel
    hasil = round(total, 2)

    # Memasukkan nilai ke dalam kolom data penyulang
    df.loc[taruh, "NILAI"] = hasil

  except IndexError:
    # Menangani IndexError
    hasil = " "
    df.loc[taruh, "NILAI"] = hasil

# Membagi data menjadi dua bagian
data_bagian_1 = df[:jumlah_baris // 2]
data_bagian_2 = df[jumlah_baris // 2:]

# Membuat multiprocessing Pool
with multiprocessing.Pool() as pool:
  # Menjalankan fungsi `hitung_total_panjang` secara paralel untuk dua bagian data
  pool.map(hitung_total_panjang, [(dari, sampai, taruh) for dari, sampai, taruh in zip(data_bagian_1['AWAL'], data_bagian_1['AKHIR'], data_bagian_1['NILAI'])])
  pool.map(hitung_total_panjang, [(dari, sampai, taruh) for dari, sampai, taruh in zip(data_bagian_2['AWAL'], data_bagian_2['AKHIR'], data_bagian_2['NILAI'])])

# Menyimpan hasil ke file Excel
df.to_excel("DATA PENYULANG_SAMPLE.xlsx", index=False)

# Menampilkan hasil ke layar
print(f'Total panjang kabel: {df["NILAI"].sum()}')
