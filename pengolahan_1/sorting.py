import pandas as pd
import numpy as np
import itertools
import threading
import time
import sys
import openpyxl

file = "data pelanggan dari kuisioner.xlsx"
data = pd.read_excel(file, header=None, sheet_name=1) #ambil sheet pertama
# data_barang = pd.read_excel(file, header=None, sheet_name='data_barang') #ambil sheet kedua


done = False
prosses = 0
#here is the animation


def animate():
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + c + '  {}'.format(prosses))
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('\rDone!     \n')


t = threading.Thread(target=animate)
t.start()

#olahan
nomor = 0
last_data = data.tail(1)
print(last_data)
lastTrx = last_data[0].values[0] #ambil id terakhir

#menyiapkan tempat untuk menyimpan smua data
# banyakBarang = data_barang.head(1).size
# dataBarang = data_barang.loc[0,:].str.upper()
# emptyData = np.zeros((lastTrx-1, banyakBarang), dtype=int)
# indexData = np.arange(1, lastTrx)
# print(indexData)
# data_sorted = pd.DataFrame(emptyData, columns=dataBarang, index=indexData)
data_sorted = pd.DataFrame()

#melakukan pengecekan pada setiap baris excel
for index, baris in data.iterrows():
    # print(index)
    nomor_transaksi = baris[0]
    nama_barang = baris[5]
    if nomor_transaksi >= 885: 
        # MULAI DARI DATA KE 885, DATANYA BERPINDAH... -_- JADI HARUS BUAT KODE LAGI
        nama_barang = baris[3]
        if nama_barang != "No. Faktur" and nama_barang != "" and nama_barang != ":":
            if nama_barang in data_sorted.columns.values:
                data_sorted.at[nomor_transaksi, nama_barang] = 1
            else:
                data_sorted[nama_barang] = np.zeros(lastTrx+1)
                data_sorted.at[nomor_transaksi, nama_barang] = 1
    else:
        if pd.isnull(baris.loc[12]) == False and pd.isnull(baris.loc[5]) == False:
            # print(nama_barang)
            if nama_barang in data_sorted.columns.values:
                data_sorted.at[nomor_transaksi, nama_barang] = 1
            else:
                data_sorted[nama_barang] = np.zeros(lastTrx+1)
                data_sorted.at[nomor_transaksi, nama_barang] = 1
        
    prosses += 1
time.sleep(1)
done = True
data_sorted.to_excel('data_sorted.xlsx')
