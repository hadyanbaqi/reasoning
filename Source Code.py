import xlrd
import xlwt

def aturanInterferensi(x, y):
    #x pengeluaran, y penghasilan
    minimum = sorted([x, y], key=lambda x: x[0])[0][0]
    x = x[1]
    y = y[1]
    result = None
    if x == 'Kecil' and y == 'Kecil':
        result = [minimum, 'Tinggi']
    elif x == 'Kecil' and y == 'Sedang':
        result = [minimum, 'Rendah']
    elif x == 'Kecil' and y == 'Besar':
        result = [minimum, 'Rendah']
    elif x == 'Kecil' and y == 'Sangat Besar':
        result = [minimum, 'Rendah']

    elif x == 'Sedang' and y == 'Kecil':
        result = [minimum, 'Tinggi']
    elif x == 'Sedang' and y == 'Sedang':
        result = [minimum, 'Tinggi']
    elif x == 'Sedang' and y == 'Besar':
        result = [minimum, 'Rendah']
    elif x == 'Sedang' and y == 'Sangat Besar':
        result = [minimum, 'Rendah']
    
    elif x == 'Besar' and y == 'Kecil':
        result = [minimum, 'Tinggi']
    elif x == 'Besar' and y == 'Sedang':
        result = [minimum, 'Tinggi']
    elif x == 'Besar' and y == 'Besar':
        result = [minimum, 'Rendah']
    elif x == 'Besar' and y == 'Sangat Besar':
        result = [minimum, 'Rendah']
        
    return result

def hitungKeanggotaanPengeluaran(x):
    temp = []
    if 0 <= x <= 2:
        temp.append([1, 'Kecil'])
    elif 2 < x <= 5:
        temp.append([(-1 * (x-5)) / 3, 'Kecil'])
    if 4 < x <= 7:
        temp.append([(x - 4) / 3, 'Sedang'])
    elif 7 < x <= 10:
        temp.append([(-1 * (x-10)) / 3, 'Sedang'])
    if 5 < x <= 10:
        temp.append([(x-5) / 5, 'Besar'])
    elif 10 <= x:
        temp.append([1, 'Besar'])
    return temp

def hitungKeanggotaanPenghasilan(x):
    temp = []
    if 0 <= x <= 2:
        temp.append([1, 'Kecil'])
    elif 2 < x <= 5:
        temp.append([(-1 * (x-2)) /3, 'Kecil'])
    if 4 < x < 6:
        temp.append([(x - 4) / 2, 'Sedang'])
    elif 6 <= x <= 8:
        temp.append([1, 'Sedang'])
    elif 8 < x <= 10:
        temp.append([(-1 * (x-10)) / 2, 'Sedang'])
    if 7 < x < 10:
        temp.append([(x-6) / 3, 'Besar'])
    elif 10 <= x <= 12:
        temp.append([1, 'Besar'])
    elif 12 < x <= 14:
        temp.append([(-1 * (x-14)) / 2, 'Besar'])
    if 10 < x < 14:
        temp.append([(x-12) / 4, 'Sangat Besar'])
    elif 14 <= x:
        temp.append([1, 'Sangat Besar'])
    return temp

def hitungAturanInterferensi(x):
    temp = []
    for penghasilan in x[4]:
        for pengeluaran in x[5]:
            temp.append(aturanInterferensi(pengeluaran, penghasilan))
    rendah = [model for model in temp if model[1] == 'Rendah']
    rendah = sorted(rendah, key=lambda x: x[0], reverse=True)
    tinggi = [model for model in temp if model[1] == 'Tinggi']
    tinggi = sorted(tinggi, key=lambda x: x[0], reverse=True)
    arrHasil = []
    if len(rendah) > 0:
        arrHasil.append(rendah[0])
    if len(tinggi) > 0:
        arrHasil.append(tinggi[0])
    return arrHasil

def hitungCenterOfGrafity(x):
    rendah = [y for y in x[6] if y[1] == 'Rendah']
    tinggi = [y for y in x[6] if y[1] == 'Tinggi']
    if len(rendah) > 0:
        rendah = rendah[0][0]
    else:
        rendah = 0
    if len(tinggi) > 0:
        tinggi = tinggi[0][0]
    else:
        tinggi = 0
    return (((10 + 20 + 30 + 40 + 50 + 60) * rendah) + ((70 + 80 + 90 + 100) * tinggi)) / ((6*rendah)+(4*tinggi))

def toString(x):
        output = 'ID : {} \t| Penghasilan : {} \t| Pengeluaran: {} \t'.format(
            x[0], x[1], x[2])
        print(output)
        print('Kelayakan Untuk Mendapat Beasiswa = {:.2f}'.format(x[3]))
        print('========================================================')

def getDataFromExcell():
    wb = xlrd.open_workbook('./Mahasiswa.xls')
    sheet = wb.sheet_by_index(0)
    temp = []
    for i in range(1, 101):
        id = int(sheet.cell_value(i, 0))
        penghasilan = sheet.cell_value(i, 1)
        pengeluaran = sheet.cell_value(i, 2)
        temp.append([id, penghasilan, pengeluaran, 0])
    return temp

def saveToFile(arrMahasiswa):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet 1')
    sheet.write(0, 0, 'id')
    for i in range(20):
        sheet.write(i+1, 0, arrMahasiswa[i][0])
    workbook.save('Bantuan.xls')

def main():
    arrMahasiswa = getDataFromExcell()
    #[id, penghasilan, pengeluaran, nilaiKelayakan, arrPenghasilan, arrPengeluaran, arrKelayakan]
    for mahasiswa in arrMahasiswa:
        mahasiswa.append(hitungKeanggotaanPenghasilan(mahasiswa[1]))
        mahasiswa.append(hitungKeanggotaanPengeluaran(mahasiswa[2]))
        mahasiswa.append(hitungAturanInterferensi(mahasiswa))
        mahasiswa[3] = hitungCenterOfGrafity(mahasiswa)

    arrMahasiswa = sorted(arrMahasiswa, key=lambda x: x[3], reverse=True)
    for i in range(20):
        toString(arrMahasiswa[i])
    saveToFile(arrMahasiswa)

main()