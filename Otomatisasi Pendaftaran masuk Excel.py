import PySimpleGUI as sg
import pandas as pd
import os

sg.theme('LightBlue')

EXCEL_FILE = 'Pendaftaran.xlsx'

# Baca file Excel jika ada, atau buat DataFrame baru jika tidak ada
if os.path.exists(EXCEL_FILE):
    try:
        df = pd.read_excel(EXCEL_FILE)
    except PermissionError:
        sg.popup_error(f"File '{EXCEL_FILE}' sedang digunakan oleh program lain. Tutup file tersebut dan coba lagi.")
        exit()
else:
    df = pd.DataFrame(columns=['Nama', 'No Hp', 'Alamat', 'Tanggal Lahir', 'Jenis Kelamin', 'Hobi'])

layout = [
    [sg.Text('Masukan Data Kamu: ')],
    [sg.Text('Nama', size=(15, 1)), sg.InputText(key='Nama')],
    [sg.Text('No Hp', size=(15, 1)), sg.InputText(key='No Hp')],
    [sg.Text('Alamat', size=(15, 1)), sg.Multiline(key='Alamat')],
    [sg.Text('Tanggal Lahir', size=(15, 1)), sg.InputText(key='Tanggal Lahir'),
     sg.CalendarButton('Kalender', target='Tanggal Lahir', format=('%d-%m-%Y'))],
    [sg.Text('Jenis Kelamin', size=(15, 1)), sg.Combo(['Pria', 'Wanita'], key='Jekel')],
    [sg.Text('Hobi', size=(15, 1)), sg.Checkbox('Belajar', key='Belajar'),
     sg.Checkbox('Membaca', key='Membaca'),
     sg.Checkbox('Musik', key='Musik'),
     sg.Checkbox('Game', key='Game'),
     sg.Checkbox('Olahraga', key='Olahraga')],
    [sg.Submit(), sg.Button('clear'), sg.Exit()]
]

window = sg.Window('Form Pendaftaran', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'clear':
        clear_input()
    if event == 'Submit':
        # Pilih hanya kolom yang sesuai
        data = {
            'Nama': values['Nama'],
            'No Hp': values['No Hp'],
            'Alamat': values['Alamat'],
            'Tanggal Lahir': values['Tanggal Lahir'],
            'Jenis Kelamin': values['Jekel'],
            'Hobi': ', '.join([hobi for hobi in ['Belajar', 'Membaca', 'Musik', 'Game', 'Olahraga'] if values[hobi]])
        }
        # Tambahkan data ke DataFrame menggunakan pd.concat
        new_row = pd.DataFrame([data])
        df = pd.concat([df, new_row], ignore_index=True)
        
        # Simpan ke Excel, tangani error jika file terkunci
        try:
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data Berhasil Disimpan!')
        except PermissionError:
            sg.popup_error(f"Gagal menyimpan file '{EXCEL_FILE}'. Pastikan file tidak sedang dibuka oleh aplikasi lain.")
        
        clear_input()

window.close()
