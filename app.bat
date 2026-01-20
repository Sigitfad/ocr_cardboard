@echo off
:: Pindah ke direktori project
cd /d "%USERPROFILE%\sigitf\Documents\Project_Inspeksi\kartonsv5"

:: Mengecek dan menginstall library dari requirements.txt
echo Memeriksa dan menginstal dependensi...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

:: Menjalankan aplikasi
echo Menjalankan aplikasi...
python main.py

pause