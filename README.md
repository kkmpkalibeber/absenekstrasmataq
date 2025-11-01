SIS Absensi Ekstrak Qur’an Wonosobo
Proyek Sistem Informasi Absensi Ekstrak Qur’an untuk SMA Takhasus Wonosobo menggunakan Google Apps Script dengan antarmuka HTML.

Struktur proyek

absensi.gs
absensiekstraku.html
appsscript.json
.clasp.json
config.js
.gitignore
README.md
Prasyarat

Akun Google dengan akses Apps Script
Node.js dan npm
Clasp: npm install -g @google/clasp
Git
Instalasi

Clone repo: git clone <url-repo>
Jalankan: cd <project-root>
Login Clasp: clasp login
Sinkronisasi lokal dengan Apps Script:
Tarik perubahan: clasp pull
Jika ada perubahan lokal: clasp push
Menjalankan secara lokal

Karena Apps Script berjalan di cloud, jalankan UI melalui HTML yang dipakai di Apps Script atau ketika di-deploy sebagai web app.
Konfigurasi lingkungan

Buat contoh file lingkungan jika diperlukan (mis. config.js berisi variabel UI)
Jangan commit kredensial sensitif.
Build / Deploy

Deploy sebagai Web App:
Buka Apps Script di Google Apps Script IDE atau gunakan clasp pour melakukan deploy.
Atur siapa yang bisa mengakses (me, anyone with the link, dll.).
Testing

Lakukan pengujian dengan akses web app dan pastikan alur absensi berjalan sesuai desain.
Kontribusi

Git flow, PR, linting, dsb.
Catatan keamanan

Hindari menyimpan rahasia di repo. Gunakan Secrets di CI/CD untuk kredensial jika ada.
