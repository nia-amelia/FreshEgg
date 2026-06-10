from flask import Flask, render_template, request, redirect
from database import get_connection
from datetime import date, datetime
from openpyxl import Workbook
from flask import send_file
import io
from datetime import datetime
from io import BytesIO

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)

from reportlab.lib.styles import getSampleStyleSheet

from flask import send_file

app = Flask(__name__)

from datetime import date

@app.route('/')
def dashboard():

    conn = get_connection()
    cur = conn.cursor()

    # Total kandang
    cur.execute("SELECT COUNT(*) FROM kandangs")
    total_kandang = cur.fetchone()[0]

    # Total ayam
    cur.execute("""
        SELECT COALESCE(SUM(jumlah_ayam),0)
        FROM kandangs
    """)
    total_ayam = cur.fetchone()[0]

    # Produksi hari ini
    cur.execute("""
        SELECT COALESCE(SUM(jumlah_telur),0)
        FROM produksi_harian
        WHERE tanggal = CURRENT_DATE
    """)
    produksi_hari_ini = cur.fetchone()[0]

    # Harga terbaru
    cur.execute("""
        SELECT harga_per_kg
        FROM harga_telur
        ORDER BY tanggal DESC
        LIMIT 1
    """)

    harga = cur.fetchone()
    harga_terbaru = harga[0] if harga else 0

    cur.execute("""
        SELECT COALESCE(SUM(total_harga),0)
        FROM penjualan
        WHERE tanggal = CURRENT_DATE
    """)

    pendapatan_hari_ini = cur.fetchone()[0]

    cur.execute("""
        SELECT COALESCE(SUM(total_harga),0)
        FROM penjualan
        WHERE DATE_TRUNC('month', tanggal)
            = DATE_TRUNC('month', CURRENT_DATE)
    """)

    pendapatan_bulan_ini = cur.fetchone()[0]

    # Produktivitas
    produktivitas = 0

    if total_ayam > 0:
        produktivitas = round(
            (produksi_hari_ini / total_ayam) * 100,
            1
        )

    # Grafik 7 hari terakhir
    cur.execute("""
        SELECT
            tanggal,
            SUM(jumlah_telur)
        FROM produksi_harian
        GROUP BY tanggal
        ORDER BY tanggal DESC
        LIMIT 7
    """)

    chart_data = cur.fetchall()

    chart_data.reverse()

    labels = [
        row[0].strftime('%d-%m')
        for row in chart_data
    ]

    values = [
        row[1]
        for row in chart_data
    ]

    cur.close()
    conn.close()

    return render_template(
        'dashboard.html',
        total_kandang=total_kandang,
        total_ayam=total_ayam,
        produksi_hari_ini=produksi_hari_ini,
        harga_terbaru=harga_terbaru,
        pendapatan_hari_ini=pendapatan_hari_ini,
        pendapatan_bulan_ini=pendapatan_bulan_ini,
        produktivitas=produktivitas,
        labels=labels,
        values=values
    )

@app.route('/kandang')
def kandang():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM kandangs
        ORDER BY id
    """)

    kandangs = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'kandang/index.html',
        kandangs=kandangs
    )

@app.route('/kandang/tambah', methods=['GET', 'POST'])
def tambah_kandang():
    if request.method == 'POST':
        nama_kandang = request.form['nama_kandang']
        umur_bulan = int(request.form['umur_bulan'])
        jumlah_ayam = int(request.form['jumlah_ayam'])

        conn = get_connection()
        cur = conn.cursor()

        cur.execute(
            "INSERT INTO kandangs (nama_kandang, umur_bulan, jumlah_ayam) VALUES (%s, %s, %s)",
            (nama_kandang, umur_bulan, jumlah_ayam)
        )

        conn.commit()
        cur.close()
        conn.close()

        return redirect('/kandang')

    return render_template('kandang/create.html')

    return render_template('produksi/index.html')

@app.route('/kandang/edit/<int:id>', methods=['GET', 'POST'])
def edit_kandang(id):

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        nama_kandang = request.form['nama_kandang']
        umur_bulan = request.form['umur_bulan']
        jumlah_ayam = request.form['jumlah_ayam']

        cur.execute("""
            UPDATE kandangs
            SET nama_kandang=%s,
                umur_bulan=%s,
                jumlah_ayam=%s
            WHERE id=%s
        """, (
            nama_kandang,
            umur_bulan,
            jumlah_ayam,
            id
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/kandang')

    cur.execute("""
        SELECT *
        FROM kandangs
        WHERE id=%s
    """, (id,))

    kandang = cur.fetchone()

    cur.close()
    conn.close()

    return render_template(
        'kandang/edit.html',
        kandang=kandang
    )

@app.route('/kandang/delete/<int:id>')
def delete_kandang(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM kandangs
        WHERE id=%s
    """, (id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect('/kandang')

@app.route('/produksi')
def produksi():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            p.id,
            p.tanggal,
            k.nama_kandang,
            p.jumlah_telur,
            k.jumlah_ayam
        FROM produksi_harian p
        JOIN kandangs k
            ON p.kandang_id = k.id
        ORDER BY p.tanggal DESC
    """)

    produksis = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'produksi/index.html',
        produksis=produksis
    )

@app.route('/produksi/tambah', methods=['GET', 'POST'])
def tambah_produksi():

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        kandang_id = request.form['kandang_id']
        jumlah_telur = request.form['jumlah_telur']

        cur.execute("""
            INSERT INTO produksi_harian
            (tanggal, kandang_id, jumlah_telur)
            VALUES (%s,%s,%s)
        """, (
            tanggal,
            kandang_id,
            jumlah_telur
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/produksi')

    cur.execute("""
        SELECT id, nama_kandang
        FROM kandangs
        ORDER BY nama_kandang
    """)

    kandangs = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'produksi/create.html',
        kandangs=kandangs
    )

@app.route('/produksi/edit/<int:id>', methods=['GET', 'POST'])
def edit_produksi(id):

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        kandang_id = request.form['kandang_id']
        jumlah_telur = request.form['jumlah_telur']

        cur.execute("""
            UPDATE produksi_harian
            SET
                tanggal=%s,
                kandang_id=%s,
                jumlah_telur=%s
            WHERE id=%s
        """, (
            tanggal,
            kandang_id,
            jumlah_telur,
            id
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/produksi')

    cur.execute("""
        SELECT *
        FROM produksi_harian
        WHERE id=%s
    """, (id,))

    produksi = cur.fetchone()

    cur.execute("""
        SELECT id, nama_kandang
        FROM kandangs
        ORDER BY nama_kandang
    """)

    kandangs = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'produksi/edit.html',
        produksi=produksi,
        kandangs=kandangs
    )

@app.route('/produksi/delete/<int:id>')
def delete_produksi(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM produksi_harian
        WHERE id=%s
    """, (id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect('/produksi')

@app.route('/harga')
def harga():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM harga_telur
        ORDER BY tanggal DESC
    """)

    hargas = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'harga/index.html',
        hargas=hargas
    )

@app.route('/harga/tambah', methods=['GET', 'POST'])
def tambah_harga():

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        harga = request.form['harga']

        cur.execute("""
            INSERT INTO harga_telur
            (tanggal, harga_per_kg)
            VALUES (%s,%s)
        """, (
            tanggal,
            harga
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/harga')

    return render_template('harga/create.html')

@app.route('/harga/edit/<int:id>', methods=['GET', 'POST'])
def edit_harga(id):

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        harga_per_kg = request.form['harga_per_kg']

        cur.execute("""
            UPDATE harga_telur
            SET
                tanggal=%s,
                harga_per_kg=%s
            WHERE id=%s
        """, (
            tanggal,
            harga_per_kg,
            id
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/harga')

    cur.execute("""
        SELECT *
        FROM harga_telur
        WHERE id=%s
    """, (id,))

    harga = cur.fetchone()

    cur.close()
    conn.close()

    return render_template(
        'harga/edit.html',
        harga=harga
    )

@app.route('/harga/delete/<int:id>')
def delete_harga(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM harga_telur
        WHERE id=%s
    """, (id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect('/harga')

@app.route('/penjualan')
def penjualan():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM penjualan
        ORDER BY tanggal DESC
    """)

    penjualans = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'penjualan/index.html',
        penjualans=penjualans
    )

@app.route('/penjualan/tambah', methods=['GET', 'POST'])
def tambah_penjualan():

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        pembeli = request.form['pembeli']
        jumlah_kg = float(request.form['jumlah_kg'])
        keterangan = request.form['keterangan']

        cur.execute("""
            SELECT harga_per_kg
            FROM harga_telur
            WHERE tanggal = %s
        """, (tanggal,))

        harga = cur.fetchone()

        if not harga:
            return "Harga telur pada tanggal tersebut belum ada"

        harga_per_kg = float(harga[0])

        total_harga = jumlah_kg * harga_per_kg

        cur.execute("""
            INSERT INTO penjualan
            (
                tanggal,
                pembeli,
                jumlah_kg,
                harga_per_kg,
                total_harga,
                keterangan
            )
            VALUES (%s,%s,%s,%s,%s,%s)
        """, (
            tanggal,
            pembeli,
            jumlah_kg,
            harga_per_kg,
            total_harga,
            keterangan
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/penjualan')

    return render_template(
        'penjualan/create.html'
    )

@app.route('/penjualan/edit/<int:id>', methods=['GET', 'POST'])
def edit_penjualan(id):

    conn = get_connection()
    cur = conn.cursor()

    if request.method == 'POST':

        tanggal = request.form['tanggal']
        pembeli = request.form['pembeli']
        jumlah_kg = float(request.form['jumlah_kg'])
        keterangan = request.form['keterangan']

        cur.execute("""
            SELECT harga_per_kg
            FROM harga_telur
            WHERE tanggal=%s
        """, (tanggal,))

        harga = cur.fetchone()

        if not harga:
            return "Harga telur pada tanggal tersebut tidak ditemukan"

        harga_per_kg = float(harga[0])
        total_harga = jumlah_kg * harga_per_kg

        cur.execute("""
            UPDATE penjualan
            SET
                tanggal=%s,
                pembeli=%s,
                jumlah_kg=%s,
                harga_per_kg=%s,
                total_harga=%s,
                keterangan=%s
            WHERE id=%s
        """, (
            tanggal,
            pembeli,
            jumlah_kg,
            harga_per_kg,
            total_harga,
            keterangan,
            id
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect('/penjualan')

    cur.execute("""
        SELECT *
        FROM penjualan
        WHERE id=%s
    """, (id,))

    penjualan = cur.fetchone()

    cur.close()
    conn.close()

    return render_template(
        'penjualan/edit.html',
        penjualan=penjualan
    )

@app.route('/penjualan/delete/<int:id>')
def delete_penjualan(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM penjualan
        WHERE id=%s
    """, (id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect('/penjualan')

@app.route('/laporan')
def laporan():

    bulan = request.args.get('bulan')

    conn = get_connection()
    cur = conn.cursor()

    if bulan:

        tahun, bulan_angka = bulan.split('-')

        # Produksi
        cur.execute("""
            SELECT COALESCE(SUM(jumlah_telur),0)
            FROM produksi_harian
            WHERE EXTRACT(YEAR FROM tanggal)=%s
            AND EXTRACT(MONTH FROM tanggal)=%s
        """, (tahun, bulan_angka))

        total_produksi = cur.fetchone()[0]

        # Penjualan
        cur.execute("""
            SELECT COALESCE(SUM(jumlah_kg),0)
            FROM penjualan
            WHERE EXTRACT(YEAR FROM tanggal)=%s
            AND EXTRACT(MONTH FROM tanggal)=%s
        """, (tahun, bulan_angka))

        total_penjualan = cur.fetchone()[0]

        # Pendapatan
        cur.execute("""
            SELECT COALESCE(SUM(total_harga),0)
            FROM penjualan
            WHERE EXTRACT(YEAR FROM tanggal)=%s
            AND EXTRACT(MONTH FROM tanggal)=%s
        """, (tahun, bulan_angka))

        total_pendapatan = cur.fetchone()[0]

    else:

        cur.execute("""
            SELECT COALESCE(SUM(jumlah_telur),0)
            FROM produksi_harian
        """)
        total_produksi = cur.fetchone()[0]

        cur.execute("""
            SELECT COALESCE(SUM(jumlah_kg),0)
            FROM penjualan
        """)
        total_penjualan = cur.fetchone()[0]

        cur.execute("""
            SELECT COALESCE(SUM(total_harga),0)
            FROM penjualan
        """)
        total_pendapatan = cur.fetchone()[0]

    cur.close()
    conn.close()

    return render_template(
        'laporan/index.html',
        total_produksi=total_produksi,
        total_penjualan=total_penjualan,
        total_pendapatan=total_pendapatan,
        bulan=bulan
    )

@app.route('/laporan/export')
def export_excel():

    conn = get_connection()
    cur = conn.cursor()

    wb = Workbook()

    ws = wb.active
    ws.title = "Laporan FreshEgg"

    ws.append([
        "Tanggal",
        "Pembeli",
        "Jumlah Kg",
        "Harga/Kg",
        "Total"
    ])

    cur.execute("""
        SELECT
            tanggal,
            pembeli,
            jumlah_kg,
            harga_per_kg,
            total_harga
        FROM penjualan
        ORDER BY tanggal DESC
    """)

    data = cur.fetchall()

    for row in data:
        ws.append(row)

    cur.close()
    conn.close()

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='Laporan_FreshEgg.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/laporan/pdf')
def export_pdf():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT COALESCE(SUM(jumlah_telur),0)
        FROM produksi_harian
    """)
    total_produksi = cur.fetchone()[0]

    cur.execute("""
        SELECT COALESCE(SUM(jumlah_kg),0)
        FROM penjualan
    """)
    total_penjualan = cur.fetchone()[0]

    cur.execute("""
        SELECT COALESCE(SUM(total_harga),0)
        FROM penjualan
    """)
    total_pendapatan = cur.fetchone()[0]

    cur.close()
    conn.close()

    buffer = BytesIO()

    doc = SimpleDocTemplate(buffer)

    styles = getSampleStyleSheet()

    elements = []

    # Judul

    elements.append(
        Paragraph(
            "🥚 FreshEgg Farm Management",
            styles['Title']
        )
    )

    elements.append(
        Paragraph(
            "Laporan Peternakan",
            styles['Heading2']
        )
    )

    elements.append(Spacer(1, 20))

    # Tabel Ringkasan

    data = [

        ["Keterangan", "Nilai"],

        [
            "Total Produksi",
            f"{total_produksi} Butir"
        ],

        [
            "Total Penjualan",
            f"{total_penjualan} Kg"
        ],

        [
            "Total Pendapatan",
            f"Rp {total_pendapatan:,.0f}"
        ]

    ]

    elements.append(

    Paragraph(

        f"Tanggal Cetak : {datetime.now().strftime('%d-%m-%Y')}",

        styles['Normal']

        )

    )

    table = Table(data, colWidths=[200, 200])

    table.setStyle(

        TableStyle([

            ('BACKGROUND', (0,0), (-1,0), colors.gold),

            ('TEXTCOLOR', (0,0), (-1,0), colors.black),

            ('GRID', (0,0), (-1,-1), 1, colors.black),

            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),

            ('ALIGN', (0,0), (-1,-1), 'CENTER')

        ])

    )

    elements.append(table)

    elements.append(Spacer(1, 30))

    elements.append(

        Paragraph(
            "Terima kasih telah menggunakan FreshEgg.",
            styles['Italic']
        )

    )

    doc.build(elements)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="laporan_freshegg.pdf",
        mimetype="application/pdf"
    )

@app.route('/pendapatan')
def pendapatan():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            tanggal,
            SUM(total_harga)
        FROM penjualan
        GROUP BY tanggal
        ORDER BY tanggal DESC
    """)

    data = cur.fetchall()

    cur.execute("""
        SELECT
            tanggal,
            SUM(total_harga)
        FROM penjualan
        GROUP BY tanggal
        ORDER BY tanggal ASC
        LIMIT 7
    """)

    grafik = cur.fetchall()

    labels = [
        str(row[0])
        for row in grafik
    ]

    values = [
        float(row[1])
        for row in grafik
    ]

    cur.close()
    conn.close()

    return render_template(
        'pendapatan/index.html',
        pendapatans=data,
        labels=labels,
        values=values
    )

if __name__ == '__main__':
    app.run(debug=True)