import numpy as np
from openpyxl import Workbook
from matplotlib.colors import to_hex
from PIL import Image, ImageDraw, ImageTk

# ==========================================
# KONSTANTA UTM
# ==========================================
A = 6378137.0
F = 1 / 298.257223563
K0 = 0.9996
E0 = 500000
N0_N, N0_S = 0.0, 10000000.0

# ==========================================
# FUNGSI PEMBANTU
# ==========================================
def hitung_zona_utm(lon):
    return int((lon + 180) / 6) + 1

def hitung_hemisphere(lat):
    return "Belahan Bumi Utara" if lat >= 0 else "Belahan Bumi Selatan"

def validasi(lat, lon):
    if not (-80 <= lat <= 84):
        raise ValueError("Latitude harus antara -80 sampai 84")
    if not (-180 <= lon <= 180):
        raise ValueError("Longitude harus antara -180 sampai 180")

def konversi_utm(lat, lon, zona):
    N0 = N0_S if lat < 0 else N0_N
    lam0 = np.radians((zona - 1) * 6 - 177)

    phi = np.radians(lat)
    lam = np.radians(lon)
    e2 = 2 * F - F**2

    A1 = 1 - e2/4 - 3*e2**2/64
    A2 = 3*e2/8 + 3*e2**2/32

    M = A*(A1*phi - A2*np.sin(2*phi))
    nu = A / np.sqrt(1 - e2*np.sin(phi)**2)
    p = lam - lam0

    E = E0 + K0 * nu * (p*np.cos(phi) + (p**3*np.cos(phi)**3)/6)
    N = K0 * (M + N0 + nu*np.sin(phi)*np.cos(phi)*p**2/2)

    return round(E, 3), round(N, 3), 0.0, 0.0016

# ==========================================
# ICON WARNA
# ==========================================
def hex_to_rgb_tuple(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def create_color_circle_from_hex(hex_color, size=18):
    from PIL import ImageTk  # penting untuk frontend
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((1, 1, size-2, size-2), fill=hex_to_rgb_tuple(hex_color))
    return ImageTk.PhotoImage(img)

# ==========================================
# CETAK EXCEL
# ==========================================
def cetak_excel(data, messagebox):
    wb = Workbook()
    ws = wb.active
    ws.title = "UTM Converted Data"

    ws.append(["Lat", "Lon", "Easting", "Northing", "Zona", "Belahan", "HEX"])

    for d in data:
        ws.append([d["lat"], d["lon"], d["E"], d["N"], d["zona"], d["hemi"], d["color"]])

    wb.save("hasil_konversi_UTM.xlsx")
    messagebox.showinfo("Excel Tersimpan", "File telah dibuat: hasil_konversi_UTM.xlsx")
