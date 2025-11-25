import tkinter as tk
from tkinter import ttk, Toplevel, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import cartopy.crs as ccrs
import cartopy.feature as cfeature
import numpy as np
from matplotlib.colors import to_hex

# BACKEND IMPORT
from backend import (
    hitung_zona_utm, hitung_hemisphere, validasi,
    konversi_utm, cetak_excel,
    create_color_circle_from_hex
)

# ==========================================
# TAMPILKAN HASIL
# ==========================================
def tampilkan_hasil(data):
    win = Toplevel(root)
    win.title("Hasil Konversi UTM")
    win.geometry("1400x850")

    frame_top = tk.Frame(win)
    frame_top.pack(fill="x", pady=5)

    frame_top.grid_columnconfigure(0, weight=1)
    frame_top.grid_columnconfigure(1, weight=8)
    frame_top.grid_columnconfigure(2, weight=1)

    tk.Button(
        frame_top,
        text="Cetak ke Excel",
        font=("Arial", 12, "bold"),
        bg="#0066CC",
        fg="white",
        command=lambda: cetak_excel(data, messagebox)
    ).grid(row=0, column=0, sticky="w", padx=10)

    tk.Label(
        frame_top,
        text="Hasil Konversi Koordinat ke Sistem UTM",
        font=("Arial", 18, "bold")
    ).grid(row=0, column=1, sticky="n", pady=5)

    frame_center = tk.Frame(win)
    frame_center.pack(fill="both", expand=True)

    fig = plt.Figure(figsize=(12, 5))
    ax = fig.add_subplot(111, projection=ccrs.PlateCarree())
    ax.stock_img()
    ax.add_feature(cfeature.COASTLINE, lw=0.5)
    ax.add_feature(cfeature.BORDERS, lw=0.5)
    ax.gridlines(draw_labels=True, linewidth=0.3)

    for L in range(-180, 181, 6):
        ax.plot([L, L], [-80, 84], color="gray", lw=0.25, transform=ccrs.PlateCarree())

    cmap = plt.cm.jet(np.linspace(0, 1, len(data)))
    win.image_cache = []

    for d, c in zip(data, cmap):
        hex_color = to_hex(c)
        d["color"] = hex_color

        ax.plot(d["lon"], d["lat"], marker="o", color=hex_color,
                markersize=8, transform=ccrs.PlateCarree())

        icon = create_color_circle_from_hex(hex_color)
        win.image_cache.append(icon)
        d["icon"] = icon

    canvas = FigureCanvasTkAgg(fig, master=frame_center)
    canvas.draw()
    canvas.get_tk_widget().pack()

    columns = ("Lat", "Lon", "Easting", "Northing", "Zona", "Belahan")
    tree = ttk.Treeview(frame_center, columns=columns, show="tree headings", height=8)
    tree.pack(fill="x", pady=10)

    tree.heading("#0", text="Warna")
    tree.column("#0", width=80)

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    for d in data:
        tree.insert("", "end", image=d["icon"],
                    values=(d["lat"], d["lon"], d["E"], d["N"], d["zona"], d["hemi"]))

# ==========================================
# TAMBAH BARIS
# ==========================================
def tambah_baris():
    row = len(entries) + 1
    e_lat = tk.Entry(frame_input, width=10)
    e_lon = tk.Entry(frame_input, width=10)
    e_lat.grid(row=row, column=0, padx=3, pady=3)
    e_lon.grid(row=row, column=1, padx=3, pady=3)
    entries.append((e_lat, e_lon))

# ==========================================
# KONVERSI SEMUA INPUT
# ==========================================
def konversi_semua():
    data = []

    for e_lat, e_lon in entries:
        if e_lat.get() == "" or e_lon.get() == "":
            continue

        try:
            lat = float(e_lat.get())
            lon = float(e_lon.get())

            validasi(lat, lon)

            zona = hitung_zona_utm(lon)
            hemi = hitung_hemisphere(lat)
            E, N, g, k = konversi_utm(lat, lon, zona)

            data.append({"lat": lat, "lon": lon, "E": E, "N": N,
                         "zona": zona, "hemi": hemi})

        except ValueError as err:
            messagebox.showerror("Error", str(err))
            return

    tampilkan_hasil(data)

# ==========================================
# GUI UTAMA
# ==========================================
root = tk.Tk()
root.title("Konversi UTM")
root.geometry("360x450")

frame_input = tk.Frame(root)
frame_input.pack(pady=10)

tk.Label(frame_input, text="Latitude").grid(row=0, column=0)
tk.Label(frame_input, text="Longitude").grid(row=0, column=1)

entries = []
for i in range(3):
    e_lat = tk.Entry(frame_input, width=10)
    e_lon = tk.Entry(frame_input, width=10)
    e_lat.grid(row=i+1, column=0, padx=3, pady=3)
    e_lon.grid(row=i+1, column=1, padx=3, pady=3)
    entries.append((e_lat, e_lon))

tk.Button(root, text="Tambah Baris", command=tambah_baris).pack(pady=5)
tk.Button(root, text="Konversi Semua", command=konversi_semua).pack(pady=5)

root.mainloop()
