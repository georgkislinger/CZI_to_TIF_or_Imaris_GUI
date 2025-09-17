#!/usr/bin/env python3
"""
Convert a multitile CZI → single-tile OME-TIFF or Imaris .ims with GUI file selection,
using the official PyImarisWriter example API (col.set_base_color) for coloring.
"""
import sys
import numpy as np
import tifffile
from aicspylibczi import CziFile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime

# ---- Import PyImarisWriter & locate Color class ----
try:
    import PyImarisWriter as pim
    PW = pim.PyImarisWriter
    ColorInfo = getattr(PW, "ColorInfo", None) or getattr(PW, "Color", None)
    if ColorInfo is None:
        raise AttributeError("PyImarisWriter.PyImarisWriter has no ColorInfo/Color")
    IMARIS_AVAILABLE = True
    print(f"[INFO] PyImarisWriter loaded; using Color class: {ColorInfo.__name__}")
except Exception as e:
    PW = None
    ColorInfo = None
    IMARIS_AVAILABLE = False
    print(f"[WARN] Imaris export disabled: {e}")

class ProgressCallback(PW.CallbackClass if IMARIS_AVAILABLE else object):
    """Progress callback for PyImarisWriter."""
    def __init__(self):
        try:
            super().__init__()
        except:
            pass
        self._last = 0
    def RecordProgress(self, progress, block_id):
        pct = int(progress * 100)
        if pct - self._last >= 5:
            self._last = pct
            print(f"[PROGRESS] {pct}%  bytes written: {block_id}")

def convert_to_ometiff(input_czi, output_path, scale=1.0):
    print("[STEP] Converting to OME-TIFF…")
    czi = CziFile(input_czi)
    sizes = dict(zip(czi.dims, czi.size))
    T, C, Z = sizes.get('T',1), sizes.get('C',1), sizes.get('Z',1)

    # infer plane size from one mosaic
    sample = czi.read_mosaic(T=0, C=0, Z=0, scale_factor=scale)
    Y, X = sample.shape[-2], sample.shape[-1]

    arr = np.zeros((T, C, Z, Y, X), dtype=sample.dtype)
    for t in range(T):
        for c in range(C):
            for z in range(Z):
                print(f"  reading T={t},C={c},Z={z}")
                mos = czi.read_mosaic(T=t, C=c, Z=z, scale_factor=scale)
                arr[t, c, z] = mos.reshape(-1, Y, X)[0]

    print(f"[INFO] Writing OME-TIFF to: {output_path}")
    tifffile.imwrite(output_path, arr, metadata={'axes':'TCZYX'}, ome=True)
    print("[DONE] OME-TIFF written.")

def convert_to_ims(input_czi, output_path, scale=1.0):
    if not IMARIS_AVAILABLE:
        raise RuntimeError("PyImarisWriter unavailable; cannot write .ims")
    print("[STEP] Converting to Imaris .ims…")
    czi = CziFile(input_czi)
    sizes = dict(zip(czi.dims, czi.size))
    T, C, Z = sizes.get('T',1), sizes.get('C',1), sizes.get('Z',1)

    # infer plane size
    sample = czi.read_mosaic(T=0, C=0, Z=0, scale_factor=scale)
    Y, X = sample.shape[-2], sample.shape[-1]

    arr = np.zeros((T, C, Z, Y, X), dtype=sample.dtype)
    for t in range(T):
        for c in range(C):
            for z in range(Z):
                print(f"  reading T={t},C={c},Z={z}")
                mos = czi.read_mosaic(T=t, C=c, Z=z, scale_factor=scale)
                arr[t, c, z] = mos.reshape(-1, Y, X)[0]

    # ask voxel sizes
    vx = simpledialog.askfloat("Voxel Size", "X (µm):", initialvalue=1.0)
    vy = simpledialog.askfloat("Voxel Size", "Y (µm):", initialvalue=1.0)
    vz = simpledialog.askfloat("Voxel Size", "Z (µm):", initialvalue=1.0)
    if None in (vx, vy, vz):
        raise RuntimeError("Voxel size entry cancelled.")

    # prepare Imaris writer
    image_size  = PW.ImageSize(x=X, y=Y, z=Z, c=C, t=T)
    dim_seq     = PW.DimensionSequence('x','y','z','c','t')
    block_size  = image_size
    sample_size = PW.ImageSize(x=1, y=1, z=1, c=1, t=1)
    options     = PW.Options()

    print("[INFO] Initializing Imaris ImageConverter…")
    conv = PW.ImageConverter(
        str(arr.dtype),
        image_size,
        sample_size,
        dim_seq,
        block_size,
        output_path,
        options,
        "CZI2IMS",
        "1.0",
        ProgressCallback()
    )

    # copy all data in one block
    print("[STEP] Copying full volume block…")
    flat = arr.ravel(order='C')
    idx  = PW.ImageSize()
    conv.CopyBlock(flat, idx)

    # build Finish parameters & colors per GitHub example
    print("[STEP] Preparing Finish() parameters & colors…")
    params = PW.Parameters()
    base_rgbs = [(1,0,0), (0,1,0), (0,0,1), (1,1,0)]
    for ci in range(C):
        params.set_channel_name(ci, f"Channel {ci}")
    time_infos = [datetime.now()]

    color_infos = []
    for ci in range(C):
        r, g, b = base_rgbs[ci % len(base_rgbs)]
        col = ColorInfo()
        col.set_base_color(PW.Color(r, g, b, 1.0))
        color_infos.append(col)

    image_extents      = PW.ImageExtents(0, 0, 0, vx*X, vy*Y, vz*Z)
    adjust_color_range = True

    print("[INFO] Finalizing .ims write…")
    conv.Finish(image_extents, params, time_infos, color_infos, adjust_color_range)
    conv.Destroy()
    print("[DONE] Imaris .ims written.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    # 1) select CZI
    in_czi = filedialog.askopenfilename(
        title="Select input CZI",
        filetypes=[("CZI files","*.czi"), ("All files","*.*")]
    )
    if not in_czi:
        messagebox.showinfo("Cancelled","No CZI selected. Exiting.")
        sys.exit()

    # 2) show metadata
    czi = CziFile(in_czi)
    messagebox.showinfo(
        "CZI Metadata",
        f"Dimensions: {czi.dims}\n"
        f"Plane counts: {dict(zip(czi.dims, czi.size))}\n"
        f"Physical pixel sizes: {getattr(czi,'physical_size',None)}"
    )

    # 3) choose format
    use_ims = False
    if IMARIS_AVAILABLE:
        use_ims = messagebox.askyesno(
            "Format",
            "Save as Imaris (.ims)?\nYes → .ims   No → OME-TIFF"
        )

    # 4) save dialog
    if use_ims:
        out = filedialog.asksaveasfilename(
            title="Save as .ims",
            defaultextension=".ims",
            filetypes=[("Imaris files","*.ims"), ("All files","*.*")]
        )
    else:
        out = filedialog.asksaveasfilename(
            title="Save as OME-TIFF",
            defaultextension=".ome.tif",
            filetypes=[("OME-TIFF","*.ome.tif"), ("All files","*.*")]
        )
    if not out:
        messagebox.showinfo("Cancelled","No output selected. Exiting.")
        sys.exit()

    # 5) convert
    try:
        if use_ims:
            convert_to_ims(in_czi, out)
        else:
            convert_to_ometiff(in_czi, out)
        messagebox.showinfo("Done","Conversion complete!")
    except Exception as e:
        print(f"[ERROR] {e}")
        messagebox.showerror("Error", str(e))
