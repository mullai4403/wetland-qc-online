import io, shutil, tempfile, threading, zipfile
from pathlib import Path
import pandas as pd
import streamlit as st
import config as cfg
import qc_engine as qc

st.set_page_config(page_title="Wetland QC Online", page_icon="🌊", layout="wide")
st.title("🌊 Wetland QC Online")
st.caption("Maharashtra Wetland Map QC — 2019 / 2021 / 2024")
st.info("Upload the reference Excel/CSV and one ZIP containing map images. Wetland ID is used as the key.")

def safe_extract_zip(data, dest):
    dest = dest.resolve(); allowed={x.lower() for x in cfg.IMAGE_EXTS}; n=0
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for info in z.infolist():
            if info.is_dir() or Path(info.filename).suffix.lower() not in allowed: continue
            name=Path(info.filename).name
            if not name: continue
            target=(dest/name).resolve()
            if dest not in target.parents: raise ValueError("Unsafe ZIP path")
            k=2
            while target.exists():
                target=dest/f"{Path(name).stem}_{k}{Path(name).suffix}"; k+=1
            with z.open(info) as src, open(target,'wb') as dst: shutil.copyfileobj(src,dst)
            n+=1
    return n

def to_df(results):
    rows=[]
    for r in results:
        rows.append({
            "Wetland ID":r.wid,"Wetland ID Check":r.checks["id"],
            "Excel Wetland Name":r.excel["name"],"Map Wetland Name":r.mapv["name"],"Wetland Name Check":r.checks["name"],
            "Excel Taluk / Tahasil":r.excel["taluk"],"Map Taluk / Tahasil":r.mapv["taluk"],"Taluk Check":r.checks["taluk"],
            "Excel 2021 Area":r.excel["a2021"],"Map 2021 Area":r.mapv["a2021"],"2021 Area Check":r.checks["a2021"],
            "Excel 2024 Area":r.excel["a2024"],"Map 2024 Area":r.mapv["a2024"],"2024 Area Check":r.checks["a2024"],
            "Excel 2019 Area":r.excel["a2019"],"Map 2019 Area":r.mapv["a2019"],"2019 Area Check":r.checks["a2019"],
            "Overall Result":r.overall,"Mismatch Fields":r.mismatch_fields,"Map Filename":r.filename,"Notes":r.notes})
    return pd.DataFrame(rows)

def zip_folder(folder):
    if not folder.exists(): return None
    b=io.BytesIO()
    with zipfile.ZipFile(b,'w',zipfile.ZIP_DEFLATED) as z:
        for p in folder.rglob('*'):
            if p.is_file(): z.write(p,p.relative_to(folder.parent))
    return b.getvalue()

c1,c2=st.columns(2)
with c1: ref_up=st.file_uploader("1. Upload reference Excel / CSV",type=["xlsx","xls","xlsm","csv"])
with c2: maps_up=st.file_uploader("2. Upload maps ZIP",type=["zip"])

if st.button("▶ RUN QC",type="primary",use_container_width=True,disabled=not(ref_up and maps_up)):
    work=Path(tempfile.mkdtemp(prefix='wetland_qc_')); maps=work/'maps'; out=work/'output'; maps.mkdir(); out.mkdir()
    try:
        ref_path=work/Path(ref_up.name).name; ref_path.write_bytes(ref_up.getvalue())
        count=safe_extract_zip(maps_up.getvalue(),maps)
        if not count: st.error("No JPG/JPEG/PNG maps found in ZIP."); st.stop()
        st.success(f"{count} map image(s) detected.")
        ref=qc.load_reference(str(ref_path)); st.caption(ref.summary())
        files=qc.list_maps(str(maps)); total=len(files)
        bar=st.progress(0,text=f"Starting {total} maps...")
        m1,m2,m3,m4=st.columns(4); b1=m1.empty(); b2=m2.empty(); b3=m3.empty(); b4=m4.empty()
        counts={"d":0,"p":0,"m":0,"c":0}; lock=threading.Lock()
        def on_result(r):
            with lock:
                counts["d"]+=1
                if r.overall==qc.PASS: counts["p"]+=1
                elif r.overall==qc.MISMATCH: counts["m"]+=1
                else: counts["c"]+=1
                d=counts["d"]; bar.progress(d/max(total,1),text=f"Processing {d}/{total} — {r.wid}")
                b1.metric("Checked",d); b2.metric("PASS",counts["p"]); b3.metric("MISMATCH",counts["m"]); b4.metric("MANUAL CHECK",counts["c"])
        results=qc.run_batch(ref,files,str(out),threading.Event(),on_start=None,on_result=on_result)
        excel=Path(qc.write_output(results,str(out))); df=to_df(results)
        st.subheader("Results")
        filt=st.selectbox("Filter",["All",qc.PASS,qc.MISMATCH,qc.MANUAL_CHECK])
        st.dataframe(df if filt=="All" else df[df["Overall Result"]==filt],use_container_width=True,hide_index=True)
        d1,d2=st.columns(2)
        with d1: st.download_button("⬇ Download QC Excel",excel.read_bytes(),file_name=cfg.OUTPUT_EXCEL_NAME,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        review=zip_folder(out/cfg.REVIEW_DIR_NAME)
        with d2:
            if review: st.download_button("⬇ Download QC Review ZIP",review,file_name="QC_Review.zip",mime="application/zip",use_container_width=True)
        dbg=out/cfg.DEBUG_LOG_NAME
        if dbg.exists(): st.download_button("Download OCR debug log",dbg.read_bytes(),file_name=cfg.DEBUG_LOG_NAME,mime="text/plain")
    finally:
        shutil.rmtree(work,ignore_errors=True)

st.divider()
st.caption("Privacy: files are processed on the server where you deploy this app. Use an organization-approved/private host for office data.")
