import os
import platform
import tempfile
import shutil
import zipfile
import time
import threading
from flask import Flask, request, send_file, render_template_string
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB max per request

INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>PPT → PDF Converter (Drag & Drop)</title>
<style>
  body{font-family:Arial,Helvetica,sans-serif;background:#f4f6f8;color:#222;display:flex;flex-direction:column;align-items:center;padding:30px}
  h1{margin-bottom:6px}
  .drop{width:100%;max-width:820px;height:260px;border:3px dashed #cbd5e1;border-radius:12px;display:flex;align-items:center;justify-content:center;background:white;cursor:pointer;transition:0.15s}
  .drop.dragover{background:#eef2ff;border-color:#7c3aed}
  .drop p{text-align:center;margin:0;padding:20px;color:#374151}
  .btn{margin-top:16px;padding:10px 16px;border-radius:8px;border:none;background:#7c3aed;color:white;cursor:pointer;font-size:15px}
  .btn:disabled{background:#a78bfa;cursor:not-allowed}
  .files{margin-top:18px;width:100%;max-width:820px}
  .filelist{background:white;padding:10px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
  .row{display:flex;justify-content:space-between;padding:6px 4px;border-bottom:1px solid #f1f5f9}
  .row:last-child{border-bottom:0}
  .small{font-size:13px;color:#6b7280}
</style>
</head>
<body>
  <h1>PPT → PDF Converter</h1>
  <div class="drop" id="drop">
    <p>Drag & drop your .ppt / .pptx files here<br>or click to select (multiple supported)</p>
  </div>
  <button class="btn" id="convertBtn" disabled>Convert & Download ZIP</button>

  <div class="files" id="filesArea" style="display:none">
    <div class="filelist" id="fileList"></div>
  </div>

<script>
const drop = document.getElementById('drop');
const fileList = document.getElementById('fileList');
const filesArea = document.getElementById('filesArea');
const convertBtn = document.getElementById('convertBtn');
let files = [];

function updateUI(){
  if(files.length === 0){
    filesArea.style.display = 'none';
    convertBtn.disabled = true;
  } else {
    filesArea.style.display = 'block';
    convertBtn.disabled = false;
  }
  fileList.innerHTML = files.map(f => {
    return `<div class="row"><div><strong>${escapeHtml(f.name)}</strong><div class="small">${formatBytes(f.size)}</div></div><div class="small">Ready</div></div>`;
  }).join('');
}

function escapeHtml(s){ 
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}
function formatBytes(bytes){
  if(bytes < 1024) return bytes + ' B';
  let kb = bytes / 1024;
  if(kb < 1024) return kb.toFixed(1) + ' KB';
  let mb = kb / 1024;
  if(mb < 1024) return mb.toFixed(1) + ' MB';
  return (mb / 1024).toFixed(1) + ' GB';
}

drop.addEventListener('click', ()=> {
  const inp = document.createElement('input');
  inp.type = 'file';
  inp.multiple = true;
  inp.accept = '.ppt,.pptx';
  inp.onchange = e => { files = files.concat(Array.from(inp.files)); updateUI(); };
  inp.click();
});
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
drop.addEventListener('dragleave', e => { drop.classList.remove('dragover'); });
drop.addEventListener('drop', e => {
  e.preventDefault();
  drop.classList.remove('dragover');
  const dropped = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith('.ppt') || f.name.toLowerCase().endsWith('.pptx'));
  files = files.concat(dropped);
  updateUI();
});

convertBtn.addEventListener('click', async () => {
  if(files.length === 0) return;
  convertBtn.disabled = true; convertBtn.textContent = 'Converting...';
  const form = new FormData();
  files.forEach(f => form.append('files', f));
  try {
    const resp = await fetch('/upload', { method:'POST', body: form });
    if(!resp.ok){ const txt = await resp.text(); alert('Server error: ' + txt); resetBtn(); return; }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'converted_pdfs.zip';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    convertBtn.textContent = 'Done ✓';
    setTimeout(()=>{ resetBtn(); files = []; updateUI(); }, 1400);
  } catch (err) { alert('Error: ' + err.message); resetBtn(); }
});

function resetBtn(){ convertBtn.disabled = false; convertBtn.textContent = 'Convert & Download ZIP'; }
</script>
</body>
</html>
"""
def convert_with_com_windows(input_path, output_path):
    try:
        import pythoncom
        from win32com.client import Dispatch
    except Exception as e:
        return False, f"pywin32 not installed or failed: {e}"

    try:
        pythoncom.CoInitialize()
        powerpoint = Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # visible, warna 2013 error aata hai

        # Window ko minimize karna try karo taaki user disturb na ho:
        # PowerPoint.Application.WindowState = 2 means minimized
        try:
            powerpoint.WindowState = 2  # minimize PowerPoint window
        except Exception:
            pass  # agar error aaye to ignore karo

        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)  # 32 = ppSaveAsPDF
        presentation.Close()
        powerpoint.Quit()
        pythoncom.CoUninitialize()
        return True, None
    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False, str(e)


@app.route("/")
def index():
    return render_template_string(INDEX_HTML)

@app.route("/upload", methods=["POST"])
def upload():
    if platform.system() != "Windows":
        return "This service currently only supports Windows with MS PowerPoint installed.", 400

    if 'files' not in request.files:
        return "No files uploaded", 400

    files = request.files.getlist('files')
    if not files:
        return "No files uploaded", 400

    tmpdir = tempfile.mkdtemp(prefix="ppt2pdf_")
    outdir = os.path.join(tmpdir, "out")
    os.makedirs(outdir, exist_ok=True)
    created = []

    try:
        for f in files:
            filename = secure_filename(f.filename) or "upload.pptx"
            if not filename.lower().endswith(('.ppt', '.pptx')):
                continue
            inpath = os.path.join(tmpdir, filename)
            f.save(inpath)
            base = os.path.splitext(filename)[0]
            pdfpath = os.path.join(outdir, base + ".pdf")

            ok, err = convert_with_com_windows(inpath, pdfpath)
            if ok and os.path.exists(pdfpath):
                created.append(pdfpath)
            else:
                return f"Conversion failed for {filename}: {err}", 500

        if not created:
            return "No valid PPT/PPTX files converted.", 400

        zip_path = os.path.join(tmpdir, "converted_pdfs.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for p in created:
                zf.write(p, arcname=os.path.basename(p))

        def cleanup_later(dirpath, delay=5):
            time.sleep(delay)
            try:
                shutil.rmtree(dirpath)
            except Exception:
                pass

        threading.Thread(target=cleanup_later, args=(tmpdir,), daemon=True).start()
        return send_file(zip_path, as_attachment=True, download_name="converted_pdfs.zip")

    except Exception as e:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass
        return f"Server error: {e}", 500

if __name__ == "__main__":
    # Only for local testing, do not expose publicly without security.
    app.run(host="0.0.0.0", port=5000, debug=True)
