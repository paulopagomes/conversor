#!/usr/bin/env python3
import os
import sys
import shutil
import subprocess
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image

# Optional: PyMuPDF for PDF->images and PDF->text without external poppler
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except Exception:
    HAS_PYMUPDF = False

# Optional: pypdf for merging PDFs
try:
    from pypdf import PdfWriter
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

OUTPUT_FORMATS = ["pdf", "jpg", "png", "webp", "tiff", "bmp", "txt"]

OFFICE_EXTS = {
    ".doc", ".docx", ".odt", ".rtf", ".txt",
    ".xls", ".xlsx", ".ods", ".csv",
    ".ppt", ".pptx", ".odp"
}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".gif", ".heic", ".heif"}

def safe_mkdir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def sanitize_basename(name: str) -> str:
    bad = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for b in bad:
        name = name.replace(b, "_")
    name = name.strip()
    return name if name else "arquivo"

def which_or_none(cmd: str):
    return shutil.which(cmd)

def run_cmd(cmd_list):
    try:
        p = subprocess.run(cmd_list, capture_output=True, text=True, check=False)
        return (p.returncode == 0, p.stdout, p.stderr)
    except Exception as e:
        return (False, "", str(e))

def convert_image_file(in_path: Path, out_path: Path, fmt: str, quality: int):
    safe_mkdir(out_path.parent)
    with Image.open(in_path) as im:
        if fmt == "jpg":
            if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                im = im.convert("RGBA")
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[-1])
                im = bg
            else:
                im = im.convert("RGB")
            im.save(out_path, "JPEG", quality=int(quality), optimize=True)
            return
        if fmt == "png":
            im.save(out_path, "PNG", optimize=True); return
        if fmt == "webp":
            im.save(out_path, "WEBP", quality=int(quality), method=6); return
        if fmt == "tiff":
            im.save(out_path, "TIFF"); return
        if fmt == "bmp":
            im.save(out_path, "BMP"); return
        raise ValueError("Formato de saída inválido.")

def image_to_pdf(in_path: Path, out_path: Path):
    safe_mkdir(out_path.parent)
    with Image.open(in_path) as im:
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert("RGB")
        im.save(out_path, "PDF", resolution=300.0)

def pdf_to_text_pymupdf(pdf_path: Path, out_txt: Path):
    if not HAS_PYMUPDF:
        raise RuntimeError("PyMuPDF não está disponível (dependência faltando).")
    safe_mkdir(out_txt.parent)
    doc = fitz.open(pdf_path)
    parts = []
    for page in doc:
        parts.append(page.get_text("text"))
    out_txt.write_text("\n".join(parts), encoding="utf-8", errors="ignore")

def pdf_to_images_pymupdf(pdf_path: Path, out_dir: Path, fmt: str, dpi: int, quality: int, prefix: str):
    if not HAS_PYMUPDF:
        raise RuntimeError("PyMuPDF não está disponível (dependência faltando).")
    safe_mkdir(out_dir)
    doc = fitz.open(pdf_path)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        # Save to PNG first, then convert if needed
        tmp = out_dir / f"{prefix}-{i:03d}.png"
        pix.save(str(tmp))

        if fmt == "png":
            continue
        out_file = out_dir / f"{prefix}-{i:03d}.{fmt}"
        convert_image_file(tmp, out_file, fmt, quality)
        try:
            tmp.unlink()
        except Exception:
            pass

def office_to_pdf(in_path: Path, out_dir: Path, log_cb):
    soffice = which_or_none("soffice") or which_or_none("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice não encontrado. (Opcional no Windows)")
    safe_mkdir(out_dir)
    cmd = [soffice, "--headless", "--nologo", "--nolockcheck", "--norestore",
           "--convert-to", "pdf", "--outdir", str(out_dir), str(in_path)]
    log_cb("$ " + " ".join(cmd))
    ok, out, err = run_cmd(cmd)
    if not ok:
        raise RuntimeError(f"Erro LibreOffice: {err.strip() or out.strip()}")
    out_pdf = out_dir / (in_path.stem + ".pdf")
    if not out_pdf.exists():
        cands = list(out_dir.glob(in_path.stem + ".pdf")) + list(out_dir.glob(in_path.stem + ".PDF"))
        if cands:
            return cands[0]
        raise RuntimeError("LibreOffice executou, mas não achei o PDF gerado.")
    return out_pdf

def merge_pdfs(pdf_paths, out_pdf: Path, log_cb):
    if not HAS_PYPDF:
        raise RuntimeError("pypdf não disponível. (dependência faltando)")
    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(str(p))
    safe_mkdir(out_pdf.parent)
    with open(out_pdf, "wb") as f:
        writer.write(f)
    log_cb(f"OK -> {out_pdf}")

class MultiFilePicker(tk.Toplevel):
    def __init__(self, master, start_dir: Path):
        super().__init__(master)
        self.title("Selecionar arquivos (multi)")
        self.geometry("760x480")
        self.minsize(760, 480)
        self.selected = []
        self.cur_dir = start_dir if start_dir.exists() else Path.home()
        self._build()
        self._refresh()
        self.transient(master)
        self.grab_set()

    def _build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Pasta:").pack(side="left")
        self.dir_var = tk.StringVar(value=str(self.cur_dir))
        ttk.Entry(top, textvariable=self.dir_var).pack(side="left", fill="x", expand=True, padx=(6,6))
        ttk.Button(top, text="Ir", command=self._go).pack(side="left")
        ttk.Button(top, text="⬆️ Subir", command=self._up).pack(side="left", padx=(6,0))

        mid = ttk.Frame(self, padding=(10,0,10,10))
        mid.pack(fill="both", expand=True)
        ttk.Label(mid, text="Use Shift+setas / Ctrl+clique para multi seleção.").pack(anchor="w", pady=(0,6))
        self.listbox = tk.Listbox(mid, selectmode="extended")
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<Double-Button-1>", self._open_dir)

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Adicionar selecionados", command=self._add_selected).pack(side="right")
        ttk.Button(bottom, text="Cancelar", command=self._cancel).pack(side="right", padx=(0,8))

    def _go(self):
        p = Path(self.dir_var.get()).expanduser()
        if p.exists() and p.is_dir():
            self.cur_dir = p
            self._refresh()
        else:
            messagebox.showerror("Pasta inválida", "Essa pasta não existe.")

    def _up(self):
        if self.cur_dir.parent != self.cur_dir:
            self.cur_dir = self.cur_dir.parent
            self._refresh()

    def _refresh(self):
        self.dir_var.set(str(self.cur_dir))
        self.listbox.delete(0, "end")
        entries = []
        try:
            for p in sorted(self.cur_dir.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower())):
                entries.append(("dir" if p.is_dir() else "file", p))
        except PermissionError:
            messagebox.showerror("Sem permissão", "Sem permissão para acessar essa pasta.")
            return
        self._entries = entries
        for kind, p in entries:
            prefix = "[DIR] " if kind == "dir" else "      "
            self.listbox.insert("end", f"{prefix}{p.name}")

    def _open_dir(self, _evt=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        kind, p = self._entries[sel[0]]
        if kind == "dir":
            self.cur_dir = p
            self._refresh()

    def _add_selected(self):
        sels = list(self.listbox.curselection())
        out = []
        for idx in sels:
            kind, p = self._entries[idx]
            if kind == "file":
                out.append(str(p))
        self.selected = out
        self.destroy()

    def _cancel(self):
        self.selected = []
        self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conversor + Juntar PDFs — Windows")
        self.geometry("980x640")
        self.minsize(980, 640)
        self.files = []
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Adicionar arquivos (multi)…", command=self.add_files_multi).pack(side="left")
        ttk.Button(top, text="Adicionar pasta…", command=self.add_folder).pack(side="left", padx=(8,0))
        ttk.Button(top, text="Remover selecionado", command=self.remove_selected).pack(side="left", padx=(8,0))
        ttk.Button(top, text="Limpar lista", command=self.clear_files).pack(side="left", padx=(8,0))

        ttk.Separator(self).pack(fill="x", padx=10, pady=8)

        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="both", expand=True)

        left = ttk.Frame(mid)
        left.pack(side="left", fill="both", expand=True)
        ttk.Label(left, text="Arquivos na fila: (Shift+setas funciona aqui e no seletor interno)").pack(anchor="w")
        self.listbox = tk.Listbox(left, height=18, selectmode="extended")
        self.listbox.pack(fill="both", expand=True, pady=(6,0))

        right = ttk.Frame(mid)
        right.pack(side="left", fill="y", padx=(14,0))

        ttk.Label(right, text="Pasta de saída:").pack(anchor="w")
        out_row = ttk.Frame(right)
        out_row.pack(fill="x", pady=(6,10))
        self.out_dir_var = tk.StringVar(value=str(Path.home() / "Conversoes"))
        ttk.Entry(out_row, textvariable=self.out_dir_var, width=38).pack(side="left", fill="x", expand=True)
        ttk.Button(out_row, text="Escolher…", command=self.choose_out_dir).pack(side="left", padx=(8,0))

        ttk.Label(right, text="Formato de saída:").pack(anchor="w")
        self.fmt_var = tk.StringVar(value="pdf")
        ttk.Combobox(right, textvariable=self.fmt_var, values=OUTPUT_FORMATS, state="readonly", width=12).pack(anchor="w", pady=(6,10))

        ttk.Label(right, text="DPI (PDF→imagem):").pack(anchor="w")
        self.dpi_var = tk.IntVar(value=200)
        ttk.Spinbox(right, from_=72, to=600, textvariable=self.dpi_var, width=12).pack(anchor="w", pady=(6,10))

        ttk.Label(right, text="Qualidade (JPG/WebP):").pack(anchor="w")
        self.quality_var = tk.IntVar(value=90)
        ttk.Spinbox(right, from_=40, to=100, textvariable=self.quality_var, width=12).pack(anchor="w", pady=(6,10))

        self.subfolder_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="Criar subpasta por arquivo", variable=self.subfolder_var).pack(anchor="w", pady=(6,14))

        ttk.Button(right, text="CONVERTER", command=self.convert_all).pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=14)
        ttk.Label(right, text="Juntar PDFs (usa TODOS PDFs da fila):").pack(anchor="w")
        self.merge_name_var = tk.StringVar(value="pdf_juntado.pdf")
        ttk.Entry(right, textvariable=self.merge_name_var, width=24).pack(anchor="w", pady=(6,8))
        ttk.Button(right, text="JUNTAR PDFs", command=self.merge_pdfs_now).pack(fill="x")

        ttk.Separator(self).pack(fill="x", padx=10, pady=8)

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="both", expand=False)
        ttk.Label(bottom, text="Log:").pack(anchor="w")
        self.log = tk.Text(bottom, height=9, wrap="word")
        self.log.pack(fill="both", expand=True, pady=(6,0))
        self.log.configure(state="disabled")

        # Warnings box
        warn = []
        if not HAS_PYMUPDF:
            warn.append("• PDF→imagem/TXT precisa do PyMuPDF (fitz).")
        if not HAS_PYPDF:
            warn.append("• Juntar PDFs precisa do pypdf.")
        if warn:
            self._log("AVISOS:\n" + "\n".join(warn))

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            if p and p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", p)
                added += 1
        if added:
            self._log(f"Adicionados {added} arquivo(s).")

    def add_files_multi(self):
        picker = MultiFilePicker(self, Path.home())
        self.wait_window(picker)
        self._add_paths(picker.selected)

    def add_folder(self):
        d = filedialog.askdirectory(title="Escolha uma pasta (vai adicionar todos os arquivos dentro)")
        if not d:
            return
        folder = Path(d)
        paths = [str(p) for p in sorted(folder.rglob("*")) if p.is_file()]
        self._add_paths(paths)

    def remove_selected(self):
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for i in reversed(sel):
            path = self.listbox.get(i)
            self.listbox.delete(i)
            try:
                self.files.remove(path)
            except ValueError:
                pass

    def clear_files(self):
        self.files = []
        self.listbox.delete(0, "end")

    def choose_out_dir(self):
        d = filedialog.askdirectory(title="Escolha a pasta de saída")
        if d:
            self.out_dir_var.set(d)

    def convert_all(self):
        if not self.files:
            messagebox.showwarning("Sem arquivos", "Adicione pelo menos um arquivo.")
            return

        out_base = Path(self.out_dir_var.get()).expanduser()
        fmt = self.fmt_var.get().lower().strip()
        dpi = int(self.dpi_var.get())
        quality = int(self.quality_var.get())
        safe_mkdir(out_base)

        ok_count = 0
        fail_count = 0

        for f in list(self.files):
            in_path = Path(f)
            if not in_path.exists():
                self._log(f"ERRO: não encontrado: {in_path}")
                fail_count += 1
                continue

            ext = in_path.suffix.lower()
            base_name = sanitize_basename(in_path.stem)

            dest_dir = out_base / base_name if True else out_base
            if self.subfolder_var.get():
                dest_dir = out_base / base_name
            else:
                dest_dir = out_base
            safe_mkdir(dest_dir)

            try:
                self._log(f"\n== {in_path.name} ==")

                if fmt == "txt":
                    if ext != ".pdf":
                        raise RuntimeError("TXT só é suportado a partir de PDF (PDF → TXT).")
                    out_txt = dest_dir / f"{base_name}.txt"
                    pdf_to_text_pymupdf(in_path, out_txt)
                    self._log(f"OK -> {out_txt}")
                    ok_count += 1
                    continue

                if fmt == "pdf":
                    out_pdf = dest_dir / f"{base_name}.pdf"

                    if ext == ".pdf":
                        shutil.copy2(in_path, out_pdf)
                        self._log(f"OK (copiado) -> {out_pdf}")
                        ok_count += 1
                        continue

                    if ext in IMAGE_EXTS:
                        image_to_pdf(in_path, out_pdf)
                        self._log(f"OK -> {out_pdf}")
                        ok_count += 1
                        continue

                    if ext in OFFICE_EXTS:
                        produced = office_to_pdf(in_path, dest_dir, self._log)
                        # rename to our pattern if possible
                        if produced != out_pdf and produced.exists():
                            try:
                                if out_pdf.exists():
                                    out_pdf.unlink()
                                produced.rename(out_pdf)
                                self._log(f"OK -> {out_pdf}")
                            except Exception:
                                self._log(f"OK -> {produced}")
                        ok_count += 1
                        continue

                    raise RuntimeError("Tipo não suportado para virar PDF (use Office/Imagem/PDF).")

                # image outputs
                if ext == ".pdf":
                    if fmt not in ("jpg","png","webp","tiff","bmp"):
                        raise RuntimeError("Para PDF, escolha saída: jpg/png/webp/tiff/bmp/txt/pdf.")
                    pdf_to_images_pymupdf(in_path, dest_dir, fmt, dpi, quality, base_name)
                    ok_count += 1
                    continue

                if ext in IMAGE_EXTS:
                    out_img = dest_dir / f"{base_name}.{fmt}"
                    convert_image_file(in_path, out_img, fmt, quality)
                    self._log(f"OK -> {out_img}")
                    ok_count += 1
                    continue

                if ext in OFFICE_EXTS:
                    raise RuntimeError("Office → imagem não disponível. Use saída PDF para Office.")

                raise RuntimeError("Tipo de arquivo não suportado para essa conversão.")

            except Exception as e:
                self._log(f"FALHOU: {in_path} -> {e}")
                fail_count += 1

        messagebox.showinfo("Concluído", f"Conversão finalizada.\nSucesso: {ok_count}\nFalhas: {fail_count}\nSaída: {out_base}")

    def merge_pdfs_now(self):
        pdfs = []
        for p in self.files:
            pp = Path(p)
            if pp.suffix.lower() == ".pdf" and pp.exists():
                pdfs.append(pp)
        if len(pdfs) < 2:
            messagebox.showwarning("Poucos PDFs", "Adicione pelo menos 2 PDFs na fila para juntar.")
            return
        out_dir = Path(self.out_dir_var.get()).expanduser()
        safe_mkdir(out_dir)
        name = self.merge_name_var.get().strip()
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        out_pdf = out_dir / (sanitize_basename(Path(name).stem) + ".pdf")
        try:
            self._log(f"\n== JUNTAR PDFs ({len(pdfs)}) ==")
            merge_pdfs(pdfs, out_pdf, self._log)
            messagebox.showinfo("PDF juntado", f"OK!\nArquivo gerado:\n{out_pdf}")
        except Exception as e:
            messagebox.showerror("Erro ao juntar PDFs", str(e))

def main():
    App().mainloop()

if __name__ == "__main__":
    main()
