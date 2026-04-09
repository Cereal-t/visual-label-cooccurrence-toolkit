
# Visual Label Analysis Tool
# Step 1: Google Vision Label Detection and Semantic Deduplication
# Step 2: Filter Label Co-occurrence and Generate the Matrix, Heatmap, and Gephi Files
# Step 3: Image Grid for Top-N Label Pairs


import math
import os
import base64
import time
import threading
import traceback
from collections import Counter
from itertools import combinations
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import requests
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


# ══════════════════════════════════════════════════════════
# Shared style constants
# ══════════════════════════════════════════════════════════

WHITE       = "#FFFFFF"
BLACK       = "#000000"
GREY_BTN    = "#E0E0E0"
GREY_TEXT   = "#555555"
BLUE_PATH   = "#E7F1FF"
BLUE_INFO   = "#0000FF"
RED_WARN    = "#CC0000"
ACCENT_BLUE = "#4A90D9"

FONT_TITLE  = ("Arial", 16, "bold")
FONT_H2     = ("Arial", 13, "bold")
FONT_LABEL  = ("Arial", 12)
FONT_BODY   = ("Arial", 12)
FONT_SMALL  = ("Arial", 11)
FONT_MONO   = ("Courier", 11)

PAD_X = 12
PAD_Y = 6


# ══════════════════════════════════════════════════════════
# Deduplication prompt text
# ══════════════════════════════════════════════════════════

PROMPT_TEXT = """Task Background: The labels below were extracted from a single image by a computer vision API (e.g. Google Vision). The API often returns labels with overlapping semantic content for the same image. For co-occurrence analysis, each retained label should represent a distinct visual element or concept, so redundant labels within the same image should be removed.
Your task: Perform intra-image semantic deduplication on the labels of each image. Compare labels only within the same image. Do not compare or remove labels across different images.

Decision priority (apply in order):
1. Hypernym-hyponym pairs
If a specific term and its broader category co-occur, keep the specific term and remove the broader one.
Example: dog + canidae -> keep dog; stratovolcano + volcanic landform -> keep stratovolcano

2. Everyday language over taxonomic jargon
When a common-language term and a taxonomic/scientific term refer to the same entity, keep the common-language term.
Example: cat + felinae -> keep cat; string instrument + chordophone -> keep string instrument

3. Near-synonyms
If two labels are semantically near-identical and convey no independent visual information, keep the more common or more visually grounded term.
Example: sunset + afterglow -> keep sunset; mist + fog -> keep fog; crystal + mineral + gemstone → keep crystal

4. Surface-form duplicates
Singular/plural or minor orthographic variants of the same concept: keep one, preferring the singular form.

Retain the following label types:
- Style or image-type labels, e.g. animation, cg artwork, watercolor painting
- Atmosphere, mood, or aesthetic labels, e.g. dusk, mystery, goth subculture
- Colour or lighting labels, e.g. orange, sunlight, backlighting
- Abstract concept or thematic labels, e.g. mythology, fiction, symmetry

Hard constraints:
- Do not add any new labels.
- Do not rewrite, translate, or rename original labels.
- Do not merge multiple labels into a new phrase.
- Only perform keep or remove on original labels.
- When in doubt, retain rather than remove.

Output:
1. Deduplicated label table (original columns preserved): image_id | label | score
2. Removal log: image_id, removed_label, kept_label, reason"""


# ══════════════════════════════════════════════════════════
# Shared UI helpers
# ══════════════════════════════════════════════════════════

def make_path_row(parent, row_idx, label_text, var, browse_cmd):
    tk.Label(parent, text=label_text, font=FONT_LABEL,
             bg=WHITE, fg=BLACK, anchor="e").grid(
        row=row_idx, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
    tk.Entry(parent, textvariable=var, font=FONT_LABEL,
             bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4).grid(
        row=row_idx, column=1, sticky="ew", pady=PAD_Y)
    if browse_cmd:
        b = tk.Button(parent, text="Browse...", command=browse_cmd,
                      font=FONT_SMALL, bg=GREY_BTN, fg=BLACK,
                      relief="flat", bd=0, padx=10, pady=4)
        b.bind("<Enter>", lambda _: b.config(bg="#C8C8C8"))
        b.bind("<Leave>", lambda _: b.config(bg=GREY_BTN))
        b.grid(row=row_idx, column=2, padx=(6, PAD_X), pady=PAD_Y)


def make_output_label(parent, row_idx, save_text, var, browse_cmd):
    out_frame = tk.Frame(parent, bg=WHITE)
    out_frame.grid(row=row_idx, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
    tk.Label(out_frame, text="Output", font=("Arial", 12, "bold"),
             bg=WHITE, fg=BLACK).pack(side="left")
    tk.Label(out_frame, text=save_text, font=FONT_LABEL,
             bg=WHITE, fg=BLACK).pack(side="left")
    tk.Entry(parent, textvariable=var, font=FONT_LABEL,
             bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4).grid(
        row=row_idx, column=1, sticky="ew", pady=PAD_Y)
    b = tk.Button(parent, text="Browse...", command=browse_cmd,
                  font=FONT_SMALL, bg=GREY_BTN, fg=BLACK,
                  relief="flat", bd=0, padx=10, pady=4)
    b.bind("<Enter>", lambda _: b.config(bg="#C8C8C8"))
    b.bind("<Leave>", lambda _: b.config(bg=GREY_BTN))
    b.grid(row=row_idx, column=2, padx=(6, PAD_X), pady=PAD_Y)


def grey_btn(parent, text, cmd):
    b = tk.Button(parent, text=text, command=cmd,
                  font=FONT_BODY, bg=GREY_BTN, fg=BLACK,
                  activebackground="#C8C8C8", activeforeground=BLACK,
                  relief="flat", bd=0, padx=14, pady=8)
    b.bind("<Enter>", lambda _: b.config(bg="#C8C8C8"))
    b.bind("<Leave>", lambda _: b.config(bg=GREY_BTN))
    return b


def text_label(parent, text, color=GREY_TEXT, font=None, wrap=840):
    tk.Label(parent, text=text, font=font or FONT_BODY,
             bg=WHITE, fg=color,
             wraplength=wrap, justify="left").pack(
        anchor="w", padx=PAD_X, pady=(2, 6))


def make_scrollable(root_frame):
    canvas = tk.Canvas(root_frame, bg=WHITE, highlightthickness=0)
    sb = tk.Scrollbar(root_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=WHITE)
    win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>",
               lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",
                lambda e: canvas.itemconfig(win_id, width=e.width))
    return inner


# ══════════════════════════════════════════════════════════
# Step 1 — Google Vision Label Detection and Semantic Deduplication
# ══════════════════════════════════════════════════════════

VISION_URL_TEMPLATE = "https://vision.googleapis.com/v1/images:annotate?key={key}"
SUPPORTED_EXTS      = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tiff"}
SAVE_EVERY          = 20


def call_vision_api(image_path, api_key):
    file_name = os.path.basename(image_path)
    image_id  = file_name
    try:
        with open(image_path, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")
        payload = {"requests": [{"image": {"content": content},
                                  "features": [{"type": "LABEL_DETECTION"}]}]}
        resp = requests.post(VISION_URL_TEMPLATE.format(key=api_key),
                             json=payload, timeout=60)
        data = resp.json()
        if "error" in data:
            return [{"image_id": image_id, "label": None, "score": None,
                     "_error": data["error"].get("message", "top-level API error")}]
        if not data.get("responses"):
            return [{"image_id": image_id, "label": None, "score": None,
                     "_error": "empty response from API"}]
        result = data["responses"][0]
        if "error" in result:
            return [{"image_id": image_id, "label": None, "score": None,
                     "_error": result["error"].get("message", "image-level API error")}]
        annotations = result.get("labelAnnotations", [])
        if not annotations:
            return [{"image_id": image_id, "label": None, "score": None,
                     "_error": "no labels detected"}]
        return [{"image_id": image_id, "label": a["description"],
                 "score": a["score"], "_error": None}
                for a in annotations]
    except Exception as e:
        return [{"image_id": image_id, "label": None, "score": None, "_error": str(e)}]


class Step1App:

    def __init__(self, root, content):
        self.root    = root    # single tk.Tk window
        self.content = content # shared scrollable inner frame

        self._running        = False
        self._prompt_visible = False

        self.v_img_folder = tk.StringVar()
        self.v_api_key    = tk.StringVar()
        self.v_out_long   = tk.StringVar()
        self._eta_var     = tk.StringVar(value="")

        title_bar = tk.Frame(self.content, bg=BLACK, pady=12)
        title_bar.pack(fill="x")
        tk.Label(title_bar,
                 text="Step 1   Google Vision Label Detection and Semantic Deduplication",
                 font=FONT_TITLE, bg=BLACK, fg=WHITE).pack(padx=24, anchor="w")
        self._build_detect_section()
        self._build_dedup_section()
        tk.Frame(self.content, bg=WHITE, height=8).pack()

    def _build_detect_section(self):
        tk.Label(self.content, text="Detect Labels", font=FONT_H2,
                 bg=WHITE, fg=BLACK).pack(anchor="w", padx=PAD_X, pady=(18, 2))
        text_label(self.content,
                   "Select an image folder and enter the Google Cloud Vision API key to extract "
                   "labels and their confidence scores. Your API key is never stored or "
                   "transmitted elsewhere.", wrap=1200)

        gf = tk.Frame(self.content, bg=WHITE)
        gf.pack(fill="x", padx=4)
        gf.columnconfigure(1, weight=1)
        make_path_row(gf, 0, "Image folder:", self.v_img_folder,
                      lambda: self._browse_folder(self.v_img_folder))
        make_path_row(gf, 1, "Google Cloud API key:", self.v_api_key, None)
        make_output_label(gf, 2, "Save raw_labels.csv to:", self.v_out_long,
                          lambda: self._browse_save(self.v_out_long, "raw_labels.csv"))

        btn_frame = tk.Frame(self.content, bg=WHITE)
        btn_frame.pack(anchor="w", padx=PAD_X, pady=(8, 4))
        grey_btn(btn_frame, "Start Detecting Labels",
                 self._run_extraction).pack(side="left")
        tk.Label(btn_frame, textvariable=self._eta_var,
                 font=FONT_SMALL, bg=WHITE, fg=GREY_TEXT).pack(side="left", padx=(14, 0))
        tk.Label(btn_frame,
                 text="  Columns: image_id | label | score   "
                      "Format: all labels for one image in a single cell separated by semicolons",
                 font=FONT_SMALL, bg=WHITE, fg=GREY_TEXT).pack(side="left", padx=(14, 0))

        self.v_img_folder.trace_add("write", self._update_eta)

        tk.Label(self.content, text="Progress log:", font=FONT_SMALL,
                 bg=WHITE, fg=GREY_TEXT).pack(anchor="w", padx=PAD_X, pady=(10, 2))
        self.log_box = scrolledtext.ScrolledText(
            self.content, height=8, font=FONT_MONO,
            bg="#F5F5F5", fg=BLACK, relief="flat", bd=1, state="disabled")
        self.log_box.pack(fill="x", padx=PAD_X, pady=(0, 6))

    def _build_dedup_section(self):
        tk.Label(self.content, text="Label Deduplication", font=FONT_H2,
                 bg=WHITE, fg=BLACK).pack(anchor="w", padx=PAD_X, pady=(18, 4))
        text_label(self.content,
                   "Google Vision often returns labels with semantic overlap for the same image. "
                   "Before running co-occurrence analysis, it is recommended to perform deduplication "
                   "on your raw labels using an AI language model or other methods. "
                   "A suggested prompt is provided below. You are encouraged to adapt it to your own label vocabulary.",
                   color=BLUE_INFO)

        self._toggle_btn = tk.Button(
            self.content,
            text="▶  Show suggested deduplication prompt",
            font=FONT_SMALL, bg="#EBF2FB", fg=ACCENT_BLUE,
            relief="flat", bd=0, padx=10, pady=6, anchor="w",
            command=self._toggle_prompt)
        self._toggle_btn.bind("<Enter>", lambda _: self._toggle_btn.config(bg="#D5E8FA"))
        self._toggle_btn.bind("<Leave>", lambda _: self._toggle_btn.config(bg="#EBF2FB"))
        self._toggle_btn.pack(anchor="w", padx=PAD_X, pady=(0, 0))

        # Fixed container — always packed here, directly below toggle button.
        # Both the expanded prompt frame and the collapsed red warning live inside
        # this container, so their position is fixed regardless of pack/pack_forget.
        self._collapsible_container = tk.Frame(self.content, bg=WHITE)
        self._collapsible_container.pack(fill="x")

        # Prompt frame (shown when expanded) — inside fixed container
        self._prompt_frame = tk.Frame(self._collapsible_container, bg=WHITE)
        prompt_box = scrolledtext.ScrolledText(
            self._prompt_frame, height=18, font=FONT_MONO,
            bg="#F8FAFD", fg=BLACK, relief="flat", bd=1, wrap="word")
        prompt_box.insert("1.0", PROMPT_TEXT)
        prompt_box.config(state="disabled")
        prompt_box.pack(fill="x", padx=PAD_X, pady=(4, 0))

        copy_btn = tk.Button(
            self._prompt_frame, text="Copy prompt to clipboard",
            font=FONT_SMALL, bg="#EBF2FB", fg=ACCENT_BLUE,
            relief="flat", bd=0, padx=10, pady=5,
            command=self._copy_prompt)
        copy_btn.bind("<Enter>", lambda _: copy_btn.config(bg="#D5E8FA"))
        copy_btn.bind("<Leave>", lambda _: copy_btn.config(bg="#EBF2FB"))
        copy_btn.pack(anchor="e", padx=PAD_X, pady=(4, 0))

        tk.Label(self._prompt_frame,
                 text="After deduplication, use deduplicated_labels.csv as the input for next step.",
                 font=FONT_BODY, bg=WHITE, fg=RED_WARN,
                 wraplength=840, justify="left").pack(anchor="w", padx=PAD_X, pady=(8, 4))

        # Collapsed-state red warning — inside fixed container
        self._red_warn_collapsed = tk.Label(
            self._collapsible_container,
            text="After deduplication, use deduplicated_labels.csv as the input for next step.",
            font=FONT_BODY, bg=WHITE, fg=RED_WARN,
            wraplength=840, justify="left")
        self._red_warn_collapsed.pack(anchor="w", padx=PAD_X, pady=(8, 4))

    def _update_eta(self, *_):
        folder = self.v_img_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            self._eta_var.set("")
            return
        try:
            count = sum(1 for f in os.listdir(folder)
                        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS)
        except Exception:
            self._eta_var.set("")
            return
        if count == 0:
            self._eta_var.set("No supported images found.")
            return
        s = count * 1.5
        if s < 60:
            eta = f"~{int(s)} seconds"
        elif s < 3600:
            eta = f"~{int(s / 60)} minutes"
        else:
            eta = f"~{int(s/3600)} h {int((s % 3600)/60)} min"
        self._eta_var.set(f"{count} images found  |  Estimated time: {eta}")

    def _toggle_prompt(self):
        self._prompt_visible = not self._prompt_visible
        if self._prompt_visible:
            self._red_warn_collapsed.pack_forget()
            self._prompt_frame.pack(fill="x", pady=(4, 0),
                                    in_=self._collapsible_container)
            self._toggle_btn.config(text="▼  Hide suggested deduplication prompt")
        else:
            self._prompt_frame.pack_forget()
            self._toggle_btn.config(text="▶  Show suggested deduplication prompt")
            self._red_warn_collapsed.pack(anchor="w", padx=PAD_X, pady=(8, 4),
                                          in_=self._collapsible_container)

    def _copy_prompt(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(PROMPT_TEXT)
        messagebox.showinfo("Copied", "Prompt copied to clipboard.")

    def _browse_folder(self, var):
        p = filedialog.askdirectory(title="Select image folder")
        if p: var.set(p)

    def _browse_save(self, var, default_name):
        p = filedialog.asksaveasfilename(
            title="Save CSV as", initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p: var.set(p)

    def _log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        self.root.update_idletasks()

    def _run_extraction(self):
        if self._running:
            messagebox.showwarning("Busy", "A task is already running. Please wait.")
            return
        folder  = self.v_img_folder.get().strip()
        api_key = self.v_api_key.get().strip()
        out_csv = self.v_out_long.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid image folder.")
            return
        if not api_key:
            messagebox.showerror("Error", "Please enter your Google Cloud API key.")
            return
        if not out_csv:
            messagebox.showerror("Error", "Please specify an output CSV path.")
            return
        self._running = True
        threading.Thread(target=self._extraction_worker,
                         args=(folder, api_key, out_csv), daemon=True).start()

    def _extraction_worker(self, folder, api_key, out_csv):
        image_files = sorted([f for f in os.listdir(folder)
                               if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS])
        total = len(image_files)
        self._log(f"Images found: {total}")
        wide_records, failed_images = [], []

        def save_wide(records, path):
            pd.DataFrame(records).to_csv(path, index=False, encoding="utf-8")

        for idx, file_name in enumerate(image_files, 1):
            rows         = call_vision_api(os.path.join(folder, file_name), api_key)
            image_failed = False
            good_rows    = []
            for r in rows:
                err = r.pop("_error", None)
                if err:
                    self._log(f"  [{idx}/{total}] WARNING  {file_name}: {err}")
                    image_failed = True
                else:
                    good_rows.append(r)
            if image_failed:
                failed_images.append(file_name)
            else:
                labels_str = ";".join(r["label"] for r in good_rows)
                scores_str = ";".join(str(round(r["score"], 6)) for r in good_rows)
                wide_records.append({"image_id": file_name,
                                     "label":    labels_str,
                                     "score":    scores_str})
                self._log(f"  [{idx}/{total}] OK  {file_name}  ({len(good_rows)} labels)")
            if idx % SAVE_EVERY == 0:
                save_wide(wide_records, out_csv)
                self._log(f"  -> Progress saved at {idx} images")
            time.sleep(0.2)

        save_wide(wide_records, out_csv)
        total_labels = sum(len(r["label"].split(";")) for r in wide_records if r["label"])
        self._running = False
        self.root.after(0, lambda: messagebox.showinfo(
            "Extraction complete",
            f"Label detection finished.\n\n"
            f"Images processed successfully : {total - len(failed_images)}\n"
            f"Total labels extracted        : {total_labels}\n"
            f"Failed images                 : {len(failed_images)}\n\n"
            f"File saved to:\n{out_csv}"))

    def _on_close(self):
        if self._running:
            messagebox.showinfo("Please wait",
                                "Label extraction is still running. "
                                "Please wait for it to finish before closing.")
            return
        self.root.destroy()


# ══════════════════════════════════════════════════════════
# Step 2 — Filter Label Co-occurrence and Generate the Matrix, Heatmap, and Gephi Files
# ══════════════════════════════════════════════════════════

def load_labels(csv_path):
    """
    Load wide-format label CSV (one row per image, labels semicolon-separated
    in the second column).
    """
    last_err = None
    for enc in ("utf-8", "utf-8-sig", "gb18030", "latin1"):
        try:
            df = pd.read_csv(csv_path, encoding=enc)
            break
        except UnicodeDecodeError as e:
            last_err = e
            continue
    else:
        raise last_err

    label_col = df.columns[1]
    df = df.dropna(subset=[label_col])
    df["labels"] = (
        df[label_col]
        .astype(str)
        .str.split(";")
        .apply(lambda xs: sorted({x.strip() for x in xs if x.strip()}))
    )
    df = df[df["labels"].map(len) > 0].reset_index(drop=True)
    return df[["labels"]]


def build_cooccurrence(df):
    total_images = len(df)
    label_counts = Counter()
    pair_counts  = Counter()
    for labels in df["labels"]:
        unique_labels = list(labels)
        label_counts.update(unique_labels)
        for a, b in combinations(sorted(unique_labels), 2):
            pair_counts[(a, b)] += 1

    all_labels = sorted(label_counts.keys())
    co_mat = pd.DataFrame(0, index=all_labels, columns=all_labels, dtype=int)
    for (a, b), c in pair_counts.items():
        co_mat.at[a, b] = c
        co_mat.at[b, a] = c
    for lbl in label_counts:
        co_mat.at[lbl, lbl] = 0

    stats_rows = []
    for lbl, c in label_counts.items():
        stats_rows.append({"type": "unigram", "label_i": lbl, "label_j": "",
                           "count": c, "total_images": total_images})
    for (a, b), c in pair_counts.items():
        stats_rows.append({"type": "bigram", "label_i": a, "label_j": b,
                           "count": c, "total_images": total_images})
    return co_mat, pd.DataFrame(stats_rows)


def compute_pmi_edges(stats, min_cooccurrence=2, min_pmi=1.0):
    uni          = stats[stats["type"] == "unigram"].set_index("label_i")["count"]
    bi           = stats[stats["type"] == "bigram"].copy()
    total_images = stats["total_images"].iloc[0]
    rows = []
    for _, row in bi.iterrows():
        a, b = row["label_i"], row["label_j"]
        co   = int(row["count"])
        if co < min_cooccurrence:
            continue
        pa, pb, pab = uni[a]/total_images, uni[b]/total_images, co/total_images
        denom = pa * pb
        if denom <= 0 or pab <= 0:
            continue
        pmi = math.log2(pab / denom)
        if pmi < min_pmi:
            continue
        rows.append({"Source": a, "Target": b, "Weight": pmi, "cooccur_count": co})

    if not rows:
        return pd.DataFrame(columns=["Source", "Target", "Weight", "cooccur_count"])

    return (pd.DataFrame(rows)
            .sort_values(["Weight", "cooccur_count"], ascending=[False, False])
            .reset_index(drop=True))


def build_pmi_matrix(labels_in_edges, stats, co_mat_filtered,
                     min_cooccurrence, min_pmi):
    uni          = stats[stats["type"] == "unigram"].set_index("label_i")["count"]
    total_images = int(stats["total_images"].iloc[0])
    n            = len(labels_in_edges)
    data         = [[""] * n for _ in range(n)]
    for i, a in enumerate(labels_in_edges):
        for j, b in enumerate(labels_in_edges):
            if i >= j:
                continue
            co = int(co_mat_filtered.at[a, b])
            if co < min_cooccurrence:
                continue
            pa, pb, pab = uni[a]/total_images, uni[b]/total_images, co/total_images
            denom = pa * pb
            if denom <= 0 or pab <= 0:
                continue
            pmi = math.log2(pab / denom)
            if pmi < min_pmi:
                continue
            val = round(pmi, 1)
            data[i][j] = val
            data[j][i] = val
    return pd.DataFrame(data, index=labels_in_edges, columns=labels_in_edges)


def build_nodes_from_edges_and_stats(edges_df, stats):
    uni             = stats[stats["type"] == "unigram"].set_index("label_i")["count"]
    labels_in_edges = set(edges_df["Source"].tolist()) | set(edges_df["Target"].tolist())
    rows = [{"Id": lbl, "frequency": int(uni[lbl])}
            for lbl in sorted(labels_in_edges) if lbl in uni.index]
    return pd.DataFrame(rows)


def plot_heatmap_from_df(df, out_path, title="Heatmap",
                         matrix_type="cooccurrence", top_n=None):
    import matplotlib
    matplotlib.use("Agg")  # non-interactive backend — safe to call from any thread
    import matplotlib.pyplot as plt
    import seaborn as sns

    df_plot = df.copy().apply(pd.to_numeric, errors="coerce")
    if top_n is not None and top_n > 0 and len(df_plot) > top_n:
        row_scores    = df_plot.sum(axis=1, skipna=True)
        sorted_labels = row_scores.sort_values(ascending=False).index.tolist()
        df_plot       = df_plot.reindex(index=sorted_labels[:top_n],
                                        columns=sorted_labels[:top_n])
        title = f"{title} (Top {top_n})"

    n         = len(df_plot)
    fig_size  = max(12, n * 0.4)
    font_size = max(5, min(8, int(200 / n)))
    annot_fmt = ".1f" if matrix_type == "pmi" else ".0f"

    fig, ax = plt.subplots(figsize=(fig_size, fig_size * 0.9))
    sns.heatmap(df_plot, ax=ax, cmap="YlOrRd",
                linewidths=0.3, linecolor="#eeeeee",
                annot=True, fmt=annot_fmt, annot_kws={"size": font_size},
                square=True, cbar_kws={"shrink": 0.8},
                mask=df_plot.isna())
    ax.set_title(title, fontsize=14, pad=15)
    ax.tick_params(axis="x", rotation=45, labelsize=font_size + 1)
    ax.tick_params(axis="y", rotation=0,  labelsize=font_size + 1)
    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()


def _unique_path(path):
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 1
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def process_csv(csv_path, out_dir, min_cooccurrence, min_pmi,
                matrix_type, heatmap_top_n=None):
    
    csv_path = Path(csv_path)
    out_dir  = Path(out_dir)

    df = load_labels(str(csv_path))
    co_mat, stats = build_cooccurrence(df)

    edges = compute_pmi_edges(stats,
                              min_cooccurrence=min_cooccurrence,
                              min_pmi=min_pmi)

    if edges.empty:
        raise ValueError(
            f"No label pairs passed the current filters.\n\n"
            f"Try lowering the minimum co-occurrence count (currently {min_cooccurrence}) "
            f"or the minimum PMI value (currently {min_pmi}).")

    labels_in_edges = sorted(
        set(edges["Source"].tolist()) | set(edges["Target"].tolist()))

    co_mat_filtered = co_mat.reindex(
        index=labels_in_edges, columns=labels_in_edges, fill_value=0).astype(int)
    co_mat_masked = co_mat_filtered.where(
        co_mat_filtered >= min_cooccurrence, other="")

    if matrix_type == "pmi":
        pmi_mat            = build_pmi_matrix(labels_in_edges, stats,
                                               co_mat_filtered,
                                               min_cooccurrence, min_pmi)
        co_path            = _unique_path(out_dir / "pmi_matrix.csv")
        pmi_mat.to_csv(co_path, encoding="utf-8-sig")
        matrix_df_for_plot = pmi_mat
        heat_title         = "PMI Heatmap"
    else:
        co_path            = _unique_path(out_dir / "cooccurrence_matrix.csv")
        co_mat_masked.to_csv(co_path, encoding="utf-8-sig")
        matrix_df_for_plot = co_mat_masked
        heat_title         = "Visual Label Co-occurrence Heatmap"

    edges_path = _unique_path(out_dir / "edges.csv")
    edges.to_csv(edges_path, index=False, encoding="utf-8-sig")

    nodes      = build_nodes_from_edges_and_stats(edges, stats)
    nodes_path = _unique_path(out_dir / "nodes.csv")
    nodes.to_csv(nodes_path, index=False, encoding="utf-8-sig")

    heatmap_path = None
    try:
        heatmap_path = _unique_path(out_dir / "heatmap.png")
        plot_heatmap_from_df(matrix_df_for_plot, heatmap_path,
                             title=heat_title, matrix_type=matrix_type,
                             top_n=heatmap_top_n)
    except ImportError:
        print("Heatmap skipped: install matplotlib and seaborn to enable.")
        heatmap_path = None
    except Exception as e:
        print(f"Heatmap error (CSV saved normally): {e}")
        heatmap_path = None

    result = {"cooccurrence_matrix": co_path,
              "edges": edges_path, "nodes": nodes_path}
    if heatmap_path and heatmap_path.exists():
        result["heatmap"] = heatmap_path
    return result


class Step2App:

    def __init__(self, root, content):
        self.root    = root
        self.content = content

        self._running = False

        self.v_csv     = tk.StringVar()
        self.v_min_co  = tk.StringVar(value="2")
        self.v_min_pmi = tk.StringVar(value="1.0")
        self.v_matrix  = tk.StringVar(value="cooccurrence")
        self.v_top_n   = tk.StringVar(value="")
        self.v_out_dir = tk.StringVar()

        title_bar = tk.Frame(self.content, bg=BLACK, pady=12)
        title_bar.pack(fill="x")
        tk.Label(title_bar,
                 text="Step 2   Filter Label Co-occurrence and Generate the Matrix, Heatmap, and Gephi Files",
                 font=FONT_TITLE, bg=BLACK, fg=WHITE).pack(padx=24, anchor="w")
        self._build_form()
        tk.Frame(self.content, bg=WHITE, height=8).pack()

    def _build_form(self):
        # Explanatory grey text
        text_label(self.content,
                   "Provide the deduplicated_labels.csv, configure filtering parameters, "
                   "and generate a matrix CSV, a heatmap, and Gephi-ready edge and node CSV files.")

        # Grid for all rows
        gf = tk.Frame(self.content, bg=WHITE)
        gf.pack(fill="x", padx=4)
        gf.columnconfigure(1, weight=1)

        # Input CSV
        make_path_row(gf, 0, "deduplicated_labels.csv:",
                      self.v_csv,
                      lambda: self._browse_open(self.v_csv))

        # Min co-occurrence
        tk.Label(gf, text="Minimum co-occurrence count for label pairs:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=1, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        tk.Entry(gf, textvariable=self.v_min_co, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).grid(
            row=1, column=1, sticky="w", pady=PAD_Y)

        # Min PMI
        tk.Label(gf, text="Minimum PMI value for label pairs:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=2, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        tk.Entry(gf, textvariable=self.v_min_pmi, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).grid(
            row=2, column=1, sticky="w", pady=PAD_Y)

        # Matrix type
        tk.Label(gf, text="Matrix type:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=3, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        radio_frame = tk.Frame(gf, bg=WHITE)
        radio_frame.grid(row=3, column=1, sticky="w", pady=PAD_Y)
        tk.Radiobutton(radio_frame, text="Co-occurrence matrix",
                       variable=self.v_matrix, value="cooccurrence",
                       font=FONT_BODY, bg=WHITE, activebackground=WHITE).pack(
            side="left", padx=(0, 20))
        tk.Radiobutton(radio_frame, text="PMI matrix",
                       variable=self.v_matrix, value="pmi",
                       font=FONT_BODY, bg=WHITE, activebackground=WHITE).pack(side="left")

        # Heatmap top-N
        tk.Label(gf, text="Number of label pairs displayed in the heatmap:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=4, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        tk.Entry(gf, textvariable=self.v_top_n, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).grid(
            row=4, column=1, sticky="w", pady=PAD_Y)

        # Hint under top-N
        tk.Label(gf,
                 text="Leave blank = all labels; enter N = top-N labels ranked by co-occurrence count or PMI",
                 font=FONT_SMALL, bg=WHITE, fg=GREY_TEXT,
                 wraplength=580, justify="left").grid(
            row=5, column=1, columnspan=2, sticky="w",
            padx=(PAD_X, 6), pady=(0, PAD_Y))

        # Output folder
        make_output_label(gf, 6,
                          " Save matrix.csv + heatmap.png + edges.csv + nodes.csv to:",
                          self.v_out_dir,
                          lambda: self._browse_folder(self.v_out_dir))

        # Start button + inline status label
        btn_row = tk.Frame(self.content, bg=WHITE)
        btn_row.pack(anchor="w", padx=PAD_X, pady=(14, 6))
        grey_btn(btn_row, "Start Analysis", self._run_analysis).pack(side="left")
        self._status_label = tk.Label(btn_row, text="", font=FONT_SMALL,
                                      bg=WHITE, fg=GREY_TEXT)
        self._status_label.pack(side="left", padx=(14, 0))

        # Red Gephi note
        tk.Label(self.content,
                 text="You can now import edges.csv and nodes.csv into Gephi "
                      "to generate the visual label co-occurrence network.",
                 font=FONT_BODY, bg=WHITE, fg=RED_WARN,
                 wraplength=840, justify="left").pack(
            anchor="w", padx=PAD_X, pady=(6, 4))

    def _browse_open(self, var):
        p = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p: var.set(p)

    def _browse_folder(self, var):
        p = filedialog.askdirectory(title="Select output folder")
        if p: var.set(p)

    def _run_analysis(self):
        if self._running:
            messagebox.showwarning("Busy", "Analysis is already running. Please wait.")
            return

        csv_path = self.v_csv.get().strip()
        out_dir  = self.v_out_dir.get().strip()

        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showerror("Error", "Please select a valid deduplicated_labels.csv file.")
            return
        if not out_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        try:
            min_co  = int(self.v_min_co.get().strip())
            min_pmi = float(self.v_min_pmi.get().strip())
        except ValueError:
            messagebox.showerror("Error",
                                 "Minimum co-occurrence must be an integer; "
                                 "minimum PMI must be a number.")
            return

        matrix_type = self.v_matrix.get().strip().lower()

        top_n_str = self.v_top_n.get().strip()
        top_n = None
        if top_n_str:
            try:
                top_n = int(top_n_str)
                if top_n < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error",
                                     "Number of label pairs in heatmap must be a positive integer.")
                return

        out_path = Path(out_dir)
        out_path.mkdir(parents=True, exist_ok=True)
        self._running = True
        self._status_label.config(text="Analysing, please wait...")

        def worker():
            err, result = None, None
            try:
                result = process_csv(csv_path, out_path,
                                     min_co, min_pmi, matrix_type, top_n)
            except Exception as exc:
                err = exc
                traceback.print_exc()

            def finish():
                self._status_label.config(text="")
                self._running = False
                if err:
                    messagebox.showerror("Error", f"Analysis failed:\n{err}")
                else:
                    matrix_label = "PMI matrix" if matrix_type == "pmi" else "Co-occurrence matrix"
                    msg = (f"Analysis complete.\n\n"
                           f"Output folder:\n{out_path}\n\n"
                           f"  {matrix_label}: {result['cooccurrence_matrix'].name}\n"
                           f"  Edge list:      {result['edges'].name}\n"
                           f"  Node list:      {result['nodes'].name}\n")
                    if "heatmap" in result:
                        msg += f"  Heatmap:        {result['heatmap'].name}\n"
                    else:
                        msg += "  Heatmap: not generated (install matplotlib and seaborn to enable)\n"
                    messagebox.showinfo("Done", msg)

            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _on_close(self):
        pass  # window lifecycle managed by MainApp


# ══════════════════════════════════════════════════════════
# Step 3 — Image Grid for Top-N Label Pairs
# ══════════════════════════════════════════════════════════

IMAGE_EXTS        = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
DEFAULT_BG        = "white"
DEFAULT_TEXT_COL  = "black"
DEFAULT_LINE      = (210, 210, 210)
DEFAULT_HEADER_BG = (245, 245, 245)


def _read_csv_fallback(csv_path: Path) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "latin1"):
        try:
            return pd.read_csv(csv_path, encoding=enc)
        except Exception as e:
            last_err = e
    raise ValueError(f"Cannot read CSV: {csv_path}\n{last_err}")


def _safe_float(x, default=0.0) -> float:
    try:
        return default if pd.isna(x) else float(x)
    except Exception:
        return default


def _safe_int(x, default=0) -> int:
    try:
        return default if pd.isna(x) else int(float(x))
    except Exception:
        return default


def _split_semi_str(s):
    if pd.isna(s):
        return []
    return [x.strip() for x in str(s).split(";") if x.strip()]


def _split_semi_float(s):
    if pd.isna(s):
        return []
    return [_safe_float(x.strip(), 0.0) for x in str(s).split(";") if x.strip()]


def _load_pil_font(size: int):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/seguiemj.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    for p in candidates:
        try:
            return ImageFont.truetype(p, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def _wrap_text(draw, text, font, max_width):
    if not text:
        return [""]
    words = text.split(" ")
    lines, current = [], words[0]
    for w in words[1:]:
        trial = current + " " + w
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2] - bbox[0] <= max_width:
            current = trial
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return lines


def _build_image_index(image_folder: Path) -> dict:
    index = {}
    for p in image_folder.rglob("*"):
        if p.is_file() and p.suffix.lower() in IMAGE_EXTS:
            index[p.name] = p
            index[p.stem] = p
    return index


def _resolve_image(image_id: str, index: dict):
    if not image_id:
        return None
    raw = str(image_id).strip()
    if raw in index:
        return index[raw]
    name = Path(raw).name
    if name in index:
        return index[name]
    stem = Path(raw).stem
    if stem in index:
        return index[stem]
    return None


def parse_grouped_csv(path: Path) -> dict:
    df = _read_csv_fallback(path)
    required = {"image_id", "label", "score"}
    if not required.issubset(df.columns):
        raise ValueError(
            f"deduplicated_labels.csv is missing required columns.\n"
            f"Required: {required}\nFound: {list(df.columns)}")
    result = {}
    for _, row in df.iterrows():
        iid    = str(row["image_id"]).strip()
        labels = _split_semi_str(row["label"])
        scores = _split_semi_float(row["score"])
        if not iid or not labels:
            continue
        if len(scores) < len(labels):
            scores += [0.0] * (len(labels) - len(scores))
        lmap = {}
        for lab, sc in zip(labels, scores):
            if lab not in lmap or sc > lmap[lab]:
                lmap[lab] = sc
        result[iid] = lmap
    return result


def parse_edge_csv(path: Path) -> pd.DataFrame:
    df = _read_csv_fallback(path)
    required = {"Source", "Target", "Weight", "cooccur_count"}
    if not required.issubset(df.columns):
        raise ValueError(
            f"edges.csv is missing required columns.\n"
            f"Required: {required}\nFound: {list(df.columns)}")
    df = df.copy()
    df["Source"]       = df["Source"].astype(str).str.strip()
    df["Target"]       = df["Target"].astype(str).str.strip()
    df["Weight"]       = df["Weight"].apply(lambda x: _safe_float(x, 0.0))
    df["cooccur_count"]= df["cooccur_count"].apply(lambda x: _safe_int(x, 0))
    return df[(df["Source"] != "") & (df["Target"] != "")]


def select_top_edges(edge_df: pd.DataFrame, sort_by: str, top_n: int) -> pd.DataFrame:
    secondary = "cooccur_count" if sort_by == "Weight" else "Weight"
    return (edge_df
            .sort_values([sort_by, secondary, "Source", "Target"],
                         ascending=[False, False, True, True])
            .head(top_n)
            .reset_index(drop=True))


def find_images_for_pair(source, target, grouped_data, image_index, max_images):
    matches = []
    for iid, lmap in grouped_data.items():
        if source in lmap and target in lmap:
            img_path = _resolve_image(iid, image_index)
            if img_path is None:
                continue
            score = _safe_float(lmap.get(source, 0)) + _safe_float(lmap.get(target, 0))
            matches.append((img_path, score))
    matches.sort(key=lambda x: (-x[1], x[0].name.lower()))
    return matches[:max_images]


def create_thumbnail(img_path: Path, thumb_w: int, thumb_h: int):
    try:
        img = Image.open(img_path).convert("RGB")
    except Exception:
        img = Image.new("RGB", (thumb_w, thumb_h), (240, 240, 240))
        d = ImageDraw.Draw(img)
        d.text((10, 10), "Load\nfailed", fill="black", font=_load_pil_font(14))
        return img
    canvas = Image.new("RGB", (thumb_w, thumb_h), "white")
    img.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    canvas.paste(img, ((thumb_w - img.width) // 2, (thumb_h - img.height) // 2))
    return canvas


def generate_grid_png(selected_edges, grouped_data, image_folder: Path,
                      output_path: Path, images_per_pair: int,
                      thumb_w: int, thumb_h: int,
                      canvas_w: int, canvas_h: int) -> Path:
    image_index = _build_image_index(image_folder)

    rows = []
    for _, row in selected_edges.iterrows():
        source  = row["Source"]
        target  = row["Target"]
        pmi     = _safe_float(row["Weight"], 0.0)
        cooccur = _safe_int(row["cooccur_count"], 0)
        imgs = find_images_for_pair(source, target, grouped_data,
                                    image_index, images_per_pair)
        rows.append({"source": source, "target": target,
                     "pmi": pmi, "cooccur": cooccur, "images": imgs})

    margin       = 24
    header_h     = 44
    row_gap      = 12
    col_gap      = 12
    pair_col_w   = 300
    pmi_col_w    = 110
    co_col_w     = 130
    images_col_w = max(200, images_per_pair * thumb_w + (images_per_pair - 1) * 8 + 20)

    min_w = margin * 2 + pair_col_w + col_gap + pmi_col_w + col_gap + co_col_w + col_gap + images_col_w
    if canvas_w < min_w:
        canvas_w = min_w

    row_h  = max(72, thumb_h + 20)
    auto_h = margin * 2 + header_h + len(rows) * row_h + max(0, len(rows) - 1) * row_gap
    if canvas_h <= 0:
        canvas_h = auto_h
    elif canvas_h < auto_h:
        canvas_h = auto_h

    img  = Image.new("RGB", (canvas_w, canvas_h), DEFAULT_BG)
    draw = ImageDraw.Draw(img)

    font_header = _load_pil_font(20)
    font_body   = _load_pil_font(18)
    font_small  = _load_pil_font(16)

    draw.rectangle([0, 0, canvas_w, margin + header_h], fill=DEFAULT_HEADER_BG)

    x_pair = margin
    x_pmi  = x_pair + pair_col_w + col_gap
    x_co   = x_pmi  + pmi_col_w  + col_gap
    x_imgs = x_co   + co_col_w   + col_gap

    draw.text((x_pair, margin), "Label Pair",    fill=DEFAULT_TEXT_COL, font=font_header)
    draw.text((x_pmi,  margin), "PMI",           fill=DEFAULT_TEXT_COL, font=font_header)
    draw.text((x_co,   margin), "Co-occurrence", fill=DEFAULT_TEXT_COL, font=font_header)
    draw.text((x_imgs, margin), "Images",        fill=DEFAULT_TEXT_COL, font=font_header)

    y_line = margin + header_h
    draw.line((margin, y_line, canvas_w - margin, y_line), fill=DEFAULT_LINE, width=2)

    current_y = y_line + 12
    for row in rows:
        row_top    = current_y
        row_bottom = row_top + row_h
        draw.line((margin, row_bottom, canvas_w - margin, row_bottom),
                  fill=DEFAULT_LINE, width=1)

        pair_text  = f"{row['source']} — {row['target']}"
        pair_lines = _wrap_text(draw, pair_text, font_body, pair_col_w - 10)
        text_y = row_top + 8
        for line in pair_lines[:3]:
            draw.text((x_pair, text_y), line, fill=DEFAULT_TEXT_COL, font=font_body)
            text_y += 24

        draw.text((x_pmi, row_top + 12), f"{row['pmi']:.4f}",
                  fill=DEFAULT_TEXT_COL, font=font_body)
        draw.text((x_co,  row_top + 12), str(row["cooccur"]),
                  fill=DEFAULT_TEXT_COL, font=font_body)

        if row["images"]:
            thumb_x = x_imgs
            thumb_y = row_top + 6
            for img_path, _ in row["images"]:
                thumb = create_thumbnail(img_path, thumb_w, thumb_h)
                img.paste(thumb, (thumb_x, thumb_y))
                draw.rectangle([thumb_x, thumb_y,
                                 thumb_x + thumb_w, thumb_y + thumb_h],
                                outline=(180, 180, 180), width=1)
                thumb_x += thumb_w + 8
        else:
            draw.text((x_imgs, row_top + 12), "No matching image found",
                      fill=(120, 120, 120), font=font_small)

        current_y = row_bottom + row_gap

    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(output_path, "PNG")
    return output_path


class Step3App:

    def __init__(self, root, content):
        self.root    = root
        self.content = content
        self._running = False

        self.v_filtered_csv    = tk.StringVar()
        self.v_edge_csv        = tk.StringVar()
        self.v_image_folder    = tk.StringVar()
        self.v_sort_by         = tk.StringVar(value="Weight")
        self.v_top_n           = tk.StringVar(value="10")
        self.v_images_per_pair = tk.StringVar(value="3")
        self.v_thumb_w         = tk.StringVar(value="120")
        self.v_thumb_h         = tk.StringVar(value="120")
        self.v_canvas_w        = tk.StringVar(value="1600")
        self.v_canvas_h        = tk.StringVar(value="0")
        self.v_output_png      = tk.StringVar()

        title_bar = tk.Frame(self.content, bg=BLACK, pady=12)
        title_bar.pack(fill="x")
        tk.Label(title_bar,
                 text="Step 3   Image Grid for Top-N Label Pairs",
                 font=FONT_TITLE, bg=BLACK, fg=WHITE).pack(padx=24, anchor="w")

        self._build_form()
        tk.Frame(self.content, bg=WHITE, height=24).pack()

    def _build_form(self):
        f = tk.Frame(self.content, bg=WHITE)
        f.pack(fill="x")

        # Explanatory grey text
        text_label(f,
                   "Provide the deduplicated label CSV, edge CSV (generated in Steps 1 and 2) "
                   "and image folder. After selecting the label-pair filtering settings, "
                   "an image grid for the top-N label pairs will be generated.")

        # Grid for all rows
        gf = tk.Frame(f, bg=WHITE)
        gf.pack(fill="x", padx=4)
        gf.columnconfigure(1, weight=1)

        # Input files
        make_path_row(gf, 0, "deduplicated_labels.csv:",
                      self.v_filtered_csv,
                      lambda: self._browse_open(self.v_filtered_csv))
        make_path_row(gf, 1, "edges.csv:",
                      self.v_edge_csv,
                      lambda: self._browse_open(self.v_edge_csv))
        make_path_row(gf, 2, "Image folder:",
                      self.v_image_folder,
                      lambda: self._browse_folder(self.v_image_folder))

        # Ranking method
        tk.Label(gf, text="Top label pair ranking method:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=3, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        radio_f = tk.Frame(gf, bg=WHITE)
        radio_f.grid(row=3, column=1, sticky="w", pady=PAD_Y)
        tk.Radiobutton(radio_f, text="By PMI",
                       variable=self.v_sort_by, value="Weight",
                       font=FONT_BODY, bg=WHITE, activebackground=WHITE).pack(
            side="left", padx=(0, 20))
        tk.Radiobutton(radio_f, text="By co-occurrence count",
                       variable=self.v_sort_by, value="cooccur_count",
                       font=FONT_BODY, bg=WHITE, activebackground=WHITE).pack(side="left")

        # Number of top label pairs
        tk.Label(gf, text="Number of top label pairs:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=4, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        tk.Entry(gf, textvariable=self.v_top_n, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).grid(
            row=4, column=1, sticky="w", pady=PAD_Y)

        # Images per label pair + hint
        tk.Label(gf, text="Images per label pair:",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=5, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        ipp_f = tk.Frame(gf, bg=WHITE)
        ipp_f.grid(row=5, column=1, sticky="w", pady=PAD_Y)
        tk.Entry(ipp_f, textvariable=self.v_images_per_pair, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).pack(side="left")
        tk.Label(ipp_f,
                 text="   Limit the number of images shown per label pair",
                 font=FONT_SMALL, bg=WHITE, fg=GREY_TEXT).pack(side="left")

        # Thumbnail size
        tk.Label(gf, text="Thumbnail size (W × H):",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=6, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        thumb_f = tk.Frame(gf, bg=WHITE)
        thumb_f.grid(row=6, column=1, sticky="w", pady=PAD_Y)
        tk.Entry(thumb_f, textvariable=self.v_thumb_w, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).pack(side="left")
        tk.Label(thumb_f, text=" × ", font=FONT_LABEL,
                 bg=WHITE, fg=BLACK).pack(side="left")
        tk.Entry(thumb_f, textvariable=self.v_thumb_h, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).pack(side="left")

        # Canvas size
        tk.Label(gf, text="Canvas size (W × H):",
                 font=FONT_LABEL, bg=WHITE, fg=BLACK, anchor="e").grid(
            row=7, column=0, sticky="e", padx=(PAD_X, 6), pady=PAD_Y)
        canvas_f = tk.Frame(gf, bg=WHITE)
        canvas_f.grid(row=7, column=1, sticky="w", pady=PAD_Y)
        tk.Entry(canvas_f, textvariable=self.v_canvas_w, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).pack(side="left")
        tk.Label(canvas_f, text=" × ", font=FONT_LABEL,
                 bg=WHITE, fg=BLACK).pack(side="left")
        tk.Entry(canvas_f, textvariable=self.v_canvas_h, font=FONT_LABEL,
                 bg=BLUE_PATH, fg=BLACK, relief="flat", bd=4, width=10).pack(side="left")
        tk.Label(canvas_f,
                 text="   Set height to 0 for automatic adjustment",
                 font=FONT_SMALL, bg=WHITE, fg=GREY_TEXT).pack(side="left")

        # Output path
        make_output_label(gf, 8, " Save image grid to:",
                          self.v_output_png,
                          lambda: self._browse_save_png(self.v_output_png))

        # Button + inline status label
        btn_row = tk.Frame(f, bg=WHITE)
        btn_row.pack(anchor="w", padx=PAD_X, pady=(14, 8))
        grey_btn(btn_row, "Generate Image Grid", self._run).pack(side="left")
        self._status_label = tk.Label(btn_row, text="", font=FONT_SMALL,
                                      bg=WHITE, fg=GREY_TEXT)
        self._status_label.pack(side="left", padx=(14, 0))

    # -- File dialogs --
    def _browse_open(self, var):
        p = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p: var.set(p)

    def _browse_folder(self, var):
        p = filedialog.askdirectory()
        if p: var.set(p)

    def _browse_save_png(self, var):
        p = filedialog.asksaveasfilename(
            initialfile="image_grid.png",
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
        if p: var.set(p)

    # Validate & run
    def _run(self):
        if self._running:
            messagebox.showwarning("Busy", "Generation is already running. Please wait.")
            return

        try:
            filtered_csv  = Path(self.v_filtered_csv.get().strip())
            edge_csv      = Path(self.v_edge_csv.get().strip())
            image_folder  = Path(self.v_image_folder.get().strip())
            output_png    = Path(self.v_output_png.get().strip())
            sort_by       = self.v_sort_by.get().strip()
            top_n         = int(self.v_top_n.get().strip())
            ipp           = int(self.v_images_per_pair.get().strip())
            thumb_w       = int(self.v_thumb_w.get().strip())
            thumb_h       = int(self.v_thumb_h.get().strip())
            canvas_w      = int(self.v_canvas_w.get().strip())
            canvas_h      = int(self.v_canvas_h.get().strip())
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid parameter value:\n{e}")
            return

        if not filtered_csv.exists():
            messagebox.showerror("Error", f"File not found:\n{filtered_csv}")
            return
        if not edge_csv.exists():
            messagebox.showerror("Error", f"File not found:\n{edge_csv}")
            return
        if not image_folder.is_dir():
            messagebox.showerror("Error", f"Folder not found:\n{image_folder}")
            return
        if not self.v_output_png.get().strip():
            messagebox.showerror("Error", "Please specify an output PNG path.")
            return
        if top_n <= 0 or ipp <= 0 or thumb_w <= 0 or thumb_h <= 0 or canvas_w <= 0:
            messagebox.showerror("Error", "All numeric values must be greater than 0.")
            return

        self._running = True
        self._status_label.config(text="Generating image grid, please wait...")

        def worker():
            err, result_path = None, None
            try:
                grouped_data   = parse_grouped_csv(filtered_csv)
                edge_df        = parse_edge_csv(edge_csv)
                selected_edges = select_top_edges(edge_df, sort_by, top_n)
                result_path    = generate_grid_png(
                    selected_edges, grouped_data, image_folder, output_png,
                    ipp, thumb_w, thumb_h, canvas_w, canvas_h)
            except Exception as exc:
                err = exc
                traceback.print_exc()

            def finish():
                self._status_label.config(text="")
                self._running = False
                if err:
                    messagebox.showerror("Error", f"Generation failed:\n{err}")
                else:
                    messagebox.showinfo("Done",
                                        f"Image grid generated successfully.\n\n"
                                        f"Saved to:\n{result_path}")

            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()


# ══════════════════════════════════════════════════════════
# Main window： single scrollable page, three sections
# ══════════════════════════════════════════════════════════

class MainApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Visual Label Co-occurrence Analysis Toolkit")
        self.root.configure(bg=WHITE)
        self.root.geometry("960x900")
        self.root.resizable(True, True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # One shared scrollable canvas for all three steps
        content = make_scrollable(self.root)

        Step1App(self.root, content)
        Step2App(self.root, content)
        Step3App(self.root, content)

        tk.Frame(content, bg=WHITE, height=40).pack()

    def _on_close(self):
        self.root.destroy()


def main():
    root = tk.Tk()
    MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
