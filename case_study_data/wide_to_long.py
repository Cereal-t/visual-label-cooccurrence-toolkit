import pandas as pd

# ── Paths ──────────────────────────────────────────────────
INPUT_PATH  = ""
OUTPUT_PATH = ""
# ──────────────────────────────────────────────────────────

# Read wide-format CSV (image_id | label | score, semicolon-separated within each cell)
df = pd.read_csv(INPUT_PATH)

# Convert to long format: one row per label
rows = []
for _, row in df.iterrows():
    labels = str(row["label"]).split(";")
    scores = str(row["score"]).split(";")
    for label, score in zip(labels, scores):
        rows.append({
            "image_id": row["image_id"],
            "label":    label.strip(),
            "score":    score.strip()
        })

long_df = pd.DataFrame(rows, columns=["image_id", "label", "score"])
long_df.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

print(f"Done. {len(long_df)} rows written to: {OUTPUT_PATH}")
