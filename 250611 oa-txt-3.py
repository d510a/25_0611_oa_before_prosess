# ===== XML 分割 + 拒絶理由通知取込み + まとめ作成 =====
import os, re, sys, logging
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# 拒絶理由通知・請求項ファイルの解析用ライブラリ
try:
    import docx2txt                 # DOCX → 文字列
except ImportError:
    docx2txt = None
try:
    import PyPDF2                   # PDF → 文字列
except ImportError:
    PyPDF2 = None

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s ─ %(message)s")

# ------------------------------
# 0) 共通：GUI 非表示 + ルート生成
# ------------------------------
root_tk = tk.Tk()
root_tk.withdraw()

# ------------------------------
# 1) XML ファイルの選択
# ------------------------------
input_file = filedialog.askopenfilename(
    title="入力する XML ファイルを選択してください",
    filetypes=[("XMLファイル", "*.xml")]
)
if not input_file:
    sys.exit()

# --- 追加／変更箇所 ---------------------------------
# XML と同じフォルダーを出力先に設定
output_folder = os.path.dirname(input_file)
# ----------------------------------------------------

# ------------------------------
# 2) XML ファイル読込みと文書候補抽出
# ------------------------------
tree         = ET.parse(input_file)
root_xml     = tree.getroot()
docs         = root_xml.findall('.//doc')
doc_list     = []
for doc in docs:
    id_elem = doc.find('.//str[@name="公開(公告)番号"]')
    if id_elem is not None and id_elem.text:
        doc_list.append((doc, id_elem.text.strip()))

if not doc_list:
    messagebox.showinfo("情報", "有効な文書が見つかりませんでした。")
    sys.exit()

total_docs          = len(doc_list)
seq_to_identifier   = {}
list_message        = "【文書一覧】\n"
for seq, (_, ident) in enumerate(doc_list, start=1):
    seq_to_identifier[seq] = ident
    list_message += f"{seq}: {ident}\n"

# ------------------------------
# 3) 役割（本願・引例）指定
# ------------------------------
used_seq = set()
# 本願
while True:
    main_seq = simpledialog.askinteger("役割指定",
                                       f"{list_message}\n本願は何番ですか？")
    if main_seq is None: sys.exit()
    if 1 <= main_seq <= total_docs:
        used_seq.add(main_seq)
        break
    messagebox.showerror("エラー", f"1～{total_docs} の範囲で入力してください。")

# 引例（複数可）
citation_mapping = {}
for i in range(1, total_docs):
    while True:
        citation_seq = simpledialog.askinteger(
            "役割指定",
            f"{list_message}\n引例{i}は何番ですか？（終了する場合は 0）"
        )
        if citation_seq is None: sys.exit()
        if citation_seq == 0:                  # 0 で入力終了
            break
        if not (1 <= citation_seq <= total_docs):
            messagebox.showerror("エラー", f"1～{total_docs} の範囲で入力してください。")
        elif citation_seq in used_seq:
            messagebox.showerror("エラー", "既に指定済みの番号です。")
        else:
            citation_mapping[i] = citation_seq
            used_seq.add(citation_seq)
            break
    if citation_seq == 0:
        break

# ------------------------------
# 4) XML → 個別テキスト化 & 本願/引例本文保存
# ------------------------------
role_assignment   = {}   # seq → ファイル名
main_text         = ""   # 本願全文
citation_texts    = {}   # {引例順: 本文}

for seq, (doc, identifier) in enumerate(doc_list, start=1):
    # 不要タグ削除
    for tag in doc.findall('.//uuid'):
        doc.remove(tag)
    for tag in doc.findall('.//str[@name="公開(公告)番号"]'):
        doc.remove(tag)

    # プレーンテキスト抽出
    text = ''.join(doc.itertext()).strip()

    # 文字数制限（日本語/中文＝104k、それ以外＝344k）
    if re.search(r'[\u3040-\u30ff\u4e00-\u9fff]', text):
        text = text[:104_000]
    else:
        text = text[:344_000]

    # 役割に応じたファイル名
    if seq == main_seq:
        file_name = f"h_{seq}_{identifier.replace(':','_')}.txt"
        main_text = text
    elif seq in citation_mapping.values():
        idx = list(citation_mapping.keys())[list(citation_mapping.values()).index(seq)]
        file_name = f"d{idx}_{seq}_{identifier.replace(':','_')}.txt"
        citation_texts[idx] = text
    else:
        file_name = f"{seq}_{identifier.replace(':','_')}.txt"

    role_assignment[seq] = file_name
    out_path = os.path.join(output_folder, file_name)
    with open(out_path, "w", encoding="utf-8") as fw:
        fw.write(text)

# ------------------------------
# 5) 拒絶理由通知ファイルを選択 & テキスト抽出
# ------------------------------
rej_file = filedialog.askopenfilename(
    title="拒絶理由通知ファイル（txt/docx/pdf）を選択してください",
    filetypes=[("通知ファイル", "*.txt *.docx *.pdf")]
)
if not rej_file:
    sys.exit()

def extract_rejection(path: str) -> str:
    """TXT / DOCX / PDF からプレーンテキストを抽出"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as fr:
            return fr.read()
    elif ext == ".docx":
        if docx2txt is None:
            messagebox.showerror("エラー", "docx2txt が未インストールです。pip install docx2txt を実行してください。")
            sys.exit()
        return docx2txt.process(path) or ""
    elif ext == ".pdf":
        if PyPDF2 is None:
            messagebox.showerror("エラー", "PyPDF2 が未インストールです。pip install PyPDF2 を実行してください。")
            sys.exit()
        with open(path, "rb") as frb:
            reader = PyPDF2.PdfReader(frb)
            return "\n".join(page.extract_text() or "" for page in reader.pages)
    else:
        messagebox.showerror("エラー", f"未対応拡張子: {ext}")
        sys.exit()

rejection_text = extract_rejection(rej_file)

# ------------------------------
# 5.5) ★追加：最新請求項の有無を確認
# ------------------------------
latest_claims_text = ""
if messagebox.askyesno("最新請求項の確認", "最新の請求項ファイルがありますか？"):
    latest_file = filedialog.askopenfilename(
        title="最新の請求項ファイル（txt/docx/pdf）を選択してください",
        filetypes=[("請求項ファイル", "*.txt *.docx *.pdf")]
    )
    if latest_file:
        latest_claims_text = extract_rejection(latest_file)

# ------------------------------
# 6) 「まとめ.txt」を作成
# ------------------------------
header = (
    "###役割\n"
    "-あなたは、自動車技術全般に精通した技術者であり、特許法にも精通した弁理士です。\n\n"
    "###指示\n"
    "-本願を読み込んでください。\n"
    "-拒絶理由通知を読み込んでください。\n"
    "-引例を読み込んでください。\n"
    "-拒絶理由通知から審査官の意図を把握してください。\n"
    "-本願の請求項を構成ごとに分割し、引例の全文と比較して対比表を作成してください。\n"
    "-本願と引例の差異を明確化し、応答案（意見書、補正書）を提案してください。\n"
    "-最新の請求項がある場合には、最新の請求項を本願の請求項に代えて利用ください。\n"
    "-最新の請求項が無い場合には、本願の請求項を利用ください。\n\n"
)

summary_lines = [
    header,
    "###本願\n",
    main_text,
    "\n###拒絶理由通知\n",
    rejection_text
]

if latest_claims_text:
    summary_lines.extend(["\n###最新の請求項\n", latest_claims_text])

for idx in sorted(citation_texts.keys()):
    summary_lines.extend([f"\n###引例{idx}\n", citation_texts[idx]])

summary_path = os.path.join(os.path.dirname(input_file), "まとめ.txt")
with open(summary_path, "w", encoding="utf-8") as fw:
    fw.write("".join(summary_lines))

messagebox.showinfo("完了", f"まとめ.txt を作成しました：\n{summary_path}")
print(f"[INFO] まとめ.txt を生成：{summary_path}")
