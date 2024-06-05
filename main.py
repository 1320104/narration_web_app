import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

def half_to_full_width(text):
    trans_table = str.maketrans(
        '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.',
        '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ．'
    )
    return text.translate(trans_table)

def process_text(content, replacements):
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content, flags=re.MULTILINE)
    return content

def remove_first_duplicate_line(content):
    lines = content.splitlines()
    pattern = r'\b\d{4}\b'
    all_numbers = re.findall(pattern, content)
    count_numbers = {num: all_numbers.count(num) for num in set(all_numbers)}
    first_occurrences = {num: False for num in count_numbers if count_numbers[num] > 1}
    new_lines = []
    for line in lines:
        numbers_in_line = re.findall(pattern, line)
        if any(num in first_occurrences and not first_occurrences[num] for num in numbers_in_line):
            for num in numbers_in_line:
                if num in first_occurrences and not first_occurrences[num]:
                    first_occurrences[num] = True
                    break
            continue
        new_lines.append(line)
    return "\n".join(new_lines)

def normalize_blank_lines(content):
    new_lines = []
    previous_line_was_blank = False
    for line in content.split('\n'):
        if line.strip():
            new_lines.append(line)
            previous_line_was_blank = False
        elif not previous_line_was_blank:
            new_lines.append(line)
            previous_line_was_blank = True
    return "\n".join(new_lines)

def create_document(input_file, output_file, template_file):
    with open(input_file, "r", encoding="utf-8") as file:
        content = file.read()

    replacements = [
        #空白の無駄なテキストレイヤー削除
        (r'V\d+, \d+\n{2,}', r''),

        #00;00;47;08 - 00;00;51;03　→　0047 - 0051
        (r'(\d{2});(\d{2});(\d{2});(\d{2})', r'\2\3'),

        #もともとセリフの前にNが入っていた場合、セリフまでの空白含めて削除
        (r'(^(?:Ｎ|N|N)[\s　]+)(?=.+\n)', r''),

        #2行目以降のセリフの頭に空白があった場合削除
        (r'^\s+(?=(?:.+\n)[\s　])', r''),

        #例えば1行目と3行目は頭に空白がなく、2行目だけある場合（頭に空白がない行に、空白がある行が挟まれている）など削除
        (r'^[\s　]+(?=\S+\n+)', r''),

        #V14, 1削除、前半のタイムコードの後ろに　N　セリフ、改行2回して後半のタイムコードの後ろにON
        (r'(\d{4})\s-\s(\d{4})\n(V\d{1,2},\s\d)\n((?:.+(?:\n|))*)', r'\1　　N　　\4\n\n\2　　ON\n'),

        #セリフの2行目以降の頭を1行目にそろえる
        (r'(^(?!.*\d{4}(?: |　)*(?:N|ON)(?: |　)*.*).+$)', r'　　　　　　　　　\1'),
    ]
    content = process_text(content, replacements)
    content = remove_first_duplicate_line(content)#Nのクリップが連続する場合の処理
    content = normalize_blank_lines(content)#remove_first_duplicate_lineの処理でできた空行の処理（すべて間を1行にする）
    content = half_to_full_width(content)#wordで文字が縦書きになるように全角に置換


    doc = Document(template_file)
    doc.add_paragraph(content)


    # 最初のパラグラフが空白の場合、削除する
    if not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)


    regex = re.compile(r'[０-９]{4}　　ＯＮ')
    for paragraph in doc.paragraphs:
        original_text = paragraph.text
        if regex.search(original_text):
            paragraph.clear()
            last_end = 0
            for match in regex.finditer(original_text):
                paragraph.add_run(original_text[last_end:match.start()])
                highlighted_run = paragraph.add_run(match.group())
                highlighted_run.font.highlight_color = WD_COLOR_INDEX.GRAY_50
                last_end = match.end()
            paragraph.add_run(original_text[last_end:])

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Hiragino Maru Gothic Pro'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Hiragino Maru Gothic Pro')
            run.font.size = Pt(10.5)

    doc.save(output_file)
    messagebox.showinfo("Success", "Document has been created successfully!")

def center_window(root, width=600, height=400):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f'{width}x{height}+{x}+{y}')

def setup_gui():
    root = tk.Tk()
    root.title("ナレーション原稿作成アプリ")
    center_window(root, 600, 300)

    input_file = tk.StringVar()
    output_file = tk.StringVar()

    # 実行ファイルのディレクトリパスを取得
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    # テンプレートファイルのパスを動的に設定
    template_file = os.path.join(base_path, '../temp.docx')

    def select_input_file():
        path = filedialog.askopenfilename(title="Select text file", filetypes=[("Text files", "*.txt")])
        if path:
            input_file.set(path)

    def select_output_file():
        path = filedialog.asksaveasfilename(title="Save the Word file", defaultextension=".docx")
        if path:
            output_file.set(path)

    def generate_word_file():
        if not input_file.get() or not output_file.get():
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        create_document(input_file.get(), output_file.get(), template_file)
        input_file.set("")
        output_file.set("")
    frame = tk.Frame(root)
    frame.place(relx=0.5, rely=0.5, anchor='center')

    tk.Button(frame, text="1. テキストファイルを選択(.txt)", font=("Hiragino Maru Gothic Pro", "20"), command=select_input_file).grid(row=0, column=0, pady=5)
    tk.Label(frame, textvariable=input_file).grid(row=1, column=0, sticky="ew", pady=5)
    tk.Button(frame, text="2. 保存場所・ファイル名(.docx)を指定", font=("Hiragino Maru Gothic Pro", "20"), command=select_output_file).grid(row=2, column=0, pady=5)
    tk.Label(frame, textvariable=output_file).grid(row=3, column=0, sticky="ew", pady=5)
    tk.Button(frame, text="3. Wordファイルを生成", font=("Hiragino Maru Gothic Pro", "20"), command=generate_word_file).grid(row=4, column=0, pady=5)

    root.mainloop()

if __name__ == "__main__":
    setup_gui()
