import os
import re

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

def create_text_file(input_file, output_file):
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
        #(r'^\s+(?=(?:.+\n)[\s　])', r''),

        #例えば1行目と3行目は頭に空白がなく、2行目だけある場合（頭に空白がない行に、空白がある行が挟まれている）など削除
        #(r'^[\s　]+(?=\S+\n+)', r''),

        #V14, 1削除、前半のタイムコードの後ろに　N　セリフ、改行2回して後半のタイムコードの後ろにON
        #(r'(\d{4})\s-\s(\d{4})\n(V\d{1,2},\s\d)\n((?:.+(?:\n|))*)', r'\1　　N　　\4\n\n\2　　ON\n'),

        #セリフの2行目以降の頭を1行目にそろえる
        #(r'(^(?!.*\d{4}(?: |　)*(?:N|ON)(?: |　)*.*).+$)', r'　　　　　　　　　\1'),
        #(r"(^(?!(\d{4}[\s　]{2})).+$)",r"　　　　　　　　　\1")
    ]
    content = process_text(content, replacements)
    content = remove_first_duplicate_line(content)
    content = normalize_blank_lines(content)
    content = half_to_full_width(content)

    with open(output_file, "w", encoding="utf-8") as output:
        output.write(content)

    print("Text file has been created successfully!")

if __name__ == "__main__":
    input_file = r"E:\sea01\Documents\Programming\python\vivia\narration\v2\txt\new2.txt"
    output_file = r"txt/test15.txt"
    create_text_file(input_file, output_file)