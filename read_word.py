from docx import Document
import os
from datetime import datetime

def is_chinese_number(text):
    # 檢查文字是否為國字數字1到10
    chinese_numbers = {'重點工作事項'}
    return text in chinese_numbers

def read_first_page_contents(word_file_path, output_file):
    try:
        # 讀取Word文件
        doc = Document(word_file_path)
        
        # 獲取檔名（不含副檔名）
        file_name = os.path.splitext(os.path.basename(word_file_path))[0]
        
        # 儲存第一頁內容
        contents = []
        read_until_reason = True  # 標記是否繼續讀取
        # 只讀取第一頁的段落
        for paragraph in doc.paragraphs:
            # 檢查段落是否為分頁符號
            if paragraph._element.getparent().tag.endswith('sectPr'):
                break  # 遇到分頁符號則停止
            
            # 檢查段落是否為超連結
            if paragraph.style.name == 'Hyperlink':
                continue  # 跳過超連結段落

            text = paragraph.text.strip()

             # 跳過特定段落
            if "緣由：" in text:
                read_until_reason = False
                continue

            if read_until_reason and text :  # 只記錄國字數字1到10的段落
                contents.append(text)
                print(f"讀取到的段落: {text}")  # 除錯訊息
        
        # 追加寫入結果到txt檔案
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f'檔案名稱：{file_name}\n')
            for line in contents:
                f.write(f'{line}\n')
            f.write(f'{"="*50}\n')  # 分隔線
        
        print(f'已成功處理 {file_name}')
        return True
    except Exception as e:
        print(f'處理 {word_file_path} 時發生錯誤: {str(e)}')
        return False

def process_directory(directory_path):
    # 確保目錄存在
    if not os.path.exists(directory_path):
        print(f'目錄 {directory_path} 不存在！')
        return
    
    # 設定輸出檔案名稱
    output_file = 'Word文件第一頁內容彙整.txt'
    
    # 創建新的輸出檔案（如果已存在則清空）
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('Word文件第一頁內容彙整\n')
        f.write(f'處理時間：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
        f.write(f'{"="*50}\n')
    
    # 計數器
    success_count = 0
    fail_count = 0
    
    # 遍歷目錄及其子目錄中的所有檔案
    for root, dirs, files in os.walk(directory_path):
        for filename in files:
            if (filename.endswith(('.docx', '.doc')) and 
                '一股週報' in filename and 
                not filename.startswith('~$')):
                
                file_path = os.path.join(root, filename)
                if read_first_page_contents(file_path, output_file):
                    success_count += 1
                else:
                    fail_count += 1
    
    # 追加寫入處理結果統計
    with open(output_file, 'a', encoding='utf-8') as f:
        f.write('\n處理結果統計：\n')
        f.write(f'成功處理：{success_count} 個檔案\n')
        f.write(f'處理失敗：{fail_count} 個檔案\n')

# 使用範例
if __name__ == '__main__':
    # 替換成您的目錄路徑
    directory = 'word_files'  # 例如: 'C:/Documents/word_files'
    process_directory(directory)
