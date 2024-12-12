from docx import Document
import os
from datetime import datetime

def read_word_titles(word_file_path, output_file):
    try:
        # 讀取Word文件
        doc = Document(word_file_path)
        
        # 獲取檔名（不含副檔名）
        file_name = os.path.splitext(os.path.basename(word_file_path))[0]
        
        # 儲存標題的列表
        titles = []
        
        # 遍歷段落
        for paragraph in doc.paragraphs:
            # 檢查段落是否為標題樣式
            if paragraph.style.name.startswith('Heading'):
                titles.append(paragraph.text)
        
        # 追加寫入結果到txt檔案
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f'\n{"="*50}\n')  # 分隔線
            f.write(f'檔案名稱：{file_name}\n')
            f.write('\n標題列表：\n')
            for i, title in enumerate(titles, 1):
                f.write(f'{i}. {title}\n')
        
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
    
    # 計定輸出檔案名稱
    output_file = 'all_word_titles.txt'
    
    # 創建新的輸出檔案（如果已存在則清空）
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('Word文件標題彙整\n')
        f.write(f'處理時間：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
    
    # 計數器
    success_count = 0
    fail_count = 0
    
    # 遍歷目錄及其子目錄中的所有檔案
    for root, dirs, files in os.walk(directory_path):
        for filename in files:
            # 檢查是否為Word檔案且檔名包含"一股週報"
            if (filename.endswith(('.docx', '.doc')) and 
                '一股週報' in filename and 
                not filename.startswith('~$')):  # 排除暫存檔
                
                file_path = os.path.join(root, filename)
                if read_word_titles(file_path, output_file):
                    success_count += 1
                else:
                    fail_count += 1
    
    # 追加寫入處理結果統計
    with open(output_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{"="*50}\n')
        f.write('\n處理結果統計：\n')
        f.write(f'成功處理: {success_count} 個檔案\n')
        f.write(f'處理失敗: {fail_count} 個檔案\n')
    
    # 輸出處理結果統計
    print('\n處理完成！')
    print(f'成功處理: {success_count} 個檔案')
    print(f'處理失敗: {fail_count} 個檔案')
    print(f'結果已寫入：{output_file}')

# 使用範例
if __name__ == '__main__':
    # 替換成您的目錄路徑
    directory = 'word_files'  # 例如: 'C:/Documents/word_files'
    process_directory(directory) 