from docx import Document
import os

def read_word_titles(word_file_path):
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
        
        # 寫入結果到txt檔案
        output_file = f'titles_{file_name}.txt'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f'檔案名稱：{file_name}\n')
            f.write('\n標題列表：\n')
            for i, title in enumerate(titles, 1):
                f.write(f'{i}. {title}\n')
        
        print(f'已成功處理 {file_name} 並寫入 {output_file}')
        return True
    except Exception as e:
        print(f'處理 {word_file_path} 時發生錯誤: {str(e)}')
        return False

def process_directory(directory_path):
    # 確保目錄存在
    if not os.path.exists(directory_path):
        print(f'目錄 {directory_path} 不存在！')
        return
    
    # 計數器
    success_count = 0
    fail_count = 0
    
    # 遍歷目錄中的所有檔案
    for filename in os.listdir(directory_path):
        if filename.endswith(('.docx', '.doc')) and not filename.startswith('~$'):  # 排除暫存檔
            file_path = os.path.join(directory_path, filename)
            if read_word_titles(file_path):
                success_count += 1
            else:
                fail_count += 1
    
    # 輸出處理結果統計
    print('\n處理完成！')
    print(f'成功處理: {success_count} 個檔案')
    print(f'處理失敗: {fail_count} 個檔案')

# 使用範例
if __name__ == '__main__':
    # 替換成您的目錄路徑
    directory = 'word_files'  # 例如: 'C:/Documents/word_files'
    process_directory(directory) 