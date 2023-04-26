import os
import re
import pandas as pd

# Đường dẫn tới thư mục chứa các file .md
dir_path = 'D:\DELL\dacn\LOLBAS'

# Tạo một list để lưu dữ liệu của các file .md
data = []

# Duyệt qua các tệp tin .md trong thư mục và các thư mục con
for root, dirs, files in os.walk(dir_path):
    for file_name in files:
        if file_name.endswith('.md'):
            # Đường dẫn đầy đủ đến tệp tin .md
            file_path = os.path.join(root, file_name)

            # Đọc nội dung của tệp tin .md
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.read()

                # Sử dụng regex để tìm nội dung từ "- Command:" đến "Description:"
                match = re.search(r"- Command:(.*?)(?=Description:)", lines, re.DOTALL)
                if match:
                    command_content = match.group(1).strip()
                    data.append([file_name, command_content])

# Tạo một DataFrame từ list data
df = pd.DataFrame(data, columns=['File_Name', 'Command_Content'])

# Ghi DataFrame vào file Excel
df.to_excel('D:\DELL\dacn\out.xlsx', index=False)