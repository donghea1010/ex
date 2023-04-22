
import pandas as pd
from gensim.models import Word2Vec
from openpyxl import Workbook

# Đọc dữ liệu từ file Excel đầu vào
df = pd.read_excel('D:\DELL\dacn\hihi.xlsx')

# Chuẩn bị dữ liệu đầu vào cho Word2Vec
sentences = df['Noi_dung'].str.split().tolist()  # Thay 'text' bằng tên cột chứa dữ liệu cần xử lý
model = Word2Vec(sentences, min_count=1)  # Tạo mô hình Word2Vec với độ tần suất xuất hiện tối thiểu là 1

# Lưu kết quả vào file Excel
wb = Workbook()
ws = wb.active
ws.append(['Word', 'Embedding'])
for word in model.wv.index_to_key:
    embedding = model.wv[word]
    ws.append([word] + embedding.tolist())
wb.save('ket_qua.xlsx')

