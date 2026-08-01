import json
import os
from pptx import Presentation
# Import hàm xử lý đã được mình nâng cấp từ app.py
from app import inject_content_to_pptx

def main():
    template_path = "BAN_MAU.pptx"
    output_path = "Ket_Qua_Demo_Khong_Can_API.pptx"
    
    print("Dang tien hanh ghep du lieu mau (co anh AI) vao template...")
    
    # Giả lập dữ liệu JSON mà backend sẽ sinh ra (kèm URL ảnh)
    mock_json_script = json.dumps([
        {
            "title": "Chương 11: Linear Regression",
            "bullets": [
                "Hồi quy tuyến tính là phương pháp mô hình hóa mối quan hệ giữa biến phụ thuộc và biến độc lập.",
                "Mô hình đơn giản nhất có dạng y = ax + b, trong đó a là độ dốc, b là tung độ gốc.",
                "Được ứng dụng rộng rãi trong dự báo kinh tế, phân tích dữ liệu và học máy."
            ],
            "speaker_notes": "Chào các bạn, hôm nay chúng ta sẽ tìm hiểu về hồi quy tuyến tính...",
            "image_url": "https://image.pollinations.ai/prompt/Linear%20Regression%20graph%20modern%20corporate%20style?width=400&height=300&nologo=true"
        },
        {
            "title": "Hệ số tương quan Pearson (r)",
            "bullets": [
                "Hệ số r đo lường mức độ tương quan tuyến tính giữa hai biến số liên tục.",
                "Giá trị r nằm trong khoảng từ -1 đến +1. r = 0 nghĩa là không có tương quan.",
                "Nếu r > 0: Tương quan thuận. Nếu r < 0: Tương quan nghịch."
            ],
            "speaker_notes": "Tiếp theo là phần hệ số Pearson, một chỉ số cực kỳ quan trọng...",
            "image_url": "https://image.pollinations.ai/prompt/Pearson%20correlation%20chart%20minimalist?width=400&height=300&nologo=true"
        }
    ], ensure_ascii=False)
    
    try:
        inject_content_to_pptx(template_path, mock_json_script, output_path)
        print(f"Thanh cong! Da tao ra file {output_path} (da test chen chu auto-fit va chen anh AI)")
    except Exception as e:
        print(f"Loi: {repr(e)}")

if __name__ == "__main__":
    main()
