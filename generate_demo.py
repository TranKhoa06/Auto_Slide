import json
from app import generate_native_pptx

def main():
    print("Generating demo natively...")
    
    mock_json_script = json.dumps([
        {
            "title": "Chương 11: Linear Regression",
            "bullets": [
                "Hồi quy tuyến tính là phương pháp mô hình hóa mối quan hệ giữa biến phụ thuộc và biến độc lập.",
                "Mô hình đơn giản nhất có dạng y = ax + b, trong đó a là độ dốc, b là tung độ gốc."
            ],
            "speaker_notes": "Chào các bạn...",
            "image_url": "https://image.pollinations.ai/prompt/Linear%20Regression%20graph%20modern%20corporate%20style?width=400&height=300&nologo=true"
        },
        {
            "title": "Hệ số tương quan Pearson (r)",
            "bullets": [
                "Hệ số r đo lường mức độ tương quan tuyến tính giữa hai biến số liên tục.",
                "Giá trị r nằm trong khoảng từ -1 đến +1. r = 0 nghĩa là không có tương quan."
            ],
            "speaker_notes": "Tiếp theo là phần hệ số Pearson...",
            "image_url": "https://image.pollinations.ai/prompt/Pearson%20correlation%20chart%20minimalist?width=400&height=300&nologo=true"
        }
    ], ensure_ascii=False)
    
    output_path = "Native_Demo_Result.pptx"
    try:
        generate_native_pptx(mock_json_script, output_path)
        print(f"Success! Generated {output_path}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
