import requests
import zipfile
import re
import sys
import os

API_URL = "http://127.0.0.1:8000/generate"
PDF_FILE = "Chapter 11-Simple Linear Regression and Correlation.pdf"
PPTX_FILE = "BAN_MAU.pptx"
OUTPUT_FILE = "test_output.pptx"

def run_test():
    print("🚀 Bắt đầu giả lập người dùng gửi request lên máy chủ...")
    
    if not os.path.exists(PDF_FILE) or not os.path.exists(PPTX_FILE):
        print("❌ LỖI: Không tìm thấy file mẫu để test.")
        sys.exit(1)

    with open(PDF_FILE, 'rb') as pdf, open(PPTX_FILE, 'rb') as pptx:
        files = {
            'pdf_file': ('sample.pdf', pdf, 'application/pdf'),
            'template_file': ('template.pptx', pptx, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
        }
        data = {
            'slide_count': '5',
            'api_key': '',
            'auto_image': 'false'
        }
        
        try:
            response = requests.post(API_URL, files=files, data=data, timeout=120)
        except Exception as e:
            print(f"❌ LỖI KẾT NỐI MÁY CHỦ: {e}")
            sys.exit(1)
            
    if response.status_code != 200:
        print(f"❌ LỖI BACKEND: Trả về HTTP {response.status_code}")
        try:
            print(response.json())
        except:
            print(response.text)
        sys.exit(1)
        
    print("✅ Đã nhận được file trả về từ máy chủ. Đang lưu...")
    with open(OUTPUT_FILE, 'wb') as f:
        f.write(response.content)
        
    print("🔍 Bắt đầu bung file XML để chấm điểm...")
    
    try:
        with zipfile.ZipFile(OUTPUT_FILE, 'r') as z:
            slide_files = [f for f in z.namelist() if f.startswith('ppt/slides/slide')]
            slide_files.sort(key=lambda x: int(re.search(r'slide(\d+)\.xml', x).group(1)))
            
            passed = True
            for slide_file in slide_files[2:5]: # Kiểm tra 3 slide đầu tiên AI sinh ra
                xml = z.read(slide_file).decode('utf-8')
                texts = re.findall(r'<a:t>(.*?)</a:t>', xml)
                texts = [t.strip() for t in texts if t.strip()]
                
                print(f"\n--- Phân tích {slide_file} ---")
                print("Chữ lấy được:", texts)
                
                if not texts:
                    print(f"❌ LỖI: Slide bị trống chữ!")
                    passed = False
                    continue
                
                # Kiểm tra rác Lorem Ipsum
                if any("Lorem ipsum" in t for t in texts):
                    print(f"❌ LỖI: Còn sót chữ rác 'Lorem ipsum' trên slide!")
                    passed = False
                    
    except Exception as e:
        print(f"❌ LỖI PHÂN TÍCH XML: {e}")
        sys.exit(1)
        
    if passed:
        print("\n🎉 KẾT LUẬN: PASS 100%. File PPTX cực kỳ sạch sẽ và chuẩn mực!")
        sys.exit(0)
    else:
        print("\n💥 KẾT LUẬN: FAILED. Cần sửa lại app.py!")
        sys.exit(1)

if __name__ == "__main__":
    run_test()
