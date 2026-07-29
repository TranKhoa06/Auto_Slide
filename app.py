import os
import json
import uuid
import asyncio
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
from google import genai
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

load_dotenv()

app = FastAPI()

# Cấu hình CORS để Frontend gọi được API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Phục vụ các file tĩnh (Giao diện web)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Thư mục tạm để lưu file
TEMP_DIR = "temp"
os.makedirs(TEMP_DIR, exist_ok=True)

async def generate_script_content(pdf_path: str, slide_count: int, api_key: str):
    """BƯỚC 1: Gọi AI đọc PDF và lập Kịch bản Text"""
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
    print(f"Uploading {pdf_path} to Gemini for script generation...")
    sample_file = client.files.upload(file=pdf_path, config={'display_name': "Lecture PDF"})
    
    script_prompt = f"Bạn là một chuyên gia phân tích và thiết kế nội dung bài giảng. Dưới đây là toàn bộ tài liệu học thuật. Hãy đọc kỹ, tóm tắt và phân chia toàn bộ nội dung thành một Kịch bản Thuyết trình gồm đúng {slide_count} Slide. YÊU CẦU NGHIÊM NGẶT: Phải trích xuất và giữ lại TẤT CẢ các định nghĩa, công thức, số liệu quan trọng, không được bỏ sót thông tin cốt lõi nào. Viết rõ ràng dưới dạng Markdown (ví dụ: Slide 1: [Tiêu đề], Nội dung: ...)."
    
    print("Generating script...")
    response = client.models.generate_content(
        model='gemini-flash-latest',
        contents=[sample_file, script_prompt]
    )
    client.files.delete(name=sample_file.name)
    return response.text

def extract_layouts_from_pptx(prs):
    """Trích xuất tọa độ Bounding Boxes của toàn bộ Text Shapes"""
    layouts = []
    for i, slide in enumerate(prs.slides):
        shapes_info = []
        for j, shape in enumerate(slide.shapes):
            if shape.has_text_frame:
                shapes_info.append({
                    "id": j,
                    "x": getattr(shape, 'left', 0),
                    "y": getattr(shape, 'top', 0),
                    "w": getattr(shape, 'width', 0),
                    "h": getattr(shape, 'height', 0),
                    "original_text": shape.text.strip()[:50].replace('\n', ' ')
                })
        layouts.append({"slide_index": i, "shapes": shapes_info})
    return layouts

async def map_script_to_layout(script: str, layouts: list, api_key: str):
    """BƯỚC 2: Gọi AI để tự động gán (map) kịch bản vào các khối tọa độ (Zero-Shot)"""
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
    
    json_prompt = f"""Dưới đây là Kịch bản thuyết trình:
{script}

Và đây là sơ đồ Không gian (Bounding Boxes) của các khối văn bản trên từng slide của file thiết kế (Slide 0 là bìa, v.v.):
{json.dumps(layouts, ensure_ascii=False)}

NHIỆM VỤ CỦA BẠN (ZERO-SHOT LAYOUT MAPPING):
Hãy nhúng toàn bộ kịch bản vào các slide này. Đối với mỗi slide_index, hãy suy luận xem shape_id nào là Tiêu đề, shape_id nào là Nội dung dựa vào 'original_text' (văn bản gốc của designer trên Canva), tọa độ (y nhỏ là tiêu đề) và kích thước (w, h). 
TUYỆT ĐỐI KHÔNG nhét nội dung dài (bullet points) vào các hộp chữ vốn là Tiêu đề (dựa theo original_text).
Nếu slide thiết kế có nhiều text box, hãy mạnh dạn chia nhỏ nội dung kịch bản để phân bổ đều vào các box. Cố gắng tóm tắt gọn gàng để tránh tràn viền. Mọi giải thích dài dòng nhét vào 'speaker_notes'.
Những shape không dùng đến (như chữ rác Canva), hãy gán chuỗi rỗng "".

TRẢ VỀ ĐÚNG ĐỊNH DẠNG JSON (KHÔNG MARKDOWN):
[
    {{
        "slide_index": 0,
        "mapping": {{
            "0": "Tiêu đề Slide 1",
            "1": "Nội dung 1...",
            "2": ""
        }},
        "speaker_notes": "..."
    }},
    ...
]
"""
    print("Generating Mapping JSON...")
    response_json = client.models.generate_content(
        model='gemini-flash-latest',
        contents=json_prompt
    )
    
    try:
        json_str = response_json.text.strip()
        if json_str.startswith("```json"): json_str = json_str[7:-3]
        elif json_str.startswith("```"): json_str = json_str[3:-3]
        return json.loads(json_str)
    except Exception as e:
        print("JSON parse error")
        raise ValueError(f"AI Mapping thất bại: {str(e)}")

def inject_content_to_pptx(template_path: str, mapping_data: list, output_path: str):
    """Chèn dữ liệu JSON (đã map) vào file PPTX dựa trên Shape ID"""
    prs = Presentation(template_path)
    
    if len(prs.slides) == 0:
        raise ValueError("File Template không có slide nào!")

    def replace_text_preserve_format(shape, new_text):
        if not shape.has_text_frame: return
        paragraphs = shape.text_frame.paragraphs
        if not paragraphs: return
        
        # Chỉ chèn vào đoạn văn đầu tiên để giữ định dạng
        if len(paragraphs[0].runs) > 0:
            original_size = paragraphs[0].runs[0].font.size
            paragraphs[0].runs[0].text = new_text
            # Ép nhỏ font nếu text quá dài để chống tràn
            if len(new_text) > 100 and original_size:
                paragraphs[0].runs[0].font.size = int(original_size * 0.6)
            for r in paragraphs[0].runs[1:]:
                r.text = ""
        else:
            paragraphs[0].text = new_text
            
        # Xóa sạch toàn bộ các đoạn văn thừa bên dưới để tránh lỗi lặp chữ nhiều lần
        for p in paragraphs[1:]:
            p.text = ""

    # Tạo dictionary để truy cập mapping nhanh hơn
    map_dict = {item["slide_index"]: item for item in mapping_data}

    for i, slide in enumerate(prs.slides):
        slide_map = map_dict.get(i, None)
        
        # Nếu AI quyết định không dùng slide này (ví dụ template thừa), xóa trắng
        if not slide_map:
            for shape in slide.shapes:
                replace_text_preserve_format(shape, "")
            continue
            
        shape_mappings = slide_map.get("mapping", {})
        
        for j, shape in enumerate(slide.shapes):
            if shape.has_text_frame:
                new_text = shape_mappings.get(str(j), "")
                # Tính năng đặc biệt: Nếu AI cố tình bỏ trống, ta xóa rác
                replace_text_preserve_format(shape, new_text)
                
                # Auto-fit
                if len(new_text) > 30:
                    shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                    shape.text_frame.word_wrap = True
                    
        # Speaker notes
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = slide_map.get("speaker_notes", "")

    prs.save(output_path)


@app.post("/generate_script")
async def api_generate_script(
    pdf_file: UploadFile = File(...),
    slide_count: int = Form(...),
    api_key: str = Form("")
):
    try:
        session_id = str(uuid.uuid4())
        pdf_path = os.path.join(TEMP_DIR, f"{session_id}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(await pdf_file.read())
            
        script_text = await generate_script_content(pdf_path, slide_count, api_key)
        return {"session_id": session_id, "script": script_text}
    except Exception as e:
        print("Script Error:", str(e))
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate_pptx")
async def api_generate_pptx(
    session_id: str = Form(...),
    script: str = Form(...),
    template_file: UploadFile = File(...),
    api_key: str = Form("")
):
    try:
        pptx_temp_path = os.path.join(TEMP_DIR, f"{session_id}_template.pptx")
        output_path = os.path.join(TEMP_DIR, f"{session_id}_output.pptx")
        
        with open(pptx_temp_path, "wb") as f:
            f.write(await template_file.read())
            
        prs = Presentation(pptx_temp_path)
        layouts = extract_layouts_from_pptx(prs)
        
        mapping_data = await map_script_to_layout(script, layouts, api_key)
        inject_content_to_pptx(pptx_temp_path, mapping_data, output_path)
        
        return FileResponse(
            path=output_path, 
            filename="AutoSlide_Generated.pptx", 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        print("PPTX Error:", str(e))
        raise HTTPException(status_code=500, detail=str(e))

# Phục vụ Frontend HTML
app.mount("/", StaticFiles(directory="static", html=True), name="static")
