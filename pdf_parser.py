import os
import fitz  # PyMuPDF
from langchain_text_splitters import RecursiveCharacterTextSplitter

def extract_and_split_pdf(file_path):
    print(f"[{file_path}] 매뉴얼 읽기 시작...")
    doc = fitz.open(file_path)
    full_text = ""
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        full_text += f"\n\n--- [Page {page_num + 1}] ---\n\n" + page.get_text("text")
        
    print(f"텍스트 추출 완료 ({len(full_text)} 글자). 데이터 분할 중...")
    
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=700,
        chunk_overlap=120,
        separators=["\n\n", "Chapter", "Section", "Symptom", "■", "•"]
    )
    
    chunks = text_splitter.split_text(full_text)
    print(f"총 {len(chunks)}개의 의미 있는 데이터 조각(Chunks)으로 분할 완료.")
    return chunks
if __name__ == "__main__":
    print("🚀 롱블록 엔진 바이블(PDF) 파싱을 강제 시작합니다...")

# PDF 파일들이 모여있는 폴더 경로 (치프 메카닉님의 폴더명에 맞게 세팅)
target_folder = "./docs_heavy" 

if not os.path.exists(target_folder):
    print(f"⚠️ 에러: '{target_folder}' 폴더를 찾을 수 없습니다. 경로를 확인해주세요.")
else:
    # 폴더 안의 모든 PDF 파일을 하나씩 찾아서 꺼냅니다.
    for filename in os.listdir(target_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(target_folder, filename)
            print(f"\n▶️ 타겟 파일 스캔 중: {filename}")
            
            # 1. 치프 메카닉님이 위에서 만드신 함수를 드디어 호출(실행)합니다!
            chunks = extract_and_split_pdf(pdf_path)
            
            # -----------------------------------------------------------
            # [중요] 여기에 잘라낸 chunks를 FAISS DB에 저장하는 코드가 들어가야 합니다.
            # 아마 옆에 띄워두신 'search_engine.py' 안에 DB 저장 함수가 있을 것입니다.
            # 예: create_vector_db(chunks) 
            # -----------------------------------------------------------
            
    print("\n✅ 모든 바이블 데이터 파싱 및 청크(Chunk) 분할 작업이 완료되었습니다!")