import os
import gc
import pandas as pd
from langchain_community.document_loaders import PyPDFDirectoryLoader, PyPDFLoader # 개별 PDF 로더 추가
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_core.documents import Document

class RacingBibleEngine:
    def __init__(self):
        self.heavy_db_path = "faiss_heavy_manuals"
        print("⏳ 오프라인 V8 임베딩 엔진 가동 중... (가짜 파일 방어 모드)")
        self.embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
        
        os.makedirs("docs_heavy", exist_ok=True)
        os.makedirs("docs_light", exist_ok=True)
        
        self.vectorstore = self._boot_engine()

    def _boot_engine(self):
        if not os.path.exists(self.heavy_db_path):
            print("🚨 [1단계] 베이스 롱블록 엔진이 없습니다. 방어막을 켜고 빌드를 시작합니다.")
            heavy_vectorstore = self._build_heavy_brain()
        else:
            print("✅ [1단계] 베이스 롱블록 엔진 로딩 완료. (초고속 패스)")
            heavy_vectorstore = FAISS.load_local(self.heavy_db_path, self.embeddings, allow_dangerous_deserialization=True)

        print("🚀 [2단계] 숏블록 엔진(에릭 리포트/CSV 기록지/셋업 엑셀)을 스캔하여 융합합니다...")
        light_vectorstore = self._build_light_brain()

        if light_vectorstore is not None:
            heavy_vectorstore.merge_from(light_vectorstore)
            print("🏁 [완료] 롱블록 + 숏블록 데이터 융합 완료! 시스템 스탠바이.")

        return heavy_vectorstore

    def _clean_docs(self, docs):
        clean_list = []
        for doc in docs:
            text = str(doc.page_content).replace("\x00", "").strip()
            if len(text) > 5:
                doc.page_content = text
                clean_list.append(doc)
        return clean_list

    def _build_heavy_brain(self):
        pdf_loader = PyPDFDirectoryLoader("docs_heavy")
        pdf_docs = pdf_loader.load()
        
        pdf_docs = self._clean_docs(pdf_docs)
        print(f"📚 [Heavy] 규정집/매뉴얼 로딩: 총 {len(pdf_docs)} 페이지")
        
        text_splitter = RecursiveCharacterTextSplitter(chunk_size=900, chunk_overlap=100)
        split_docs = text_splitter.split_documents(pdf_docs)
        split_docs = self._clean_docs(split_docs)
        total_chunks = len(split_docs)
        
        vectorstore = None
        batch_size = 500
        
        for i in range(0, total_chunks, batch_size):
            batch = split_docs[i:i+batch_size]
            if not batch: continue
            try:
                if vectorstore is None:
                    vectorstore = FAISS.from_documents(batch, self.embeddings)
                else:
                    vectorstore.add_documents(batch)
            except Exception as e:
                for single_doc in batch:
                    try:
                        if vectorstore is None:
                            vectorstore = FAISS.from_documents([single_doc], self.embeddings)
                        else:
                            vectorstore.add_documents([single_doc])
                    except:
                        pass
            gc.collect()
            
        vectorstore.save_local(self.heavy_db_path)
        return vectorstore

    def _build_light_brain(self):
        light_docs = []
        
        # 1. docs_light 안의 파일들을 하나씩 안전하게 로드 (가짜 파일 무시)
        for filename in os.listdir("docs_light"):
            file_path = os.path.join("docs_light", filename)
            
            # PDF 개별 로드
            if filename.lower().endswith(".pdf"):
                try:
                    loader = PyPDFLoader(file_path)
                    light_docs.extend(loader.load())
                    print(f"   📄 [성공] {filename} 로딩 완료")
                except Exception as e:
                    print(f"   ⚠️ [경고] {filename}은(는) 가짜 또는 손상된 PDF입니다. 무시합니다.")
                    
            # CSV 개별 로드
            elif filename.lower().endswith(".csv"):
                try:
                    df_csv = pd.read_csv(file_path)
                    df_csv = df_csv.fillna("")
                    for index, row in df_csv.iterrows():
                        row_text = f"[{filename}] " + " | ".join([f"{col}: {val}" for col, val in row.items()])
                        light_docs.append(Document(page_content=row_text, metadata={"source": file_path}))
                    print(f"   📊 [성공] {filename} 로딩 완료")
                except Exception as e:
                    print(f"   ⚠️ [경고] {filename} CSV 읽기 실패. 무시합니다.")

        # 2. 바탕화면의 셋업 엑셀 데이터 로드
        excel_path = "tcr_diagnosis.xlsx"
        if os.path.exists(excel_path):
            try:
                df = pd.read_excel(excel_path)
                df = df.fillna("")
                for index, row in df.iterrows():
                    row_text = " | ".join([f"{col}: {val}" for col, val in row.items()])
                    light_docs.append(Document(page_content=row_text, metadata={"source": excel_path}))
                print("   📗 [성공] tcr_diagnosis.xlsx 셋업 데이터 로딩 완료")
            except Exception as e:
                 print(f"   ⚠️ [경고] 엑셀 로딩 실패: {e}")
                
        light_docs = self._clean_docs(light_docs)
        if not light_docs: return None
            
        text_splitter = RecursiveCharacterTextSplitter(chunk_size=900, chunk_overlap=100)
        split_docs = text_splitter.split_documents(light_docs)
        split_docs = self._clean_docs(split_docs)
        
        vectorstore = None
        if split_docs:
            for doc in split_docs:
                try:
                    if vectorstore is None:
                        vectorstore = FAISS.from_documents([doc], self.embeddings)
                    else:
                        vectorstore.add_documents([doc])
                except:
                    pass
        return vectorstore

    def query_bible(self, query, k=30):
        return self.vectorstore.similarity_search(query, k=k)