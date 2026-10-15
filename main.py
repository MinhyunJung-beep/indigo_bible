import os
from search_engine import RacingBibleEngine
from langchain_google_genai import ChatGoogleGenerativeAI

# 🔑 API 키 설정 (보안 주의!)
os.environ["GOOGLE_API_KEY"] = "AQ.Ab8RN6K64uy9gE4bggZevY0sQZ_p2VRZmiqMtTt8RRog6aNt4g"

print("=== 🏁 레이스 엔진 최종 시동 ===")
engine = RacingBibleEngine()

# 모델 설정 (가장 안정적인 최신 엔진)
# google_api_key 옵션을 빼도 os.environ에서 자동으로 키를 가져옵니다!
llm = ChatGoogleGenerativeAI(model="gemini-3.8-flash", temperature=0.2)

# 테스트 검색
search_query = "브레이크 락업 원인"
print(f"분석 질문: {search_query}")

results = engine.query_bible(search_query)
context = "\n\n".join([doc.page_content for doc in results])

prompt = f"매뉴얼을 바탕으로 질문에 한국어로 답하세요.\n매뉴얼 내용: {context}\n질문: {search_query}"

print("=== 💡 AI 분석 결과 ===")
print(llm.invoke(prompt).content)