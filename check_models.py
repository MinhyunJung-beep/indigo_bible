import google.generativeai as genai

# 🔑 아까 복사하신 그 키를 여기에 넣으세요
genai.configure(api_key="AQ.Ab8RN6K64uy9gE4bggZevY0sQZ_p2VRZmiqMtTt8RRog6aNt4g")

print("=== 🔍 서버가 허용하는 모델 리스트 ===")
try:
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(f"✅ 사용 가능한 모델: {m.name}")
except Exception as e:
    print(f"❌ 오류 발생: {e}")