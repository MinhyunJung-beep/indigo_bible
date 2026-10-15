import google.generativeai as genai

# 🔑 여기에 아까 복사하신 AQ 키를 넣으세요!
genai.configure(api_key="AQ.Ab8RN6K64uy9gE4bggZevY0sQZ_p2VRZmiqMtTt8RRog6aNt4g")

# 사용 가능한 모델 목록을 스캔합니다.
for m in genai.list_models():
    if 'generateContent' in m.supported_generation_methods:
        print(f"✅ 사용 가능한 모델: {m.name}")