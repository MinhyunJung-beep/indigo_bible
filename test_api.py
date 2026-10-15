import google.generativeai as genai
import os

# 🔑 [필수] 여기에 아까 복사하신 키를 정확히 넣으세요
genai.configure(api_key="AQ.Ab8RN6K64uy9gE4bggZevY0sQZ_p2VRZmiqMtTt8RRog6aNt4g")

model = genai.GenerativeModel('gemini-1.5-flash')
response = model.generate_content("안녕하세요, API 테스트 중입니다.")

print("=== 💡 API 통신 결과 ===")
print(response.text)