import openai
import os
from decouple import config

openai.api_key = config("OPENAI_API_KEY")

def extract_data_with_openai(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1000
        )
        return response["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"OpenAI API Error: {e}")
        return None