import os
import json
import time
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

SYSTEM_ROLE = (
    "You are a professional mobile game localizer (English to Simplified Chinese). "
    "Expertise: gaming terminology, UI/UX constraints, and mobile gaming slang. "
    "IMPORTANT: Preserve all special characters like vertical tabs (\\u000b), "
    "newlines (\\n), and specific spacing. Do not clean up the formatting. "
    "Task: Translate values to Simplified Chinese. Keep keys unchanged. "
    "Output: Return a valid JSON object."
)

BATCH_SIZE = 30

def translate_all(texts):
    if not texts:
        print("Список текстов пуст.")
        return []

    total_texts = len(texts)
    batches = [texts[i:i + BATCH_SIZE] for i in range(0, total_texts, BATCH_SIZE)]
    translated_result = []
    
    total_prompt = 0
    total_completion = 0
    
    print(f"\n{'='*20}")
    print(f"🚀 НАЧАЛО ПЕРЕВОДА")
    print(f"Всего строк: {total_texts} (Всего батчей: {len(batches)})")
    print(f"{'='*20}\n")

    for i, batch in enumerate(batches):
        print(f"⏳ Обработка батча {i+1}/{len(batches)}... (Строк в батче: {len(batch)})")
        
        translations, usage = _translate_batch(batch)
        translated_result.extend(translations)
        
        if usage:
            total_prompt += usage.prompt_tokens
            total_completion += usage.completion_tokens
        
        if i < len(batches) - 1:
            time.sleep(1) 

    cost = (total_prompt * 1.75 / 1_000_000) + (total_completion * 14.00 / 1_000_000)

    print(f"\n\n{'='*20}")
    print(f"✅ ПЕРЕВОД ЗАВЕРШЕН")
    print(f"Итого обработано строк: {len(translated_result)}/{total_texts}")
    print(f"Токены: {total_prompt + total_completion} | Общая стоимость: ${cost:.4f}")
    print(f"{'='*20}\n")

    return translated_result

def _translate_batch(batch_texts):
    payload = {f"item_{i}": text for i, text in enumerate(batch_texts)}
    
    try:
        response = client.chat.completions.create(
            model="gpt-5.2",
            messages=[
                {"role": "system", "content": SYSTEM_ROLE},
                {"role": "user", "content": f"Translate these items:\n{json.dumps(payload, ensure_ascii=False)}"}
            ],
            response_format={"type": "json_object"},
            temperature=0.3
        )

        content = response.choices[0].message.content
        if content:
            translated_data = json.loads(content)
        else:
            translated_data = {}

        batch_results = [translated_data.get(f"item_{i}", batch_texts[i]) for i in range(len(batch_texts))]

        return batch_results, response.usage
    except Exception as e:
        print(f"Ошибка при обработке батча: {e}")
        return batch_texts, None
