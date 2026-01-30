import os
import json
import logging
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

def translate_all(texts, model="gpt-5.2", batch_size=30):
    """Принимает список текстов и возвращает перевод."""

    if not texts:
        print("Список текстов пуст.")
        return []

    total_texts = len(texts)
    batches = [texts[i:i + batch_size] for i in range(0, total_texts, batch_size)]
    translated_result = []
    
    print(f"\n{'='*20}")
    print(f"🚀 НАЧАЛО ПЕРЕВОДА")
    print(f"Всего строк: {total_texts}")
    print(f"Размер батча: {batch_size}")
    print(f"{'='*20}\n")

    for i, batch in enumerate(batches):
        current_count = len(translated_result)
        logging.info(f"Translating batch {i+1}/{len(batches)} (size: {len(batch)})")
        
        print(f"⏳ Обработка батча {i+1}/{len(batches)}... (Переведено: {current_count}/{total_texts})", end="\r")
        
        translations = _translate_batch(batch, model)
        translated_result.extend(translations)
        
        if i < len(batches) - 1:
            time.sleep(1) 

    print(f"\n\n{'='*20}")
    print(f"✅ ПЕРЕВОД ЗАВЕРШЕН")
    print(f"Итого обработано: {len(translated_result)}/{total_texts}")
    print(f"{'='*20}\n")

    return translated_result

def _translate_batch(batch_texts, model):
    payload = {f"item_{i}": text for i, text in enumerate(batch_texts)}
    
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_ROLE},
                {"role": "user", "content": f"Translate these items:\n{json.dumps(payload, ensure_ascii=False)}"}
            ],
            response_format={"type": "json_object"},
            temperature=0.3
        )

        translated_data = json.loads(response.choices[0].message.content)
        batch_results = [translated_data.get(f"item_{i}", batch_texts[i]) for i in range(len(batch_texts))]
        
        _log_usage(response.usage)
        return batch_results

    except Exception as e:
        logging.error(f"Error in _translate_batch: {e}")
        return batch_texts

def _log_usage(usage):
    """Расчет стоимости."""
    cost = (usage.prompt_tokens * 1.75 / 1_000_000) + (usage.completion_tokens * 14.00 / 1_000_000)
    logging.info(f"Tokens: {usage.total_tokens} | Cost: ${cost:.4f}")