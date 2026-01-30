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
    """
    Основная точка входа. Принимает список текстов, 
    сама бьет на батчи и возвращает полный перевод.
    """
    if not texts:
        return []

    # 1. Инкапсулированное разбиение на батчи
    batches = [texts[i:i + batch_size] for i in range(0, len(texts), batch_size)]
    translated_result = []

    for i, batch in enumerate(batches):
        logging.info(f"Translating batch {i+1}/{len(batches)} (size: {len(batch)})")
        
        translations = _translate_batch(batch, model)
        translated_result.extend(translations)
        
        if i < len(batches) - 1:
            time.sleep(1) 

    return translated_result

def _translate_batch(batch_texts, model):
    """Внутренняя функция для обработки одного батча."""
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
        
        # Собираем результаты, подставляя оригинал при отсутствии ключа
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