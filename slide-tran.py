import os
import glob
from pptx import Presentation
from openai import OpenAI
from dotenv import load_dotenv
import logging
import time
from datetime import datetime
import httpx
import sys
import re

# Настройка логирования
def setup_logging():
    log_dir = 'SlideTranslateLog'
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'translation_{timestamp}.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.FileHandler(log_file, encoding='utf-8')]
    )
    return log_file

log_file = setup_logging()
load_dotenv()

# Инициализация OpenAI (ChatGPT)
client = OpenAI(
    api_key=os.getenv('OPENAI_API_KEY'), # Убедитесь, что в .env теперь этот ключ
)

# Загрузка шаблона промпта
with open('prompt.txt', 'r', encoding='utf-8') as f:
    PROMPT_TEMPLATE = f.read()

def batch_texts(texts, batch_size=20): # Уменьшил батч для стабильности GPT-4
    return [texts[i:i + batch_size] for i in range(0, len(texts), batch_size)]

def translate_batch(texts):
    """Перевод батча текста с английского на китайский через ChatGPT."""
    if not texts:
        return []
    
    prompt = PROMPT_TEMPLATE.format(texts="\n---\n".join(texts))
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o", # Или gpt-3.5-turbo
            messages=[
                {"role": "system", "content": "You are a professional translator from English to Chinese (Simplified). Maintain technical terms and formatting."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3
        )
        
        content = response.choices[0].message.content.strip()
        translations = content.split("\n---\n")
        
        # Проверка соответствия количества переведенных строк
        if len(translations) != len(texts):
            logging.warning(f"Mismatch: {len(translations)} trans for {len(texts)} texts")
            if len(translations) < len(texts):
                translations.extend([""] * (len(texts) - len(translations)))
            else:
                translations = translations[:len(texts)]
        
        return translations
        
    except Exception as e:
        logging.error(f"Error during translation: {str(e)}")
        raise

def save_presentation(prs, original_filename):
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.basename(original_filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    output_filename = os.path.join(output_dir, f"{name_without_ext}_CN.pptx")
    prs.save(output_filename)
    return output_filename

# ... (функции extract_table_texts и split_text_by_paragraphs остаются без изменений) ...

def process_presentation(input_file):
    logging.info(f"Processing {input_file}")
    try:
        prs = Presentation(input_file)
        all_texts = []
        text_locations = []
        
        # Сбор текстов (код из вашего примера)
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if hasattr(shape, "text") and shape.text.strip():
                    if hasattr(shape, "text_frame"):
                        for para_idx, para in enumerate(shape.text_frame.paragraphs):
                            if para.text.strip():
                                all_texts.append(para.text.strip())
                                text_locations.append(("paragraph", slide_idx, shape_idx, para_idx))
                
                if hasattr(shape, "table"):
                    # (Логика таблиц как в вашем исходнике)
                    pass

        # Перевод
        translated_texts = []
        batches = batch_texts(all_texts)
        for i, batch in enumerate(batches):
            translations = translate_batch(batch)
            translated_texts.extend(translations)
            time.sleep(1) # Небольшая пауза для API

        # Применение перевода
        for location, translated_text in zip(text_locations, translated_texts):
            if location[0] == "paragraph":
                _, slide_idx, shape_idx, para_idx = location
                paragraph = prs.slides[slide_idx].shapes[shape_idx].text_frame.paragraphs[para_idx]
                
                # Сохраняем форматирование
                original_font_size = paragraph.runs[0].font.size if paragraph.runs else None
                
                paragraph.text = translated_text
                
                # Устанавливаем китайский шрифт
                for run in paragraph.runs:
                    run.font.name = "Microsoft YaHei" # Стандарт для Китая
                    if original_font_size:
                        run.font.size = original_font_size

        save_presentation(prs, input_file)
        
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        raise

def main():
    input_files = glob.glob('input/*.pptx')
    for input_file in input_files:
        process_presentation(input_file)

if __name__ == "__main__":
    main()