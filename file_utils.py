import os

def save_presentation(prs, original_filename):
    """Сохраняет презентацию с созданием уникального имени файла."""

    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.basename(original_filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    counter = 1
    while True:
        suffix = "" if counter == 1 else f" ({counter})"
        output_filename = os.path.join(output_dir, f"{name_without_ext}_cn{suffix}.pptx")
        
        if not os.path.exists(output_filename):
            break
        counter += 1

    try:
        prs.save(output_filename)
        print(f"✅ Файл сохранен: {output_filename}")
        return output_filename
    except Exception as e:
        print(f"Ошибка при записи: {e}")
        raise
