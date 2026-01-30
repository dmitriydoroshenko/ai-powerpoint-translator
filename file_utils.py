import os
import logging

def save_presentation(prs, original_filename):
    """Сохраняет презентацию с обработкой ошибок и созданием уникального имени файла."""
    # Создаем директорию для вывода, если она не существует
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    
    # Формируем базовое имя выходного файла
    base_name = os.path.basename(original_filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    # Пытаемся сохранить файл, подбирая имя, если файл занят или уже существует
    counter = 1
    while True:
        if counter == 1:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_cn.pptx")
        else:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_cn_{counter}.pptx")
        
        try:
            prs.save(output_filename)
            logging.info(f"Презентация успешно сохранена: {output_filename}")
            print(f"✅ Файл сохранен: {output_filename}")
            return output_filename
        except PermissionError:
            logging.warning(f"Ошибка доступа при сохранении {output_filename}. Возможно, файл открыт в PowerPoint.")
            logging.info("Пожалуйста, закройте файл в PowerPoint, если он открыт.")
            print(f"Ошибка доступа при сохранении {output_filename}. Возможно, файл открыт в PowerPoint.")
            counter += 1
            if counter > 5:  # Ограничение количества попыток
                raise Exception(f"Не удалось сохранить презентацию после {counter-1} попыток. Убедитесь, что файл не открыт в сторонних программах.")
        except Exception as e:
            logging.error(f"Ошибка при сохранении презентации: {str(e)}")
            raise