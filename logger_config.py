import os
import logging
from datetime import datetime

def setup_logging():
    """Настройка системы логирования."""
    # Создаем директорию для логов, если её нет
    log_dir = 'SlideTranslateLog'
    os.makedirs(log_dir, exist_ok=True)
    
    # Создаем временную метку для названия файла
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    log_file = os.path.join(log_dir, f'{timestamp}.log')
    
    # Конфигурация
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
        ]
    )
    
    return log_file