import logging


def setup_logging():
    # Создаем логгер
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Создаем обработчик для записи логов в файл
    file_handler = logging.FileHandler('app.log', mode='a')  # 'w' для перезаписи файла каждый раз, 'a' для добавления в файл
    file_handler.setLevel(logging.INFO)

    # Создаем обработчик для вывода логов в консоль
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Определяем формат логов
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] [%(message)s]', datefmt='%Y-%m-%d %H:%M:%S')

    # Применяем формат к обработчикам
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Добавляем обработчики к логгеру
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

