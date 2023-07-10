import pandas as pd
import openai
import os
import time
from dotenv import load_dotenv
import yaml
from tqdm import tqdm
import logging

load_dotenv('.env')

with open('config.yaml', 'r') as tmp_cfg:
    cfg = yaml.safe_load(tmp_cfg)

# Настройка логгера
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('log.txt'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

logger.info(f'Translation language: {cfg["settings"]["translate_lang"]}')
logger.info(f'Excel file: {cfg["settings"]["xls_file"]}')
logger.info(f'Sheet name: {cfg["settings"]["sheet_name"]}')

# Установка ключа API OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')


def request_translation(texts, target_language):
    prompt = 'Translate the following texts from English to ' + target_language + ':\n'
    prompt += '\n'.join([f'"{text}"' for text in texts]) + '\nTranslate:'

    response = openai.Completion.create(
        engine='text-davinci-003',
        prompt=prompt,
        temperature=0.3,
        max_tokens=100,
        n=1,
        stop=None
    )
    print(response['choices'][0]['text'])
    return response.choices


# Функция для запроса перевода с помощью модели ChatGPT
def translate_batch(batch, target_language):
    response_choices = request_translation(batch, target_language)

    translations = []
    for choice in response_choices:
        if choice.text and ':' in choice.text:
            translation = choice.text.strip().split(':')[1].strip()
            translations.append(translation)
        else:
            translations.append('')

    # Удаление специальных символов Unicode
    translations = [text.encode('utf-8').decode('unicode_escape') for text in translations]

    return translations


# Загрузка данных из файла Excel
data = pd.read_excel(cfg['settings']['xls_file'], sheet_name=cfg['settings']['sheet_name'])

logger.info(f'Data from Excel:')
logger.info(f'{data.head()}\n')

# Параметры для разбиения на пакеты
rows_per_batch = cfg['settings']['rows_per_batch']
total_rows = len(data)
num_batches = (total_rows + rows_per_batch - 1) // rows_per_batch

# Использование tqdm для отображения прогресса
with tqdm(total=total_rows, ncols=80, desc='Translating') as pbar:
    for batch_num in range(num_batches):
        start_idx = batch_num * rows_per_batch
        end_idx = min(start_idx + rows_per_batch, total_rows)
        batch_A = data[cfg['settings']['source_column_A']][start_idx:end_idx]

        translations = translate_batch(batch_A, cfg['settings']['translate_lang'])

        # Запись результата перевода в столбец с сохранением в Excel файл
        new_file_name = cfg['settings']['new_file_name']
        save_column_A = cfg['settings']['save_column_A']  # Specify the save column for translation A
        save_column_B = cfg['settings']['save_column_B']  # Specify the save column for translation B
        data[save_column_A] = translated_text_A
        data[save_column_B] = translated_text_B
        data.to_excel(new_file_name, index=False)

        pbar.update(len(batch_A))
        time.sleep(cfg['settings']['batch_interval'])

logger.info(f'Translation completed. Result saved to: {cfg["settings"]["xls_file"]}')
