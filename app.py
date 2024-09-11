from flask import Flask, request, send_file, render_template
import pandas as pd
import requests
from io import BytesIO
import time

app = Flask(__name__)


def shorten_url_vk(url, retries=3, delay=5):
    vk_api_url = "https://api.vk.com/method/utils.getShortLink"
    params = {
        'url': url,
        'private': 0,
        'access_token': ' ',  # Замените на ваш токен
        'v': '5.199'
    }

    for attempt in range(retries):
        try:
            response = requests.get(vk_api_url, params=params, timeout=10)
            data = response.json()

            if 'response' in data:
                return data['response']['short_url']
            elif 'error' in data:
                return f"Ошибка: {data['error']['error_msg']}"
        except requests.exceptions.RequestException as e:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                return f"Ошибка: не удалось сократить ссылку после {retries} попыток. {str(e)}"

    return "Ошибка: Не удалось получить короткую ссылку"


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']

    if file.filename.endswith('.xlsx'):
        engine = 'openpyxl'
    else:
        return "Ошибка: Пожалуйста, загрузите файл в формате .xlsx"

    try:
        df = pd.read_excel(file, engine=engine)
    except Exception as e:
        return f"Ошибка при обработке файла: {str(e)}"


    df['Short URL'] = df.iloc[:, 0].apply(shorten_url_vk)


    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    return send_file(output, download_name="shortened_links.xlsx", as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
