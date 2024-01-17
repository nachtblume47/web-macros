from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
import pandas as pd
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# Подключаем CORS для разрешения запросов из любых источников
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/upload/")
async def upload_form(request: Request):
    return templates.TemplateResponse("upload_form.html", {"request": request})


@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    try:
        # Считываем загруженный файл в pandas DataFrame
        df = pd.read_excel(BytesIO(file.file.read()))

        # Загружаем словарь
        replacement_df = pd.read_excel('dict.xlsx')
        print(replacement_df)

        # Переименовываем столбец
        df = df.rename(columns={'Баркод': 'Barcode'})
        df = df.rename(columns={'баркод товара': 'Barcode'})
        df = df.rename(columns={'barcode': 'Barcode'})
        df = df.rename(columns={'баркод': 'Barcode'})
        df = df.rename(columns={'Последний баркод': 'Barcode'})

        # Объединяем DataFrames
        df_merged = pd.merge(df, replacement_df, on='Barcode', how='left')
        df_merged = df_merged.fillna('')

        # Сохраняем результат в новый Excel файл
        output_filename = 'extended.xlsx'
        df_merged.to_excel(output_filename, index=False)

        # Возвращаем обработанный файл
        return FileResponse(output_filename, filename=output_filename)
    except Exception as e:
        return {"error": str(e)}
