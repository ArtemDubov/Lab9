"""FastAPI-бэкенд для преобразования координат из Excel в отчёт Markdown/DOCX."""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from io import BytesIO
import pandas as pd
from sympy import symbols, Matrix, pi, latex
from datetime import datetime
import subprocess
import tempfile
import os

app = FastAPI(
    title="Coordinate Transformation API",
    description="API для преобразования координат из Excel-файла в отчёт",
    version="1.0.0"
)


def transform_coordinates(df, initial_system, final_system, params):
    # Числовые параметры перехода
    dx, dy, dz = params['dx'], params['dy'], params['dz']
    wx, wy, wz = params['wx'], params['wy'], params['wz']
    m = params['m']

    # Перевод угловых секунд в радианы
    wx_r = wx / 3600 * pi / 180
    wy_r = wy / 3600 * pi / 180
    wz_r = wz / 3600 * pi / 180

    # Символьные переменные и матрицы
    X, Y, Z = symbols('X Y Z')
    V = Matrix([X, Y, Z])
    D = Matrix([dx, dy, dz])
    M = Matrix([
        [1,  wz_r, -wy_r],
        [-wz_r, 1,  wx_r],
        [wy_r, -wx_r, 1]
    ])
    formula = (1 + m) * M * V + D

    # Применяем формулу к каждой строке
    new_x, new_y, new_z = [], [], []
    for _, row in df.iterrows():
        xv = row[initial_system[0]]
        yv = row[initial_system[1]]
        zv = row[initial_system[2]]
        res = formula.subs({X: xv, Y: yv, Z: zv})
        new_x.append(float(res[0]))
        new_y.append(float(res[1]))
        new_z.append(float(res[2]))

    df_out = df.copy()
    df_out[final_system[0]] = new_x
    df_out[final_system[1]] = new_y
    df_out[final_system[2]] = new_z
    return df_out, formula


def generate_markdown_report(df_transformed, formula, initial_system, final_system, params):
    # Формула в LaTeX
    formula_latex = latex(formula)

    # Таблица координат до и после
    table_lines = ['| Name | X | Y | Z | X_new | Y_new | Z_new |',
                   '| --- | --- | --- | --- | --- | --- | --- |']
    for _, r in df_transformed.iterrows():
        name = r['Name'] if 'Name' in df_transformed.columns else '-'
        table_lines.append(
            f"| {name} | {r[initial_system[0]]:.6f} | {r[initial_system[1]]:.6f} | "
            f"{r[initial_system[2]]:.6f} | {r[final_system[0]]:.6f} | "
            f"{r[final_system[1]]:.6f} | {r[final_system[2]]:.6f} |"
        )
    table_md = '\n'.join(table_lines)

    # Статистика
    stats = df_transformed[[final_system[0], final_system[1], final_system[2]]].describe().to_string()

    report = f"""# Отчёт по преобразованию координат

Дата создания: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Параметры преобразования

- dx = {params['dx']}
- dy = {params['dy']}
- dz = {params['dz']}
- wx = {params['wx']}
- wy = {params['wy']}
- wz = {params['wz']}
- m = {params['m']}

## Формула преобразования

$$ {formula_latex} $$

## Таблица координат до и после преобразования

{table_md}

## Общая статистика

    {stats}
"""
    return report


def _read_and_validate(file_bytes):
    # Чтение Excel и проверка наличия нужных столбцов
    df = pd.read_excel(BytesIO(file_bytes))
    for col in ('X', 'Y', 'Z'):
        if col not in df.columns:
            raise HTTPException(status_code=400, detail=f"Столбец {col} не найден в файле")
    return df


@app.get("/")
def read_root():
    return {
        "message": "Coordinate Transformation API работает",
        "endpoints": {
            "/process-excel/": "POST - Excel -> отчёт в Markdown",
            "/process-excel-docx/": "POST - Excel -> отчёт в DOCX"
        }
    }


@app.post("/process-excel/")
async def process_excel(
    file: UploadFile = File(...),
    dx: float = Form(23.557),
    dy: float = Form(-140.844),
    dz: float = Form(-79.778),
    wx: float = Form(-0.00230),
    wy: float = Form(-0.34646),
    wz: float = Form(-0.79421),
    m: float = Form(-0.00000228)
):
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        raise HTTPException(status_code=400, detail="Поддерживаются только файлы .xlsx и .xls")

    contents = await file.read()
    df = _read_and_validate(contents)

    initial_system = ('X', 'Y', 'Z')
    final_system = ('X_new', 'Y_new', 'Z_new')
    params = {'dx': dx, 'dy': dy, 'dz': dz, 'wx': wx, 'wy': wy, 'wz': wz, 'm': m}

    df_out, formula = transform_coordinates(df, initial_system, final_system, params)
    report = generate_markdown_report(df_out, formula, initial_system, final_system, params)

    output = BytesIO(report.encode('utf-8'))
    output.seek(0)
    filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    return StreamingResponse(
        output,
        media_type="text/markdown",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.post("/process-excel-docx/")
async def process_excel_docx(
    file: UploadFile = File(...),
    dx: float = Form(23.557),
    dy: float = Form(-140.844),
    dz: float = Form(-79.778),
    wx: float = Form(-0.00230),
    wy: float = Form(-0.34646),
    wz: float = Form(-0.79421),
    m: float = Form(-0.00000228)
):
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        raise HTTPException(status_code=400, detail="Поддерживаются только файлы .xlsx и .xls")

    contents = await file.read()
    df = _read_and_validate(contents)

    initial_system = ('X', 'Y', 'Z')
    final_system = ('X_new', 'Y_new', 'Z_new')
    params = {'dx': dx, 'dy': dy, 'dz': dz, 'wx': wx, 'wy': wy, 'wz': wz, 'm': m}

    df_out, formula = transform_coordinates(df, initial_system, final_system, params)
    report = generate_markdown_report(df_out, formula, initial_system, final_system, params)

    # Конвертация Markdown в DOCX через pandoc
    with tempfile.TemporaryDirectory() as tmpdir:
        md_path = os.path.join(tmpdir, 'report.md')
        docx_path = os.path.join(tmpdir, 'report.docx')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(report)
        subprocess.run(['pandoc', md_path, '-o', docx_path], check=True)
        with open(docx_path, 'rb') as f:
            docx_bytes = f.read()

    output = BytesIO(docx_bytes)
    output.seek(0)
    filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )