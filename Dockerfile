# Railway build recipe. Railway reads this file automatically.
# Uses a small Python image, installs the app's packages, then Chromium for the PDF step.
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
 && python -m playwright install --with-deps chromium

COPY . .

# Railway sets $PORT. 2 workers is plenty for an internal tool; PDF can take a few seconds.
CMD gunicorn app:app --bind 0.0.0.0:${PORT:-8000} --workers 2 --timeout 180
