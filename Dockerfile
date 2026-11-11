FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Seoul
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY inventory.py app.py test_inventory.py test_app.py ./
COPY templates ./templates
EXPOSE 8000
# 엑셀 파일 하나를 여러 요청이 함께 쓰므로 프로세스(worker)는 1개만 두고, 동시 접속은 스레드로 처리한다
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "1", "--threads", "4", "app:create_app()"]
