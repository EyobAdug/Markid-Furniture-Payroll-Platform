FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn
COPY app.py schema.sql ./
ENV MARKID_MODE=server MARKID_DB=/data/markid.db MARKID_HTTPS=1 PORT=8000
VOLUME /data
EXPOSE 8000
# one worker + threads: SQLite is a single-writer database
CMD ["sh", "-c", "gunicorn -w 1 --threads 8 --timeout 300 -b 0.0.0.0:${PORT} app:app"]
