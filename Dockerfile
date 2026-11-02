FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 HLS_DATA_DIR=/data PORT=8080
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN useradd -m app && mkdir -p /data && chown -R app:app /data /app && chmod +x /app/docker-entrypoint.sh
USER app
VOLUME /data
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s CMD python3 -c "import urllib.request,os;urllib.request.urlopen('http://localhost:'+os.environ.get('PORT','8080')+'/')" || exit 1
ENTRYPOINT ["/app/docker-entrypoint.sh"]
CMD ["python3", "main.py"]
