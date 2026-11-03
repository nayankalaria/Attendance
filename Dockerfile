FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    DATA_DIR=/data \
    DISPLAY=:99 \
    PORT=10000

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    xvfb x11vnc novnc websockify nginx supervisor \
    ca-certificates fonts-liberation fonts-noto-color-emoji \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt \
    && python -m playwright install --with-deps chromium

COPY app.py students.xlsx ./
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY supervisord.conf /etc/supervisor/conf.d/attendance.conf

RUN mkdir -p /data /run/nginx /var/log/supervisor \
    && chmod 700 /data

EXPOSE 10000

CMD ["/bin/sh", "-c", "if [ -z \"$VNC_PASSWORD\" ] || [ ${#VNC_PASSWORD} -lt 8 ]; then echo 'ERROR: Set VNC_PASSWORD to at least 8 characters in Render environment variables.'; exit 1; fi; exec /usr/bin/supervisord -n -c /etc/supervisor/supervisord.conf"]
