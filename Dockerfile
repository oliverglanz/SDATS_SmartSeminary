FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    SMART_SEMINARY_PORT=8501 \
    SMART_SEMINARY_ENABLE_CLOSE_BUTTON=false

WORKDIR /app

RUN useradd --create-home --shell /bin/bash appuser

COPY programs/requirements_local_app.txt /tmp/requirements_local_app.txt
RUN pip install --upgrade pip \
    && pip install -r /tmp/requirements_local_app.txt

COPY 0_source_files /app/0_source_files
COPY programs /app/programs

RUN mkdir -p /app/outputs_app \
    && chown -R appuser:appuser /app

USER appuser

WORKDIR /app/programs

EXPOSE 8501

CMD ["sh", "-c", "python -m streamlit run app/streamlit_app.py --server.headless=true --server.address=0.0.0.0 --server.port=${SMART_SEMINARY_PORT}"]
