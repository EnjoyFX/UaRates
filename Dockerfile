FROM python:3.7-alpine

WORKDIR uarates

COPY requirements.txt uarates.py ./

RUN pip install -r requirements.txt && \
    python uarates.py EUR,USD
