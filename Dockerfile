FROM python:3.7-alpine AS base

RUN apk update && apk add --no-cache  \
    git \
    musl-dev

RUN pip install wheel

# for alpine based OS need to install git:
# RUN apk add --no-cache git

RUN git clone https://github.com/EnjoyFX/uarates.git

WORKDIR uarates

RUN pip install -r requirements.txt && \
    python uarates.py EUR,USD
