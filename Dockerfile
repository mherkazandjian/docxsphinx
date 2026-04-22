# syntax=docker/dockerfile:1.6
ARG PYTHON_VERSION=3.12
FROM python:${PYTHON_VERSION}-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        git \
        make \
        pandoc \
 && rm -rf /var/lib/apt/lists/*

ARG UID=1000
ARG GID=1000
RUN groupadd -g ${GID} dev \
 && useradd -m -u ${UID} -g ${GID} -s /bin/bash dev

WORKDIR /workspace

COPY --chown=dev:dev pyproject.toml README.md ./
COPY --chown=dev:dev src/ ./src/

RUN pip install -e ".[dev]"

USER dev
CMD ["bash"]
