# syntax = docker/dockerfile:1.7-labs
ARG PYTHON_VERSION="3.13"
FROM mcr.microsoft.com/devcontainers/python:${PYTHON_VERSION}

USER root

RUN apt-get update \
    && apt-get install --no-install-recommends -y \
        build-essential \
        ca-certificates \
        curl \
        git \
        just \
    && rm -rf /var/lib/apt/lists/*

USER vscode
WORKDIR /workspace

RUN mkdir -p /home/vscode/.cache/uv

ENV PATH="/home/vscode/.local/bin:${PATH}" \
    UV_LINK_MODE=copy \
    UV_PROJECT_ENVIRONMENT=/home/vscode/.venv \
    UV_NO_MANAGED_PYTHON=1 \
    UV_PYTHON_DOWNLOADS=never

RUN curl -LsSf https://astral.sh/uv/install.sh | sh

RUN uv tool install prek >/dev/null

CMD ["sleep", "infinity"]
