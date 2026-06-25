FROM rocker/tidyverse:4.4.2

# uv install
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv
COPY --from=ghcr.io/astral-sh/uv:latest /uvx /usr/local/bin/uvx

# System deps for Python packages
RUN apt-get update && apt-get install -y \
    python3 \
    python3-dev \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /project

# Python deps via uv
COPY pyproject.toml .
RUN uv sync

# R packages for ITS
RUN Rscript -e "install.packages(c('readxl','dplyr','tidyr','ggplot2','segmented','nlme','lubridate','writexl'), repos='https://cloud.r-project.org')"

COPY scripts/ scripts/
COPY data/ data/
