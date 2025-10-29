ARG PYTHON_VERSION=3.9
ARG DEBIAN_VERSION=bookworm

FROM python:${PYTHON_VERSION}-${DEBIAN_VERSION} AS build

WORKDIR /usr/src/app

COPY requirements.txt .

RUN apt-get -y update && \
    apt-get install --no-install-recommends --no-install-suggests -y \
        build-essential && \
    apt-get autoremove && \
    rm -rf /var/lib/apt/lists/*

RUN python -m venv --prompt spscs venv && \
    ./venv/bin/python -m pip install --no-cache-dir -r requirements.txt


ARG PYTHON_VERSION
ARG DEBIAN_VERSION

FROM python:${PYTHON_VERSION}-slim-${DEBIAN_VERSION} AS release

WORKDIR /usr/src/app

RUN apt-get -y update && \
    apt-get install --no-install-recommends --no-install-suggests -y \
        libtk8.6 && \
    apt-get autoremove && \
    rm -rf /var/lib/apt/lists/*

COPY --from=build /usr/src/app/venv ./venv

COPY . .
    
RUN mkdir -p downloads ternary_plots uploads /tmp/spscs && \
    chmod -R g=u downloads ternary_plots uploads /tmp/spscs && \
    chmod +x start.sh

ENV NUM_WORKERS=2 \
    MPLCONFIGDIR=/tmp/spscs

EXPOSE 8000

CMD ["./start.sh"]
