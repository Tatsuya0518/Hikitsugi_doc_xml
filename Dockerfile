FROM rust:slim AS builder

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        pkg-config \
        libssl-dev \
        libbz2-dev \
        liblzma-dev \
        libzstd-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

RUN cargo build --release

FROM debian:bookworm-slim AS runtime

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        libssl3 \
        libbz2-1.0 \
        liblzma5 \
        libzstd1 \
        python3 \
        python3-pip \
        python3-openpyxl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY --from=builder /app/target/release/hikitsugi-doc-xml /usr/local/bin/hikitsugi-doc-xml
COPY --from=builder /app/static ./static
COPY --from=builder /app/templates ./templates

ENV PORT=5000
EXPOSE 5000

CMD ["hikitsugi-doc-xml"]
