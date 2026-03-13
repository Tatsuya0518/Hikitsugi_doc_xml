FROM rust:1.75-slim AS builder

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
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY --from=builder /app/target/release/hikitsugi-doc-xml /usr/local/bin/hikitsugi-doc-xml
COPY --from=builder /app/static ./static
COPY --from=builder /app/templates ./templates
COPY --from=builder /app/data ./data
COPY --from=builder /app/cache ./cache

ENV PORT=5000
EXPOSE 5000

CMD ["hikitsugi-doc-xml"]
