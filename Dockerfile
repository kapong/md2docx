# syntax=docker/dockerfile:1

# Build arguments for multi-arch
ARG TARGETARCH

# Base stage
FROM rust:1-slim AS base
RUN apt-get update && apt-get install -y musl-tools pkg-config libssl-dev && rm -rf /var/lib/apt/lists/*
RUN cargo install cargo-chef --locked

# Planner stage
FROM base AS planner
WORKDIR /app
COPY . .
RUN cargo chef prepare --recipe-path recipe.json

# Builder stage
FROM base AS builder
ARG TARGETARCH
WORKDIR /app

# Add the appropriate musl target based on architecture
RUN case "$TARGETARCH" in \
    amd64) rustup target add x86_64-unknown-linux-musl ;; \
    arm64) rustup target add aarch64-unknown-linux-musl ;; \
    esac

COPY --from=planner /app/recipe.json recipe.json

# Build dependencies
ENV RUSTFLAGS="-C target-feature=+crt-static"
RUN case "$TARGETARCH" in \
    amd64) cargo chef cook --release --target x86_64-unknown-linux-musl --recipe-path recipe.json ;; \
    arm64) cargo chef cook --release --target aarch64-unknown-linux-musl --recipe-path recipe.json ;; \
    esac

# Build application
COPY . .
RUN case "$TARGETARCH" in \
    amd64) cargo build --release --target x86_64-unknown-linux-musl --bin md2docx && \
    mv target/x86_64-unknown-linux-musl/release/md2docx /md2docx ;; \
    arm64) cargo build --release --target aarch64-unknown-linux-musl --bin md2docx && \
    mv target/aarch64-unknown-linux-musl/release/md2docx /md2docx ;; \
    esac

RUN echo "md2docx:x:1000:1000::/workspace:/sbin/nologin" > /etc/passwd

# Runtime stage (scratch)
FROM scratch
WORKDIR /workspace

# Copy certificates (required for HTTPS)
COPY --from=builder /etc/ssl/certs/ca-certificates.crt /etc/ssl/certs/
# Copy user information
COPY --from=builder /etc/passwd /etc/passwd
# Copy static binary
COPY --from=builder /md2docx /md2docx

# Run as non-root (user 1000)
USER md2docx
ENTRYPOINT ["/md2docx"]
CMD ["build", "-d", "docs/", "-o", "output/output.docx"]