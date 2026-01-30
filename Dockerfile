# syntax=docker/dockerfile:1
# Base stage - use generic linux (Debian) to avoid self-referential static linking issues
FROM --platform=linux/amd64 rust:1-slim AS base
RUN apt-get update && apt-get install -y musl-tools pkg-config libssl-dev && rm -rf /var/lib/apt/lists/*
RUN rustup target add x86_64-unknown-linux-musl
RUN cargo install cargo-chef --locked

# Planner stage
FROM --platform=linux/amd64 base AS planner
WORKDIR /app
COPY . .
RUN cargo chef prepare --recipe-path recipe.json

# Builder stage
FROM --platform=linux/amd64 base AS builder
WORKDIR /app
COPY --from=planner /app/recipe.json recipe.json
# Build dependencies - separate target allows static linking without breaking proc-macros
ENV RUSTFLAGS="-C target-feature=+crt-static"
RUN cargo chef cook --release --target x86_64-unknown-linux-musl --recipe-path recipe.json

# Build application
COPY . .
RUN cargo build --release --target x86_64-unknown-linux-musl --bin md2docx
RUN echo "md2docx:x:1000:1000::/workspace:/sbin/nologin" > /etc/passwd

# Runtime stage (scratch)
FROM --platform=linux/amd64 scratch
WORKDIR /workspace

# Copy certificates (required for HTTPS)
COPY --from=builder /etc/ssl/certs/ca-certificates.crt /etc/ssl/certs/
# Copy user information
COPY --from=builder /etc/passwd /etc/passwd

# Copy static binary
COPY --from=builder /app/target/x86_64-unknown-linux-musl/release/md2docx /md2docx

# Run as non-root (user 1000)
USER md2docx
ENTRYPOINT ["/md2docx"]
CMD ["--help"]