# The media probe job. Built by Cloud Build from this folder (scripts/deploy-media-probe.sh).
# Debian trixie's ffmpeg (7.1) reads ProRes in MOV and MXF, H.264 and H.265.
FROM node:24-trixie-slim
RUN apt-get update \
  && apt-get install -y --no-install-recommends ffmpeg ca-certificates \
  && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json ./
COPY src ./src
USER node
CMD ["node", "src/main.js"]
