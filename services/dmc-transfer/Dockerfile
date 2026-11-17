# The DMC transfer job. Built by Cloud Build from this folder (scripts/deploy-dmc-transfer.sh).
FROM node:24-slim
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY src ./src
USER node
CMD ["node", "src/main.js"]
