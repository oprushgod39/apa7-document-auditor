# --- Build stage -------------------------------------------------------------
FROM node:22-alpine AS build
WORKDIR /app
COPY package.json ./
COPY server/package.json server/
COPY web/package.json web/
RUN npm install --no-audit --no-fund
COPY server server
COPY web web
RUN npm run build --workspace=server && npm run build --workspace=web

# --- Runtime stage ------------------------------------------------------------
FROM node:22-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY package.json ./
COPY server/package.json server/
RUN npm install --omit=dev --workspace=server --no-audit --no-fund
COPY --from=build /app/server/dist server/dist
COPY --from=build /app/web/dist web/dist
ENV PORT=8000 \
    FRONTEND_DIST_DIR=/app/web/dist \
    STORAGE_DIR=/tmp/apa7-auditor-storage
EXPOSE 8000
USER node
CMD ["node", "server/dist/index.js"]
