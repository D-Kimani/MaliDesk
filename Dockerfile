# API image. Build the frontend separately with `npm --prefix frontend run build`.
FROM node:22-alpine
WORKDIR /app/server
COPY server/package.json ./
RUN npm install --omit=dev --no-audit --no-fund
COPY server/ ./
ENV NODE_ENV=production
EXPOSE 3001
CMD ["node","index.js"]
