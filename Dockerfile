# Build stage — compile TypeScript
FROM node:22-slim AS build

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src/ src/

RUN npm run build

# Production stage — runtime only
FROM node:22-slim

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY --from=build /app/dist/ dist/

EXPOSE 8030

ENV PORT=8030

CMD ["node", "dist/index.js"]
