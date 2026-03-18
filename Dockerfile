FROM node:22-slim

WORKDIR /app

# Copy package files and install production dependencies
COPY package*.json ./
RUN npm ci --production

# Copy pre-built TypeScript output
COPY dist/ dist/

EXPOSE 8030

ENV PORT=8030

CMD ["node", "dist/index.js"]
