FROM node:20-alpine
WORKDIR /app

# Instala dependencias primero (mejor caché de capas)
COPY package*.json ./
RUN npm install --omit=dev

# Copia el código (los secretos .env y los .json de credenciales quedan
# EXCLUIDOS por .dockerignore; se inyectan en runtime via env_file + volumen).
COPY . .

EXPOSE 3000
ENV PORT=3000
CMD ["node", "index.js"]
