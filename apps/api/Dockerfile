FROM node:20-alpine

WORKDIR /app

COPY package.json ./
COPY apps/api/package.json ./apps/api/package.json

RUN npm install

COPY apps/api ./apps/api
COPY templates ./templates

WORKDIR /app/apps/api

RUN npm run build

EXPOSE 4000

CMD ["npm", "run", "start"]
