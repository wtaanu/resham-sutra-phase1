FROM node:20-alpine

WORKDIR /app

COPY package.json ./
COPY apps/web/package.json ./apps/web/package.json

RUN npm install

COPY apps/web ./apps/web

WORKDIR /app/apps/web

ARG VITE_API_URL=https://api.reshamsutra.com
ENV VITE_API_URL=$VITE_API_URL

RUN npm run build

EXPOSE 3000

CMD ["npm", "run", "preview"]
