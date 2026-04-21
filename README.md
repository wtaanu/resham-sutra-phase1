# Resham Sutra Platform

Starter monorepo for the Resham Sutra enquiry-to-order system.

## Apps

- `apps/web` - React + TypeScript admin interface
- `apps/api` - Node.js + TypeScript API

## Local infrastructure

- `docker-compose.yml` runs:
  - `web`
  - `api`
  - `mysql`
  - `n8n`

## Getting started

1. Copy `.env.example` to `.env`.
2. Run `npm install`.
3. Run `docker compose up --build`.

The frontend will be available at `http://localhost:3000`.
The API will be available at `http://localhost:4000`.
The n8n editor will be available at `http://localhost:5678`.

