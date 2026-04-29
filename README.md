# libreoffice-converter

Headless LibreOffice service that converts uploaded `.pptx` files into `.pdf`.
Used as the backend for [Attention Optimizer](https://presentation-coach-two.vercel.app).

## Endpoints

- `GET /` — health
- `POST /convert` — body = raw .pptx (or multipart `file`); returns PDF

## Deploy

Designed for Render.com free tier. Just connect the repo and click Deploy.
