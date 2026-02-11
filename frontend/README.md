# Legal Document Formatter — React (Vite + TipTap)

Run the FastAPI backend from the **project root** first, then start the frontend.

## Backend (from project root)

```bash
cd path/to/formatter
pip install fastapi uvicorn python-multipart  # if not already installed
uvicorn api.main:app --reload --port 8000
```

## Frontend

```bash
cd frontend
npm install
npm run dev
```

Open http://localhost:5173. The dev server proxies `/api` to http://localhost:8000.

## Build

```bash
cd frontend
npm run build
```

Static files go to `frontend/dist`. Point your backend or nginx at this folder to serve the SPA, or use `npm run preview` to test the build locally.
