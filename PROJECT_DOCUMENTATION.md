# Metalman Auto-Engineer — Complete Project Documentation

> **Purpose**: A complete, code-verified snapshot of this project — what it does, how it's built, and how it actually behaves today — written to be handed to an AI assistant (or a new team member) as full context for an open-ended brainstorm. Every technical claim below was checked against the actual source files, not assumed. Where something in older planning docs turned out to be outdated (e.g. a status, a model name, a status flag), that's called out explicitly rather than silently corrected.

---

## 1. What the product is

Metalman Auto-Engineer automates the creation of PPAP/APQP manufacturing engineering documents. Given **one CAD assembly (STEP file)**, **one feasibility/master-data spreadsheet**, and a folder of **2D drawing PDFs** for a single manufactured part, it generates a stack of manufacturing/quality documents as Excel workbooks — using CAD geometry processing, computer vision, and a vision-LLM fallback to extract and clean up part imagery along the way.

It's built as a **FastAPI (Python) backend** + **React/Vite (TypeScript) frontend**, deployable via Docker to either an Oracle Cloud VM or Render.

Login (dev/demo credentials, auto-seeded into the DB on every backend start — see `main.py`):

| Email | Password | Role |
|---|---|---|
| `admin@metalman.io` | `admin123` | Admin |
| `demo@metalman.io` | `password123` | Engineer |

---

## 2. Tech stack

### Backend
- **FastAPI + Uvicorn** — the web API and app server
- **SQLAlchemy + SQLite** — used *only* for auth (`User`/`Role`/`Permission` tables); there is no `Job`/`Task`/`Part` table anywhere
- **Pydantic** — request/response validation
- **python-jose + passlib (pbkdf2_sha256)** — JWT auth (HS256, 8h expiry)
- **pandas + openpyxl** — the real workhorse: parses feasibility spreadsheets, writes every generated Excel document (styling, formulas, embedded images)
- **PyMuPDF (`fitz`)** — PDF rendering to images, PDF text search, SVG→PNG rasterization
- **OpenCV (`cv2`, headless)** — contour-based "surgical cleanup" of CAD/drawing views
- **Pillow** — image crop/trim/centering
- **cadquery 2.7.0 + cadquery-ocp (OpenCascade/XCAF bindings)** — the primary CAD engine: STEP import/export, STL export (3D viewer), SVG rendering with hidden-line removal, and direct traversal of the STEP assembly tree to pull out named sub-parts
- **FreeCAD** (external executable) — a **fallback only**, and only reachable on Windows; the code explicitly returns `None` for the FreeCAD command on non-Windows, so it **never runs in the actual Linux production deployment**
- **Vision AI** — `services/ai_provider.py` switches between:
  - `AI_PROVIDER=ollama` (default): an OpenAI-SDK-compatible client pointed at `OLLAMA_BASE_URL`, model **`gemma4:31b-cloud`** (Ollama Cloud, not a local model)
  - `AI_PROVIDER=gemini`: Google's `google-generativeai`, model **`gemini-2.0-flash`** — used on deployments where Ollama isn't reachable
  - This is invoked only as a **last-resort tier** — see §5.1 below
- **pywin32** — Windows-only, drives real desktop Excel via COM automation for high-fidelity PDF preview / "Edit in Excel"
- **celery + redis** — present in `requirements.txt` but **not actually used anywhere**. `worker/tasks.py` is a dead `MockCelery` stub. The real background job execution is FastAPI's in-process `BackgroundTasks`. Neither service is defined in either `docker-compose.yml` or `docker-compose.prod.yml` — confirmed by reading both files directly.

### Frontend
- **React + TypeScript + Vite**
- **shadcn/ui** (Radix UI primitives) + **Tailwind CSS** — "Steel Blue + Safety Orange" industrial B2B palette, Space Grotesk/Inter/JetBrains Mono fonts
- **react-router-dom v6**
- **@react-three/fiber + @react-three/drei + three.js** — renders the generated STL in a live, rotating 3D viewer (`CADWireframe.tsx`)
- **pdfjs-dist** — client-side PDF rendering for the manual "Cropper" tool
- **lucide-react** — icon set used throughout, including the illustrated landing-page diagrams
- **axios / fetch** — all data fetching is raw REST calls + polling; no WebSockets/SSE anywhere
- Installed **but confirmed never actually used anywhere in the source**:
  - `@tanstack/react-query` — initialized at the app root, zero `useQuery`/`useMutation` calls exist
  - `zod` + `react-hook-form` — present in `package.json`; no `zodResolver`, no `z.object(...)`, no schema validation calls found anywhere in `frontend/src`

### Infra / Deployment
- **Docker** — separate `backend` and `frontend` images, wired together by `docker-compose.yml` (verified: only two services defined — `backend` and `frontend`, no `redis`/`celery` service in either compose file)
- **Nginx** (inside the frontend container) — reverse-proxies `/api` and `/outputs` to the backend so only one public port is needed
- **SQLite** file, persisted via the `backend_data` Docker volume
- Two deployment targets referenced in the repo: an **Oracle Cloud VM** (`DEPLOYMENT.md`, Docker Compose) and **Render** (the frontend's `fetchWithRetry` has explicit cold-start handling language for Render's free tier)

---

## 3. Backend architecture

### 3.1 `backend/main.py`
- Creates the FastAPI app, mounts `routes.router` at `/api` and `auth.router` at `/api/auth`.
- `CORSMiddleware` is wide open (`allow_origins=["*"]`).
- On import (not a startup hook), seeds the two demo users above into SQLite every process start.
- Mounts `outputs/` as static files at `/outputs` — **no auth check**, so any generated document is downloadable if you know/guess the filename.
- Also mounts `stitch_suite/` (see §6.6) at `/stitch-demo`.
- Docker entrypoint runs `uvicorn main:app --workers 1` — **load-bearing**, not incidental (see §8.2).

### 3.2 `backend/api/routes.py` — endpoints and the pipeline

Module-level global state (not per-request, not per-tenant):
- `TASKS: dict` — `task_id → {status, progress, message, stl_url, files, error}`, persisted to `tasks_db.json` on every write, no file locking.
- `current_welding_row` — a single shared cursor into one global welding-checksheet workbook.
- `upload_retention.json` — keeps the 3 most-recently-completed jobs' upload scratch files; `outputs/` is never cleaned up.

**Endpoints:**

| Endpoint | Method | Purpose |
|---|---|---|
| `/api/analyze` | POST | Pipeline entry point. Takes `cad_file`, `feasibility_file`, `drawings[]`, and an `outputs` flag-set (which of the 7 documents to generate). Mints `task_id = uuid4()[:8]`, schedules the pipeline as a `BackgroundTasks` job, returns immediately. |
| `/api/status/{task_id}` | GET | Polled every 1.5s by the frontend. |
| `/api/correct` | POST | Applies hand-typed cell corrections to an already-generated file. |
| `/api/preview/{filename}` | GET | High-fidelity preview via Excel COM (Windows only) or a pandas-HTML fallback. |
| `/api/edit/{filename}` | POST | Launches desktop Excel **on the server**, not the client. |
| `/api/save-crop`, `/api/welding/set-reference-image`, `/api/download-excel`, `/api/reset-excel` | — | The standalone "Cropper" tool — see §5 (Welding Check Sheet). |

**The pipeline itself** (`process_engineering_task`, runs synchronously inside the background task):
1. Resolve canonical `(part_no, part_desc)` once (`utils/part_info.py`), reused by every downstream generator.
2. `cad_service.export_step_to_stl()` — STL for the live 3D viewer.
3. `bom_generator.generate_bom_excel()` — BOM + isometric image extraction (populates a shared image cache reused by later steps).
4. `pfd_renderer.*` — 4 PFD sheets (the deepest CAD-geometry step).
5. `tooling_generator.generate_tooling_list()`.
6. `fitment_generator.generate_fitment_check_sheet()` (reuses BOM's cached images).
7. `pfmea_generator.generate_pfmea_excel()`.
8. `control_plan_generator.generate_control_plan_excel()` — **confirmed currently live and wired in** (`routes.py` line ~362, gated only by whether the `control_plan` output was requested). Depends on PFD having already run in the same job.
9. `fixture_plan_generator.generate_fixture_plan()`.

All sequential, single-threaded, single-process — no parallelism across parts or across document types within one part.

### 3.3 `backend/api/auth.py`
Fully implemented JWT flow — but `get_current_user`/`check_role` are **never applied to any business endpoint**. Combined with open CORS and the unauthenticated `/outputs` mount, the API is functionally open regardless of frontend login state. The frontend's `ProtectedRoute` is a client-only redirect guard, not real access control.

### 3.4 `backend/services/` — the 7 document generators
| File | Generates |
|---|---|
| `bom_generator.py` | Bill of Materials + the isometric-image extraction/cleanup pipeline (reused by 2 other endpoints too) |
| `pfd_renderer.py` (the largest, ~1100 lines) | Process Flow Diagram — the real CAD-geometry engine |
| `tooling_generator.py` | Fixture/Tooling Master List |
| `fitment_generator.py` | Fitment Checksheet |
| `pfmea_generator.py` | Process FMEA (Assembly / Sheet Metal / BOP sheets) |
| `control_plan_generator.py` | Control Plan (4 sheets, cross-references the PFD output) |
| `fixture_plan_generator.py` | Fixture PM Master Plan |
| `ai_provider.py` | Not a generator — the single chokepoint for vision-LLM calls, shared by BOM/welding |
| `correction_service.py` | Applies hand-typed cell corrections post-generation |
| `cad_slicer.py`, `feasibility_parser.py`, `tooling_extractor.py` | **Dead placeholder stubs**, unused — evidence of an earlier, more decoupled design that was never finished |

### 3.5 `backend/assets/` — the knowledge base and templates
- Blank Excel **templates** every generator copies before filling (BOM, Control Plan, PFMEA, PFD, Fitment, Fixture Plan, Tooling, Welding Checksheet).
- `control_plan_dicts/{assy,bop,sheetmetal}/*.json` and `pfmea_dicts/{assembly,bop,sheetmetal}/*.json` — the **actual engineering knowledge base**: per-process-type standard failure modes / control checkpoints, applied identically regardless of specific part number. This is the part of the system that's already naturally reusable across parts.

### 3.6 Database & persistence
- SQLite schema is **only** `User`/`Role`/`Permission` — no job/task/document tables exist.
- Job state lives entirely in the in-memory `TASKS` dict + a flat, unlocked `tasks_db.json` snapshot. `correction_service.py` independently re-reads that same JSON file (a second, separately-synced reader).
- File storage is local disk only — `uploads/` (namespaced by `{task_id}_...`, garbage-collected after 3 completed jobs) and `outputs/` (namespaced by **`{part_no}`**, not `task_id` — so two jobs for the same part number can overwrite each other's files, and nothing is ever cleaned up).

---

## 4. The 7 generated documents (detailed pipeline per document)

*(This section mirrors the verified content of `Process_Details_and_Feasibility_Updated.docx`, already produced and saved in this project's root — see that file for the full table with "Required Documents," "Accuracy," and "Technical Challenges" columns per sheet.)*

1. **Bill of Materials** — pandas extracts part data from the feasibility sheet. Isometric images are located via a 3-tier cascade: (1) deterministic PDF text search for a "REFERENCE ISOMETRIC VIEW" caption — no AI call; (2) OpenCV contour detection if no caption; (3) only as a last resort, a grid-overlay sent to the vision-LLM. Then OpenCV thresholding/contour/dilation strips dimension lines and text before the cleaned image is embedded via openpyxl.
2. **Process Flow Diagram** (4 sheets: Sub Assy Index, Sub Assy, Sheetmetal Parts, BOP & Hardware) — resolves the operation→parts routing matrix (with recursive phantom/sub-assembly resolution), then extracts per-operation 3D geometry via **cadquery-ocp first** (cross-platform, runs in production), falling back to a FreeCAD subprocess only on Windows dev if OCP fails.
3. **Fixture/Tooling Master List** — pulls fixture photos already embedded inside the feasibility workbook itself (fuzzy row-offset matching to each "Fixture No"), rather than generating new images.
4. **Fitment Checksheet** — reuses the cleaned CAD images generated during the BOM step, laid out per-part with dynamic row replication.
5. **Process FMEA** (3 sheets: Assembly, Sheet Metal, BOP) — dict/template-driven from the `pfmea_dicts/` knowledge base, with formula-based RPN calculation. Severity/Occurrence/Detection scoring itself is still manual — it requires domain expertise the tool doesn't automate.
6. **Control Plan** (4 sheets, mirrors PFD's structure) — **live today**, not on hold. Cross-references the already-generated PFD output and the `control_plan_dicts/` knowledge base.
7. **Fixture PM Master Plan** — builds a monthly preventive-maintenance schedule with day-wise (1–31) tracking from the master data sheet.

**Separately, outside the main `/api/analyze` pipeline**: the **Welding Check Sheet** tool (the "Cropper") lets an operator drag-crop weld symbols from a ballooned drawing PDF directly into Excel. It's real and working, but it operates on **one single shared global workbook** with no per-user/per-part scoping — a concurrency risk worth knowing about before multiple people use it at once.

**Not implemented anywhere in the codebase** (confirmed by direct search): **Checking Aids** and **PDI Report** — both appear in older planning docs as "YES" but don't exist as generators today.

---

## 5. Frontend architecture

### 5.1 Routing (`App.tsx`)
`/` (public landing page) → `/login` (public) → `/dashboard`, `/correction`, `/cropper`, `/stitch` (all `ProtectedRoute`-wrapped, client-side-only auth check) → `*` 404.

### 5.2 Landing page (`pages/Index.tsx` + `components/landing/`)
Recently redesigned specifically to make the product comprehensible to non-manufacturing demo audiences (who previously couldn't follow jargon like PPAP/APQP/BOM). Current structure:
- **Hero** — headline/CTA unchanged, but the old fake static "progress bar" mock was replaced with `ProcessDiagram.tsx`: a real illustrated 3-tier diagram (3 inputs → automation engine → 7 real output documents, sourced from the same names as `Dashboard.tsx`'s `OUTPUTS` array so the two never disagree), visible at all breakpoints (not desktop-only like the original).
- **`WhatIsPpap.tsx`** — a dedicated, illustration-driven explainer section: two parallel icon timelines (APQP's 5 real phases vs. PPAP's document checklist converging into a "Proof Package / Approved" stamp), animated in via a small custom `useInView` scroll-reveal hook, converging into a shared "Part accepted by the customer" node.
- **`SoftwarePipeline.tsx`** (the `#features` section) — a 5-node illustrated pipeline (Input Files → Template Library → Knowledge Base → Processing → Generated Documents) where **hovering each node reveals real, grounded detail** via a `HoverCard` popover: exact template/document names, and which real industry standards apply (ASME Y14.5, AIAG-VDA FMEA Handbook, AIAG APQP & Control Plan Manual, IATF 16949) — tagged only where a genuine standard applies, not fabricated across the board.
- **`ApqpDeepDive.tsx` / `PpapDeepDive.tsx`** — full-screen dialogs (triggered by "See all 5 phases" / "See all 18 elements" buttons) that drill into the *real* AIAG APQP 5-phase structure and the *real* AIAG 18-element PPAP checklist, with only the phases/elements Metalman's tool genuinely automates flagged — the other elements shown honestly as outside current scope (no overclaiming).
- All of this reuses the existing design system (`--primary` Steel Blue, `--accent` Safety Orange, existing Tailwind keyframes) — no new dependencies were added for any of it; `HoverCard` was an already-installed-but-previously-unused shadcn primitive.

### 5.3 Dashboard (`pages/Dashboard.tsx`) — the core single-part workflow
Models **exactly one job at a time** — `feasibilityFile`, `cadFile`, `drawingFiles[]`, a 7-item `outputs` toggle set, a single `taskId`. No job history/list/queue UI anywhere. `ProcessingView.tsx` polls `/api/status/{taskId}` every 1.5s, streams the STL into `CADWireframe.tsx` as soon as it's ready, and on completion shows Preview/Download actions per file.

### 5.4 Other pages
- **`Correction.tsx`** — queue up multiple cell-level corrections against a completed job, submitted in bulk to `/api/correct`. The document→column mapping is duplicated between this file and the backend's `correction_service.py`.
- **`Cropper.tsx`** / **`PdfCropper.tsx`** — the standalone Welding Check Sheet tool described above.
- **`StitchDemo.tsx`** — renders `stitch_suite/`, a folder of **static AI-generated UI mockups** (login variants, an AI assistant chat screen, a BI dashboard, a task board, etc.) — product/UX explorations, not working code, served as an in-app demo gallery.

### 5.5 State & communication
No Redux/Zustand/Context — plain `useState` per page. REST + polling only, no real-time channel. `react-query` is installed but genuinely unused anywhere (verified), so there's no existing caching/invalidation layer a future feature could build on without adopting real usage of it.

---

## 6. Known architecture constraints (verified, not guessed)

These matter for almost any "what should we build next" conversation:

1. **No real task queue despite the dependency being present.** Celery/Redis are dead weight; real jobs run via FastAPI `BackgroundTasks` in the same process as the HTTP server.
2. **`--workers 1` is load-bearing.** Job state lives in per-process memory (`TASKS` dict); multiple worker processes would each have an inconsistent view of task state and break status polling.
3. **No job/task database table** — only `User`/`Role`/`Permission` exist in SQLite.
4. **Output filenames are keyed by part number, not job ID** — two jobs for the same part can silently overwrite each other's files.
5. **One global mutable singleton for the Welding/Cropper tool** — shared across every user, no per-task/part scoping, hardcoded 27-row capacity.
6. **Local filesystem is the only storage layer** — no S3/GCS abstraction; `/outputs` is a public, unauthenticated static mount.
7. **Auth exists but isn't enforced anywhere on the pipeline** — JWT gates the frontend only; the API itself is open.
8. **Two features require a colocated, licensed, Windows Excel install** (`/api/preview` COM automation, `/api/edit`'s `os.startfile`) — can't run in a Linux container fleet beyond their existing lower-fidelity fallbacks.
9. **CAD extraction's FreeCAD fallback is silently unavailable in production** (Linux) — only the OCP path works there; any assembly that needed the FreeCAD fallback in Windows dev testing should be spot-checked in prod.
10. **No output retention policy** — `outputs/` grows forever; only `uploads/` scratch is garbage-collected.
11. **Vision-AI calls are unbatched**, one drawing at a time, inline inside the sequential per-BOM-row loop, with per-job (not cross-job) image caching that's defeated by a `force_rewrite=True` default.
12. **The frontend models exactly one job at a time by construction** — no array/list dimension in Dashboard state, no job history UI, and the one library that could support a "job list" view (`react-query`) is installed but unused.
13. **Dead scaffolding** (`worker/tasks.py`, `cad_slicer.py`, `feasibility_parser.py`, `tooling_extractor.py`) hints at an earlier, more decoupled design that was abandoned mid-build.

## 7. What's already well-built and reusable
- The PFMEA/Control Plan JSON knowledge-base dicts already encode *process-type* knowledge (sheet metal / assembly / BOP), not per-part data — this layer generalizes to "many parts" as-is.
- `utils/image_utils.place_image_centered()` and the BOM isometric-extraction cascade are genuinely reusable, part-agnostic infrastructure.
- `utils/part_info.get_canonical_part_info()` already solved "single source of truth for part identity within one job" — a natural next step is promoting that into a first-class persisted entity.
- `ai_provider.py` is already provider-agnostic (Ollama/Gemini) and could be made batching-aware without touching call sites.
- The landing page's PPAP/APQP drill-downs are a good model for "explain accurately, flag what's not built yet" — worth carrying into any future customer-facing material.

---

## 8. Suggested framing for a brainstorm
Depending on what you want to explore with this document, natural threads to pull on:
- **Scaling one part → many parts**: real job queue, per-job output namespacing, object storage, enforced auth/multi-tenancy, batch UI.
- **Closing the PPAP coverage gap**: today's 7 documents cover 4 of the AIAG 18-element checklist (Process Flow Diagram, Process FMEA, Control Plan, Dimensional Results) — which of the remaining 14 are realistically automatable vs. inherently manual (lab tests, physical samples)?
- **Productizing the demo-ready landing page work** into an actual public marketing site, or turning the `stitch_suite` mockups into real features.
- **Security hardening** before this goes anywhere near a real customer network — enforcing the JWT/RBAC that already exists but isn't wired in.
