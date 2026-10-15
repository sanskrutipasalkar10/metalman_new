# Metalman Auto-Engineer — Technical Context for Scaling Brainstorm

> **Purpose of this document**: This is a complete technical snapshot of the Metalman Auto-Engineer application as it exists today, written to be handed to an AI assistant (or a new engineer) as context for a brainstorm on one specific question: **how do we evolve this from a system that processes one part at a time to one that can process many parts** (batch/fleet scale)? Every claim below is grounded in the actual code (file paths, function names, line-level behavior), not a guess.

---

## 1. What the app does (one sentence)

Given one CAD assembly (STEP file) + one feasibility spreadsheet + a folder of 2D drawing PDFs for a single manufactured part, the system automatically generates a stack of manufacturing/quality engineering documents (BOM, Process Flow Diagram, Tooling List, Fitment Checksheet, PFMEA, Control Plan, Fixture Plan) as Excel workbooks, using a mix of CAD geometry processing, computer vision, and vision-LLM calls to extract and clean up part imagery.

It is a **FastAPI (Python) backend** + **React/Vite (TypeScript) frontend**, currently deployed as a single Docker container (backend) + static frontend (per `render.yaml` / `DEPLOYMENT.md`), targeting Render.com.

---

## 2. Tech stack

**Backend**
- **FastAPI** + **uvicorn** — HTTP API and app server (`main.py`, `--workers 1`, see §6.2)
- **SQLAlchemy + SQLite** — auth only (`User`/`Role`/`Permission` tables); no job/task/part tables exist in the DB
- **Pydantic** — request/response models
- **python-jose + passlib** — JWT auth (HS256, pbkdf2_sha256)
- **pandas / openpyxl** — the actual workhorse for reading feasibility spreadsheets and writing every generated Excel document (styling, formulas, embedded images)
- **PyMuPDF (`fitz`)** — PDF rendering to images, PDF text search, SVG→PNG rasterization
- **OpenCV (`cv2`, headless)** — contour-based image isolation ("surgical cleanup" of CAD/drawing views)
- **Pillow** — image crop/trim/centering
- **cadquery 2.7.0 + cadquery-ocp (OCCT bindings)** — STEP import/export, STL export for the 3D viewer, SVG rendering with hidden-line removal, and (via OCP's XCAF API) direct traversal of the STEP assembly tree to extract named sub-parts
- **FreeCAD** (external executable, Windows-only) — fallback CAD renderer if the OCP path fails; explicitly disabled on non-Windows
- **openai SDK** (pointed at an Ollama-compatible endpoint) and **google-generativeai (Gemini)** — interchangeable vision-LLM backends behind one abstraction (`services/ai_provider.py`), used only as a last-resort tier for locating an isometric view on a drawing when deterministic/CV methods fail
- **pywin32** (Windows only) — drives real desktop Excel via COM automation for high-fidelity PDF preview and "Edit in Excel"
- **celery + redis** — **present in requirements.txt but entirely unused** (`worker/tasks.py` is a dead `MockCelery` stub); the real background job execution is FastAPI's in-process `BackgroundTasks`

**Frontend**
- **React + Vite + TypeScript**, **shadcn/ui** (Radix-based component library), **Tailwind CSS**
- **react-router-dom v6** — client-side routing
- **@tanstack/react-query** — wired at the app root but **not actually used anywhere** (all data fetching is raw `fetch`/`axios`)
- **@react-three/fiber + three-stdlib** — renders the generated STL in-browser (3D wireframe viewer)
- **pdfjs-dist** — client-side PDF rendering for the manual "Cropper" tool
- No Redux/Zustand/Context — plain `useState` per page; no WebSockets/SSE, communication is REST + polling every 1.5s

---

## 3. High-level shape

A FastAPI monolith with:
- a single SQLite DB used **only** for auth,
- a **stateful in-memory Python dict + flat JSON file** as the de facto "jobs table" (no real job/task database model),
- a battery of Excel/CAD/CV/vision-LLM "generator" services that each independently parse the feasibility spreadsheet and produce one output workbook,
- local filesystem as the only storage layer (`uploads/`, `outputs/`),
- and a React SPA that uploads one job, polls its status, and lets a user preview/download/hand-correct the results.

There's also a `stitch_suite/` folder of **static AI-generated HTML/Tailwind UI mockups** (login variants, an AI assistant chat screen, an FMEA recommendations screen, a BI/analytics dashboard, a task management board, etc.) — these are **product/UX explorations, not working code**, served as a demo gallery inside the app. They're a useful signal for what future modules the product team has already been envisioning (batch dashboards, fleet-level analytics, task management), which is directly relevant to a "many parts" conversation even though none of it is implemented.

---

## 4. Backend deep dive

### 4.1 `backend/main.py`
- Creates the FastAPI app, mounts two routers: `routes.router` at `/api`, `auth.router` at `/api/auth`.
- `CORSMiddleware` is wide open (`allow_origins=["*"]`, `allow_credentials=True`) — no per-tenant origin control.
- On module import (not a startup hook), calls `Base.metadata.create_all()` and seeds two hardcoded users (`admin@metalman.io` / `demo@metalman.io`) every process start.
- Mounts `outputs/` as static files at `/outputs` (**no auth check** — any generated document is downloadable by anyone who knows/guesses the filename) and `stitch_suite/` at `/stitch-demo`.
- In dev, auto-picks a free port from 8001 and writes `frontend/.env.development.local` with the resulting API URL — a local-machine convenience that assumes backend and frontend share a filesystem.
- Docker entrypoint runs `uvicorn main:app --workers 1` — **this is load-bearing, not incidental** (see §6.2).

### 4.2 `backend/api/routes.py` — endpoints and pipeline

**Module-level global state** (not per-request, not per-tenant):
- `TASKS: dict` — `task_id → {status, progress, message, stl_url, files: [], error}`, persisted to `backend/tasks_db.json` on every mutation, with no file locking.
- `current_welding_row` — a single shared integer cursor into one shared "welding check sheet" workbook (see the Cropper tool, below).
- `upload_retention.json` — tracks the 3 most-recently-completed task_ids; anything older gets its `uploads/` scratch files deleted (`outputs/` is never cleaned up — unbounded growth).

**Endpoints:**

| Endpoint | Method | Purpose |
|---|---|---|
| `/api/demo-folders` | GET | Lists `stitch_suite/` subfolders for the UI mockup gallery |
| `/api/analyze` | POST | **Pipeline entry point.** Multipart upload of `cad_file`, `feasibility_file`, `drawings[]`, and an `outputs` JSON flag-set. Mints `task_id = uuid4()[:8]`, saves files to disk, schedules `process_engineering_task` as a `BackgroundTasks` job, returns `task_id` immediately. |
| `/api/status/{task_id}` | GET | Returns `TASKS[task_id]` verbatim. Polled every 1.5s by the frontend. |
| `/api/correct` | POST | Applies hand-typed cell corrections to an already-generated Excel file, keyed by `task_id` + document type + cell address. Uses a **second, independent** read of `tasks_db.json` (not the in-memory `TASKS` dict) — these two can theoretically disagree. |
| `/api/preview/{filename}` | GET | "High-fidelity" preview: on Windows, drives real desktop Excel via COM (`win32com`) to export a PDF; otherwise falls back to a pandas-generated HTML table. **Requires a licensed Excel install colocated with the backend.** |
| `/api/edit/{filename}` | POST | Calls `os.startfile()` — **launches desktop Excel on the server machine**, not the user's machine. Only coherent because backend and browser currently run on the same box in practice. |
| `/api/save-crop`, `/api/welding/set-reference-image`, `/api/download-excel`, `/api/reset-excel` | POST/GET | The manual "Cropper" tool — harvests weld-symbol crops from drawings into **one single shared global Excel file** (`welding_crops.xlsx`), completely independent of `task_id`, shared by every user, hardcoded to 27 rows capacity. |

**The pipeline itself** (`process_engineering_task`, runs synchronously inside the background task, all in one `try/except`):

0. Extract `part_no`/`part_desc` from the feasibility file — one canonical identity (`utils/part_info.get_canonical_part_info`) reused by every downstream generator (this replaced earlier code where each generator re-derived part identity independently and could disagree).
1. `services/cad_service.export_step_to_stl()` — cadquery STEP→STL export, powers the browser 3D viewer.
2. `services/bom_generator.generate_bom_excel()` — BOM Excel + per-part isometric image extraction (see §4.3.2). Populates a shared temp image cache reused by later steps.
3. `services/pfd_renderer.*` — 4 sequential Excel sheet writes (Sub-Assembly Index, Sub-Assy flow, Sheet Metal Parts, BOP list) — the deepest CAD-geometry step (see §4.3.3).
4. `services/tooling_generator.generate_tooling_list()`.
5. `services/fitment_generator.generate_fitment_check_sheet()` (reuses BOM's cached images).
6. `services/pfmea_generator.generate_pfmea_excel()` — driven by static JSON knowledge-base dicts.
7. `services/control_plan_generator.generate_control_plan_excel()` — **hard-depends on PFD having already run** in the same job (re-reads the PFD output file).
8. `services/fixture_plan_generator.generate_fixture_plan()`.

Progress percentages (15/60/92/95/97/98/99/99.5/100) are hardcoded here **and independently duplicated** in the frontend's `ProcessingView.tsx` to render matching stage labels — no shared schema/contract between the two today.

Everything above runs **single-threaded, sequential, in one Python process** per job — no parallelism across parts, and none across the document types within one part either.

### 4.3 `backend/services/` — business logic inventory

| File | Role |
|---|---|
| `ai_provider.py` | Single chokepoint for vision-LLM calls (`get_vision_completion`); switches between Ollama-cloud (OpenAI-SDK-compatible client) and Gemini via `AI_PROVIDER` env var |
| `bom_generator.py` (607 lines) | BOM Excel generation + the isometric-image extraction/cleanup pipeline (reused by 3 different endpoints) |
| `cad_service.py` | STEP→STL export for the 3D viewer |
| `pfd_renderer.py` (1099 lines) | **The real CAD-geometry engine**: STEP assembly traversal via cadquery-ocp/OCCT, per-operation part extraction, SVG/PNG rendering, 4-sheet PFD Excel |
| `pfmea_generator.py` (506 lines) | Dict/template-driven FMEA builder — reads static JSON "knowledge base" files, not AI-generated |
| `control_plan_generator.py` (814 lines) | Structurally parallel to PFMEA, cross-referenced against the already-generated PFD output |
| `fitment_generator.py`, `fixture_plan_generator.py`, `tooling_generator.py` | Same overall shape: parse feasibility Excel → copy blank template → fill data → embed images |
| `correction_service.py` | Applies hand-typed cell corrections onto generated output files |
| `cad_slicer.py`, `feasibility_parser.py`, `tooling_extractor.py` | **Dead placeholder stubs**, unused — evidence of an earlier, more decoupled design that was never finished |

**The isometric-image extraction pipeline** (used by BOM generation and the standalone welding-reference endpoint) is a 3-tier cascade:
1. Deterministic text search on the PDF for a "REFERENCE ISOMETRIC VIEW" caption (PyMuPDF) — no AI call.
2. OpenCV contour detection on the likely region if no caption is found.
3. **Only as a last resort**: draw a lettered/numbered grid over the page, send it to the vision LLM, parse the returned grid-cell bounds.

Then, regardless of which tier located the view, OpenCV contour/threshold/dilation logic ("surgical cleanup") strips dimension lines and text to produce a clean part image, which is embedded into the target Excel cell with pixel-accurate centering (`utils/image_utils.place_image_centered`).

**The CAD-geometry pipeline** (`pfd_renderer.py`) parses the feasibility spreadsheet into an operation → part-number "routing matrix" (with recursive sub-assembly/phantom-part resolution), then uses `cadquery-ocp`'s OCCT/XCAF bindings to traverse the STEP assembly tree, match named components to the routing matrix, and build a per-manufacturing-stage 3D compound — falling back to a Windows-only FreeCAD subprocess if the OCP path fails.

### 4.4 Knowledge base (`backend/assets/`)
- Blank Excel **templates** every generator copies before filling (BOM, Control Plan, PFMEA, PFD, Fitment, Fixture Plan, Tooling, Welding Checksheet).
- `control_plan_dicts/{assy,bop,sheetmetal}/*.json` and `pfmea_dicts/{assembly,bop,sheetmetal}/*.json` — **static, per-process-type engineering knowledge** (standard failure modes, controls, RPN logic), applied identically regardless of the specific part number. This is the part of the system that is **already** naturally shared/reusable across many parts, since it encodes process-type knowledge rather than per-part data.

### 4.5 Database & persistence model
- SQLite (`sqlite:///./metalman.db`) via SQLAlchemy, `connect_args={"check_same_thread": False}` — itself a concurrency smell (SQLite doesn't support real concurrent writers).
- Schema is **only** `User`, `Role`, `Permission` (+ join table). **No `Job`/`Task`/`Part`/`Document` tables exist.**
- All job state lives in the in-memory `TASKS` dict (source of truth during a request) + `tasks_db.json` (durability snapshot, no locking, no schema versioning).
- File storage is local disk only: `uploads/` (namespaced by `{task_id}_...`, garbage-collected after 3 completed jobs) and `outputs/` (namespaced by **`{part_no}`**, not `task_id`, and never cleaned up).

### 4.6 Auth
- Fully implemented JWT flow (`api/auth.py`: OAuth2 password flow, HS256, 8h expiry, pbkdf2_sha256 hashing) — but `get_current_user`/`check_role` are **never wired into any business endpoint** (`/analyze`, `/status`, `/correct`, `/preview`, `/edit`, `/save-crop`, etc. all have zero `Depends()` auth). Combined with open CORS and the unauthenticated `/outputs` static mount, the API is effectively open regardless of frontend login state. The frontend's `ProtectedRoute` is a client-only redirect guard, not real access control.

---

## 5. Frontend deep dive (`frontend/src`)

**Routing**: `/` (marketing landing), `/login`, `/dashboard` (protected), `/correction` (protected), `/cropper` (protected), `/stitch` (protected, UI mockup gallery), `*` → 404.

**Core workflow — `Dashboard.tsx`**: models **exactly one job at a time**. State is `feasibilityFile`, `cadFile`, `drawingFiles[]`, `outputs` (7 boolean toggles), `processing`, and a single `taskId`. There is no list/array dimension anywhere in this state — no job history, no batch upload, no queue view. `handleGenerate()` posts the multipart form to `/api/analyze` and switches to `ProcessingView`.

**`ProcessingView.tsx`**: polls `/api/status/{taskId}` every 1.5s (with exponential-backoff retry for Render cold starts), renders a progress bar against the same hardcoded percentage thresholds the backend uses, streams the STL into `CADWireframe.tsx` as soon as available, and on completion shows a file list with Preview/Download actions plus "Start New Project" (resets Dashboard state for the *next single* part).

**`Correction.tsx`** (742 lines): lets a user queue up multiple cell-level corrections against a completed job's output files, with a document→column mapping that is **duplicated** between this file and the backend's `correction_service.py` (front and back must be kept in sync by hand).

**`PdfCropper.tsx`**: the manual welding-checksheet tool — entirely independent of `taskId`, operates on the one global shared workbook.

**State management**: no Redux/Zustand/Context, just local `useState`; `@tanstack/react-query` is installed and initialized but **unused** — there's no existing caching/invalidation layer a future "job list" feature could build on without adopting real usage of it (or another state manager).

**Communication pattern**: REST + polling only. No WebSockets/SSE anywhere.

---

## 6. End-to-end trace for one part (concrete path)

1. `Dashboard.tsx` → `POST /api/analyze` (cad_file, feasibility_file, drawings[], outputs flags).
2. `routes.py` mints `task_id`, saves uploads, schedules `process_engineering_task` via `BackgroundTasks`, returns immediately.
3. `ProcessingView.tsx` begins polling `GET /api/status/{task_id}` every 1.5s.
4. `utils/part_info.get_canonical_part_info()` fixes `(part_no, part_desc)` for the whole job.
5. `cad_service.export_step_to_stl()` → STL streamed into the 3D viewer.
6. `bom_generator.generate_bom_excel()` → BOM Excel + isometric images (PyMuPDF render → deterministic/CV/vision-LLM location cascade → OpenCV cleanup → openpyxl embed).
7. `pfd_renderer.*` → cadquery-ocp/OCCT STEP traversal → per-operation 3D compounds → SVG/PNG render → 4-sheet PFD Excel.
8. Tooling / Fitment / PFMEA / Control Plan / Fixture Plan generators run in sequence, each doing its own pandas/openpyxl pass, some reusing BOM's cached images, Control Plan reusing PFD's output.
9. `TASKS[task_id]` → `"completed"`, `files[]` populated; old jobs' upload scratch files purged (outputs kept forever).
10. Frontend shows the file list; Preview/Download/Edit-in-Excel/Correction flows as described above.

---

## 7. Concrete scaling constraints (code-verified, not inferred)

These are the load-bearing facts for a "one part → many parts" conversation:

1. **No task queue despite the dependency being present.** `worker/tasks.py` is a dead `MockCelery` stub (celery/redis are in `requirements.txt` but never actually imported anywhere else). The real pipeline runs via `FastAPI.BackgroundTasks` in the same process/event loop as the HTTP server — N concurrent uploads today means N background coroutines racing for one process's CPU/threads, no concurrency limit, no retries, no dead-letter handling, no cross-process distribution.

2. **`--workers 1` is load-bearing, not incidental.** The Dockerfile comments explicitly explain why: `TASKS` lives in per-process memory; multiple worker processes would each have an inconsistent view of task state and break status polling. This caps the whole backend to one OS process — no horizontal or even multi-worker vertical scaling without a real shared state store first.

3. **No real job/task database table.** The only DB tables are `User`/`Role`/`Permission`. Job state is an in-memory dict snapshotted to a flat, unlocked JSON file, with a second independent reader in `correction_service.py` that can theoretically see stale data.

4. **Output filenames are keyed by part number, not job ID.** E.g. `BOM_Result_{part_no}.xlsx` — two jobs for the same part number (re-run, correction, or genuinely concurrent) **silently overwrite each other's files**, with no per-job output namespace.

5. **One global mutable singleton for the welding/Cropper tool** — a single shared `welding_crops.xlsx` and a single shared cursor variable, with no per-user/per-part scoping and a hardcoded 27-row capacity.

6. **Local filesystem is the only storage layer.** No S3/GCS/blob storage abstraction — scaling to multiple backend instances needs either a shared network filesystem or a storage-layer rewrite. The `/outputs` static mount also means any generated document is publicly downloadable with zero auth if the filename is known/guessed.

7. **Auth is implemented but not enforced anywhere on the pipeline.** JWT auth exists fully but no endpoint actually requires it; combined with wide-open CORS, the API is functionally public. This needs fixing before genuine multi-tenant scaling (jobs need to be scoped to a user/org, which requires enforced auth as a prerequisite).

8. **Two features require a colocated, licensed, Windows Excel install** (`/api/preview` PDF export and `/api/edit`'s `os.startfile`) — these cannot run in a Linux container fleet at all beyond their existing lower-fidelity fallbacks, and even on Windows, concurrent requests would contend for one COM-automated Excel instance with no real serialization.

9. **CAD extraction has a silent Windows-only fallback.** If the primary cross-platform `cadquery-ocp` path fails on some assembly, the FreeCAD-subprocess fallback is unavailable on Linux/Docker (Render) — parts that fully render in Windows dev may quietly produce incomplete PFD sheets in production.

10. **No output retention policy — unbounded disk growth.** Only `uploads/` scratch files are garbage-collected (3-job retention window); `outputs/` grows forever. Fine at demo scale, a real problem at "many parts" scale.

11. **Vision-AI calls are unbatched and inline**, one drawing at a time, inside the sequential per-BOM-row loop, with no caching across different jobs referencing the same underlying part drawing (and the existing per-job cache is defeated by a `force_rewrite=True` default in the current wiring).

12. **The frontend models exactly one job at a time by construction.** No array/list dimension in Dashboard state, no job history/list UI, no batch-upload UI — and the one library that could support a "job list" view with caching (`react-query`) is installed but never actually used.

13. **Hardcoded progress-percentage duplication** between backend (`process_engineering_task`) and frontend (`ProcessingView.tsx`) — any change to the pipeline's stage set requires synchronized manual edits in two places, with no shared schema/contract.

14. **Dead scaffolding hints at an earlier, more decoupled design that was abandoned** (`worker/tasks.py`, `cad_slicer.py`, `feasibility_parser.py`, `tooling_extractor.py` are all unused stubs) — worth knowing so a redesign doesn't have to guess whether these were intentional or vestigial.

---

## 8. What's *already* naturally shareable across many parts

Worth noting as a counterbalance to §7 — not everything needs to be rebuilt:

- The **PFMEA/Control Plan JSON knowledge-base dicts** (`assets/{pfmea,control_plan}_dicts/`) already encode process-type knowledge (sheet metal / assembly / BOP), not per-part data — this layer generalizes to "many parts" as-is.
- **`utils/image_utils.place_image_centered()`** and the isometric-extraction cascade in `bom_generator.py` are genuinely reusable, part-agnostic infrastructure, not per-part special-casing.
- **`utils/part_info.get_canonical_part_info()`** already solved the "single source of truth for part identity within a job" problem — the natural next step is making that identity a first-class, persisted entity (a `Part`/`Job` row) rather than a tuple threaded through function arguments.
- The **`ai_provider.py`** abstraction is already provider-agnostic (Ollama/Gemini) and could be made batching-aware without touching call sites.

---

## 9. Suggested framing for the brainstorm

Given the above, the "one part → many parts" conversation likely needs to cover, roughly in dependency order:

1. **A real job/data model** — `Job`/`Part`/`Document` tables (Postgres, not SQLite), replacing the in-memory `TASKS` dict + JSON file.
2. **A real task queue** — actually wiring up Celery+Redis (already a dependency!) or an alternative (RQ, arq, cloud task queue), enabling multiple workers and horizontal scaling, removing the `--workers 1` constraint.
3. **Per-job output namespacing** — fixing the part-number-keyed filename collision problem.
4. **Object storage** — moving `uploads/`/`outputs/` off local disk so multiple backend instances can share state.
5. **Enforcing auth on the pipeline** — prerequisite for any real multi-tenancy (jobs scoped to user/org).
6. **Batch ingestion UX** — both API shape (submit N parts, get N job IDs) and frontend (a job list/queue view, finally giving `react-query` a purpose).
7. **Vision-LLM batching/caching strategy** across parts that share sub-components.
8. **Decoupling the Windows-only desktop-Excel features** (preview PDF export, edit-in-Excel) from the core pipeline so the "many parts" batch engine isn't gated on a licensed Windows box.
9. **Retention/lifecycle policy** for `outputs/` once volume is no longer demo-scale.

This document intentionally stops at *describing the current system* — the actual scaling architecture (queue choice, DB schema, storage provider, deployment topology) is the brainstorm itself.
