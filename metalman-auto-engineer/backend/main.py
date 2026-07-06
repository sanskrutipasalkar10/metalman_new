from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from api import routes, auth
from database import engine, SessionLocal
import models
import logging
import sys
from passlib.context import CryptContext

# ── Logging ──────────────────────────────────────────────────────────────────
# Force all print() / logging calls to appear immediately in the terminal.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
logging.getLogger("uvicorn").setLevel(logging.INFO)
logging.getLogger("uvicorn.access").setLevel(logging.INFO)
logging.getLogger("uvicorn.error").setLevel(logging.INFO)
# Make sure stdout is unbuffered so print() shows up immediately
import functools
print = functools.partial(print, flush=True)
# ─────────────────────────────────────────────────────────────────────────────

# Create tables
models.Base.metadata.create_all(bind=engine)

# Seed Data
pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")
db = SessionLocal()
try:
    # Roles
    admin_role = db.query(models.Role).filter(models.Role.name == "Admin").first()
    if not admin_role:
        admin_role = models.Role(name="Admin")
        db.add(admin_role)
    
    engineer_role = db.query(models.Role).filter(models.Role.name == "Engineer").first()
    if not engineer_role:
        engineer_role = models.Role(name="Engineer")
        db.add(engineer_role)
    
    db.commit()

    # Admin User
    admin_user = db.query(models.User).filter(models.User.email == "admin@metalman.io").first()
    if not admin_user:
        admin_user = models.User(
            username="admin",
            email="admin@metalman.io",
            hashed_password=pwd_context.hash("admin123"),
            role_id=admin_role.id
        )
        db.add(admin_user)
    
    # Engineer User
    eng_user = db.query(models.User).filter(models.User.email == "demo@metalman.io").first()
    if not eng_user:
        eng_user = models.User(
            username="engineer",
            email="demo@metalman.io",
            hashed_password=pwd_context.hash("password123"),
            role_id=engineer_role.id
        )
        db.add(eng_user)
    
    db.commit()
finally:
    db.close()

app = FastAPI(
    title="Metalman Auto-Engineer API",
    description="The Python FastAPI Brain for Lights-Out Engineering Automation",
    version="1.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust this for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routes
app.include_router(routes.router, prefix="/api")
app.include_router(auth.router, prefix="/api/auth", tags=["auth"])

import os
outputs_dir = "/data/outputs" if os.path.exists("/data") else "outputs"
os.makedirs(outputs_dir, exist_ok=True)
app.mount("/outputs", StaticFiles(directory=outputs_dir), name="outputs")

# Serve Stitch Demo statically
app.mount("/stitch-demo", StaticFiles(directory="stitch_suite"), name="stitch-demo")

@app.get("/")
async def root():
    return {
        "status": "online",
        "message": "Metalman Auto-Engineer API is operational",
        "version": "1.0.0"
    }

if __name__ == "__main__":
    import uvicorn
    import socket

    def find_free_port(start_port: int) -> int:
        current_port = start_port
        while current_port < 65535:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind(("0.0.0.0", current_port))
                    return current_port
                except OSError:
                    current_port += 1
        return start_port

    base_port = int(os.environ.get("PORT", 8001))
    port = find_free_port(base_port)

    # Write port configuration to the frontend's local development environment file
    # This enables Vite to hot-reload and connect to the correct backend port automatically.
    try:
        frontend_env_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "frontend", ".env.development.local"))
        with open(frontend_env_path, "w") as f:
            f.write(f"VITE_API_URL=http://127.0.0.1:{port}\n")
        print(f"Automatically configured frontend .env.development.local with VITE_API_URL=http://127.0.0.1:{port}")
    except Exception as e:
        print(f"Could not update frontend local env file: {e}")

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=False,
        log_level="info",
        access_log=True,
    )
