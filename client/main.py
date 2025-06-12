from fastapi import FastAPI, Request, Form, HTTPException, Depends
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from fastapi.security import OAuth2PasswordBearer
from jose import jwt, JWTError
from pydantic import BaseModel
import openpyxl
import bcrypt
from datetime import datetime, timedelta

import os

# === Config ===
SECRET_KEY = "plm_secret"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60

# === Init App ===
app = FastAPI()
templates = Jinja2Templates(directory="templates")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="login")

# === In-memory sheet-driven DB ===
users_db = {}
modules_db = {}
permissions_db = {}

def load_sheets():
    global users_db, modules_db, permissions_db

    wb_main = openpyxl.load_workbook("app_sheets/main.xlsx")
    refs_sheet = wb_main["refs"]
    refs = {row[1].value: row[0].value for row in refs_sheet.iter_rows(min_row=2)}

    wb_users = openpyxl.load_workbook(f"app_sheets/{refs['users']}")
    sheet = wb_users["users"]
    users_db = {
        row[1].value: {
            "username": row[1].value,
            "password_hash": row[2].value,
            "role": row[3].value,
        } for row in sheet.iter_rows(min_row=2)
    }

    wb_modules = openpyxl.load_workbook(f"app_sheets/{refs['modules']}")
    sheet = wb_modules["modules"]
    modules_db = {
        row[0].value: {
            "id": row[0].value,
            "name": row[1].value,
            "description": row[2].value,
        } for row in sheet.iter_rows(min_row=2)
    }

    wb_perms = openpyxl.load_workbook(f"app_sheets/{refs['permissions']}")
    sheet = wb_perms["permissions"]
    permissions_db = {
        row[0].value: row[1].value.split(",") if row[1].value != "all" else "all"
        for row in sheet.iter_rows(min_row=2)
    }

# === Auth ===
def authenticate_user(username, password):
    user = users_db.get(username)
    if not user or not bcrypt.checkpw(password.encode(), user["password_hash"].encode()):
        return None
    return user

def create_token(data: dict):
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

async def get_current_user(token: str = Depends(oauth2_scheme)):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        user = users_db.get(payload.get("sub"))
        if not user:
            raise HTTPException(status_code=401)
        return user
    except JWTError:
        raise HTTPException(status_code=401)

# === Routes ===
@app.on_event("startup")
def startup():
    load_sheets()

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

@app.post("/login")
async def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    user = authenticate_user(username, password)
    if not user:
        return templates.TemplateResponse("login.html", {"request": request, "error": "Invalid login"})
    token = create_token({"sub": user["username"]})
    response = RedirectResponse("/dashboard", status_code=302)
    response.set_cookie("token", token)
    return response

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    token = request.cookies.get("token")
    user = await get_current_user(token)
    allowed = permissions_db.get(user["role"])
    visible = modules_db.values() if allowed == "all" else [modules_db[mid] for mid in allowed]
    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "user": user,
        "modules": visible
    })

@app.post("/admin/reload")
async def admin_reload(request: Request):
    token = request.cookies.get("token")
    user = await get_current_user(token)
    if user["role"] != "admin":
        raise HTTPException(status_code=403)
    load_sheets()
    return RedirectResponse("/dashboard", status_code=302)
