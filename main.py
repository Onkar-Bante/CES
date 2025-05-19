# main.py

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError
from starlette.exceptions import HTTPException as StarletteHTTPException
from routers import company_router, employee_router, attendance_router, test_router
import math
import json
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Dynamic Company & Employee Management")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust as needed for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Custom JSON encoder for handling NaN values
class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
        return super().default(obj)

# Custom exception handler for JSON serialization errors
@app.exception_handler(ValueError)
async def value_error_handler(request: Request, exc: ValueError):
    if "Out of range float values are not JSON compliant" in str(exc):
        # Return a more specific error message
        return JSONResponse(
            status_code=500,
            content={"detail": "The response contains NaN or Infinity values which cannot be serialized to JSON. This is typically caused by missing or invalid numeric data."}
        )
    # For other ValueError exceptions, re-raise
    raise exc

# Include your routers
app.include_router(company_router.router)
app.include_router(employee_router.router)
app.include_router(attendance_router.router)
app.include_router(test_router.router)

@app.get("/")
def read_root():
    return {"message": "Welcome to Dynamic Company Management System"}