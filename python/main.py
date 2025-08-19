from fastapi import FastAPI
from pydantic import BaseModel
from typing import List
import os
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or ["*"] for all origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class SheetData(BaseModel):
    name: str
    csv: str
    
class SheetsRequest(BaseModel):
    sheets: List[SheetData]

@app.get("/")
def read_root():
    return {"message": "Hello World"}


@app.post("/upload-csvs")
async def upload_csvs(request: SheetsRequest):
    os.makedirs("uploads", exist_ok=True)

    for sheet in request.sheets:
        file_path = os.path.join("uploads", f"{sheet.name}.csv")
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(sheet.csv)

    return {"status": "success", "message": f"{len(request.sheets)} sheets saved"}


# Request schema


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)