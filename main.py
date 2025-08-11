from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import os
from starlette.middleware.cors import CORSMiddleware

# Create a FastAPI instance
app = FastAPI()

# Add CORS middleware to allow cross-origin requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

# Mount the static files (e.g., images, videos)
# This assumes your files (index.html, background.mp4, sc1.jpg, etc.) are in the same directory as main.py
app.mount("/static", StaticFiles(directory="."), name="static")

# Jinja2 templates setup
templates = Jinja2Templates(directory=".")

# Excel file configuration
excel_file_name = "results.xlsx"

# Pydantic model for incoming quiz data
class QuizResult(BaseModel):
    rollno1: str
    name1: str
    college1: str
    rollno2: Optional[str] = None
    name2: Optional[str] = None
    college2: Optional[str] = None
    score: int
    elapsedTime: int
    answers: List[str]

# Route to serve the HTML file
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Route to handle quiz results submission
@app.post("/save-result")
async def save_result(result: QuizResult):
    print("Received new result:", result.dict())
    
    # Prepare the data for Excel
    new_result = {
        'Participant 1 Roll No': result.rollno1,
        'Participant 1 Name': result.name1,
        'Participant 1 College': result.college1,
        'Participant 2 Roll No': result.rollno2,
        'Participant 2 Name': result.name2,
        'Participant 2 College': result.college2,
        'Total Score': result.score,
        'Time Taken (s)': result.elapsedTime
    }

    # Add answers to the result dictionary dynamically
    for i, answer in enumerate(result.answers):
        new_result[f'Answer Q{i + 1}'] = answer

    # Create a DataFrame from the new result
    new_df = pd.DataFrame([new_result])

    try:
        if os.path.exists(excel_file_name):
            # If the file exists, append the new data without writing the header again
            existing_df = pd.read_excel(excel_file_name)
            merged_df = pd.concat([existing_df, new_df], ignore_index=True)
            merged_df.to_excel(excel_file_name, index=False)
            print("Result appended to existing Excel file.")
        else:
            # If the file doesn't exist, create a new one with headers
            new_df.to_excel(excel_file_name, index=False)
            print("Result saved to a new Excel file.")
        
        return {"message": "Result saved successfully"}
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return {"message": f"Error saving result: {e}"}

# Route to download the results file (optional)
@app.get("/download-results")
async def download_results():
    if os.path.exists(excel_file_name):
        return FileResponse(excel_file_name, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=excel_file_name)
    else:
        return {"message": "Results file not found."}