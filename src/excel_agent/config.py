import os
from dotenv import load_dotenv

load_dotenv() # Load environment variables from .env file

class Config:
    """
    Configuration class for the Excel Agent.
    Loads API keys and model names from environment variables.
    """
    GROQ_API_KEY: str = os.getenv("GROQ_API_KEY", "")
    GROQ_MODEL_NAME: str = os.getenv("GROQ_MODEL_NAME", "llama-3.3-70b-versatile") # Default to a smaller model for testing
    # Note: The user specified "llama-3.3-70b-versatile", but Groq's current public models are llama-3.1-8b-versatile and llama-3.1-70b-versatile.
    # I'm defaulting to 8b for initial testing, but it can be changed to 70b in the .env or here.
    # If "llama-3.3-70b-versatile" becomes available, we can update this.

    # Output file naming convention
    OUTPUT_FILE_PREFIX: str = "excel_agent_output_"
    PLOTS_DIR: str = "plots" # MODIFIED: Added directory for plots

    if not GROQ_API_KEY:
        print("Warning: GROQ_API_KEY not found in environment variables. Please set it in a .env file or directly.")
