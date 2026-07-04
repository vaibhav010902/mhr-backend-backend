from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    APP_NAME: str = "MHR Analytics Platform"
    MONGO_URI: str = "mongodb://localhost:27017"
    DATABASE_NAME: str = "mhr_platform"
    ALLOWED_ORIGINS: str = "http://localhost:5173,http://127.0.0.1:5173"

    class Config:
        env_file = ".env"


settings = Settings()
