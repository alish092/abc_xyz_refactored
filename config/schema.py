from pydantic import BaseModel, Field

class Thresholds(BaseModel):
    A: float = Field(default=0.8, ge=0.0, le=1.0)
    B: float = Field(default=0.95, ge=0.0, le=1.0)
    X: float = Field(default=0.1, ge=0.0)
    Y: float = Field(default=0.25, ge=0.0)

class AppConfig(BaseModel):
    thresholds: Thresholds = Thresholds()  # ✅ Значение по умолчанию
