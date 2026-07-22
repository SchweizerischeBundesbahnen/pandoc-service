from pydantic import BaseModel, Field


class VersionSchema(BaseModel):
    apiVersion: int = Field(title="API Version", description="API version for compatibility checking with docx-exporter")
    python: str = Field()
    pandoc: str | None = Field()
    pandocService: str | None = Field()
    timestamp: str | None = Field()
    chromium: str | None = Field()
