from pydantic import BaseModel, Field


class VersionSchema(BaseModel):
    python: str = Field()
    pandoc: str | None = Field()
    pandocService: str | None = Field()
    timestamp: str | None = Field()
