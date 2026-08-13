from pydantic import BaseModel, Field


class VersionSchema(BaseModel):
    apiVersion: int = Field(title="API Version", description="API version for compatibility checking with docx-exporter")  # noqa: N815 - this is the JSON field name clients read
    python: str = Field()
    pandoc: str | None = Field()
    pandocService: str | None = Field()  # noqa: N815 - this is the JSON field name clients read
    timestamp: str | None = Field()
    chromium: str | None = Field()
