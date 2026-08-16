from pydantic import BaseModel, Field


class VersionSchema(BaseModel):
    # This is the JSON field name clients read.
    apiVersion: int = Field(title="API Version", description="API version for compatibility checking with docx-exporter")  # noqa: N815
    python: str = Field()
    pandoc: str | None = Field()
    # This is the JSON field name clients read.
    pandocService: str | None = Field()  # noqa: N815
    timestamp: str | None = Field()
    chromium: str | None = Field()
