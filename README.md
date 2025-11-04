[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=bugs)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Code Smells](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=code_smells)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=coverage)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Duplicated Lines (%)](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=duplicated_lines_density)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=ncloc)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Reliability Rating](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=reliability_rating)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=security_rating)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=sqale_rating)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)
[![Vulnerabilities](https://sonarcloud.io/api/project_badges/measure?project=SchweizerischeBundesbahnen_pandoc-service&metric=vulnerabilities)](https://sonarcloud.io/summary/new_code?id=SchweizerischeBundesbahnen_pandoc-service)

# Pandoc Service

A Dockerized service providing a REST API interface to leverage [Pandoc](https://pandoc.org/)'s functionality for converting documents
from one format into another.

## Features

- Simple REST API to access [Pandoc](https://pandoc.org/)
- Direct subprocess calls to the pandoc binary (no Python module dependency)
- Compatible with amd64 and arm64 architectures
- Easily deployable via Docker

## Getting Started

### Installation

To install the latest version of the Pandoc Service, run the following command:

```bash
docker pull ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

### Running the Service

To start the Pandoc service container, execute:

```bash
  docker run --init --detach \
    --publish 9082:9082 \
    --name pandoc-service \
    --env REQUEST_BODY_LIMIT_MB=500 \
    ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

The service will be accessible on port 9082.

The REQUEST_BODY_LIMIT_MB environment variable sets the maximum allowed size (in megabytes) for uploaded files or request bodies processed by the Pandoc service. The default is 500 MB.

### Using as a Base Image

To extend or customize the service, use it as a base image in the Dockerfile:

```Dockerfile
FROM ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

## Development

### Building the Docker Image

To build the Docker image from the source with a custom version, use:

```bash
  docker build \
    --build-arg APP_IMAGE_VERSION=0.0.0 \
    --file Dockerfile \
    --tag pandoc-service:0.0.0 .
```

Replace 0.0.0 with the desired version number.

### Running the Development Container

To start the Docker container with your custom-built image:

```bash
  docker run --init --detach \
    --publish 9082:9082 \
    --name pandoc-service \
    pandoc-service:0.0.0
```

### Stopping the Container

To stop the running container, execute:

```bash
  docker container stop pandoc-service
```

### Testing

The project includes several test methods to ensure functionality.

#### Container Structure Test
```bash
container-structure-test test --image pandoc-service:local --config ./tests/container/container-structure-test.yaml
```

#### Docker Image Smoke Test
To test the Docker image build and API functionality:
```bash
bash tests/shell/test_pandoc_service.sh
```
This script builds the image, starts a container, and performs tests on all endpoints.

#### Python Tests
```bash
# Prepare testing
poetry install
```

```bash
# Run all Python tests
poetry run pytest -v
```
```bash
# Run a specific test
poetry run pytest tests/test_docx_post_process.py -v
```

#### Tox
```bash
# Run all test pytest and linting
poetry run tox
```

#### Pre-commit
```bash
poetry run pre-commit run --all
```

For more detailed testing information, see the [tests README](tests/README.md).

### Access service

Pandoc Service provides the following endpoints:

------------------------------------------------------------------------------------------

#### Getting version info

<details>
  <summary>
    <code>GET</code> <code>/version</code>
  </summary>

##### Responses

> | HTTP code | Content-Type       | Response                                                                                                       |
> |-----------|--------------------|----------------------------------------------------------------------------------------------------------------|
> | `200`     | `application/json` | `{ "python": "3.12.5", "timestamp": "2024-09-23T12:23:09Z", "pandoc": "3.6.2", "pandocService": "0.0.0" }` |

##### Example cURL

> ```bash
>  curl -X GET -H "Content-Type: application/json" http://localhost:9082/version
> ```

</details>

------------------------------------------------------------------------------------------

#### Getting openapi info

<details>
  <summary>
    <code>GET</code> <code>/static/openapi.json</code>
  </summary>

##### Responses

> | HTTP code | Content-Type       | Response      |
> |-----------|--------------------|---------------|
> | `200`     | `application/json` | openapi.json  |

##### Example cURL

> ```bash
>  curl -X GET -H "Content-Type: application/json" http://localhost:9082/static/openapi.json
> ```
</details>

------------------------------------------------------------------------------------------

#### Getting docx template

<details>
  <summary>
    <code>GET</code> <code>/docx-template</code>
  </summary>

##### Responses

> | HTTP code | Content-Type                                                              | Response                 |
> |-----------|---------------------------------------------------------------------------|--------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` | binary document content  |

##### Example cURL

> ```bash
>  curl -X GET -H "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document" http://localhost:9082/docx-template
> ```

</details>

------------------------------------------------------------------------------------------

#### Convert HTML to DOCX

<details>
  <summary>
    <code>POST</code> <code>/convert/html/to/docx</code>
  </summary>

##### Parameters

> | Parameter name       | Type     | Data type | Description                                                                                                     |
> |----------------------|----------|-----------|-----------------------------------------------------------------------------------------------------------------|
> | encoding             | optional | string    | Encoding of provided HTML (default: utf-8)                                                                      |
> | file_name            | optional | string    | Output filename (default: converted-document.pdf)                                                               |
> | paper_size           | optional | string    | Paper size for the output document. Supported values: A5, A4, A3, B5, B4, JIS_B5, JIS_B4, LETTER, LEGAL, LEDGER |
> | orientation          | optional | string    | Page orientation. Supported values: portrait, landscape                                                         |

##### Responses

> | HTTP code | Content-Type                                                               | Response                     |
> |-----------|----------------------------------------------------------------------------|------------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.wordprocessingml.document`  | DOCX document (binary data)  |
> | `400`     | `plain/text`                                                               | Error message with exception |
> | `500`     | `plain/text`                                                               | Error message with exception |

##### Example cURL

> ```bash
> curl -X POST -H "Content-Type: application/html" --data @input_html http://localhost:9082/convert/html/to/docx --output output.docx
> ```
>
> With custom paper size and orientation:
> ```bash
> curl -X POST -H "Content-Type: application/html" --data @input_html "http://localhost:9082/convert/html/to/docx?paper_size=A4&orientation=landscape" --output output.docx
> ```

</details>

------------------------------------------------------------------------------------------

#### Convert HTML to DOCX with custom template

<details>
  <summary>
    <code>POST</code> <code>/convert/html/to/docx-with-template</code>
  </summary>

##### Parameters

> | Parameter name       | Type     | Data type | Description                                                                                                     |
> |----------------------|----------|-----------|-----------------------------------------------------------------------------------------------------------------|
> | source               | required | file      | Source HTML content as multipart/form-data                                                                      |
> | template             | optional | file      | Custom DOCX template file as multipart/form-data                                                                |
> | encoding             | optional | string    | Encoding of provided HTML (default: utf-8)                                                                      |
> | file_name            | optional | string    | Output filename (default: converted-document.docx)                                                              |
> | paper_size           | optional | string    | Paper size for the output document. Supported values: A5, A4, A3, B5, B4, JIS_B5, JIS_B4, LETTER, LEGAL, LEDGER |
> | orientation          | optional | string    | Page orientation. Supported values: portrait, landscape                                                         |

##### Responses

> | HTTP code | Content-Type                                                               | Response                     |
> |-----------|----------------------------------------------------------------------------|------------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.wordprocessingml.document`  | DOCX document (binary data)  |
> | `400`     | `plain/text`                                                               | Error message with exception |
> | `500`     | `plain/text`                                                               | Error message with exception |

##### Example cURL

> ```bash
> curl -X POST -F "source=@input.html" -F "template=@custom-template.docx" http://localhost:9082/convert/html/to/docx-with-template --output output.docx
> ```
>
> With custom paper size and orientation:
> ```bash
> curl -X POST -F "source=@input.html" -F "template=@custom-template.docx" "http://localhost:9082/convert/html/to/docx-with-template?paper_size=A4&orientation=landscape" --output output.docx
> ```

</details>
