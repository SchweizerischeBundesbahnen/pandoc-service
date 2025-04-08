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

A Dockerized service providing a REST API interface to leverage Pandoc's functionality for converting documents
from one format into another.

## Features

- Simple REST API to access Pandoc
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
  docker run --detach \
    --publish 9082:9082 \
    --name pandoc-service \
    ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

The service will be accessible on port 9082.

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
  docker run --detach \
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

#### container-structure-test
```bash
docker build -t pandoc-service:local .
```
```bash
container-structure-test test --image pandoc-service:local --config ./tests/container/container-structure-test.yaml
```
#### tox
```bash
poetry run tox
```
#### pytest (for debugging)
```bash
# all tests
poetry run pytest
```
```bash
# a specific test
poetry run pytest tests/test_docx_post_process.py -v
```
#### pre-commit
```bash
poetry run pre-commit run --all
```

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

#### Convert HTML to PDF

<details>
  <summary>
    <code>POST</code> <code>/convert/html/to/docx</code>
  </summary>

##### Parameters

> | Parameter name       | Type     | Data type | Description                                                          |
> |----------------------|----------|-----------|----------------------------------------------------------------------|
> | encoding             | optional | string    | Encoding of provided HTML (default: utf-8)                           |
> | file_name            | optional | string    | Output filename (default: converted-document.pdf)                    |

##### Responses

> | HTTP code | Content-Type      | Response                     |
> |-----------|-------------------|------------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` | DOCX document (binary data)  |
> | `400`     | `plain/text`      | Error message with exception |
> | `500`     | `plain/text`      | Error message with exception |

##### Example cURL

> ```bash
> curl -X POST -H "Content-Type: application/html" --data @input_html http://localhost:9082/convert/html/to/docx --output output.docx
> ```

</details>
