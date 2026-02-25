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
- Prometheus metrics endpoint on dedicated port (9182) for monitoring and Grafana integration
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
  --publish 9182:9182 \
  --name pandoc-service \
  --env REQUEST_BODY_LIMIT_MB=500 \
  ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

The service will be accessible on port 9082, and Prometheus metrics on port 9182.

The REQUEST_BODY_LIMIT_MB environment variable sets the maximum allowed size (in megabytes) for uploaded files or request bodies processed by the Pandoc service. The default is 500 MB.

### Using as a Base Image

To extend or customize the service, use it as a base image in the Dockerfile:

```Dockerfile
FROM ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
```

### Prometheus & Grafana Integration

The service exposes Prometheus-compatible metrics for comprehensive monitoring and observability through Grafana dashboards.

**Metrics Endpoint:** `/metrics` on port 9182 (dedicated metrics port)

> **Security:** The metrics endpoint is served on a separate port (9182) from the main API (9082). This allows network-level isolation using security groups or firewall rules to restrict metrics access to your Prometheus server only.

**Available Metrics:**

**Conversion Metrics:**
- `pandoc_conversions_total` - Total successful conversions (labeled by source/target format)
- `pandoc_conversion_failures_total` - Total failed conversions (labeled by source/target format)
- `pandoc_conversion_error_rate_percent` - Conversion error rate as percentage
- `pandoc_template_conversions_total` - Total conversions using custom templates

**Performance Metrics:**
- `pandoc_conversion_duration_seconds` - Conversion time histogram (labeled by format)
- `pandoc_subprocess_duration_seconds` - Pandoc subprocess execution time histogram
- `pandoc_post_processing_duration_seconds` - DOCX/PPTX post-processing time histogram
- `avg_pandoc_conversion_time_seconds` - Average conversion time

**Size Metrics:**
- `pandoc_request_body_bytes` - Input document size histogram
- `pandoc_response_body_bytes` - Output document size histogram

**Service Metrics:**
- `uptime_seconds` - Service uptime
- `active_conversions` - Current active conversion count
- `pandoc_info` - Service and pandoc version information

**HTTP Metrics (via prometheus-fastapi-instrumentator):**
- `http_request_duration_seconds` - HTTP request duration histogram
- `http_requests_inprogress` - Current in-flight requests
- `http_requests_total` - Total HTTP requests

**Prometheus Configuration Example:**

```yaml
scrape_configs:
  - job_name: 'pandoc-service'
    static_configs:
      - targets: ['pandoc-service:9182']  # Metrics on dedicated port
    metrics_path: '/metrics'
    scrape_interval: 15s
    scrape_timeout: 10s
```

**Grafana Dashboard Queries:**

```promql
# Conversion rate (requests per second)
rate(pandoc_conversions_total[5m])

# Error rate percentage
sum(rate(pandoc_conversion_failures_total[5m])) / sum(rate(pandoc_conversions_total[5m])) * 100

# 95th percentile conversion duration
histogram_quantile(0.95, rate(pandoc_conversion_duration_seconds_bucket[5m]))

# Conversions by format
sum by (source_format, target_format) (pandoc_conversions_total)
```

**Docker Compose Example with Prometheus & Grafana:**

```yaml
services:
  pandoc-service:
    image: ghcr.io/schweizerischebundesbahnen/pandoc-service:latest
    init: true
    ports:
      - "9082:9082"   # Main API
      - "9182:9182"   # Metrics endpoint

  prometheus:
    image: prom/prometheus:latest
    ports:
      - "9090:9090"
    volumes:
      - ./prometheus.yml:/etc/prometheus/prometheus.yml
      - prometheus-data:/prometheus

  grafana:
    image: grafana/grafana:latest
    ports:
      - "3000:3000"
    volumes:
      - grafana-data:/var/lib/grafana
    environment:
      - GF_SECURITY_ADMIN_PASSWORD=admin

volumes:
  prometheus-data:
  grafana-data:
```

**Pre-configured Monitoring Stack:**

For a complete, production-ready monitoring setup with pre-configured Prometheus, Grafana, and dashboards:

```bash
cd monitoring
./start-monitoring.sh
```

This will start the Pandoc service, Prometheus, and Grafana with a pre-built dashboard. Access Grafana at http://localhost:3000 (admin/admin) and view the dashboard at http://localhost:3000/d/pandoc-service.

For detailed setup instructions and configuration options, see [monitoring/README.md](monitoring/README.md).

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
  --publish 9182:9182 \
  --network weasyprint_network \
  --name pandoc \
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
uv sync --group dev --group test
```

```bash
# Run all Python tests
uv run pytest -v
```
```bash
# Run a specific test
uv run pytest tests/test_docx_post_process.py -v
```

#### Tox
```bash
# Run all test pytest and linting
uv run tox
```

#### Pre-commit
```bash
uv run pre-commit run --all
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
> | `200`     | `application/json` | `{ "python": "3.14.0", "timestamp": "2024-09-23T12:23:09Z", "pandoc": "3.6.2", "pandocService": "0.0.0" }` |

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

------------------------------------------------------------------------------------------

#### Getting pptx template

<details>
  <summary>
    <code>GET</code> <code>/pptx-template</code>
  </summary>

##### Responses

> | HTTP code | Content-Type                                                                   | Response                      |
> |-----------|--------------------------------------------------------------------------------|-------------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.presentationml.presentation`    | binary presentation content   |

##### Example cURL

> ```bash
>  curl -X GET http://localhost:9082/pptx-template --output reference.pptx
> ```

</details>

------------------------------------------------------------------------------------------

#### Convert to PPTX with custom template

<details>
  <summary>
    <code>POST</code> <code>/convert/{source_format}/to/pptx-with-template</code>
  </summary>

##### Parameters

> | Parameter name       | Type     | Data type | Description                                                                                          |
> |----------------------|----------|-----------|------------------------------------------------------------------------------------------------------|
> | source               | required | file      | Source content as multipart/form-data                                                                |
> | template             | optional | file      | Custom PPTX template file as multipart/form-data                                                     |
> | encoding             | optional | string    | Encoding of provided source content (default: utf-8)                                                 |
> | file_name            | optional | string    | Output filename (default: converted-document.pptx)                                                   |
> | slide_size           | optional | string    | Slide size for the presentation. Supported values: 16:9, WIDESCREEN, 4:3, A3, A4, LETTER, LEDGER   |

##### Responses

> | HTTP code | Content-Type                                                                   | Response                      |
> |-----------|--------------------------------------------------------------------------------|-------------------------------|
> | `200`     | `application/vnd.openxmlformats-officedocument.presentationml.presentation`    | PPTX presentation (binary)    |
> | `400`     | `plain/text`                                                                   | Error message with exception  |
> | `500`     | `plain/text`                                                                   | Error message with exception  |

##### Example cURL

> ```bash
> curl -X POST -F "source=@input.html" http://localhost:9082/convert/html/to/pptx-with-template --output output.pptx
> ```
>
> With custom template:
> ```bash
> curl -X POST -F "source=@input.md" -F "template=@custom-template.pptx" http://localhost:9082/convert/markdown/to/pptx-with-template --output output.pptx
> ```
>
> With custom slide size:
> ```bash
> curl -X POST -F "source=@input.html" "http://localhost:9082/convert/html/to/pptx-with-template?slide_size=16:9" --output output.pptx
> ```
>
> With both template and slide size:
> ```bash
> curl -X POST -F "source=@input.md" -F "template=@custom-template.pptx" "http://localhost:9082/convert/markdown/to/pptx-with-template?slide_size=4:3" --output output.pptx
> ```

</details>

### Supported Special Variables

The document processor recognizes several special placeholder variables that can be automatically replaced with dynamic Word fields during HTML to DOCX conversion:

#### Table of Contents Placeholders

- **TOC_PLACEHOLDER** - Automatically replaced with a Word Table of Contents (TOC) field that lists all headings in your document
- **TOF_PLACEHOLDER** - Automatically replaced with a Word Table of Figures (TOF) field that lists all figure captions
- **TOT_PLACEHOLDER** - Automatically replaced with a Word Table of Tables (TOT) field that lists all table captions
