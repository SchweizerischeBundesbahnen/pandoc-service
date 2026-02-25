# Pandoc Service Monitoring

Comprehensive monitoring setup with Prometheus and Grafana for Pandoc Service.

## Quick Start

### 1. Start Monitoring Stack

```bash
./start-monitoring.sh
```

This will:
- Build the Pandoc service Docker image
- Start Prometheus, Grafana, and Pandoc service
- Wait for all services to be healthy
- Generate initial test traffic
- Display access URLs

### 2. Generate Test Load

```bash
# Generate 100 requests with 10 concurrent workers
./generate-load.sh

# Custom load: 500 requests with 20 concurrent workers
./generate-load.sh 500 20
```

### 3. View Metrics

#### Grafana Dashboard
- URL: http://localhost:3000/d/pandoc-service
- Username: `admin`
- Password: `admin`
- Pre-configured with all metrics

#### Prometheus
- URL: http://localhost:9090
- Query metrics directly
- View targets: http://localhost:9090/targets

### 4. Stop Monitoring Stack

```bash
./stop-monitoring.sh
```

## Architecture

```
┌─────────────────┐
│  Pandoc         │
│  Service        │──────┐
│  :9082          │      │
│  Metrics :9182  │      │
└─────────────────┘      │
                         │ Scrapes /metrics
                         ▼
                  ┌─────────────┐
                  │ Prometheus  │
                  │ :9090       │
                  └─────────────┘
                         │
                         │ Data source
                         ▼
                  ┌─────────────┐
                  │  Grafana    │
                  │  :3000      │
                  └─────────────┘
```

## Available Metrics

### Conversion Metrics
- `pandoc_conversions_total` - Total successful conversions (labeled by source/target format)
- `pandoc_conversion_failures_total` - Total failed conversions (labeled by source/target format)
- `pandoc_conversion_error_rate_percent` - Conversion error rate as percentage
- `pandoc_template_conversions_total` - Total conversions using custom templates

### Performance Metrics
- `pandoc_conversion_duration_seconds` - Conversion time histogram (labeled by format)
- `pandoc_subprocess_duration_seconds` - Pandoc subprocess execution time histogram
- `pandoc_post_processing_duration_seconds` - DOCX/PPTX post-processing time histogram
- `avg_pandoc_conversion_time_seconds` - Average conversion time

### Size Metrics
- `pandoc_request_body_bytes` - Input document size histogram
- `pandoc_response_body_bytes` - Output document size histogram

### Service Metrics
- `uptime_seconds` - Service uptime
- `active_conversions` - Current active conversion count
- `pandoc_info` - Service and pandoc version information

## Grafana Dashboard Panels

1. **Status** - Service up/down status
2. **Uptime** - Total service uptime
3. **Active Conversions** - Current concurrent operations
4. **Error Rate** - Conversion error percentage
5. **Total Conversions** - Cumulative successful conversions
6. **Avg Conversion Time** - Average conversion duration
7. **Total Failures** - Cumulative failed conversions
8. **Template Conversions** - Conversions using templates
9. **Conversion Rate** - Conversions per second over time
10. **Duration (p50/p95)** - Conversion latency percentiles
11. **Request/Response Sizes** - Document size trends
12. **Subprocess & Post-processing** - Internal timing breakdown

## Prometheus Query Examples

### Conversion Rate (requests/sec)
```promql
rate(pandoc_conversions_total[5m])
```

### Error Rate (%)
```promql
rate(pandoc_conversion_failures_total[5m]) / rate(pandoc_conversions_total[5m]) * 100
```

### P95 Response Time
```promql
histogram_quantile(0.95, rate(pandoc_conversion_duration_seconds_bucket[5m]))
```

### Conversions by Format
```promql
sum by (source_format, target_format) (pandoc_conversions_total)
```

## Configuration Files

### Prometheus Configuration
- File: `prometheus.yml`
- Scrape interval: 10 seconds
- Scrape timeout: 5 seconds

### Grafana Provisioning
- Datasources: `grafana/provisioning/datasources/`
- Dashboards: `grafana/provisioning/dashboards/`
- Dashboard JSON: `grafana/dashboards/pandoc-service.json`

### Docker Compose
- File: `docker-compose.yml`
- Services: pandoc-service, prometheus, grafana
- Networks: monitoring (bridge)
- Volumes: prometheus-data, grafana-data

## Troubleshooting

### Services not starting
```bash
# Check Docker logs
docker compose -f docker-compose.yml logs

# Check individual service
docker compose -f docker-compose.yml logs pandoc-service
docker compose -f docker-compose.yml logs prometheus
docker compose -f docker-compose.yml logs grafana
```

### Metrics not appearing
1. Check Prometheus targets: http://localhost:9090/targets
2. Verify service is exposing metrics: http://localhost:9182/metrics
3. Check Grafana datasource configuration

### Dashboard not loading
1. Verify Grafana provisioning: `docker compose -f docker-compose.yml logs grafana`
2. Check dashboard exists: http://localhost:3000/dashboards
3. Reimport dashboard manually if needed

## Clean Up

### Stop services but keep data
```bash
./stop-monitoring.sh
# Choose 'N' when asked about removing volumes
```

### Stop services and remove all data
```bash
./stop-monitoring.sh
# Choose 'Y' when asked about removing volumes
```

### Manual cleanup
```bash
docker compose -f docker-compose.yml down -v
docker rmi pandoc-service:dev
```

## Load Testing

Use the built-in load generator:
```bash
# Light load: 100 requests, 10 concurrent
./generate-load.sh

# Medium load: 500 requests, 20 concurrent
./generate-load.sh 500 20

# Heavy load: 2000 requests, 50 concurrent
./generate-load.sh 2000 50
```

## Access URLs Summary

| Service | URL | Credentials |
|---------|-----|-------------|
| Pandoc Service | http://localhost:9082 | - |
| API Docs | http://localhost:9082/api/docs | - |
| Raw Metrics | http://localhost:9182/metrics | - |
| Prometheus | http://localhost:9090 | - |
| Grafana | http://localhost:3000 | admin/admin |
| Grafana Dashboard | http://localhost:3000/d/pandoc-service | admin/admin |
