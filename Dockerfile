FROM python:3.12-slim
WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir mcp openpyxl && pip install --no-cache-dir --no-deps -e .

# OCI labels required by the official MCP Registry validator
LABEL io.modelcontextprotocol.server.name="io.github.diyaluo/excel-toolkit-mcp" \
      io.modelcontextprotocol.mcp.version="0.1.0" \
      org.opencontainers.image.title="excel-toolkit-mcp"

ENTRYPOINT ["python", "-m", "excel_toolkit_mcp.server"]
