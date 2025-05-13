# Agent Docs

Este proyecto es un servidor que permite la creación y edición de documentos Word (.docx) utilizando `python-docx`. Proporciona herramientas para interactuar con documentos a través de un servidor basado en `FastMCP`.

## Requisitos

- Python 3.11 o superior (verificado en el archivo `.python-version`).
- Dependencias especificadas en `pyproject.toml`:
  - `mcp[cli]>=1.8.0`
  - `python-docx>=1.1.2`

## Instalación

1. Clona este repositorio en tu máquina local.
2. Asegúrate de tener Python 3.11 instalado.
3. Instala las dependencias ejecutando:

   ```bash
   pip install -r requirements.txt