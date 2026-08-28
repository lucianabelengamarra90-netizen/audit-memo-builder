# Audit Memo Builder

Herramienta web para Auditoria Interna.

## Ejecucion Local

```bash
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
python app.py
```

Abrir: http://127.0.0.1:5000

## Despliegue en Render

1. Subir a GitHub
2. Crear Web Service en Render
3. Conectar repositorio
4. Deploy automatico con render.yaml

## Estructura

- app.py: Backend Flask
- templates/index.html: Interfaz
- static/styles.css: Diseño
- static/app.js: Interactividad
- requirements.txt: Dependencias
- render.yaml: Config Render

## Flujo

1. Informacion → Cargar archivos
2. Datos → Completar campos
3. Hechos → Validar
4. Tareas → Registrar
5. Hallazgos → Definir
6. Memo → Generar
7. Revision → Editar
8. Exportar → Excel

v1.0
