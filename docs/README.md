# GenCer
Generador de certificados a partir de un archivo Excel
# GenCer - Generador de Certificados

Sistema automático para generar certificados en PDF desde un archivo Excel almacenado en OneDrive/SharePoint, utilizando Microsoft Graph API.

## 🚀 Características

- **Lectura automática** de Excel desde OneDrive
- **Generación masiva** de certificados PDF usando plantilla Word
- **Organización automática** por empresa en carpetas de OneDrive
- **Prevención de duplicados** mediante control de estado
- **Logging completo** para seguimiento y debugging
- **Interfaz de línea de comandos** fácil de usar

## 📋 Requisitos Previos

1. **Python 3.8+**
2. **Aplicación registrada en Microsoft Azure** con permisos para Microsoft Graph
3. **Archivo Excel** en OneDrive con la estructura requerida
4. **Plantilla Word** (.docx) para los certificados

## 🛠️ Instalación

1. **Clonar el repositorio**
```bash
git clone <tu-repositorio>
cd GenCer
```

2. **Crear entorno virtual**
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# o
venv\Scripts\activate     # Windows
```

3. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

4. **Configurar variables de entorno**
```bash
cp .env.example .env
# Editar .env con tus credenciales
```

5. **Colocar archivos requeridos**
- `plantilla.docx` - Tu plantilla de certificado
- `.env` - Configurado con tus credenciales

## ⚙️ Configuración

### Archivo .env
```bash
# Microsoft Graph API
CLIENT_ID=tu_client_id
CLIENT_SECRET=tu_client_secret
TENANT_ID=tu_tenant_id

# OneDrive
EXCEL_FILE_ID=id_del_archivo_excel
CERTIFICADOS_FOLDER_ID=01WIY7HEN2WE6VD2WDPFG3UMFYFXUVC25K
```

### Estructura del Excel
El archivo Excel debe contener al menos estas columnas:

| nombre | cedula | compañia | certificado | ... |
|--------|--------|----------|-------------|-----|
| Juan Pérez | 12345678 | Empresa A | no | ... |
| María González | 87654321 | Empresa B | si | ... |

- **certificado**: 'no' para generar, 'si' para omitir
- **compañia**: nombre de la empresa (para organizar carpetas)
- **nombre, cedula**: datos para el certificado

### Plantilla Word
La plantilla `plantilla.docx` debe contener marcadores que serán reemplazados:
- `{nombre