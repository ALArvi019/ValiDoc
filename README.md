# Validoc - PDF Generator with Digital Signature

[Español](#español) | [English](#english)

---

## English

### Overview

Validoc is a desktop application that generates personalized PDF documents from Word templates and CSV data, with optional digital signature capabilities. The application features a user-friendly GUI built with Python's tkinter library.

### Features

- **Template-based PDF Generation**: Use Word (.docx) templates with field placeholders
- **CSV Data Integration**: Populate templates with data from CSV files  
- **Digital Signatures**: Sign PDFs with digital certificates (.pfx, .p12)
- **Cross-platform Support**: Works on Windows and Linux
- **Batch Processing**: Generate multiple PDFs from a single template and dataset
- **Progress Tracking**: Real-time progress updates and operation logs
- **Signature Stamps**: Add verification stamps to signed documents

### System Requirements

- Python 3.6 or higher
- Windows or Linux operating system
- LibreOffice (for Linux users)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ALArvi019/Validoc.git
   cd Validoc
   ```

2. **Install Python dependencies:**
   ```bash
   pip install python-docx
   pip install docx2pdf
   pip install endesive
   pip install cryptography
   ```

3. **Windows-specific dependency:**
   ```bash
   pip install pywin32
   ```

4. **Linux users:** Ensure LibreOffice is installed:
   ```bash
   sudo apt-get install libreoffice
   ```

### Usage

1. **Run the application:**
   ```bash
   python main.py
   ```

2. **Configure your documents:**
   - **Word Template**: Select a .docx file with field placeholders marked with `«field_name»`
   - **CSV Data**: Choose a CSV file with column headers matching your template fields
   - **Certificate** (optional): Select a .pfx or .p12 certificate file for digital signing

3. **Set up digital signing** (optional):
   - Check "Firmar documento con certificado" (Sign document with certificate)
   - Enter your certificate password
   - Select your certificate file

4. **Generate PDFs:**
   - Click "Generar PDFs" (Generate PDFs)
   - Monitor progress in the log window
   - Find generated PDFs in the same directory as your template

### Template Format

Create Word templates with field placeholders using the format `«field_name»`. For example:

```
Dear «name»,

Your account balance is «balance».
Transaction date: «date»

Best regards,
«company_name»
```

### CSV Format

Ensure your CSV file has headers that match the field names in your template:

```csv
name,balance,date,company_name
John Doe,1250.50,2024-01-15,ABC Corp
Jane Smith,890.75,2024-01-16,ABC Corp
```

### Troubleshooting

- **PDF conversion fails on Linux**: Ensure LibreOffice is properly installed
- **Certificate errors**: Verify certificate file format and password
- **Template fields not replaced**: Check that CSV headers exactly match template field names
- **Permission errors**: Run with appropriate file system permissions

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

### License

This project is open source. Please check the repository for specific license terms.

---

## Español

### Descripción General

Validoc es una aplicación de escritorio que genera documentos PDF personalizados a partir de plantillas de Word y datos CSV, con capacidades opcionales de firma digital. La aplicación cuenta con una interfaz gráfica fácil de usar construida con la librería tkinter de Python.

### Características

- **Generación de PDF basada en plantillas**: Usa plantillas de Word (.docx) con marcadores de posición para campos
- **Integración de datos CSV**: Rellena plantillas con datos de archivos CSV
- **Firmas Digitales**: Firma PDFs con certificados digitales (.pfx, .p12)
- **Soporte multiplataforma**: Funciona en Windows y Linux
- **Procesamiento por lotes**: Genera múltiples PDFs desde una sola plantilla y conjunto de datos
- **Seguimiento de progreso**: Actualizaciones de progreso en tiempo real y registros de operación
- **Sellos de firma**: Añade sellos de verificación a documentos firmados

### Requisitos del Sistema

- Python 3.6 o superior
- Sistema operativo Windows o Linux
- LibreOffice (para usuarios de Linux)

### Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/ALArvi019/Validoc.git
   cd Validoc
   ```

2. **Instalar dependencias de Python:**
   ```bash
   pip install python-docx
   pip install docx2pdf
   pip install endesive
   pip install cryptography
   ```

3. **Dependencia específica para Windows:**
   ```bash
   pip install pywin32
   ```

4. **Usuarios de Linux:** Asegurar que LibreOffice esté instalado:
   ```bash
   sudo apt-get install libreoffice
   ```

### Uso

1. **Ejecutar la aplicación:**
   ```bash
   python main.py
   ```

2. **Configurar sus documentos:**
   - **Plantilla Word**: Seleccionar un archivo .docx con marcadores de campo marcados con `«nombre_campo»`
   - **Datos CSV**: Elegir un archivo CSV con encabezados de columna que coincidan con los campos de su plantilla
   - **Certificado** (opcional): Seleccionar un archivo de certificado .pfx o .p12 para firma digital

3. **Configurar firma digital** (opcional):
   - Marcar "Firmar documento con certificado"
   - Introducir la contraseña de su certificado
   - Seleccionar su archivo de certificado

4. **Generar PDFs:**
   - Hacer clic en "Generar PDFs"
   - Monitorear el progreso en la ventana de registro
   - Encontrar los PDFs generados en el mismo directorio que su plantilla

### Formato de Plantilla

Crear plantillas de Word con marcadores de posición usando el formato `«nombre_campo»`. Por ejemplo:

```
Estimado/a «nombre»,

Su saldo de cuenta es «saldo».
Fecha de transacción: «fecha»

Saludos cordiales,
«nombre_empresa»
```

### Formato CSV

Asegurar que su archivo CSV tenga encabezados que coincidan con los nombres de campo en su plantilla:

```csv
nombre,saldo,fecha,nombre_empresa
Juan Pérez,1250.50,2024-01-15,Empresa ABC
María García,890.75,2024-01-16,Empresa ABC
```

### Solución de Problemas

- **La conversión de PDF falla en Linux**: Asegurar que LibreOffice esté correctamente instalado
- **Errores de certificado**: Verificar el formato del archivo de certificado y la contraseña
- **Los campos de plantilla no se reemplazan**: Verificar que los encabezados CSV coincidan exactamente con los nombres de campo de la plantilla
- **Errores de permisos**: Ejecutar con permisos apropiados del sistema de archivos

### Contribuir

1. Hacer fork del repositorio
2. Crear una rama de característica
3. Realizar sus cambios
4. Probar exhaustivamente
5. Enviar un pull request

### Licencia

Este proyecto es de código abierto. Por favor, revisar el repositorio para términos específicos de licencia.